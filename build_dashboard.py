# -*- coding: utf-8 -*-
"""
build_dashboard.py  —  Live Project static dashboard builder  (v2)
──────────────────────────────────────────────────────────────────
Legge i risultati prodotti da daily_forecast.py e genera UNA pagina HTML
autoconsistente (dati inclusi) da pubblicare su GitHub Pages.

Perché statica: non c'è nessun processo da tenere vivo, quindi non c'è nulla
che possa "addormentarsi" come su Streamlit Community Cloud. È solo un file.

Uso:
    python build_dashboard.py                # genera il sito in SITE_DIR
    python build_dashboard.py --publish      # genera + git commit + git push
    python build_dashboard.py --open         # genera + apre nel browser
    python build_dashboard.py --demo         # genera con dati finti (test)
    python build_dashboard.py --data-root X  # usa un'altra radice dati

Input (prodotti da daily_forecast.py):
    /Volumes/ESTERNO/forecasts/forecasts_all.xlsx     ← tutti i run
    /Volumes/ESTERNO/official_series/<var>_<tag>.parquet
    /Volumes/ESTERNO/processed/<sub>_daily_mentions.parquet

Output:
    <SITE_DIR>/index.html      dashboard completa (dati embedded)
    <SITE_DIR>/plotly.min.js   libreria grafici, scaricata una volta
    <SITE_DIR>/data.json       stessi dati, per chi li vuole riusare
    <SITE_DIR>/forecasts.csv   master file in CSV (link di download)
"""

import os
import re
import io
import sys
import glob
import json
import math
import shutil
import tarfile
import argparse
import subprocess
import webbrowser
import urllib.request
from collections import defaultdict
from datetime import datetime, timezone

import numpy as np
import pandas as pd

# =====================================================================
# CONFIG — EDIT HERE
# =====================================================================
PROJECT_DIR = "/Users/luigilongo/Desktop/live_project"
DATA_ROOT   = "/Volumes/ESTERNO"

PROCESSED_PATH = os.path.join(DATA_ROOT, "processed")
SERIES_PATH    = os.path.join(DATA_ROOT, "official_series")
FORECAST_PATH  = os.path.join(DATA_ROOT, "forecasts")
MASTER_XLSX    = os.path.join(FORECAST_PATH, "forecasts_all.xlsx")

SITE_DIR = os.path.join(PROJECT_DIR, "dashboard-site")

SITE_TITLE    = "Live Macro Forecasts"
SITE_TAGLINE  = "Can language models forecast the next official statistic?"
SITE_SUBTITLE = ("Every morning three language models forecast the next official release of "
                 "eight US and euro-area indicators. When the statistic is published, the "
                 "forecast is scored against it. Everything is in real-time.")
SITE_FOOTER   = ("Research project · official data: OECD, ECB Data Portal, Eurostat, FRED · "
                 "conversations: Reddit via Arctic Shift")
SITE_DISCLAIMER = ("This website represents my own views and not those of the European "
                   "Commission. For more information, write to me at "
                   "luigi.longo[at]ec.europa.eu")

MAX_REDDIT_SERIES = 7      # oltre questo, i subreddit finiscono in "Other"
MAX_HINDCASTS     = 6      # quante previsioni passate disegnare sul grafico
WINSOR            = (0.25, 4.0)   # clip dei ratio prima della media geometrica

# La libreria dei grafici viene scaricata UNA volta e messa nel repo accanto
# a index.html: la pagina non dipende da nessuna CDN esterna (utile se la rete
# dell'ufficio le blocca) e continua a funzionare anche fra dieci anni.
PLOTLY_VERSION = "2.35.2"
PLOTLY_FILE    = "plotly.min.js"

VARIABLES = {
    "us_cpi_yoy":  {"short": "US inflation",        "area": "US", "unit": "% YoY"},
    "ea_hicp_yoy": {"short": "Euro area inflation", "area": "EA", "unit": "% YoY"},
    "us_unrate":   {"short": "US unemployment",     "area": "US", "unit": "%"},
    "ea_unrate":   {"short": "Euro area unemployment", "area": "EA", "unit": "%"},
    "us_gdp_qoq":  {"short": "US GDP growth",       "area": "US", "unit": "% QoQ"},
    "ea_gdp_qoq":  {"short": "Euro area GDP growth", "area": "EA", "unit": "% QoQ"},
    "us_ip_yoy":   {"short": "US industrial production", "area": "US", "unit": "% YoY"},
    "ea_ip_yoy":   {"short": "Euro area industrial production", "area": "EA", "unit": "% YoY"},
}


# =====================================================================
# HELPERS — periodi
# =====================================================================
def period_freq(label: str) -> str:
    return "Q" if isinstance(label, str) and "Q" in label.upper() else "M"


def period_to_ts(label):
    if not isinstance(label, str) or not label.strip():
        return None
    s = label.strip().upper()
    m = re.match(r"^(\d{4})[-\s]?Q([1-4])$", s)
    if m:
        return pd.Timestamp(year=int(m.group(1)), month=3 * (int(m.group(2)) - 1) + 1, day=1)
    m = re.match(r"^(\d{4})[-/](\d{1,2})", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)
    try:
        return pd.Timestamp(pd.to_datetime(s)).normalize().replace(day=1)
    except Exception:
        return None


def ts_to_period(ts, freq: str) -> str:
    ts = pd.Timestamp(ts)
    return ts.strftime("%Y-%m") if freq == "M" else f"{ts.year}-Q{ts.quarter}"


def next_period(label: str) -> str:
    """Periodo successivo, calcolato dai NOSTRI dati (non da quello che dice l'LLM)."""
    freq = period_freq(label)
    ts = period_to_ts(label)
    if ts is None:
        return ""
    nxt = ts + (pd.offsets.MonthBegin(1) if freq == "M"
                else pd.offsets.QuarterBegin(startingMonth=1))
    return ts_to_period(nxt, freq)


def pretty_period(label: str) -> str:
    ts = period_to_ts(label)
    if ts is None:
        return label
    return ts.strftime("%B %Y") if period_freq(label) == "M" else f"Q{ts.quarter} {ts.year}"


def model_label(provider, model) -> str:
    return f"{provider}/{model}"


def short_model(mid: str) -> str:
    return mid.split("/")[-1]


# =====================================================================
# LOAD
# =====================================================================
def load_master(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Non trovo {path}. Lancia prima daily_forecast.py, "
            f"oppure usa --demo per una prova con dati finti.")
    df = pd.read_excel(path)
    for col in ["provider", "model", "variable", "variable_name", "forecast_type",
                "last_official_period", "target_period", "reasoning"]:
        if col not in df.columns:
            df[col] = ""
    for col in ["point_forecast", "last_official_value"]:
        df[col] = pd.to_numeric(df.get(col), errors="coerce")
    if "used_web_search" not in df.columns:
        df["used_web_search"] = False
    df["run_date"] = pd.to_datetime(df["run_date"], errors="coerce")
    df = df.dropna(subset=["run_date", "variable"])
    df["model_id"] = [model_label(p, m) for p, m in zip(df["provider"], df["model"])]
    # target canonico: calcolato dall'ultimo dato ufficiale, non dal testo dell'LLM
    df["target"] = df["last_official_period"].astype(str).apply(next_period)
    bad = df["target"] == ""
    if bad.any():
        df.loc[bad, "target"] = df.loc[bad, "target_period"].astype(str)
    df = df.sort_values("run_date").drop_duplicates(
        subset=["run_date", "model_id", "variable", "forecast_type", "target"], keep="last")
    return df.reset_index(drop=True)


def load_official_series(series_dir: str) -> dict:
    out = {}
    if not os.path.isdir(series_dir):
        return out
    for var in VARIABLES:
        files = sorted(glob.glob(os.path.join(series_dir, f"{var}_*.parquet")))
        if not files:
            continue
        try:
            df = pd.read_parquet(files[-1])
        except Exception as e:
            print(f"  [warn] {var}: {type(e).__name__}: {e}")
            continue
        df["date"] = pd.to_datetime(df["date"]).dt.tz_localize(None)
        df["value"] = pd.to_numeric(df["value"], errors="coerce")
        out[var] = df.dropna(subset=["value"]).sort_values("date").reset_index(drop=True)
    return out


def load_reddit_mentions(processed_dir: str) -> pd.DataFrame:
    files = sorted(glob.glob(os.path.join(processed_dir, "*_daily_mentions.parquet")))
    frames = []
    for f in files:
        sub = os.path.basename(f).replace("_daily_mentions.parquet", "")
        try:
            d = pd.read_parquet(f)
        except Exception:
            continue
        if d.empty:
            continue
        d["subreddit"] = sub
        frames.append(d)
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    df["date"] = pd.to_datetime(df["date"]).dt.tz_localize(None).dt.normalize()
    return df


# =====================================================================
# SCOREBOARD
# =====================================================================
def _skill(rows):
    """rows = [(variable, abs_err, abs_rw_err)] → media geometrica dei ratio.

    Media GEOMETRICA e non aritmetica perché sono rapporti: con una sola
    osservazione per variabile, una variabile in cui il random walk ha
    azzeccato quasi esattamente produrrebbe un ratio enorme che da solo
    deciderebbe la classifica. I ratio sono anche winsorizzati a [0.25, 4].
    """
    by_var = defaultdict(list)
    for var, e, rw in rows:
        by_var[var].append((e, rw))
    ratios, clipped = [], 0
    for var, lst in by_var.items():
        mae = float(np.mean([a for a, _ in lst]))
        rw = float(np.mean([b for _, b in lst]))
        if rw <= 0:
            continue                      # benchmark perfetto: ratio indefinito
        r = mae / rw
        rc = min(max(r, WINSOR[0]), WINSOR[1])
        if rc != r:
            clipped += 1
        ratios.append(rc)
    if not ratios:
        return None, 0, 0
    return float(np.exp(np.mean(np.log(ratios)))), len(ratios), clipped


def build_scoreboard(obs_by_variant):
    """obs_by_variant: {(model, type): [ {var, err, rw} ]} → righe ordinate."""
    base = {}
    for key, obs in obs_by_variant.items():
        s, nvars, clipped = _skill([(o["var"], abs(o["err"]), abs(o["rw"])) for o in obs])
        base[key] = s

    rows = []
    for key, obs in obs_by_variant.items():
        mid, ftype = key
        s, nvars, clipped = _skill([(o["var"], abs(o["err"]), abs(o["rw"])) for o in obs])
        wins = sum(1 for o in obs if abs(o["err"]) < abs(o["rw"]))
        others = sorted([v for k, v in base.items() if k != key and v is not None])

        # rank range leave-one-out: togli UNA osservazione alla volta e guarda
        # dove finisce il modello in classifica. Non è un test statistico,
        # è il modo più onesto di dire "con otto punti la classifica balla".
        lo = hi = None
        if s is not None and len(obs) > 1:
            ranks = []
            for i in range(len(obs)):
                sub = [(o["var"], abs(o["err"]), abs(o["rw"]))
                       for j, o in enumerate(obs) if j != i]
                s_i, _, _ = _skill(sub)
                if s_i is None:
                    continue
                ranks.append(1 + sum(1 for v in others if v < s_i))
            if ranks:
                lo, hi = min(ranks), max(ranks)

        rows.append({
            "model": mid, "type": ftype,
            "label": short_model(mid) + (" · cond." if ftype == "conditional" else ""),
            "skill": None if s is None else round(s, 3),
            "n": len(obs), "vars": nvars, "wins": wins, "clipped": clipped,
            "loo_lo": lo, "loo_hi": hi,
            "mae": round(float(np.mean([abs(o["err"]) for o in obs])), 3) if obs else None,
            "bias": round(float(np.mean([o["err"] for o in obs])), 3) if obs else None,
            "rw": round(float(np.mean([abs(o["rw"]) for o in obs])), 3) if obs else None,
        })
    rows.sort(key=lambda d: (d["skill"] is None, d["skill"]))
    for i, r in enumerate(rows):
        r["rank"] = i + 1
    return rows


# =====================================================================
# PAYLOAD
# =====================================================================
def build_payload(master: pd.DataFrame, series: dict, reddit: pd.DataFrame) -> dict:
    last_run = master["run_date"].max()
    model_ids = sorted(master["model_id"].dropna().unique().tolist())

    realized = {}
    for var, sdf in series.items():
        freq = "Q" if "qoq" in var else "M"
        for _, r in sdf.iterrows():
            realized[(var, ts_to_period(r["date"], freq))] = float(r["value"])

    master = master.copy()
    master["realized"] = [realized.get((v, t)) for v, t in
                          zip(master["variable"], master["target"])]
    master["error"] = master["point_forecast"] - master["realized"]
    master["rw_error"] = master["last_official_value"] - master["realized"]

    obs_by_variant = defaultdict(list)
    evidence, variables, tiles = {}, {}, []
    pending_targets = 0

    for var, meta in VARIABLES.items():
        sub = master[master["variable"] == var]
        sdf = series.get(var)
        if sub.empty and (sdf is None or sdf.empty):
            continue

        name = meta["short"]
        long_name = name
        if not sub.empty:
            nm = sub["variable_name"].dropna().astype(str)
            if len(nm) and nm.iloc[-1].strip():
                long_name = nm.iloc[-1].strip()

        freq = "Q" if "qoq" in var else "M"
        hist = []
        if sdf is not None and not sdf.empty:
            hist = [[ts_to_period(r["date"], freq), round(float(r["value"]), 3)]
                    for _, r in sdf.iterrows()]

        latest = pd.DataFrame()
        if not sub.empty:
            last_var_run = sub["run_date"].max()
            latest = sub[(sub["run_date"] == last_var_run) & sub["point_forecast"].notna()]

        fcs, target, last_off, fc_run = [], "", None, ""
        if not latest.empty:
            target = str(latest["target"].iloc[0])
            fc_run = pd.Timestamp(latest["run_date"].iloc[0]).strftime("%Y-%m-%d")
            lv = latest["last_official_value"].iloc[0]
            last_off = {"period": str(latest["last_official_period"].iloc[0]),
                        "value": float(lv) if pd.notna(lv) else None}
            for _, r in latest.iterrows():
                fcs.append({
                    "model": r["model_id"], "type": r["forecast_type"],
                    "label": short_model(r["model_id"]) +
                             (" · cond." if r["forecast_type"] == "conditional" else ""),
                    "value": round(float(r["point_forecast"]), 3),
                    "web": bool(r.get("used_web_search", False)),
                    "reasoning": str(r.get("reasoning", "")).strip()[:2600],
                })
        elif hist:
            last_off = {"period": hist[-1][0], "value": hist[-1][1]}
            target = next_period(hist[-1][0])

        stale = bool(fc_run) and fc_run != last_run.strftime("%Y-%m-%d")
        if not stale and target and target not in {h[0] for h in hist}:
            pending_targets += 1

        # ── revisioni ───────────────────────────────────────────────
        revisions = {}
        if not sub.empty:
            for tgt, g in sub[sub["point_forecast"].notna()].groupby("target"):
                lines = []
                for (mid, ftype), gg in g.groupby(["model_id", "forecast_type"]):
                    gg = gg.sort_values("run_date")
                    lines.append({
                        "model": mid, "type": ftype,
                        "label": short_model(mid) + (" · cond." if ftype == "conditional" else ""),
                        "x": [d.strftime("%Y-%m-%d") for d in gg["run_date"]],
                        "y": [round(float(v), 3) for v in gg["point_forecast"]],
                    })
                if lines:
                    revisions[str(tgt)] = {"lines": lines,
                                           "realized": realized.get((var, str(tgt)))}

        # ── previsioni già valutate (hindcast) ──────────────────────
        ev = sub[sub["error"].notna()]
        hind, track, ev_targets = [], [], {}
        if not ev.empty:
            for (mid, ftype), g in ev.groupby(["model_id", "forecast_type"]):
                g = g.sort_values("run_date").drop_duplicates("target", keep="last")
                for _, r in g.iterrows():
                    rec = {"var": var, "target": str(r["target"]),
                           "err": float(r["error"]),
                           "rw": float(r["rw_error"]) if pd.notna(r["rw_error"]) else float("nan")}
                    if not math.isnan(rec["rw"]):
                        obs_by_variant[(mid, ftype)].append(rec)
                    hind.append({"model": mid, "type": ftype, "t": str(r["target"]),
                                 "f": round(float(r["point_forecast"]), 3),
                                 "a": round(float(r["realized"]), 3),
                                 "rw": None if pd.isna(r["rw_error"]) else round(abs(float(r["rw_error"])), 3)})
                    ev_targets.setdefault(str(r["target"]), {
                        "t": str(r["target"]), "a": round(float(r["realized"]), 3),
                        "rw": None if pd.isna(r["rw_error"]) else round(float(r["rw_error"]), 3),
                        "marks": []})
                    ev_targets[str(r["target"])]["marks"].append({
                        "model": mid, "type": ftype,
                        "err": round(float(r["error"]), 3)})
                mae = float(np.mean(np.abs(g["error"])))
                rwv = g["rw_error"].dropna()
                track.append({"model": mid, "type": ftype, "n": int(len(g)),
                              "label": short_model(mid) +
                                       (" · cond." if ftype == "conditional" else ""),
                              "mae": round(mae, 3),
                              "rw": round(float(np.mean(np.abs(rwv))), 3) if len(rwv) else None,
                              # ogni singola previsione valutata: con n=1 o 2 la media
                              # da sola nasconderebbe quanto è sottile l'evidenza
                              "points": [{"t": str(t), "e": round(float(e), 3)}
                                         for t, e in zip(g["target"], g["error"])]})
            track.sort(key=lambda d: d["mae"])
            hind.sort(key=lambda d: d["t"])
            hind = [h for h in hind
                    if h["t"] in sorted({x["t"] for x in hind})[-MAX_HINDCASTS:]]
            if ev_targets:
                evidence[var] = {"name": long_name, "unit": meta["unit"],
                                 "targets": [ev_targets[k] for k in sorted(ev_targets)]}

        variables[var] = {
            "name": long_name, "short": meta["short"], "unit": meta["unit"],
            "area": meta["area"], "freq": freq, "stale": stale,
            "history": hist, "spark": [h[1] for h in hist[-24:]],
            "last_official": last_off, "target": target,
            "target_pretty": pretty_period(target) if target else "",
            "forecast_run": fc_run, "forecasts": fcs,
            "hindcasts": hind, "track": track, "revisions": revisions,
        }

        if last_off and fcs:
            vals = [f["value"] for f in fcs]
            tiles.append({
                "var": var, "short": meta["short"], "name": long_name, "unit": meta["unit"],
                "last_period": last_off["period"],
                "last_period_pretty": pretty_period(last_off["period"]),
                "last_value": last_off["value"],
                "target": target, "target_pretty": pretty_period(target) if target else "",
                "consensus": round(float(np.median(vals)), 2),
                "lo": round(float(np.min(vals)), 2), "hi": round(float(np.max(vals)), 2),
                "n": len(vals), "stale": stale, "run": fc_run,
            })

    scoreboard = build_scoreboard(obs_by_variant) if obs_by_variant else []
    scored_targets = len({(o["var"], o["target"]) for obs in obs_by_variant.values() for o in obs})

    # hero: la prossima uscita fra le variabili vive
    hero = None
    for t in tiles:
        if not t["stale"]:
            hero = t
            break

    # ── Reddit (dentro il cassetto metodologico) ────────────────────
    reddit_payload = None
    if reddit is not None and not reddit.empty:
        tot = reddit.groupby("subreddit")["mentions_total"].sum().sort_values(ascending=False)
        keep = list(tot.head(MAX_REDDIT_SERIES).index)
        r = reddit.copy()
        r["grp"] = np.where(r["subreddit"].isin(keep), "r/" + r["subreddit"], "Other")
        piv = r.pivot_table(index="date", columns="grp", values="mentions_total",
                            aggfunc="sum").fillna(0.0).sort_index().tail(400)
        order = ["r/" + k for k in keep if "r/" + k in piv.columns]
        if "Other" in piv.columns:
            order.append("Other")
        piv = piv[order]
        roll = piv.rolling(7, min_periods=2).mean().round(2)
        reddit_payload = {
            "dates": [d.strftime("%Y-%m-%d") for d in piv.index],
            "series": [{"name": c, "y": [None if pd.isna(v) else float(v) for v in roll[c]]}
                       for c in piv.columns],
            "n_subs": int(reddit["subreddit"].nunique()),
        }

    runs = sorted(master["run_date"].dt.strftime("%Y-%m-%d").unique().tolist())
    return {
        "generated_at": datetime.now(timezone.utc).strftime("%d %b %Y, %H:%M UTC"),
        "last_run": last_run.strftime("%Y-%m-%d"),
        "first_run": runs[0] if runs else "",
        "n_runs": len(runs),
        "models": model_ids,
        "model_short": {m: short_model(m) for m in model_ids},
        "hero": hero,
        "tiles": tiles,
        "variables": variables,
        "scoreboard": scoreboard,
        "scored_targets": scored_targets,
        "pending_targets": pending_targets,
        "evidence": evidence,
        "reddit": reddit_payload,
        "n_forecasts": int(master["point_forecast"].notna().sum()),
        "n_stale": sum(1 for v in variables.values() if v["stale"]),
    }


# =====================================================================
# PLOTLY VENDORING
# =====================================================================
def ensure_plotly(site_dir: str) -> bool:
    dest = os.path.join(site_dir, PLOTLY_FILE)
    if os.path.exists(dest) and os.path.getsize(dest) > 500_000:
        return True
    for url in [f"https://cdn.plot.ly/plotly-{PLOTLY_VERSION}.min.js",
                f"https://cdn.jsdelivr.net/npm/plotly.js-dist-min@{PLOTLY_VERSION}/plotly.min.js"]:
        try:
            with urllib.request.urlopen(url, timeout=60) as r:
                data = r.read()
            if len(data) > 500_000:
                with open(dest, "wb") as f:
                    f.write(data)
                print(f"  ✓ {PLOTLY_FILE} scaricato ({len(data)/1024/1024:.1f} MB)")
                return True
        except Exception as e:
            print(f"  · {url.split('/')[2]} non raggiungibile ({type(e).__name__})")
    try:
        url = (f"https://registry.npmjs.org/plotly.js-dist-min/-/"
               f"plotly.js-dist-min-{PLOTLY_VERSION}.tgz")
        with urllib.request.urlopen(url, timeout=120) as r:
            buf = io.BytesIO(r.read())
        with tarfile.open(fileobj=buf, mode="r:gz") as tf:
            with open(dest, "wb") as f:
                shutil.copyfileobj(tf.extractfile("package/plotly.min.js"), f)
        print(f"  ✓ {PLOTLY_FILE} scaricato da npm ({os.path.getsize(dest)/1024/1024:.1f} MB)")
        return True
    except Exception as e:
        print(f"  ⚠ non sono riuscito a scaricare Plotly ({type(e).__name__}); uso la CDN.")
        return False


def render_html(payload: dict, local_plotly: bool) -> str:
    tpl = HTML_TEMPLATE
    tpl = tpl.replace("__TITLE__", SITE_TITLE)
    tpl = tpl.replace("__TAGLINE__", SITE_TAGLINE)
    tpl = tpl.replace("__SUBTITLE__", SITE_SUBTITLE)
    tpl = tpl.replace("__FOOTER__", SITE_FOOTER)
    tpl = tpl.replace("__DISCLAIMER__", SITE_DISCLAIMER)
    tpl = tpl.replace("__PLOTLY_SRC__", PLOTLY_FILE if local_plotly
                      else f"https://cdn.plot.ly/plotly-{PLOTLY_VERSION}.min.js")
    tpl = tpl.replace("__PAYLOAD__",
                      json.dumps(payload, ensure_ascii=False, allow_nan=False)
                          .replace("</", "<\\/"))
    return tpl


HTML_TEMPLATE = r"""<!doctype html>
<html lang="en" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__ — __TAGLINE__</title>
<meta name="description" content="__SUBTITLE__">
<script src="__PLOTLY_SRC__" charset="utf-8"></script>
<style>
:root{
  color-scheme: light;
  --surface-1:#fcfcfb; --page:#f9f9f7;
  --text-primary:#0b0b0b; --text-secondary:#52514e; --muted:#898781;
  --grid:#e1e0d9; --axis:#c3c2b7; --border:rgba(11,11,11,0.10);
  --series-1:#2a78d6; --series-2:#eb6834; --series-3:#1baf7a;
  --warning:#fab219;
}
@media (prefers-color-scheme: dark){
  :root:where(:not([data-theme="light"])){
    color-scheme: dark;
    --surface-1:#1a1a19; --page:#0d0d0d;
    --text-primary:#ffffff; --text-secondary:#c3c2b7; --muted:#898781;
    --grid:#2c2c2a; --axis:#383835; --border:rgba(255,255,255,0.10);
    --series-1:#3987e5; --series-2:#d95926; --series-3:#199e70;
  }
}
:root[data-theme="dark"]{
  color-scheme: dark;
  --surface-1:#1a1a19; --page:#0d0d0d;
  --text-primary:#ffffff; --text-secondary:#c3c2b7; --muted:#898781;
  --grid:#2c2c2a; --axis:#383835; --border:rgba(255,255,255,0.10);
  --series-1:#3987e5; --series-2:#d95926; --series-3:#199e70;
}
*{box-sizing:border-box}
body{margin:0;background:var(--page);color:var(--text-primary);
  font-family:system-ui,-apple-system,"Segoe UI",sans-serif;font-size:15px;line-height:1.5}
.wrap{max-width:1120px;margin:0 auto;padding:0 20px 64px}

/* ── masthead ─────────────────────────────────────────────── */
header{padding:32px 0 20px;border-bottom:1px solid var(--border);margin-bottom:26px}
.topbar{display:flex;justify-content:space-between;align-items:baseline;gap:16px}
.topbar #theme{flex:0 0 auto}
h1{font-size:26px;margin:0;letter-spacing:-0.015em;flex:1 1 auto;min-width:0}
h1 span{color:var(--text-secondary);font-weight:400}
.stand{color:var(--text-secondary);max-width:66ch;margin:8px 0 0;font-size:14.5px}
.hero{font-size:21px;font-weight:500;max-width:62ch;margin:18px 0 0;line-height:1.35}
.hero b{font-weight:600}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:15px}
.chip{font-size:12.5px;color:var(--text-secondary);border:1px solid var(--border);
  border-radius:99px;padding:2px 10px;white-space:nowrap}
.chip.warn{border-color:var(--warning)}

h2{font-size:19px;margin:38px 0 4px;letter-spacing:-0.01em}
h2:first-of-type{margin-top:6px}
.hint{color:var(--text-secondary);font-size:13.5px;margin:0 0 14px;max-width:74ch}

button,select{font:inherit;font-size:13px;color:var(--text-primary);background:var(--surface-1);
  border:1px solid var(--border);border-radius:8px;padding:6px 11px;cursor:pointer}
button:hover,select:hover{border-color:var(--axis)}
button[aria-pressed="true"]:not(.tile){background:var(--text-primary);color:var(--surface-1);
  border-color:var(--text-primary)}
.linkbtn{border:none;background:none;padding:0;color:var(--text-secondary);
  text-decoration:underline;font-size:13px;cursor:pointer}
.linkbtn:hover{color:var(--text-primary)}

/* ── legenda condivisa ────────────────────────────────────── */
.mlegend{display:flex;flex-wrap:wrap;gap:14px;align-items:center;font-size:12px;
  color:var(--text-secondary);margin:0 0 12px}
.sw{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px;vertical-align:0}
.mlegend .div{width:1px;height:13px;background:var(--border)}

/* ── tiles ────────────────────────────────────────────────── */
.tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(202px,1fr));gap:10px}
.tile{border:1px solid var(--border);border-radius:12px;background:var(--surface-1);
  padding:12px 13px;text-align:left;width:100%;min-height:112px;position:relative}
.tile:hover{border-color:var(--axis)}
.tile[aria-pressed="true"]{box-shadow:inset 0 0 0 2px var(--text-primary);
  border-color:transparent}
.tile[aria-pressed="true"] .tname{color:var(--text-primary)}
.tile .tname{font-size:12.5px;color:var(--text-secondary);line-height:1.28;min-height:2.5em}
.tile .tval{font-size:26px;font-weight:600;letter-spacing:-0.02em;margin:6px 0 1px}
.tile .tdelta{font-size:13px;font-weight:600;color:var(--text-secondary);margin-left:6px}
.tile .tmeta{font-size:11.5px;color:var(--muted);font-variant-numeric:tabular-nums}
.tile .spark{display:block;margin-top:8px}
.tile.stale{border-style:dashed}
.tile.stale .tval{font-size:15px;font-weight:400;color:var(--text-secondary);letter-spacing:0}
.pill{display:inline-block;font-size:10.5px;border:1px solid var(--warning);border-radius:99px;
  padding:0 6px;color:var(--text-secondary);margin-left:6px;vertical-align:1px;white-space:nowrap}

/* ── panel ────────────────────────────────────────────────── */
.panel{background:var(--surface-1);border:1px solid var(--border);border-radius:12px;
  padding:16px;margin-top:14px}
.controls{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:12px}
.controls .div{width:1px;height:20px;background:var(--border);margin:0 3px}
.controls label{font-size:12.5px;color:var(--text-secondary);display:flex;
  align-items:center;gap:6px;cursor:pointer}
.body{height:470px}
.rankbody{height:380px}
.split2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media (max-width:900px){.split2{grid-template-columns:1fr;gap:24px}}
.subh{font-size:13.5px;font-weight:600;margin:0 0 2px}
.subnote{font-weight:400;color:var(--muted);font-size:12px;margin-left:6px}
.scrollbody{height:470px;overflow-y:auto;padding-right:6px}
.cap{font-size:12px;color:var(--muted);margin:10px 0 0}
.warnstrip{border:1px solid var(--warning);border-radius:8px;padding:9px 12px;
  font-size:13px;color:var(--text-secondary);margin-bottom:12px}

/* ── rail dei valori ──────────────────────────────────────── */
.rail{display:grid;grid-template-columns:repeat(auto-fit,minmax(168px,1fr));gap:6px 14px;
  margin-top:12px;padding-top:12px;border-top:1px solid var(--border)}
.rail div{font-size:12.5px;color:var(--text-secondary);font-variant-numeric:tabular-nums}
.rail b{color:var(--text-primary);font-weight:600}

/* ── evidence bar ─────────────────────────────────────────── */
.ebar{height:4px;border-radius:2px;background:var(--grid);overflow:hidden;margin-bottom:6px}
.ebar i{display:block;height:100%;background:var(--text-primary)}

table{border-collapse:collapse;width:100%;font-size:13px;font-variant-numeric:tabular-nums}
th,td{text-align:right;padding:7px 9px;border-bottom:1px solid var(--border)}
th:first-child,td:first-child{text-align:left;font-variant-numeric:normal}
th{color:var(--text-secondary);font-weight:600;white-space:nowrap}
details{margin-top:12px}
summary{cursor:pointer;color:var(--text-secondary);font-size:13px;padding:5px 0}
summary:hover{color:var(--text-primary)}
.reason{border-left:2px solid var(--border);padding:1px 0 1px 13px;margin:0 0 16px}
.reason .rh{font-size:12.5px;color:var(--text-secondary);margin-bottom:4px}
.reason .rb{font-size:13.5px}
.tag{display:inline-block;font-size:11px;border:1px solid var(--border);border-radius:99px;
  padding:1px 7px;color:var(--text-secondary);margin-left:6px;vertical-align:1px}
.quote{font-size:13.5px;border-left:2px solid var(--border);padding-left:13px;margin:14px 0 0;
  color:var(--text-secondary)}
.method p{font-size:13.5px;color:var(--text-secondary);max-width:76ch}
.method h3{font-size:14px;margin:20px 0 4px}
footer{margin-top:44px;padding-top:16px;border-top:1px solid var(--border);
  color:var(--muted);font-size:12.5px}
footer .disclaimer{color:var(--text-secondary);font-size:13px;max-width:70ch;
  margin:0 0 12px}
footer a{color:var(--text-secondary)}
.empty{color:var(--muted);font-size:13.5px;padding:24px 0;text-align:center}
@media (max-width:820px){.body,.scrollbody{height:400px}}
@media (max-width:640px){
  .body,.scrollbody{height:360px}h1{font-size:22px}.hero{font-size:18px}
  .tiles{grid-template-columns:repeat(auto-fill,minmax(150px,1fr))}
}
</style>
</head>
<body>
<div class="wrap">

<header>
  <div class="topbar">
    <h1>__TITLE__ <span>— __TAGLINE__</span></h1>
    <button id="theme" aria-label="Toggle theme">◐ Theme</button>
  </div>
  <p class="stand">__SUBTITLE__</p>
  <p class="hero" id="hero"></p>
  <div class="chips" id="chips"></div>
</header>

<h2>The next release</h2>
<p class="hint">Pick an indicator. The chart shows its published history and, in the shaded
   zone, where each model thinks the next release will land.</p>
<div class="mlegend" id="legend1"></div>
<div class="tiles" id="tiles"></div>

<div class="panel">
  <div class="controls">
    <button class="view" data-view="level"     aria-pressed="true">Forecast</button>
    <button class="view" data-view="revisions" aria-pressed="false">Revisions</button>
    <button class="view" data-view="reasoning" aria-pressed="false">Reasoning</button>
    <button class="view" data-view="table"     aria-pressed="false">Data</button>
    <span class="div"></span>
    <span id="rangeGroup">
      <button class="rng" data-r="2"  aria-pressed="true">2y</button>
      <button class="rng" data-r="5"  aria-pressed="false">5y</button>
      <button class="rng" data-r="10" aria-pressed="false">10y</button>
      <button class="rng" data-r="0"  aria-pressed="false">All</button>
    </span>
    <label id="pastWrap"><input type="checkbox" id="showPast" checked> Show past forecasts</label>
  </div>
  <div id="warnStrip"></div>
  <div id="mainBody" class="body"></div>
  <div id="rail" class="rail"></div>
  <p class="cap" id="cap"></p>
  <div id="quote"></div>
</div>

<h2>Scoreboard — has any model beaten a naive forecast?</h2>
<p class="hint">The benchmark is doing nothing: assume the next release equals the last
   published one. A model to the left of the dotted line has been closer than that
   benchmark so far; to the right, further away.</p>
<div class="mlegend" id="legend2"></div>
<div class="panel">
  <div class="ebar"><i id="ebarFill"></i></div>
  <p class="cap" id="ebarCap" style="margin:0 0 14px"></p>
  <div class="split2">
    <div>
      <div class="subh">All indicators together</div>
      <div id="chartRank" class="rankbody"></div>
    </div>
    <div>
      <div class="subh">Just <span id="vrName"></span>
        <span class="subnote">follows the indicator selected above</span></div>
      <div id="chartVarRank" class="rankbody"></div>
    </div>
  </div>
  <details id="evDetails"><summary>Show every scored forecast, indicator by indicator</summary>
    <div id="evidence"></div></details>
  <details><summary>Show the numbers</summary><div id="tblRank"></div></details>
</div>

<footer>
  <p class="disclaimer">__DISCLAIMER__</p>
  <p>__FOOTER__</p>
  <p>Full data: <a href="forecasts.csv" download>forecasts.csv</a> ·
     <a href="data.json" download>data.json</a> ·
     static page, rebuilt after every run — no server, nothing to fall asleep.</p>
</footer>

</div>

<script id="payload" type="application/json">__PAYLOAD__</script>
<script>
const D = JSON.parse(document.getElementById('payload').textContent);
const $ = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const esc = s => String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const fmt = v => (v===null||v===undefined||Number.isNaN(v)) ? '–' : Number(v).toFixed(2);

/* tre tinte per i tre modelli, due FORME per le due varianti: sei colori
   categorici non passerebbero il gate di distinguibilità, tre sì. */
const MODEL_COLOR = {};
function refreshColors(){
  const p = [css('--series-1'), css('--series-2'), css('--series-3')];
  D.models.forEach((m,i) => MODEL_COLOR[m] = p[i % p.length]);
}
refreshColors();
const SYM  = t => t === 'conditional' ? 'square'  : 'diamond';
const GLYPH= t => t === 'conditional' ? '■' : '◆';

const VARS = Object.keys(D.variables);
let curVar = (D.hero && D.hero.var) || VARS[0];
let curView = 'level';
let curRange = 2;
let showPast = true;

/* ── periodi ────────────────────────────────────────────────── */
function pTs(label){
  if(!label) return null;
  const q = /^(\d{4})-Q([1-4])$/.exec(label);
  if(q) return Date.UTC(+q[1], (+q[2]-1)*3, 1);
  const m = /^(\d{4})-(\d{2})$/.exec(label);
  if(m) return Date.UTC(+m[1], +m[2]-1, 1);
  return Date.parse(label);
}
const pDate = l => new Date(pTs(l));

function baseLayout(extra){
  const L = {
    paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
    font:{family:'system-ui,-apple-system,"Segoe UI",sans-serif', size:12.5,
          color:css('--text-secondary')},
    margin:{l:56,r:18,t:26,b:44}, hovermode:'closest',
    hoverlabel:{bgcolor:css('--surface-1'), bordercolor:css('--border'),
                font:{color:css('--text-primary'), size:12.5}},
    xaxis:{gridcolor:css('--grid'), linecolor:css('--axis'), zeroline:false,
           showgrid:false, tickfont:{color:css('--muted')}},
    yaxis:{gridcolor:css('--grid'), linecolor:'rgba(0,0,0,0)', zeroline:false,
           showgrid:true, automargin:true, tickfont:{color:css('--muted')}},
    legend:{orientation:'h', y:-0.16, x:0, font:{size:11.5, color:css('--text-secondary')},
            bgcolor:'rgba(0,0,0,0)'},
    showlegend:false
  };
  return Object.assign(L, extra||{});
}
const CFG = {displayModeBar:false, responsive:true};
const NARROW = () => window.innerWidth < 720;
function plot(id, traces, layout){
  const el = document.getElementById(id);
  if(el.style.height === 'auto') el.style.height = '';   /* solo se l'avevamo tolta noi */
  Plotly.react(id, traces, layout, CFG);
}
/* un grafico senza dati non deve lasciare in pagina un buco alto mezzo schermo */
function empty(id, msg){
  Plotly.purge(id);
  const el = document.getElementById(id);
  el.style.height = 'auto';
  el.innerHTML = '<p class="empty">' + msg + '</p>';
}

/* ── legenda condivisa ──────────────────────────────────────── */
function legendHTML(){
  const models = D.models.map(m =>
    `<span><span class="sw" style="background:${MODEL_COLOR[m]}"></span>${esc(D.model_short[m])}</span>`).join('');
  return models + '<span class="div"></span>' +
    '<span>◆ unconditional</span><span>■ Reddit-conditioned</span>';
}

/* ── masthead ───────────────────────────────────────────────── */
function renderHead(){
  const h = D.hero;
  $('#hero').innerHTML = h
    ? `Next up: <b>${esc(h.short)}</b> for ${esc(h.target_pretty)}. ` +
      `The models put it at <b>${fmt(h.consensus)}%</b>, against <b>${fmt(h.last_value)}%</b> ` +
      `in ${esc(h.last_period_pretty)}.`
    : 'No current forecast available.';
  const chips = [
    `${D.n_runs} daily runs since ${D.first_run}`,
    `${D.models.length} models · ${D.n_forecasts} forecasts`,
    `${D.scored_targets} releases scored so far`,
  ];
  let html = chips.map(c => `<span class="chip">${esc(c)}</span>`).join('');
  if(D.n_stale) html += `<span class="chip warn">⚠ ${D.n_stale} series stopped updating</span>`;
  html += `<span class="chip">built ${esc(D.generated_at)}</span>`;
  $('#chips').innerHTML = html;
}

/* ── tiles ──────────────────────────────────────────────────── */
function sparkSVG(vals, muted){
  if(!vals || vals.length < 2) return '';
  const w=62,h=18,lo=Math.min(...vals),hi=Math.max(...vals),rg=(hi-lo)||1;
  const pts = vals.map((v,i) =>
    `${(i/(vals.length-1)*(w-3)+1.5).toFixed(1)},${(h-2-((v-lo)/rg)*(h-4)).toFixed(1)}`);
  const last = pts[pts.length-1].split(',');
  return `<svg class="spark" width="${w}" height="${h}" aria-hidden="true">
    <polyline fill="none" stroke="${muted?css('--muted'):css('--axis')}" stroke-width="1.5"
      points="${pts.join(' ')}"/>
    <circle cx="${last[0]}" cy="${last[1]}" r="2.6"
      fill="${muted?css('--muted'):css('--text-primary')}"/></svg>`;
}

function renderTiles(){
  $('#tiles').innerHTML = D.tiles.map(t => {
    const v = D.variables[t.var];
    if(t.stale){
      return `<button class="tile stale" data-var="${t.var}" aria-pressed="${t.var===curVar}">
        <div class="tname">${esc(t.short)}<span class="pill">paused</span></div>
        <div class="tval">no release since ${esc(t.last_period)}</div>
        <div class="tmeta">forecasting stopped ${esc(t.run)}</div>
        ${sparkSVG(v.spark, true)}</button>`;
    }
    const d = (t.last_value===null) ? null : t.consensus - t.last_value;
    const dTxt = d===null ? '' : `<span class="tdelta">${
      d>0.005?'▲':(d<-0.005?'▼':'■')} ${Math.abs(d).toFixed(2)}</span>`;
    return `<button class="tile" data-var="${t.var}" aria-pressed="${t.var===curVar}">
      <div class="tname">${esc(t.short)}</div>
      <div class="tval">${fmt(t.consensus)}${dTxt}</div>
      <div class="tmeta">${esc(t.target)} · ${t.n} forecasts · ${fmt(t.lo)}–${fmt(t.hi)}</div>
      ${sparkSVG(v.spark, false)}</button>`;
  }).join('');
  $$('.tile').forEach(b => b.onclick = () => {
    curVar = b.dataset.var; syncTiles(); renderPanel(); renderVarRank(); });
}
function syncTiles(){
  $$('.tile').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.var === curVar)));
}

/* ── il grafico principale ──────────────────────────────────── */
function levelChart(){
  const v = D.variables[curVar];
  const per = v.freq === 'M' ? 1 : 3;
  let hist = v.history;
  if(curRange > 0) hist = hist.slice(-Math.round(curRange*12/per));
  if(!hist.length){ empty('mainBody', 'Official series not available.');
    $('#rail').innerHTML = ''; $('#cap').textContent = ''; return; }

  const lastP = hist[hist.length-1][0], lastV = hist[hist.length-1][1];
  const msL = pTs(lastP), lo = pTs(hist[0][0]);
  const tgt = (!v.stale && v.target) ? v.target : null;
  const msT = tgt ? pTs(tgt) : null;
  const step = msT ? Math.max(msT - msL, 1) : 1;

  /* La zona di previsione occupa sempre almeno l'8% della larghezza, a
     qualunque zoom: è questo che permette di avere UN grafico invece di due.
     Allarga il passo oltre il tempo realmente trascorso, quindi la banda
     ombreggiata e la didascalia non sono decorative — servono a dirlo. */
  let R = msT ? Math.min(Math.max((msL - 0.08*lo)/0.92, msT + 0.4*step), msT + 12*step)
              : msL + 0.02*(msL - lo);
  const xr = [new Date(lo - 0.01*(R - lo)), new Date(R)];

  const nvis = hist.length;
  const traces = [{
    x: hist.map(r => pDate(r[0])), y: hist.map(r => r[1]),
    type:'scatter', mode: nvis < 40 ? 'lines+markers' : 'lines', name:'Official data',
    line:{color:css('--text-primary'), width:2},
    marker:{size:5.5, color:css('--text-primary'),
            line:{color:css('--surface-1'), width:1.5}},
    hovertemplate:'%{x|%b %Y}<br><b>%{y:.2f}</b><extra>published</extra>'
  }];

  /* previsioni passate già valutate: la memoria del grafico */
  if(showPast && v.hindcasts.length){
    const segX=[], segY=[];
    v.hindcasts.forEach(h => { segX.push(pDate(h.t), pDate(h.t), null);
                               segY.push(h.f, h.a, null); });
    traces.push({x:segX, y:segY, type:'scatter', mode:'lines',
      line:{color:css('--axis'), width:1}, hoverinfo:'skip', showlegend:false});
    D.models.forEach(m => {
      const hs = v.hindcasts.filter(h => h.model === m);
      if(!hs.length) return;
      traces.push({
        x: hs.map(h => pDate(h.t)), y: hs.map(h => h.f), type:'scatter', mode:'markers',
        marker:{size:8, color:MODEL_COLOR[m], opacity:0.5,
                symbol: hs.map(h => SYM(h.type)+'-open'),
                line:{color:MODEL_COLOR[m], width:1.5}},
        showlegend:false,
        customdata: hs.map(h => [h.a, (h.f-h.a), h.rw===null?'–':h.rw.toFixed(2)]),
        hovertemplate: D.model_short[m]+'<br>forecast <b>%{y:.2f}</b> for %{x|%b %Y}'+
          '<br>published %{customdata[0]:.2f} · miss %{customdata[1]:+.2f}'+
          '<br>naive missed %{customdata[2]}<extra>scored</extra>'
      });
    });
  }

  /* la previsione corrente: un punto per modello-variante */
  if(tgt && v.forecasts.length){
    const med = v.forecasts.map(f => f.value).sort((a,b)=>a-b);
    const mid = med.length%2 ? med[(med.length-1)/2]
                             : (med[med.length/2-1]+med[med.length/2])/2;
    traces.push({x:[pDate(lastP), pDate(tgt)], y:[lastV, mid], type:'scatter', mode:'lines',
      line:{color:css('--text-secondary'), width:1.5, dash:'dot'},
      hoverinfo:'skip', showlegend:false});

    /* Previsioni troppo vicine si coprono a vicenda: le scosto di poco in
       orizzontale, a gruppi. Sono tutte per lo stesso periodo — lo scarto è
       solo per poterle vedere, e la didascalia lo dice. */
    const ys = hist.map(r => r[1]).concat(v.forecasts.map(f => f.value));
    if(showPast && v.hindcasts.length) v.hindcasts.forEach(h => ys.push(h.f, h.a));
    const yrange = Math.max(...ys) - Math.min(...ys) || 1;
    const thr = 0.025 * yrange;
    const byVal = v.forecasts.slice().sort((a,b) => a.value - b.value);
    const dx = new Map();
    let i = 0, nudged = false;
    while(i < byVal.length){
      let j = i + 1;
      while(j < byVal.length && byVal[j].value - byVal[j-1].value < thr) j++;
      const k = j - i;
      for(let q = i; q < j; q++){
        const off = k > 1 ? (q - i - (k-1)/2) * 0.30 * step : 0;
        dx.set(byVal[q], off);
        if(off !== 0) nudged = true;
      }
      i = j;
    }
    window._nudged = nudged;

    v.forecasts.forEach(f => {
      const tr = v.track.find(t => t.model===f.model && t.type===f.type);
      const rec = tr ? `<br>on this indicator: ${tr.n} scored · MAE ${fmt(tr.mae)}` +
                       (tr.rw!==null?` · naive ${fmt(tr.rw)}`:'') : '';
      traces.push({
        x:[new Date(msT + (dx.get(f) || 0))], y:[f.value],
        type:'scatter', mode:'markers', name:f.label,
        marker:{size:13, color:MODEL_COLOR[f.model], symbol:SYM(f.type),
                line:{color:css('--surface-1'), width:2}},
        showlegend:false,
        hovertemplate:`<b>%{y:.2f}</b> for ${esc(v.target_pretty)}${rec}` +
                      (f.web?'<br>web search: yes':'') +
                      `<extra>${esc(f.label)}</extra>`
      });
    });
  }

  const shapes = [], annos = [];
  if(msT){
    shapes.push({type:'rect', xref:'x', yref:'paper', y0:0, y1:1,
      x0:new Date(msL), x1:xr[1], fillcolor:css('--text-primary'), opacity:0.045,
      line:{width:0}, layer:'below'});
    /* ancorata al bordo destro del plot: centrata sulla banda sborderebbe */
    annos.push({x:1, y:1, xref:'paper', yref:'paper', xanchor:'right', yanchor:'bottom',
      text: NARROW() ? esc(v.target)+' · forecast' : esc(v.target_pretty)+' · not yet published',
      showarrow:false, font:{size:11, color:css('--muted')}});
  }

  plot('mainBody', traces, baseLayout({
    xaxis:Object.assign(baseLayout().xaxis, {range:xr}),
    yaxis:Object.assign(baseLayout().yaxis, {ticksuffix:'%  '}),
    shapes, annotations:annos
  }));

  /* rail: i sei numeri, leggibili senza passarci sopra col mouse */
  if(tgt && v.forecasts.length){
    const sorted = v.forecasts.slice().sort((a,b) => b.value - a.value);
    $('#rail').innerHTML = sorted.map(f =>
      `<div><span style="color:${MODEL_COLOR[f.model]}">${GLYPH(f.type)}</span> ` +
      `${esc(D.model_short[f.model])} <b>${fmt(f.value)}</b></div>`).join('');
  } else $('#rail').innerHTML = '';

  $('#cap').textContent = msT
    ? `The shaded zone is one ${v.freq==='M'?'month':'quarter'} ahead, drawn wider than ` +
      `elapsed time so the forecasts stay readable at every zoom level.` +
      (window._nudged
        ? ` Forecasts that land on nearly the same value are nudged sideways to stop them ` +
          `hiding each other — they are all for the same period.` : '') +
      (showPast && v.hindcasts.length
        ? ` Hollow marks are earlier forecasts, joined to what was actually published.` : '')
    : '';
}

/* ── revisioni ──────────────────────────────────────────────── */
function revisionsView(){
  const v = D.variables[curVar];
  const keys = Object.keys(v.revisions).sort().reverse();
  if(!keys.length){
    empty('mainBody', 'Revisions appear once a target has been forecast on more than one day.');
    $('#rail').innerHTML = ''; $('#cap').textContent = ''; return; }
  const sel = window._revTarget && v.revisions[window._revTarget] ? window._revTarget : keys[0];
  const R = v.revisions[sel];
  const traces = R.lines.map(l => ({
    x: l.x.map(d => new Date(d)), y: l.y, type:'scatter',
    mode: l.y.length > 1 ? 'lines+markers' : 'markers', name: l.label,
    line:{color:MODEL_COLOR[l.model], width:2,
          dash: l.type==='conditional' ? 'dash' : 'solid'},
    marker:{color:MODEL_COLOR[l.model], size:8, symbol:SYM(l.type),
            line:{color:css('--surface-1'), width:2}},
    hovertemplate:'%{x|%d %b}<br><b>%{y:.2f}</b><extra>'+esc(l.label)+'</extra>'
  }));
  const allX = R.lines.flatMap(l => l.x.map(d => Date.parse(d)));
  const lo = Math.min(...allX), hi = Math.max(...allX), DAY = 864e5;
  const pad = Math.max((hi-lo)*0.06, 2*DAY);
  if(R.realized !== null && R.realized !== undefined){
    traces.push({x:[new Date(lo-pad), new Date(hi+pad)], y:[R.realized, R.realized],
      type:'scatter', mode:'lines', name:'published',
      line:{color:css('--text-primary'), width:1.6, dash:'dash'},
      hovertemplate:'published: <b>%{y:.2f}</b><extra></extra>'});
  }
  plot('mainBody', traces, baseLayout({
    showlegend:true,
    xaxis:Object.assign(baseLayout().xaxis, {range:[new Date(lo-pad), new Date(hi+pad)],
      tickformat:'%d %b', dtick:(hi-lo+2*pad) < 12*DAY ? DAY : undefined,
      title:{text:'run date', font:{size:11.5}}})
  }));
  $('#rail').innerHTML = '';
  $('#cap').innerHTML = `Target period: <select id="revSel">` +
    keys.map(k => `<option value="${k}" ${k===sel?'selected':''}>${esc(k)}</option>`).join('') +
    `</select> — each line is one model's forecast of that same release, ` +
    `revised run after run.` +
    (R.realized!==null && R.realized!==undefined
      ? ` The dashed line is the value eventually published.` : '');
  $('#revSel').onchange = e => { window._revTarget = e.target.value; renderPanel(); };
}

/* ── reasoning / tabella ────────────────────────────────────── */
function reasoningView(){
  const v = D.variables[curVar];
  const w = v.forecasts.filter(f => f.reasoning && f.reasoning.length > 20);
  $('#mainBody').className = 'scrollbody';
  $('#mainBody').innerHTML = w.length ? w.map(f => `
    <div class="reason" style="border-left-color:${MODEL_COLOR[f.model]}">
      <div class="rh"><b>${esc(D.model_short[f.model])}</b>
        <span class="tag">${f.type==='conditional'?'Reddit-conditioned':'unconditional'}</span>
        <span class="tag">forecast ${fmt(f.value)}</span>
        ${f.web?'<span class="tag">web search</span>':''}</div>
      <div class="rb">${esc(f.reasoning)}</div></div>`).join('')
    : '<p class="empty">No reasoning text for this indicator.</p>';
  $('#rail').innerHTML = ''; $('#cap').textContent = '';
}

function tableView(){
  const v = D.variables[curVar];
  const hist = v.history.slice(-12).reverse();
  let t1 = v.forecasts.length ? `<table><thead><tr><th>Model</th><th>Type</th>
      <th>Forecast ${esc(v.target)}</th><th>vs last</th><th>Web search</th></tr></thead><tbody>` +
    v.forecasts.map(f => `<tr><td><span class="sw" style="background:${MODEL_COLOR[f.model]}"></span>${
      esc(D.model_short[f.model])}</td><td>${f.type==='conditional'?'conditional':'uncond.'}</td>
      <td>${fmt(f.value)}</td><td>${v.last_official && v.last_official.value!==null
        ? (f.value - v.last_official.value>=0?'+':'') + (f.value - v.last_official.value).toFixed(2) : '–'}</td>
      <td>${f.web?'yes':'no'}</td></tr>`).join('') + '</tbody></table>' : '';
  let t2 = `<table style="margin-top:16px"><thead><tr><th>Period</th>
      <th>Published (${esc(v.unit)})</th></tr></thead><tbody>` +
    hist.map(h => `<tr><td>${esc(h[0])}</td><td>${fmt(h[1])}</td></tr>`).join('') +
    '</tbody></table>';
  $('#mainBody').className = 'scrollbody';
  $('#mainBody').innerHTML = t1 + t2;
  $('#rail').innerHTML = ''; $('#cap').textContent = '';
}

/* ── pannello ───────────────────────────────────────────────── */
function renderPanel(){
  const v = D.variables[curVar];
  $('#rangeGroup').style.display = curView==='level' ? '' : 'none';
  $('#pastWrap').style.display   = (curView==='level' && v.hindcasts.length) ? '' : 'none';
  $('#warnStrip').innerHTML = v.stale
    ? `<div class="warnstrip">⚠ <b>${esc(v.short)}</b> has not been published since
       ${esc(v.history.length ? v.history[v.history.length-1][0] : '—')}, so the pipeline
       stopped forecasting it on ${esc(v.forecast_run)}. No current forecast is shown.</div>` : '';
  $('#quote').innerHTML = '';

  const mb = $('#mainBody');
  if(curView !== 'level' && curView !== 'revisions'){
    Plotly.purge('mainBody'); mb.style.height = '';
  } else {
    mb.className = 'body';
    /* tornando da una vista HTML il div contiene ancora la tabella:
       Plotly.react ci disegnerebbe sopra invece di sostituirla */
    if(!mb._fullLayout) mb.innerHTML = '';
  }

  if(curView === 'level'){ levelChart(); renderQuote(); }
  else if(curView === 'revisions') revisionsView();
  else if(curView === 'reasoning') reasoningView();
  else tableView();
}

function renderQuote(){
  const v = D.variables[curVar];
  if(v.stale) return;
  const w = v.forecasts.filter(f => f.reasoning && f.reasoning.length > 40);
  if(!w.length) return;
  const vals = v.forecasts.map(f => f.value).sort((a,b)=>a-b);
  const mid = vals.length%2 ? vals[(vals.length-1)/2] : (vals[vals.length/2-1]+vals[vals.length/2])/2;
  const pick = w.slice().sort((a,b) => Math.abs(a.value-mid) - Math.abs(b.value-mid))[0];
  const first = (pick.reasoning.split(/(?<=[.!?])\s+/)[0] || '').slice(0, 300);
  $('#quote').innerHTML = `<div class="quote" style="border-left-color:${MODEL_COLOR[pick.model]}">
    “${esc(first)}”<br><span style="font-size:12.5px;color:var(--muted)">
    <span class="sw" style="background:${MODEL_COLOR[pick.model]}"></span>
    ${esc(D.model_short[pick.model])} · ${pick.type==='conditional'?'Reddit-conditioned':'unconditional'}
    — <button class="linkbtn" id="readAll">read all ${v.forecasts.length} →</button></span></div>`;
  $('#readAll').onclick = () => setView('reasoning');
}

/* ── scoreboard ─────────────────────────────────────────────── */
function renderRank(){
  const rows = (D.scoreboard || []).filter(r => r.skill !== null);
  const tot = D.scored_targets + D.pending_targets;
  $('#ebarFill').style.width = tot ? (100*D.scored_targets/tot).toFixed(0)+'%' : '0%';
  $('#ebarCap').textContent = D.scored_targets === 0
    ? `Nothing scored yet: ${D.pending_targets} forecasts are waiting for their release.`
    : `${D.scored_targets} releases have been published and scored; ${D.pending_targets} more ` +
      `forecasts are waiting. Each model has ${rows.length?rows[0].n:0} scored forecasts — ` +
      `few enough to count by hand, and far too few to separate these models.`;

  if(!rows.length){
    empty('chartRank', 'No release has been published yet after a forecast. ' +
      'The scoreboard fills in on its own as the statistics come out.');
    $('#tblRank').innerHTML = ''; return;
  }
  const nar0 = ($('#chartRank').clientWidth || 800) < 640;
  const shortLab = r => {
    let s = D.model_short[r.model];
    if(nar0 && s.length > 10) s = s.split('-').slice(0,2).join('-');
    return s + (r.type==='conditional' ? ' · cond.' : '');
  };
  const disp = rows.slice().reverse();          /* migliore in alto */
  const labels = disp.map(shortLab);
  const traces = [{
    x: disp.map(r => r.skill), y: labels, type:'scatter', mode:'markers',
    marker:{size: nar0 ? 13 : 15, color: disp.map(r => MODEL_COLOR[r.model]),
            symbol: disp.map(r => SYM(r.type)),
            line:{color:css('--surface-1'), width:2}},
    customdata: disp.map(r => [r.wins, r.n, r.mae, r.rw, r.loo_lo, r.loo_hi]),
    hovertemplate:'skill <b>%{x:.2f}</b> × the naive error'+
      '<br>beat the benchmark on %{customdata[0]} of %{customdata[1]} forecasts'+
      '<br>MAE %{customdata[2]:.2f} vs naive %{customdata[3]:.2f}'+
      '<extra>%{y}</extra>', showlegend:false
  }];
  const segX=[], segY=[];
  disp.forEach(r => { const L = shortLab(r);
    segX.push(1, r.skill, null); segY.push(L, L, null); });
  traces.unshift({x:segX, y:segY, type:'scatter', mode:'lines',
    line:{color:css('--axis'), width:1.5}, hoverinfo:'skip', showlegend:false});

  const nar = ($('#chartRank').clientWidth || 800) < 640;
  plot('chartRank', traces, baseLayout({
    margin: nar ? {l:104, r:22, t:30, b:50} : {l:190, r:120, t:46, b:52},
    xaxis:{type:'log', range:[Math.log10(0.42), Math.log10(2.6)],
      tickvals: nar ? [0.5,1,2] : [0.5,0.71,1,1.41,2],
      ticktext: nar ? ['2× closer','naive','2× worse']
                    : ['2× closer','1.4× closer','same as naive','1.4× worse','2× worse'],
      showgrid:false, zeroline:false, linecolor:css('--axis'),
      tickfont:{color:css('--muted'), size: nar ? 10.5 : 11.5}},
    yaxis:{automargin:true, tickfont:{color:css('--text-secondary'), size: nar ? 10.5 : 12},
      linecolor:'rgba(0,0,0,0)', showgrid:false},
    shapes:[
      {type:'rect', xref:'x', yref:'paper', x0:0.42, x1:1, y0:0, y1:1,
       fillcolor:css('--text-primary'), opacity:0.03, line:{width:0}, layer:'below'},
      {type:'line', x0:1, x1:1, xref:'x', yref:'paper', y0:0, y1:1,
       line:{color:css('--axis'), width:1, dash:'dot'}}
    ],
    annotations:[
      {x:Math.log10(0.45), y:0, xref:'x', yref:'paper',
       text: nar ? '' : 'beats doing nothing',
       showarrow:false, xanchor:'left', yanchor:'bottom',
       font:{size:11, color:css('--muted')}},
      {x:0, y:1.10, xref:'paper', yref:'paper', xanchor:'left',
       showarrow:false, align:'left', visible: !nar,
       text:`Each dot is one model, scored on ${rows[0].n} forecasts. `+
            `That is not enough to rank them — read it as a first look, not a verdict.`,
       font:{size:11.5, color:css('--text-secondary')}}
    ]
  }));

  $('#tblRank').innerHTML = `<table><thead><tr><th>Model</th><th>Type</th>
    <th>Indicators</th><th>Scored</th><th>MAE</th><th>Bias</th><th>Naive MAE</th>
    <th>Skill</th><th>Beat naive</th><th>Rank if one dropped</th></tr></thead><tbody>` +
    (D.scoreboard||[]).map(r => `<tr>
      <td><span class="sw" style="background:${MODEL_COLOR[r.model]}"></span>${esc(D.model_short[r.model])}</td>
      <td>${r.type==='conditional'?'conditional':'uncond.'}</td>
      <td>${r.vars}</td><td>${r.n}</td><td>${fmt(r.mae)}</td><td>${fmt(r.bias)}</td>
      <td>${fmt(r.rw)}</td><td>${r.skill===null?'–':fmt(r.skill)}</td>
      <td>${r.wins} / ${r.n}</td>
      <td>${r.loo_lo===null?'–':(r.loo_lo===r.loo_hi?r.loo_lo:r.loo_lo+'–'+r.loo_hi)}</td>
    </tr>`).join('') + '</tbody></table>' +
    `<p class="cap">Skill = geometric mean across indicators of (model MAE ÷ naive MAE),
     each ratio clipped to [0.25, 4]. "Rank if one dropped" recomputes the ranking with a
     single scored forecast removed, one at a time: the spread is how much the ordering
     depends on any one observation.</p>`;
}

/* ── scoreboard della singola variabile ─────────────────────────
   Qui n è 1 o 2 per modello: una media sarebbe un numero solo travestito
   da classifica. Quindi ordino per errore medio ma DISEGNO ogni singola
   previsione, così si vede su quanto poco si sta giudicando. */
function renderVarRank(){
  const v = D.variables[curVar];
  $('#vrName').textContent = v.short;
  const rows = (v.track || []).slice().reverse();   /* migliore in alto */
  if(!rows.length){
    empty('chartVarRank', 'No release of this indicator has been published yet ' +
      'after a forecast.');
    return;
  }
  const nar = ($('#chartVarRank').clientWidth || 500) < 640;
  const lab = r => {
    let s = D.model_short[r.model];
    if(nar && s.length > 10) s = s.split('-').slice(0,2).join('-');
    return s + (r.type==='conditional' ? ' · cond.' : '');
  };
  const rw = rows.map(r => r.rw).find(x => x !== null && x !== undefined);
  const ns = rows.map(r => r.n), nmin = Math.min(...ns), nmax = Math.max(...ns);

  const traces = [];
  /* il segmento va dal benchmark al valore del modello: la lunghezza È il divario */
  const segX=[], segY=[];
  rows.forEach(r => { const L = lab(r);
    segX.push(rw === undefined ? 0 : rw, r.mae, null); segY.push(L, L, null); });
  traces.push({x:segX, y:segY, type:'scatter', mode:'lines',
    line:{color:css('--axis'), width:1.5}, hoverinfo:'skip', showlegend:false});

  /* ogni previsione valutata, in chiaro */
  D.models.forEach(m => {
    const xs=[], ys=[], syms=[], cd=[];
    rows.filter(r => r.model===m).forEach(r => r.points.forEach(pt => {
      xs.push(Math.abs(pt.e)); ys.push(lab(r)); syms.push(SYM(r.type));
      cd.push([pt.t, pt.e]);
    }));
    if(xs.length) traces.push({x:xs, y:ys, type:'scatter', mode:'markers',
      marker:{size:9, color:MODEL_COLOR[m], opacity:0.45, symbol:syms,
              line:{color:css('--surface-1'), width:1.5}},
      customdata:cd, showlegend:false,
      hovertemplate:'%{customdata[0]}: missed by %{customdata[1]:+.2f} pp<extra></extra>'});
  });
  /* la media, il mark grande */
  rows.forEach(r => traces.push({
    x:[r.mae], y:[lab(r)], type:'scatter', mode:'markers',
    marker:{size:14, color:MODEL_COLOR[r.model], symbol:SYM(r.type),
            line:{color:css('--surface-1'), width:2}},
    customdata:[[r.n, r.rw]], showlegend:false,
    hovertemplate:'average miss <b>%{x:.2f} pp</b> over %{customdata[0]} forecast(s)'+
      '<br>naive missed %{customdata[1]:.2f} pp<extra>'+esc(lab(r))+'</extra>'}));

  const shapes = [], annos = [];
  if(rw !== undefined){
    shapes.push({type:'line', x0:rw, x1:rw, xref:'x', yref:'paper', y0:0, y1:1,
      line:{color:css('--axis'), width:1, dash:'dot'}});
    shapes.push({type:'rect', xref:'x', yref:'paper', x0:0, x1:rw, y0:0, y1:1,
      fillcolor:css('--text-primary'), opacity:0.03, line:{width:0}, layer:'below'});
    annos.push({x:rw, y:1, xref:'x', yref:'paper', yanchor:'bottom', xanchor:'left',
      text:' naive missed '+fmt(rw), showarrow:false,
      font:{size:11, color:css('--muted')}});
  }
  plot('chartVarRank', traces, baseLayout({
    margin: nar ? {l:104, r:26, t:30, b:50} : {l:150, r:34, t:30, b:52},
    xaxis:{rangemode:'tozero', showgrid:false, zeroline:false,
      linecolor:css('--axis'), ticksuffix:' pp',
      tickfont:{color:css('--muted'), size: nar ? 10.5 : 11.5},
      title:{text:'average miss · faint marks are the individual forecasts (' +
             (nmin===nmax ? 'n = '+nmin : 'n = '+nmin+'–'+nmax) + ' per model)',
             font:{size:11, color:css('--muted')}}},
    yaxis:{automargin:true, showgrid:false, linecolor:'rgba(0,0,0,0)',
      tickfont:{color:css('--text-secondary'), size: nar ? 10.5 : 12}},
    shapes, annotations:annos
  }));
}

/* ── evidence strip (una tacca per previsione valutata) ─────── */
let evidenceDrawn = false;
function renderEvidence(){
  const host = $('#evidence');
  const keys = Object.keys(D.evidence || {});
  if(!keys.length){ host.innerHTML = '<p class="empty">Nothing scored yet.</p>'; return; }
  host.innerHTML = `<p class="cap" style="margin:6px 0 10px">Distance from the published value,
    in percentage points. Each panel has its own scale. ✕ marks where the naive
    benchmark landed. Left of the line is an undershoot, right an overshoot.</p>` +
    keys.map((k,i) => `<div style="font-size:12.5px;color:var(--text-secondary);margin:14px 0 2px">
      ${esc(D.evidence[k].name)}</div><div id="ev${i}"></div>`).join('');
  keys.forEach((k,i) => {
    const E = D.evidence[k];
    const traces = [];
    D.models.forEach(m => {
      const xs=[], ys=[], syms=[];
      E.targets.forEach(t => t.marks.filter(mk => mk.model===m).forEach(mk => {
        xs.push(mk.err); ys.push(t.t); syms.push(SYM(mk.type));
      }));
      if(xs.length) traces.push({x:xs, y:ys, type:'scatter', mode:'markers',
        marker:{size:10, color:MODEL_COLOR[m], symbol:syms,
                line:{color:css('--surface-1'), width:2}},
        hovertemplate:'miss %{x:+.2f}<extra>'+esc(D.model_short[m])+'</extra>', showlegend:false});
    });
    const rwx = E.targets.filter(t => t.rw!==null).map(t => t.rw);
    const rwy = E.targets.filter(t => t.rw!==null).map(t => t.t);
    if(rwx.length) traces.push({x:rwx, y:rwy, type:'scatter', mode:'markers',
      marker:{size:10, color:css('--muted'), symbol:'x-thin',
              line:{color:css('--muted'), width:2}},
      hovertemplate:'naive missed %{x:+.2f}<extra>benchmark</extra>', showlegend:false});
    Plotly.react('ev'+i, traces, baseLayout({
      height: 48 + 34*E.targets.length + 30,
      margin:{l:78, r:24, t:22, b:32},
      xaxis:Object.assign(baseLayout().xaxis, {showgrid:false, zeroline:false,
        ticksuffix:' pp'}),
      yaxis:{type:'category', automargin:true, showgrid:false,
        linecolor:'rgba(0,0,0,0)', tickfont:{color:css('--muted'), size:11.5}},
      shapes:[{type:'line', x0:0, x1:0, xref:'x', yref:'paper', y0:0, y1:1,
               line:{color:css('--text-primary'), width:1}}],
      annotations:[{x:0, y:1, xref:'x', yref:'paper', text:'published', showarrow:false,
        yanchor:'bottom', font:{size:10.5, color:css('--muted')}}]
    }), CFG);
  });
  evidenceDrawn = true;
}

/* ── orchestrazione ─────────────────────────────────────────── */
function setView(v){
  curView = v;
  $$('.view').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.view===v)));
  renderPanel();
}
function renderAll(){
  refreshColors();
  $('#legend1').innerHTML = legendHTML();
  $('#legend2').innerHTML = legendHTML();
  renderHead(); renderTiles(); renderPanel(); renderRank(); renderVarRank();
  if(evidenceDrawn) renderEvidence();
}

$$('.view').forEach(b => b.onclick = () => setView(b.dataset.view));
$$('.rng').forEach(b => b.onclick = () => {
  curRange = +b.dataset.r;
  $$('.rng').forEach(o => o.setAttribute('aria-pressed', String(o===b)));
  renderPanel();
});
$('#showPast').onchange = e => { showPast = e.target.checked; renderPanel(); };
$('#evDetails').addEventListener('toggle', e => {
  if(e.target.open && !evidenceDrawn) renderEvidence();
});
$('#theme').onclick = () => {
  const cur = document.documentElement.getAttribute('data-theme');
  document.documentElement.setAttribute('data-theme', cur==='dark' ? 'light' : 'dark');
  renderAll();
};
if(window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches)
  document.documentElement.setAttribute('data-theme','dark');

let _rz, _w = window.innerWidth;
window.addEventListener('resize', () => {
  if(Math.abs(window.innerWidth - _w) < 40) return;
  _w = window.innerWidth;
  clearTimeout(_rz);
  _rz = setTimeout(() => { renderPanel(); renderRank(); renderVarRank(); }, 250);
});

renderAll();
</script>
</body>
</html>
"""


# =====================================================================
# DEMO
# =====================================================================
def make_demo():
    rng = np.random.default_rng(7)
    models = [("jrc", "llama-3.3-70b-instruct"), ("jrc", "gpt-oss-120b"), ("ec", "gpt-5.1")]
    series, rows = {}, []
    for var in VARIABLES:
        q = "qoq" in var
        dates = pd.date_range("2011-01-01", "2026-07-01", freq="QS" if q else "MS")
        base = 0.5 if q else (2.5 if "cpi" in var or "hicp" in var else 5.0)
        vals = base + np.cumsum(rng.normal(0, 0.12, len(dates))) * (0.4 if q else 1.0)
        series[var] = pd.DataFrame({"date": dates, "value": np.round(vals, 2)})
    for rd in pd.date_range("2025-10-01", "2026-08-20", freq="7D"):
        for var in VARIABLES:
            sdf = series[var]
            freq = "Q" if "qoq" in var else "M"
            avail = sdf[sdf["date"] <= rd - pd.Timedelta(days=35)]
            if avail.empty:
                continue
            last = avail.iloc[-1]
            for prov, mod in models:
                for ftype in ("unconditional", "conditional"):
                    rows.append({
                        "run_date": rd.strftime("%Y-%m-%d"), "provider": prov, "model": mod,
                        "variable": var, "variable_name": VARIABLES[var]["short"],
                        "forecast_type": ftype,
                        "last_official_period": ts_to_period(last["date"], freq),
                        "last_official_value": float(last["value"]),
                        "point_forecast": round(float(last["value"] + rng.normal(0, 0.22)), 3),
                        "target_period": next_period(ts_to_period(last["date"], freq)),
                        "used_web_search": bool(rng.random() > 0.5),
                        "reasoning": ("Demo text. The series shows a gradual slowdown; base "
                                      "effects in energy and cooling domestic demand point to a "
                                      "reading close to the previous one."),
                    })
    subs = ["inflation", "economics", "europe", "italy", "wallstreetbets",
            "personalfinance", "investing", "germany", "france"]
    days = pd.date_range("2026-05-01", "2026-08-20", freq="D")
    reddit = pd.concat([
        pd.DataFrame({"date": days, "subreddit": s,
                      "mentions_total": rng.poisson(40/(i+1)+5, len(days))})
        for i, s in enumerate(subs)], ignore_index=True)
    return pd.DataFrame(rows), series, reddit


def load_master_from_df(df: pd.DataFrame) -> pd.DataFrame:
    tmp = os.path.join(os.path.expanduser("~"), ".live_project_demo_master.xlsx")
    df.to_excel(tmp, index=False)
    out = load_master(tmp)
    os.remove(tmp)
    return out


# =====================================================================
# PUBLISH
# =====================================================================
def publish(site_dir: str, msg: str):
    if not os.path.isdir(os.path.join(site_dir, ".git")):
        print(f"  ⚠  {site_dir} non è un repo git — salto il push. Vedi SETUP_dashboard.md.")
        return
    def git(*a):
        return subprocess.run(["git", "-C", site_dir, *a], capture_output=True, text=True)
    git("add", "-A")
    if not git("status", "--porcelain").stdout.strip():
        print("  ↷ nessuna modifica da pubblicare")
        return
    c = git("commit", "-m", msg)
    if c.returncode != 0:
        print("  ⚠ commit fallito:", c.stderr.strip() or c.stdout.strip());  return
    p = git("push")
    if p.returncode != 0:
        print("  ⚠ push fallito:", p.stderr.strip() or p.stdout.strip());  return
    print("  ✓ pubblicato su GitHub Pages (1–2 min per vederlo online)")


# =====================================================================
# MAIN
# =====================================================================
def main():
    ap = argparse.ArgumentParser(description="Genera la dashboard statica del Live Project")
    ap.add_argument("--publish", action="store_true", help="git commit + push dopo la build")
    ap.add_argument("--open", action="store_true", help="apri l'HTML nel browser")
    ap.add_argument("--demo", action="store_true", help="usa dati finti (test del layout)")
    ap.add_argument("--out", default=None, help="cartella di output (default: SITE_DIR)")
    ap.add_argument("--data-root", default=None, help="usa un'altra radice dati")
    args = ap.parse_args()

    global PROCESSED_PATH, SERIES_PATH, FORECAST_PATH, MASTER_XLSX
    if args.data_root:
        PROCESSED_PATH = os.path.join(args.data_root, "processed")
        SERIES_PATH    = os.path.join(args.data_root, "official_series")
        FORECAST_PATH  = os.path.join(args.data_root, "forecasts")
        MASTER_XLSX    = os.path.join(FORECAST_PATH, "forecasts_all.xlsx")

    site_dir = args.out or SITE_DIR
    os.makedirs(site_dir, exist_ok=True)
    print(f"── Costruzione dashboard → {site_dir}")

    if args.demo:
        m, series, reddit = make_demo()
        master = load_master_from_df(m)
    else:
        master = load_master(MASTER_XLSX)
        series = load_official_series(SERIES_PATH)
        reddit = load_reddit_mentions(PROCESSED_PATH)

    print(f"  · {len(master)} righe di forecast, {len(series)} serie ufficiali, "
          f"{0 if reddit is None or reddit.empty else reddit['subreddit'].nunique()} subreddit")

    payload = build_payload(master, series, reddit)
    html = render_html(payload, ensure_plotly(site_dir))

    index_path = os.path.join(site_dir, "index.html")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(html)
    with open(os.path.join(site_dir, "data.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    cols = [c for c in ["run_date", "provider", "model", "variable", "forecast_type",
                        "target", "point_forecast", "last_official_period",
                        "last_official_value", "used_web_search"] if c in master.columns]
    master[cols].to_csv(os.path.join(site_dir, "forecasts.csv"), index=False)
    open(os.path.join(site_dir, ".nojekyll"), "a").close()

    print(f"  ✓ index.html ({os.path.getsize(index_path)/1024:.0f} KB), data.json, forecasts.csv")
    if args.publish:
        publish(site_dir, f"dashboard update {payload['last_run']}")
    if args.open:
        webbrowser.open("file://" + os.path.abspath(index_path))


if __name__ == "__main__":
    main()
