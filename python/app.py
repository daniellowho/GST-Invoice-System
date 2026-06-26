"""
GSTR 2B Reconciliation — Desktop App (pywebview host)
======================================================
Opens a native desktop window showing ui.html. All file processing is done
by engine.py (Python). The JavaScript UI only collects files and displays
results — every reconciliation calculation happens in Python.

The Api class below is exposed to JavaScript as `window.pywebview.api`.
"""

import base64
import json
import os
import sys
import threading
from datetime import datetime
from pathlib import Path

import webview

import engine

APP_NAME = "GSTR 2B Reconciliation"

# Small persistent state file — independent of the chosen output folder, so it
# survives even if the user later changes "Change location...". Used purely
# for reminders (last carry-forward upload, last month's pending payments).
_STATE_DIR = Path.home() / ".gstr2b_recon"
_STATE_FILE = _STATE_DIR / "state.json"


def _load_app_state():
    try:
        return json.loads(_STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_app_state(state):
    try:
        _STATE_DIR.mkdir(parents=True, exist_ok=True)
        _STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")
    except Exception:
        pass  # reminders are best-effort; never let this break the app


def _resource_path(rel):
    """
    Resolve a bundled resource path.

    - Frozen (.exe built by PyInstaller): files added via --add-data live at
      sys._MEIPASS, so use that.
    - Running from source: app.py lives in python\\, but ui.html sits one level
      up in the project root. Check the script folder first, then its parent.
    """
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, rel)
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, rel),                       # same folder
        os.path.join(os.path.dirname(here), rel),      # parent folder (project root)
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return candidates[0]  # fall back to the first; error surfaces clearly if missing


def _default_output_dir():
    """Default place to save outputs: ~/Documents/GSTR2B Outputs (created on demand)."""
    docs = Path.home() / "Documents"
    base = docs if docs.exists() else Path.home()
    out = base / "GSTR2B Outputs"
    return str(out)


class Api:
    """Methods here are callable from JS via window.pywebview.api.<name>(...)."""

    def __init__(self):
        # In-memory working state for the current month
        self._reset_state()
        self._output_dir = _default_output_dir()
        self._app_state = _load_app_state()

    def _reset_state(self):
        self.df_2b_latest = None
        self.df_tally_latest = None
        self.df_tally_gt = None
        self.df_2b_pending = None
        self.df_tally_pending = None
        self.stage3 = None        # dict of dataframes + stats
        self.raw = {"raw2b": None, "rawtally": None, "cf2b": None, "cftally": None, "excl": None}

    # ── Helpers ───────────────────────────────────────────────────────────────
    @staticmethod
    def _decode(data_url_or_b64):
        """Accept either a bare base64 string or a data: URL and return bytes."""
        if not data_url_or_b64:
            return None
        s = data_url_or_b64
        if s.startswith("data:"):
            s = s.split(",", 1)[1]
        return base64.b64decode(s)

    # ── Output folder management ──────────────────────────────────────────────
    def get_output_dir(self):
        return self._output_dir

    def choose_output_dir(self):
        """Open a native folder picker; return the chosen path (or current if cancelled)."""
        win = webview.windows[0]
        result = win.create_file_dialog(webview.FOLDER_DIALOG)
        if result and len(result) > 0:
            self._output_dir = result[0]
        return self._output_dir

    # ── File intake (called when user drops/picks a file) ─────────────────────
    def load_raw_2b(self, b64, filename=None):
        try:
            data = self._decode(b64)
            self.raw["raw2b"] = data
            count = engine.pd.read_excel(engine.io.BytesIO(data), header=None).shape[0]
            period = None
            try:
                df_tmp, _ = engine.clean_2b(data)
                period = engine.detect_2b_period(df_tmp)
            except Exception:
                period = None  # detection is best-effort; cleaning errors surface in Stage 1
            return {"ok": True, "rows": int(count), "period": period}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def load_raw_tally(self, b64, filename=None):
        try:
            data = self._decode(b64)
            self.raw["rawtally"] = data
            count = engine.pd.read_excel(engine.io.BytesIO(data), header=None).shape[0]
            period = None
            try:
                df_tmp, _, _ = engine.clean_tally(data)
                period = engine.detect_tally_period(df_tmp)
            except Exception:
                period = None
            return {"ok": True, "rows": int(count), "period": period}
        except Exception as e:
            return {"ok": False, "error": str(e), "hint": "Please use the Tally template shown in the app: row 1 should be the headers, row 2 should be example values, and the fields should stay in the same order."}

    def load_cf_2b(self, b64, filename=None):
        try:
            self.raw["cf2b"] = self._decode(b64)
            self._app_state["last_cf_2b"] = {"when": datetime.utcnow().isoformat(), "filename": filename}
            _save_app_state(self._app_state)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def load_cf_tally(self, b64, filename=None):
        try:
            self.raw["cftally"] = self._decode(b64)
            self._app_state["last_cf_tally"] = {"when": datetime.utcnow().isoformat(), "filename": filename}
            _save_app_state(self._app_state)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def get_tally_template(self):
        try:
            data = engine.build_tally_template_bytes()
            return {"ok": True, "filename": "Tally Template.xlsx", "b64": base64.b64encode(data).decode("ascii")}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def load_exclusion(self, b64):
        try:
            data = self._decode(b64)
            self.raw["excl"] = data
            gstins = engine.load_exclusion(data)
            return {"ok": True, "count": len(gstins)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def clear_file(self, which):
        if which in self.raw:
            self.raw[which] = None
        return {"ok": True}

    # ── Stage runners ─────────────────────────────────────────────────────────
    def run_stage1(self):
        try:
            if not self.raw["raw2b"] or not self.raw["rawtally"]:
                return {"ok": False, "error": "Upload both raw files first."}
            df2b, s2b = engine.clean_2b(self.raw["raw2b"])
            dft, dft_gt, st = engine.clean_tally(self.raw["rawtally"])
            self.df_2b_latest = df2b
            self.df_tally_latest = dft
            self.df_tally_gt = dft_gt
            log = (f"2B Latest: read {s2b['read']} \u2192 dropped {s2b['dropped']} "
                   f"(blank GSTIN) \u2192 written {s2b['written']}\n"
                   f"Tally Latest: read {st['read']} \u2192 dropped {st['dropped']} "
                   f"(grand total/blank) \u2192 written {st['written']}")
            return {
                "ok": True, "log": log,
                "preview2b": engine.preview_records(df2b, 8),
                "cols2b": list(df2b.columns),
                "previewTally": engine.preview_records(dft, 8),
                "colsTally": list(dft.columns),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def run_stage2(self):
        try:
            if self.df_2b_latest is None:
                return {"ok": False, "error": "Run Stage 1 first."}
            m2b, i2 = engine.merge_2b(self.df_2b_latest, self.raw["cf2b"])
            mt, it = engine.merge_tally(self.df_tally_latest, self.raw["cftally"])
            self.df_2b_pending = m2b
            self.df_tally_pending = mt

            def desc(info, label):
                if info["merged"]:
                    return f"{label}: {info['rows']} rows (merged {info['cf_rows']} CF + {info['latest_rows']} latest)"
                return f"{label}: {info['rows']} rows (latest only, no CF)"

            log = desc(i2, "2B All Pending") + "\n" + desc(it, "Tally All Pending")
            return {
                "ok": True, "log": log,
                "preview2b": engine.preview_records(m2b, 8),
                "cols2b": list(m2b.columns),
                "previewTally": engine.preview_records(mt, 8),
                "colsTally": list(mt.columns),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def run_stage3(self):
        try:
            if self.df_2b_pending is None:
                return {"ok": False, "error": "Run Stage 2 first."}
            excl = engine.load_exclusion(self.raw["excl"]) if self.raw["excl"] else set()
            recon, ntbc, cf2b, cft, stats = engine.reconcile(
                self.df_2b_pending, self.df_tally_pending, excl)
            self.stage3 = {"recon": recon, "ntbc": ntbc, "cf2b": cf2b, "cft": cft, "stats": stats}
            return {
                "ok": True, "stats": stats,
                "previewRecon": engine.preview_records(recon, 8),
                "colsRecon": list(recon.columns),
                "previewNtbc": engine.preview_records(ntbc, 8),
                "colsNtbc": list(ntbc.columns),
                "previewCf2b": engine.preview_records(cf2b, 8),
                "colsCf2b": list(cf2b.columns),
                "previewCft": engine.preview_records(cft, 8),
                "colsCft": list(cft.columns),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── Single-button pipeline: Stage 1 → 2 → 3 in one call ───────────────────
    def run_all(self):
        """Run the whole reconciliation end to end and return a step log + results."""
        try:
            if not self.raw["raw2b"] or not self.raw["rawtally"]:
                return {"ok": False, "error": "Upload both raw files first."}

            steps = []

            # Stage 1 — clean
            df2b, s2b = engine.clean_2b(self.raw["raw2b"])
            dft, dft_gt, st = engine.clean_tally(self.raw["rawtally"])
            self.df_2b_latest, self.df_tally_latest, self.df_tally_gt = df2b, dft, dft_gt
            steps.append(f"Cleaned 2B: {s2b['written']} rows kept "
                         f"({s2b['dropped']} dropped). Tally: {st['written']} rows kept "
                         f"({st['dropped']} grand-total/blank dropped).")

            # Stage 2 — merge
            m2b, i2 = engine.merge_2b(df2b, self.raw["cf2b"])
            mt, it = engine.merge_tally(dft, self.raw["cftally"])
            self.df_2b_pending, self.df_tally_pending = m2b, mt
            steps.append(f"Prepared pending lists: 2B {i2['rows']} rows, "
                         f"Tally {it['rows']} rows"
                         + (" (carry-forward merged)." if (i2['merged'] or it['merged']) else "."))

            # Stage 3 — reconcile
            excl = engine.load_exclusion(self.raw["excl"]) if self.raw["excl"] else set()
            recon, ntbc, cf2b, cft, stats = engine.reconcile(m2b, mt, excl)
            self.stage3 = {"recon": recon, "ntbc": ntbc, "cf2b": cf2b, "cft": cft, "stats": stats}
            steps.append(f"Reconciled: {stats['reconciled']} matched, "
                         f"{stats['cf2b']} + {stats['cfTally']} carried forward, "
                         f"{stats['ntbc']} not to be claimed.")

            return {
                "ok": True, "stats": stats, "steps": steps,
                "tabs": self._full_tabs(),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _full_tabs(self, limit=500):
        """Return the three output datasets (capped rows) for the results view."""
        s = self.stage3
        return {
            "summary_reconciled": {
                "title": "Reconciled entries",
                "rows": engine.preview_records(s["recon"], limit),
                "cols": list(s["recon"].columns),
                "total": len(s["recon"]),
            },
            "summary_ntbc": {
                "title": "Not to be claimed",
                "rows": engine.preview_records(s["ntbc"], limit),
                "cols": list(s["ntbc"].columns),
                "total": len(s["ntbc"]),
            },
            "cf_2b": {
                "title": "2B carry forward",
                "rows": engine.preview_records(s["cf2b"], limit),
                "cols": list(s["cf2b"].columns),
                "total": len(s["cf2b"]),
            },
            "cf_tally": {
                "title": "Tally carry forward",
                "rows": engine.preview_records(s["cft"], limit),
                "cols": list(s["cft"].columns),
                "total": len(s["cft"]),
            },
        }

    def get_results(self):
        """Return current results for the Results tab (or null if not run)."""
        if self.stage3 is None:
            return {"ok": False}
        return {"ok": True, "stats": self.stage3["stats"], "tabs": self._full_tabs()}

    # ── Reminders (Stage 2 popup + pending-payments banner) ────────────────────
    def get_reminders(self):
        def info(key):
            d = self._app_state.get(key)
            if not d:
                return None
            days_ago = None
            try:
                days_ago = (datetime.utcnow() - datetime.fromisoformat(d["when"])).days
            except Exception:
                pass
            return {"when": d.get("when"), "filename": d.get("filename"), "days_ago": days_ago}

        pending = self._app_state.get("last_run")
        pending_out = None
        if pending and pending.get("cf2b_count", 0) > 0:
            days_ago = None
            try:
                days_ago = (datetime.utcnow() - datetime.fromisoformat(pending["when"])).days
            except Exception:
                pass
            pending_out = {**pending, "days_ago": days_ago}

        return {
            "cf_last": {"raw2b": info("last_cf_2b"), "rawtally": info("last_cf_tally")},
            "pending": pending_out,
        }

    # ── Save outputs to disk ───────────────────────────────────────────────────
    def save_outputs(self, month):
        """
        Write output files into a tidy month folder:

            <output_dir>/<month>/
                Summary/
                    GSTR 2B Summary <month>.xlsx
                Carry Forward/
                    2B Carry Forward <month>.xlsx
                    Tally Carry Forward <month>.xlsx
        """
        try:
            if self.stage3 is None:
                return {"ok": False, "error": "Run the process before saving."}
            month = (month or "output").strip() or "output"
            month_dir   = Path(self._output_dir) / month
            summary_dir = month_dir / "Summary"
            cf_dir      = month_dir / "Carry Forward"
            summary_dir.mkdir(parents=True, exist_ok=True)
            cf_dir.mkdir(parents=True, exist_ok=True)

            s = self.stage3
            written = []

            def write(folder, fname, data):
                p = folder / fname
                with open(p, "wb") as f:
                    f.write(data)
                written.append(str(p))

            write(summary_dir, f"GSTR 2B Summary {month}.xlsx",
                  engine.build_summary_bytes(s["recon"], s["ntbc"]))
            write(cf_dir, f"2B Carry Forward {month}.xlsx",
                  engine.build_2b_cf_bytes(s["cf2b"]))
            write(cf_dir, f"Tally Carry Forward {month}.xlsx",
                  engine.build_tally_cf_bytes(s["cft"]))

            try:
                cf2b_amount = float(engine.pd.to_numeric(
                    s["cf2b"]["Invoice Value(\u20b9)"], errors="coerce").fillna(0).sum()) if len(s["cf2b"]) else 0.0
            except Exception:
                cf2b_amount = 0.0
            self._app_state["last_run"] = {
                "month": month,
                "when": datetime.utcnow().isoformat(),
                "cf2b_count": len(s["cf2b"]),
                "cf_tally_count": len(s["cft"]),
                "ntbc_count": len(s["ntbc"]),
                "cf2b_amount": cf2b_amount,
            }
            _save_app_state(self._app_state)

            return {"ok": True, "dir": str(month_dir), "files": written}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_output_dir(self, month=None):
        """Open the output folder in the system file explorer."""
        try:
            target = Path(self._output_dir)
            if month:
                cand = target / month
                if cand.exists():
                    target = cand
            target.mkdir(parents=True, exist_ok=True)
            p = str(target)
            if sys.platform.startswith("win"):
                os.startfile(p)  # noqa: S606  (Windows only)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.call(["open", p])
            else:
                import subprocess
                subprocess.call(["xdg-open", p])
            return {"ok": True, "dir": p}
        except Exception as e:
            return {"ok": False, "error": str(e)}


def main():
    api = Api()
    html_path = _resource_path("ui.html")
    window = webview.create_window(
        APP_NAME,
        url=html_path,
        js_api=api,
        width=1180,
        height=820,
        min_size=(900, 640),
    )
    webview.start()


if __name__ == "__main__":
    main()
