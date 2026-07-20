"""
generate_audio_ppt.py
Reads a PowerPoint, extracts slide notes, synthesizes speech via AWS Polly,
and embeds the audio onto each slide for Storyline 360 import.

Run with no arguments to open the GUI. Pass CLI arguments for headless use.
"""

import argparse
import json
import logging
import os
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import boto3
from botocore.config import Config as BotoConfig
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

log = logging.getLogger("polly-tts")

# ── Voice presets ────────────────────────────────────────────────────────────
VOICE_PRESETS = {
    "mandarin": {"voice": "Zhiyu", "lang": "cmn-CN"},
    "english": {"voice": "Matthew", "lang": "en-US"},
}

PRESET_LABELS = {
    "mandarin": "Mandarin (Zhiyu)",
    "english": "English (Matthew)",
}

CONFIG_FILE = "config.json"


# ── Config helpers ───────────────────────────────────────────────────────────
def get_config_path() -> Path:
    """Return the path to config.json next to the executable or script."""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    return base / CONFIG_FILE


def load_config() -> dict:
    """Load config.json if it exists."""
    cfg_path = get_config_path()
    if cfg_path.exists():
        with open(cfg_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(config: dict) -> None:
    """Write config back to config.json."""
    cfg_path = get_config_path()
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)


_EMPTY_PRONUNCIATIONS = {k: [] for k in VOICE_PRESETS}


def load_pronunciations(config: dict | None = None) -> dict:
    """Return per-voice pronunciation correction lists from config.

    Returns a deep copy so callers can mutate safely without affecting the
    module-level default or the on-disk config.
    """
    if config is None:
        config = load_config()
    src = config.get("pronunciations", _EMPTY_PRONUNCIATIONS)
    return {k: [dict(item) for item in v] for k, v in src.items()}


def save_pronunciations(pronunciations: dict, config: dict | None = None) -> None:
    """Persist pronunciation corrections into config.json."""
    if config is None:
        config = load_config()
    config["pronunciations"] = pronunciations
    save_config(config)


def create_polly_client(config: dict = None):
    """Create a boto3 Polly client, optionally using credentials from config."""
    kwargs = {
        "config": BotoConfig(retries={"max_attempts": 3, "mode": "adaptive"}),
    }
    if config:
        aws = config.get("aws", {})
        key_id = aws.get("access_key_id", "")
        secret = aws.get("secret_access_key", "")
        if key_id and secret and not key_id.startswith("YOUR_"):
            kwargs["aws_access_key_id"] = key_id
            kwargs["aws_secret_access_key"] = secret
            if aws.get("region"):
                kwargs["region_name"] = aws["region"]
    return boto3.client("polly", **kwargs)


# ── Core logic ───────────────────────────────────────────────────────────────
def parse_slide_range(spec: str, total: int) -> list[int]:
    """Parse '1,3,5-8' into sorted 0-indexed slide indices."""
    indices: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            lo, hi = part.split("-", 1)
            lo, hi = int(lo), int(hi)
            if lo < 1 or hi > total or lo > hi:
                raise ValueError(f"Invalid range {part} for presentation with {total} slides")
            indices.update(range(lo - 1, hi))
        else:
            n = int(part)
            if n < 1 or n > total:
                raise ValueError(f"Slide {n} out of range (1-{total})")
            indices.add(n - 1)
    return sorted(indices)


def detect_language(filename: str) -> str:
    """Heuristically detect voice preset from the filename.

    Returns the preset key (``"mandarin"`` or ``"english"``) based on the
    ratio of CJK characters to total non-whitespace characters.  If the
    filename is predominantly Chinese, ``"mandarin"`` is returned; otherwise
    ``"english"``.
    """
    import unicodedata
    text = Path(filename).stem
    cjk = total = 0
    for ch in text:
        if ch.isspace():
            continue
        total += 1
        if unicodedata.category(ch).startswith(("Lo",)):
            # "Lo" = Letter, other — covers CJK ideographs, kana, hangul, etc.
            cjk += 1
    if total == 0:
        return "mandarin"
    return "mandarin" if cjk / total > 0.3 else "english"


def apply_corrections(text: str, corrections: list[dict]) -> str:
    """Apply pronunciation corrections to *text* using longest-match-first.

    *corrections* is a list of dicts with keys ``original`` and ``replacement``.
    Overlapping matches are resolved by preferring the longest ``original`` string.
    """
    if not corrections:
        return text
    # Sort longest-original-first so greediest match wins
    active = sorted(corrections, key=lambda c: len(c["original"]), reverse=True)
    for entry in active:
        text = text.replace(entry["original"], entry["replacement"])
    return text


def synthesize_audio(polly_client, text: str, output_path: str,
                     voice_id: str, engine: str, lang_code: str) -> bool:
    """Call AWS Polly and save the result as MP3. Returns True on success."""
    response = polly_client.synthesize_speech(
        Engine=engine,
        OutputFormat="mp3",
        Text=text,
        LanguageCode=lang_code,
        VoiceId=voice_id,
    )
    if "AudioStream" in response:
        with open(output_path, "wb") as f:
            f.write(response["AudioStream"].read())
        return True
    return False


def embed_audio_on_slide(slide, audio_path: str) -> None:
    """Embed an MP3 onto a slide using add_movie() then patch timing for autoplay."""
    from lxml import etree

    slide.shapes.add_movie(
        audio_path,
        left=Inches(-2.0),
        top=Inches(-2.0),
        width=Inches(1.0),
        height=Inches(1.0),
        poster_frame_image=None,
        mime_type="audio/mpeg",
    )

    timing = slide._element.find(qn("p:timing"))
    if timing is not None:
        for cond in timing.iter(qn("p:cond")):
            if cond.get("delay") == "indefinite":
                cond.set("delay", "0")


def normalize_audio_for_storyline(output_path: Path) -> bool:
    """Open in PowerPoint via COM, nudge each shape to trigger media normalization."""
    abs_path = str(output_path.resolve())
    log.info("Normalizing audio for Storyline via PowerPoint COM ...")
    try:
        import time

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.Presentations.Open(abs_path, ReadOnly=False)

        for slide_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(slide_idx)
            slide.Select()
            for shape_idx in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(shape_idx)
                try:
                    mt = shape.MediaType
                except Exception:
                    mt = 0
                if mt == 0:
                    continue
                orig_left = shape.Left
                shape.Left = orig_left + 1
                time.sleep(0.05)
                shape.Left = orig_left
                log.debug("Nudged shape '%s' on slide %d", shape.Name, slide_idx)

        pres.Save()
        pres.Close()
        ppt.Quit()
        pythoncom.CoUninitialize()
        log.info("PowerPoint normalization complete")
        return True
    except Exception as e:
        log.warning("PowerPoint COM normalization failed: %s", e)
        try:
            pres.Close()
        except Exception:
            pass
        try:
            ppt.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return False


def process_pptx(input_path: Path, output_path: Path, voice_id: str, engine: str,
                 lang_code: str, slides_spec: str | None = None,
                 callback=None, config: dict = None,
                 pronunciations: list[dict] | None = None) -> dict:
    """
    Main processing pipeline. Returns a summary dict.
    callback(message) is called with status updates.
    *pronunciations* is the correction list for the active voice; if provided
    a *copy* of each slide's notes is corrected before being sent to Polly.
    """
    def emit(msg):
        log.info(msg)
        if callback:
            callback(msg)

    prs = Presentation(str(input_path))
    total = len(prs.slides)
    emit(f"Loaded '{input_path.name}' — {total} slide(s)")

    if slides_spec:
        slide_indices = parse_slide_range(slides_spec, total)
    else:
        slide_indices = list(range(total))

    emit(f"Processing {len(slide_indices)} slide(s): "
         + ", ".join(str(i + 1) for i in slide_indices))

    config = config or load_config()
    polly_client = create_polly_client(config)
    emit("AWS Polly client ready")

    success_count = skip_count = fail_count = 0

    with tempfile.TemporaryDirectory(prefix="polly_tts_") as tmp_dir:
        for i, idx in enumerate(slide_indices):
            slide = prs.slides[idx]
            slide_num = idx + 1

            notes_text = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes_text = slide.notes_slide.notes_text_frame.text.strip()

            if not notes_text:
                emit(f"Slide {slide_num}: no notes — skipping")
                skip_count += 1
                continue

            tts_text = apply_corrections(notes_text, pronunciations or [])

            emit(f"Slide {slide_num}: synthesizing {len(tts_text)} chars ... ({i + 1}/{len(slide_indices)})")
            audio_path = os.path.join(tmp_dir, f"slide_{slide_num}.mp3")

            try:
                ok = synthesize_audio(polly_client, tts_text, audio_path,
                                      voice_id, engine, lang_code)
                if not ok:
                    emit(f"Slide {slide_num}: Polly returned no audio stream")
                    fail_count += 1
                    continue
            except Exception as e:
                emit(f"Slide {slide_num}: Polly API error: {e}")
                fail_count += 1
                continue

            embed_audio_on_slide(slide, audio_path)
            emit(f"Slide {slide_num}: audio embedded")
            success_count += 1

    prs.save(str(output_path))
    emit(f"Saved: {output_path}")

    normalize_audio_for_storyline(output_path)

    emit(f"Done — {success_count} succeeded, {skip_count} skipped (no notes), {fail_count} failed")
    return {"success": success_count, "skipped": skip_count, "failed": fail_count}


# ── CLI ──────────────────────────────────────────────────────────────────────
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Add AWS Polly TTS narration to PowerPoint slides.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "voice presets:\n"
            "  mandarin   -> Zhiyu     / cmn-CN\n"
            "  english    -> Matthew   / en-US\n"
            "\nRun with no arguments to open the GUI."
        ),
    )
    p.add_argument("input", nargs="?", help="Path to the input .pptx file")
    p.add_argument("-o", "--output", help="Output file path (default: narrated_<input>)")
    p.add_argument("--voice", default="Zhiyu", help="Polly voice ID (default: Zhiyu)")
    p.add_argument("--engine", choices=["standard", "neural"], default="neural",
                   help="TTS engine (default: neural)")
    p.add_argument("--lang", default="cmn-CN", help="Language code (default: cmn-CN)")
    p.add_argument("--preset", choices=list(VOICE_PRESETS), default=None,
                   help="Use a named voice preset instead of --voice/--lang")
    p.add_argument("--slides", default=None,
                   help="Slide numbers or ranges, e.g. '1,3,5-8' (default: all)")
    p.add_argument("--dry-run", action="store_true",
                   help="Show extracted notes without calling AWS")
    p.add_argument("--keep-audio", action="store_true",
                   help="Keep temp MP3 files after processing")
    p.add_argument("--gui", action="store_true", help="Force open the GUI")
    p.add_argument("-v", "--verbose", action="store_true", help="Enable debug logging")
    return p


def cli_main(args=None):
    parsed = build_parser().parse_args(args)

    level = logging.DEBUG if parsed.verbose else logging.INFO
    logging.basicConfig(level=level, format="[%(levelname)s] %(message)s")

    if not parsed.input and not parsed.gui:
        gui_main()
        return

    voice_id = parsed.voice
    lang_code = parsed.lang
    if parsed.preset:
        preset = VOICE_PRESETS[parsed.preset]
        voice_id = preset["voice"]
        lang_code = preset["lang"]

    input_path = Path(parsed.input)
    if not input_path.exists():
        log.error("File not found: %s", input_path)
        sys.exit(1)

    if parsed.output:
        output_path = Path(parsed.output)
    else:
        stem = input_path.stem
        suffix = input_path.suffix
        output_path = input_path.with_name(f"narrated_{stem}{suffix}")
        i = 1
        while output_path.exists():
            output_path = input_path.with_name(f"narrated_{stem}_{i}{suffix}")
            i += 1

    if parsed.dry_run:
        prs = Presentation(str(input_path))
        total = len(prs.slides)
        indices = parse_slide_range(parsed.slides, total) if parsed.slides else list(range(total))
        for idx in indices:
            slide = prs.slides[idx]
            notes = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                preview = notes[:200].replace("\n", " ")
                log.info("Slide %d — %d chars: %s%s", idx + 1, len(notes),
                         preview, "..." if len(notes) > 200 else "")
            else:
                log.info("Slide %d — (no notes)", idx + 1)
        log.info("Dry run complete. No audio generated.")
        return

    process_pptx(input_path, output_path, voice_id, parsed.engine, lang_code,
                 slides_spec=parsed.slides)


# ── Pronunciation Corrections Dialog ──────────────────────────────────────────
class CorrectionsDialog:
    """Separate window for managing per-voice pronunciation corrections."""

    def __init__(self, parent: tk.Tk, config: dict):
        self.config = config
        self.pronunciations = load_pronunciations(config)
        self.active_voice = tk.StringVar(value="mandarin")

        self.win = tk.Toplevel(parent)
        self.win.title("Pronunciation Corrections")
        self.win.geometry("560x480")
        self.win.resizable(False, False)
        self.win.transient(parent)
        self.win.grab_set()

        self._rows: list[dict] = []
        self._build_ui()
        self._refresh_rows()

        self.active_voice.trace_add("write", lambda *_: self._refresh_rows())

    # ── UI construction ──────────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # Voice selector
        top = ttk.Frame(self.win)
        top.pack(fill="x", **pad)
        ttk.Label(top, text="Voice:").pack(side="left", padx=(0, 6))
        voices = [f"{k} — {PRESET_LABELS[k]}" for k in VOICE_PRESETS]
        cb = ttk.Combobox(top, textvariable=self.active_voice, state="readonly",
                          width=30, values=voices)
        cb.pack(side="left")
        cb.current(0)

        # Scrollable list area
        list_frame = ttk.LabelFrame(self.win, text="Corrections", padding=4)
        list_frame.pack(fill="both", expand=True, **pad)

        self._canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical",
                                  command=self._canvas.yview)
        self._inner = ttk.Frame(self._canvas)

        self._inner.bind("<Configure>",
                         lambda _: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        # Column headers
        hdr = ttk.Frame(self._inner)
        hdr.pack(fill="x", pady=(0, 2))
        ttk.Label(hdr, text="Original", width=18, anchor="w").pack(side="left", padx=(4, 2))
        ttk.Label(hdr, text="", width=3).pack(side="left")
        ttk.Label(hdr, text="Replacement", width=18, anchor="w").pack(side="left", padx=(2, 2))

        # Add row
        add_frame = ttk.Frame(self.win)
        add_frame.pack(fill="x", **pad)

        self._add_orig = tk.StringVar()
        self._add_repl = tk.StringVar()
        ttk.Label(add_frame, text="New:").pack(side="left", padx=(0, 4))
        ttk.Entry(add_frame, textvariable=self._add_orig, width=16).pack(
            side="left", padx=(0, 4))
        ttk.Label(add_frame, text="→").pack(side="left", padx=(0, 4))
        ttk.Entry(add_frame, textvariable=self._add_repl, width=16).pack(
            side="left", padx=(0, 8))
        ttk.Button(add_frame, text="Add", command=self._on_add).pack(side="left")

        # Bottom buttons
        btn_frame = ttk.Frame(self.win)
        btn_frame.pack(fill="x", **pad)
        ttk.Button(btn_frame, text="Preview", command=self._on_preview).pack(side="left")
        ttk.Button(btn_frame, text="Save & Close", command=self._on_save_close).pack(
            side="right")

    # ── Row management ───────────────────────────────────────────────────
    def _current_list(self) -> list[dict]:
        key = self.active_voice.get().split(" — ")[0]
        return self.pronunciations.setdefault(key, [])

    def _refresh_rows(self):
        for entry in self._rows:
            entry["frame"].destroy()
        self._rows.clear()

        for i, item in enumerate(self._current_list()):
            self._add_row(i, item)
        self._canvas.yview_moveto(0)

    def _add_row(self, idx: int, item: dict):
        frame = ttk.Frame(self._inner)
        frame.pack(fill="x", pady=1)

        orig_var = tk.StringVar(value=item["original"])
        repl_var = tk.StringVar(value=item["replacement"])

        orig_entry = ttk.Entry(frame, textvariable=orig_var, width=18)
        orig_entry.pack(side="left", padx=(4, 2))
        ttk.Label(frame, text="→").pack(side="left", padx=(0, 2))
        repl_entry = ttk.Entry(frame, textvariable=repl_var, width=18)
        repl_entry.pack(side="left", padx=(2, 2))

        del_btn = ttk.Button(frame, text="✕", width=3,
                             command=lambda: self._delete(idx))
        del_btn.pack(side="left", padx=(8, 4))

        entry = {
            "frame": frame,
            "orig_var": orig_var,
            "repl_var": repl_var,
            "orig_entry": orig_entry,
            "repl_entry": repl_entry,
        }
        self._rows.append(entry)

        orig_var.trace_add("write", lambda *_a, i=idx, v=orig_var: self._sync_item(i, "original", v.get()))
        repl_var.trace_add("write", lambda *_a, i=idx, v=repl_var: self._sync_item(i, "replacement", v.get()))

    def _sync_item(self, idx: int, field: str, value: str):
        lst = self._current_list()
        if idx < len(lst):
            lst[idx][field] = value

    def _delete(self, idx: int):
        lst = self._current_list()
        if idx < len(lst):
            del lst[idx]
        self._refresh_rows()

    def _on_add(self):
        orig = self._add_orig.get().strip()
        repl = self._add_repl.get().strip()
        if not orig or not repl:
            messagebox.showwarning("Missing values", "Both fields are required.",
                                   parent=self.win)
            return
        self._current_list().append({
            "original": orig,
            "replacement": repl,
        })
        self._add_orig.set("")
        self._add_repl.set("")
        self._refresh_rows()

    # ── Preview ──────────────────────────────────────────────────────────
    def _on_preview(self):
        input_file = self._get_input_file()
        if not input_file:
            return

        key = self.active_voice.get().split(" — ")[0]
        corrections = self.pronunciations.get(key, [])

        try:
            prs = Presentation(str(input_file))
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open PPTX:\n{e}", parent=self.win)
            return

        preview_win = tk.Toplevel(self.win)
        preview_win.title(f"Preview — {key}")
        preview_win.geometry("640x400")
        preview_win.transient(self.win)

        txt = tk.Text(preview_win, wrap="word", font=("Consolas", 9), state="normal")
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.tag_configure("highlight", background="#ffff80", foreground="#000000")
        txt.tag_configure("slide_hdr", font=("Consolas", 9, "bold"))

        total_replacements = 0
        PREVIEW_LEN = 60

        for idx, slide in enumerate(prs.slides):
            notes = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()
            if not notes:
                continue

            corrected = apply_corrections(notes, corrections)
            if corrected == notes:
                continue

            # Count replacements
            active = sorted(corrections, key=lambda c: len(c["original"]), reverse=True)
            slide_count = 0
            tmp = notes
            for entry in active:
                before = tmp
                tmp = tmp.replace(entry["original"], entry["replacement"])
                slide_count += before.count(entry["original"])
            total_replacements += slide_count

            hdr = f"── Slide {idx + 1} ({slide_count} replacement{'s' if slide_count != 1 else ''}) ──\n"
            txt.insert("end", hdr, "slide_hdr")

            # Before line
            before_preview = notes[:PREVIEW_LEN]
            if len(notes) > PREVIEW_LEN:
                before_preview += "..."
            txt.insert("end", f"  Before: {before_preview}\n")

            # After line with highlighted replacements
            txt.insert("end", "  After:  ")
            corr_trunc = corrected[:PREVIEW_LEN]
            spans = self._compute_highlight_spans(notes, corrected, active)
            pos = 0
            for span_start, span_end in spans:
                if span_start > pos:
                    txt.insert("end", corr_trunc[pos:span_start])
                txt.insert("end", corr_trunc[span_start:span_end], "highlight")
                pos = span_end
            if pos < len(corr_trunc):
                txt.insert("end", corr_trunc[pos:])
            if len(corrected) > PREVIEW_LEN:
                txt.insert("end", "...")
            txt.insert("end", "\n\n")

        if total_replacements == 0:
            txt.insert("end", "No replacements to preview — no notes match any correction.\n")

        preview_win.title(f"Preview — {key} ({total_replacements} replacement"
                          f"{'s' if total_replacements != 1 else ''})")
        txt.configure(state="disabled")

    @staticmethod
    def _compute_highlight_spans(original: str, corrected: str,
                                 corrections: list[dict]) -> list[tuple[int, int]]:
        """Return (start, end) spans in *corrected* that were changed by replacements.

        Walks a working copy of *original*, applying each replacement
        sequentially.  Positions found via ``str.find`` on the mutated copy
        are already in corrected-text coordinates, so no offset tracking
        is needed.
        """
        spans: list[tuple[int, int]] = []
        text = original
        for entry in corrections:
            orig_str = entry["original"]
            repl_str = entry["replacement"]
            i = 0
            while i < len(text):
                idx = text.find(orig_str, i)
                if idx == -1:
                    break
                spans.append((idx, idx + len(repl_str)))
                text = text[:idx] + repl_str + text[idx + len(orig_str):]
                i = idx + len(repl_str)
        spans.sort()
        return spans

    def _get_input_file(self) -> str | None:
        app = getattr(self.win.master, "_ppttts_app", None)
        if app:
            val = app.input_path.get().strip()
            if val and Path(val).exists():
                return val
        messagebox.showwarning("No file", "Select a PowerPoint file first.",
                               parent=self.win)
        return None

    # ── Save & close ─────────────────────────────────────────────────────
    def _on_save_close(self):
        save_pronunciations(self.pronunciations, self.config)
        self.win.destroy()


# ── GUI ──────────────────────────────────────────────────────────────────────
class PPTTTSApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PPT TTS — PowerPoint Narrator")
        self.root.geometry("640x520")
        self.root.resizable(False, False)

        self.input_path = tk.StringVar()
        self.preset_var = tk.StringVar(value="mandarin")
        self.slides_var = tk.StringVar()
        self.processing = False
        self.config = load_config()
        self.root._ppttts_app = self  # let the dialog find us via root

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 4}

        # ── File selection ───────────────────────────────────────────────
        file_frame = ttk.LabelFrame(self.root, text="PowerPoint File", padding=8)
        file_frame.pack(fill="x", **pad)

        ttk.Entry(file_frame, textvariable=self.input_path, width=60).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(file_frame, text="Browse ...", command=self._browse_file).pack(
            side="right")

        # ── Options ─────────────────────────────────────────────────────
        opt_frame = ttk.LabelFrame(self.root, text="Options", padding=8)
        opt_frame.pack(fill="x", **pad)

        ttk.Label(opt_frame, text="Voice:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        preset_combo = ttk.Combobox(
            opt_frame, textvariable=self.preset_var, state="readonly", width=28,
            values=[f"{k} — {v}" for k, v in PRESET_LABELS.items()],
        )
        preset_combo.grid(row=0, column=1, sticky="w")
        # Map display label back to key
        self._preset_keys = list(PRESET_LABELS.keys())
        self._preset_combo = preset_combo
        preset_combo.current(0)

        ttk.Label(opt_frame, text="Slides (opt):").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        slides_entry = ttk.Entry(opt_frame, textvariable=self.slides_var, width=28)
        slides_entry.grid(row=1, column=1, sticky="w", pady=(6, 0))
        ttk.Label(opt_frame, text="e.g. 1,3,5-8").grid(
            row=1, column=2, sticky="w", padx=(6, 0), pady=(6, 0))

        ttk.Button(opt_frame, text="Corrections...",
                   command=self._open_corrections).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # ── Generate button ─────────────────────────────────────────────
        self.generate_btn = ttk.Button(
            self.root, text="Generate", command=self._on_generate)
        self.generate_btn.pack(pady=8)

        # ── Log area ────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self.root, text="Status", padding=8)
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = tk.Text(log_frame, height=12, state="disabled",
                                wrap="word", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical",
                                  command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

    def _browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")])
        if path:
            self.input_path.set(path)
            detected = detect_language(path)
            idx = self._preset_keys.index(detected) if detected in self._preset_keys else 0
            self._preset_combo.current(idx)

    def _open_corrections(self):
        CorrectionsDialog(self.root, self.config)

    def _log(self, msg: str):
        def _append():
            self.log_text.configure(state="disabled")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.configure(state="disabled")
            self.log_text.see("end")
        # Accumulate logs in a buffer and display them
        if not hasattr(self, "_log_buffer"):
            self._log_buffer = []
        self._log_buffer.append(msg)
        full = "\n".join(self._log_buffer)
        self.root.after(0, _append_set, self, full)

    def _on_generate(self):
        if self.processing:
            return

        input_file = self.input_path.get().strip()
        if not input_file:
            messagebox.showwarning("Missing file", "Please select a PowerPoint file.")
            return

        if not Path(input_file).exists():
            messagebox.showerror("File not found", f"Cannot find:\n{input_file}")
            return

        # Resolve preset
        preset_key = self._preset_keys[0]  # default
        combo_val = self.preset_var.get()
        for k in self._preset_keys:
            if combo_val.startswith(k):
                preset_key = k
                break
        preset = VOICE_PRESETS[preset_key]

        input_path = Path(input_file)
        output_path = input_path.with_name(f"narrated_{input_path.stem}{input_path.suffix}")
        i = 1
        while output_path.exists():
            output_path = input_path.with_name(f"narrated_{input_path.stem}_{i}{input_path.suffix}")
            i += 1

        slides_spec = self.slides_var.get().strip() or None

        self.processing = True
        self.generate_btn.configure(state="disabled")
        self._log_buffer = []

        def worker():
            try:
                # Validate slide range early
                if slides_spec:
                    prs = Presentation(str(input_path))
                    parse_slide_range(slides_spec, len(prs.slides))

                all_pronunciations = load_pronunciations(self.config)
                voice_corrections = all_pronunciations.get(preset_key, [])

                process_pptx(
                    input_path, output_path,
                    voice_id=preset["voice"],
                    engine="neural",
                    lang_code=preset["lang"],
                    slides_spec=slides_spec,
                    callback=self._log,
                    pronunciations=voice_corrections,
                )
                self.root.after(0, lambda: messagebox.showinfo(
                    "Complete", f"Saved to:\n{output_path}"))
            except Exception as e:
                self._log(f"ERROR: {e}")
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", str(e)))
            finally:
                self.processing = False
                self.root.after(0, lambda: self.generate_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


def _append_set(app, full_text):
    app.log_text.configure(state="normal")
    app.log_text.delete("1.0", "end")
    app.log_text.insert("end", full_text)
    app.log_text.configure(state="disabled")
    app.log_text.see("end")


def gui_main():
    root = tk.Tk()
    PPTTTSApp(root)
    root.mainloop()


if __name__ == "__main__":
    cli_main()
