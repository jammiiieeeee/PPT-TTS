"""
generate_audio_ppt.py
Reads a PowerPoint, extracts slide notes, synthesizes speech via AWS Polly,
and embeds the audio onto each slide for Storyline 360 import.
"""

import argparse
import logging
import os
import sys
import tempfile
from pathlib import Path

import boto3
from botocore.config import Config as BotoConfig
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

log = logging.getLogger("polly-tts")

# ── Voice presets ────────────────────────────────────────────────────────────
VOICE_PRESETS = {
    "cantonese": {"voice": "Hiujin",    "lang": "yue-CN"},
    "mandarin":  {"voice": "Zhiyu",     "lang": "cmn-CN"},
    "english":   {"voice": "Matthew",   "lang": "en-US"},
}


def parse_slide_range(spec: str, total: int) -> list[int]:
    """Parse '1,3,5-8' into sorted 0-indexed slide indices."""
    indices: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            lo, hi = part.split("-", 1)
            lo, hi = int(lo), int(hi)
            if lo < 1 or hi > total or lo > hi:
                log.error("Invalid range %s for presentation with %d slides", part, total)
                sys.exit(1)
            indices.update(range(lo - 1, hi))
        else:
            n = int(part)
            if n < 1 or n > total:
                log.error("Slide %d out of range (1-%d)", n, total)
                sys.exit(1)
            indices.add(n - 1)
    return sorted(indices)


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

    # Use the official add_movie() method — it creates correct PPTX structure
    # that PowerPoint and Storyline 360 both recognize.
    movie = slide.shapes.add_movie(
        audio_path,
        left=Inches(-2.0),
        top=Inches(-2.0),
        width=Inches(1.0),
        height=Inches(1.0),
        poster_frame_image=None,
        mime_type="audio/mpeg",
    )

    # Patch p:timing: change delay="indefinite" (click-to-play) to delay="0" (autoplay)
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
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.Presentations.Open(abs_path, ReadOnly=False)

        for slide_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(slide_idx)
            slide.Select()
            for shape_idx in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(shape_idx)
                # Skip shapes that are not media (no MediaType attribute or
                # MediaType 0 = ppObjectTypeNone). Audio added via add_movie
                # appears as a media shape.
                try:
                    mt = shape.MediaType
                except Exception:
                    mt = 0
                if mt == 0:
                    continue
                # Nudge shape 1 EMU right then back — triggers PowerPoint
                # to process the media and normalize the XML.
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


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Add AWS Polly TTS narration to PowerPoint slides.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "voice presets:\n"
            "  cantonese  -> Hiujin    / yue-CN\n"
            "  mandarin   -> Zhiyu     / cmn-CN\n"
            "  english    -> Matthew   / en-US\n"
        ),
    )
    p.add_argument("input", help="Path to the input .pptx file")
    p.add_argument("-o", "--output", help="Output file path (default: narrated_<input>)")
    p.add_argument("--voice", default="Hiujin", help="Polly voice ID (default: Hiujin)")
    p.add_argument("--engine", choices=["standard", "neural"], default="neural",
                   help="TTS engine (default: neural)")
    p.add_argument("--lang", default="yue-CN", help="Language code (default: yue-CN)")
    p.add_argument("--preset", choices=list(VOICE_PRESETS), default=None,
                   help="Use a named voice preset instead of --voice/--lang")
    p.add_argument("--slides", default=None,
                   help="Slide numbers or ranges, e.g. '1,3,5-8' (default: all)")
    p.add_argument("--dry-run", action="store_true",
                   help="Show extracted notes without calling AWS")
    p.add_argument("--keep-audio", action="store_true",
                   help="Keep temp MP3 files after processing")
    p.add_argument("-v", "--verbose", action="store_true", help="Enable debug logging")
    return p


def main() -> None:
    args = build_parser().parse_args()

    # ── Logging ──────────────────────────────────────────────────────────────
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="[%(levelname)s] %(message)s",
    )

    # ── Resolve voice settings ───────────────────────────────────────────────
    voice_id = args.voice
    lang_code = args.lang
    if args.preset:
        preset = VOICE_PRESETS[args.preset]
        voice_id = preset["voice"]
        lang_code = preset["lang"]
        log.info("Using preset '%s': voice=%s, lang=%s", args.preset, voice_id, lang_code)

    engine = args.engine

    # ── Load presentation ────────────────────────────────────────────────────
    input_path = Path(args.input)
    if not input_path.exists():
        log.error("File not found: %s", input_path)
        sys.exit(1)

    if args.output:
        output_path = Path(args.output)
    else:
        stem = input_path.stem
        suffix = input_path.suffix
        output_path = input_path.with_name(f"narrated_{stem}{suffix}")
        i = 1
        while output_path.exists():
            output_path = input_path.with_name(f"narrated_{stem}_{i}{suffix}")
            i += 1

    try:
        prs = Presentation(str(input_path))
    except Exception as e:
        log.error("Failed to load PPTX: %s", e)
        sys.exit(1)

    total = len(prs.slides)
    log.info("Loaded '%s' — %d slide(s)", input_path.name, total)

    # ── Determine which slides to process ────────────────────────────────────
    if args.slides:
        slide_indices = parse_slide_range(args.slides, total)
    else:
        slide_indices = list(range(total))

    log.info("Processing %d slide(s): %s", len(slide_indices),
             ", ".join(str(i + 1) for i in slide_indices))

    # ── Dry-run mode ─────────────────────────────────────────────────────────
    if args.dry_run:
        for idx in slide_indices:
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

    # ── AWS Polly client with retries ────────────────────────────────────────
    polly_client = boto3.client(
        "polly",
        config=BotoConfig(
            retries={"max_attempts": 3, "mode": "adaptive"},
        ),
    )
    log.info("AWS Polly client ready (region: %s)",
             polly_client.meta.region_name or "(default)")

    # ── Temp directory for audio files ───────────────────────────────────────
    with tempfile.TemporaryDirectory(prefix="polly_tts_") as tmp_dir:
        success_count = 0
        skip_count = 0
        fail_count = 0

        for idx in slide_indices:
            slide = prs.slides[idx]
            slide_num = idx + 1

            # Extract notes
            notes_text = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes_text = slide.notes_slide.notes_text_frame.text.strip()

            if not notes_text:
                log.warning("Slide %d: no notes — skipping", slide_num)
                skip_count += 1
                continue

            # Synthesize
            audio_path = os.path.join(tmp_dir, f"slide_{slide_num}.mp3")
            log.info("Slide %d: synthesizing %d chars ...", slide_num, len(notes_text))

            try:
                ok = synthesize_audio(polly_client, notes_text, audio_path,
                                      voice_id, engine, lang_code)
                if not ok:
                    log.error("Slide %d: Polly returned no audio stream", slide_num)
                    fail_count += 1
                    continue
            except Exception as e:
                log.error("Slide %d: Polly API error: %s", slide_num, e)
                fail_count += 1
                continue

            # Embed off-screen (Storyline detects and pulls to timeline)
            embed_audio_on_slide(slide, audio_path)
            log.info("Slide %d: audio embedded", slide_num)
            success_count += 1

            # Optionally copy to persistent location
            if args.keep_audio:
                keep_dir = Path(tmp_dir).parent / "polly_audio"
                keep_dir.mkdir(exist_ok=True)
                import shutil
                shutil.copy2(audio_path, keep_dir / f"slide_{slide_num}.mp3")

    # ── Save ─────────────────────────────────────────────────────────────────
    prs.save(str(output_path))
    log.info("Saved: %s", output_path)

    # ── Normalize via PowerPoint COM (triggers Storyline recognition) ────────
    normalize_audio_for_storyline(output_path)

    # ── Summary ──────────────────────────────────────────────────────────────
    log.info("Done — %d succeeded, %d skipped (no notes), %d failed",
             success_count, skip_count, fail_count)

    if args.keep_audio:
        log.info("Temp audio kept at: %s", Path("polly_audio").resolve())


if __name__ == "__main__":
    main()
