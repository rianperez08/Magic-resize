#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pptx_pair_morph_to_video.py

Takes two PPTX files:
- original 16:9
- resized 4:3

Builds a new 2-slide presentation:
  Slide 1 = first slide of original PPTX
  Slide 2 = first slide of resized PPTX (with Morph transition)
Then exports to MP4 using PowerPoint's CreateVideo.

Windows-only (requires Microsoft PowerPoint desktop).

Usage:
  python pptx_pair_morph_to_video.py \
      --orig original.pptx \
      --resized resized.pptx \
      --out_video output.mp4 \
      --seconds 5 \
      --fps 30 \
      --height 1080 \
      --quality 100 \
      --morph_mode object
"""

import argparse
import os
import sys
import time

import win32com.client as win32

# PowerPoint Morph entry effects (late-binding-friendly numeric values)
# See: ppEffectMorphByObject / Word / Char :contentReference[oaicite:0]{index=0}
PP_EFFECT_MORPH_BY_OBJECT = 3954
PP_EFFECT_MORPH_BY_WORD = 3955
PP_EFFECT_MORPH_BY_CHAR = 3956

# CreateVideoStatus values (PpMediaTaskStatus) :contentReference[oaicite:1]{index=1}
PP_MEDIA_STATUS_NONE = 0
PP_MEDIA_STATUS_QUEUED = 1
PP_MEDIA_STATUS_INPROGRESS = 2
PP_MEDIA_STATUS_DONE = 3
PP_MEDIA_STATUS_FAILED = 4


def parse_args():
    p = argparse.ArgumentParser(description="Build 2-slide Morph deck and export to MP4.")
    p.add_argument("--orig", required=True, help="Path to original 16:9 PPTX")
    p.add_argument("--resized", required=True, help="Path to resized 4:3 PPTX")
    p.add_argument("--out_video", required=True, help="Path to output MP4 file")
    p.add_argument("--seconds", type=float, default=5.0, help="Morph duration for slide 2")
    p.add_argument("--fps", type=int, default=30, help="Frames per second")
    p.add_argument("--height", type=int, default=1080, help="Vertical resolution")
    p.add_argument("--quality", type=int, default=100, help="Quality (1–100)")
    p.add_argument(
        "--morph_mode",
        choices=["object", "word", "char"],
        default="object",
        help="Morph mode"
    )
    return p.parse_args()


def main():
    args = parse_args()

    orig = os.path.abspath(args.orig)
    resized = os.path.abspath(args.resized)
    out_video = os.path.abspath(args.out_video)

    if not os.path.exists(orig):
        print(f"[ERROR] Original PPTX not found: {orig}")
        sys.exit(1)
    if not os.path.exists(resized):
        print(f"[ERROR] Resized PPTX not found: {resized}")
        sys.exit(1)

    out_dir = os.path.dirname(out_video) or "."
    os.makedirs(out_dir, exist_ok=True)

    # Temporary base PPTX for the 2-slide deck (optional, but safer for CreateVideo)
    base_pptx = os.path.splitext(out_video)[0] + "_morph_base.pptx"

    morph_effect = {
        "object": PP_EFFECT_MORPH_BY_OBJECT,
        "word": PP_EFFECT_MORPH_BY_WORD,
        "char": PP_EFFECT_MORPH_BY_CHAR,
    }[args.morph_mode]

    print("[INFO] Launching PowerPoint...")
    try:
        pp = win32.gencache.EnsureDispatch("PowerPoint.Application")
    except Exception:
        pp = win32.Dispatch("PowerPoint.Application")

    pp.Visible = True  # or False if you want it hidden

    pres_orig = None
    pres_resized = None
    pres_out = None

    try:
        print(f"[INFO] Opening original: {orig}")
        pres_orig = pp.Presentations.Open(orig, WithWindow=False)

        print(f"[INFO] Opening resized: {resized}")
        pres_resized = pp.Presentations.Open(resized, WithWindow=False)

        if pres_orig.Slides.Count < 1:
            print("[ERROR] Original PPTX has no slides.")
            sys.exit(2)
        if pres_resized.Slides.Count < 1:
            print("[ERROR] Resized PPTX has no slides.")
            sys.exit(3)

        print("[INFO] Creating new 2-slide presentation...")
        pres_out = pp.Presentations.Add()

        # Copy slide 1 from original
        pres_orig.Slides(1).Copy()
        pres_out.Slides.Paste(1)

        # Copy slide 1 from resized
        pres_resized.Slides(1).Copy()
        pres_out.Slides.Paste(2)

        if pres_out.Slides.Count != 2:
            print(f"[ERROR] Expected 2 slides, got {pres_out.Slides.Count}.")
            sys.exit(4)

        s1 = pres_out.Slides(1)
        s2 = pres_out.Slides(2)

        # Transition timings
        # Slide 1: show immediately, no wait
        s1.SlideShowTransition.AdvanceOnTime = True
        s1.SlideShowTransition.AdvanceTime = 0

        # Slide 2: let Morph handle the duration, then end
        s2.SlideShowTransition.AdvanceOnTime = True
        s2.SlideShowTransition.AdvanceTime = 0

        # Apply Morph on slide 2
        try:
            s2.SlideShowTransition.EntryEffect = morph_effect
        except Exception as e:
            print(
                "[WARN] Could not set Morph entry effect directly; "
                "check Office version. Error:", e
            )

        try:
            s2.SlideShowTransition.Duration = float(args.seconds)
        except Exception as e:
            print("[WARN] Could not set transition Duration:", e)

        print(f"[INFO] Saving base 2-slide deck to: {base_pptx}")
        pres_out.SaveAs(base_pptx)

        # Start video export using timings & transitions
        print("[INFO] Starting CreateVideo export...")
        pres_out.CreateVideo(
            out_video,
            True,                 # UseTimingsAndNarrations
            args.seconds,         # DefaultSlideDuration (fallback)
            args.height,          # VertResolution
            args.fps,             # FramesPerSecond
            args.quality          # Quality
        )  # :contentReference[oaicite:2]{index=2}

        # Poll CreateVideoStatus
        status_names = {
            PP_MEDIA_STATUS_NONE: "None",
            PP_MEDIA_STATUS_QUEUED: "Queued",
            PP_MEDIA_STATUS_INPROGRESS: "InProgress",
            PP_MEDIA_STATUS_DONE: "Done",
            PP_MEDIA_STATUS_FAILED: "Failed",
        }

        while True:
            status = pres_out.CreateVideoStatus
            name = status_names.get(status, f"Unknown({status})")
            print(f"[INFO] CreateVideoStatus: {name}")

            if status == PP_MEDIA_STATUS_DONE:
                print("[INFO] Video export completed.")
                break
            if status == PP_MEDIA_STATUS_FAILED:
                print("[ERROR] Video export FAILED.")
                sys.exit(5)

            time.sleep(1.5)

    finally:
        # Close everything
        for p in (pres_out, pres_orig, pres_resized):
            try:
                if p is not None:
                    p.Close()
            except Exception:
                pass

        try:
            pp.Quit()
        except Exception:
            pass

    print(f"[OK] Saved video: {out_video}")


if __name__ == "__main__":
    main()
