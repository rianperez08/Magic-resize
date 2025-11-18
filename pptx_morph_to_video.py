import argparse
import os
import time

import win32com.client


def build_combined_pptx(app, input_16_9, input_4_3, temp_pptx_path,
                        slide1_duration, transition_duration):
    """
    Build a combined presentation using 16:9 as base and 4:3 as layout reference:

    - Open 16:9 PPTX as base.
    - Keep only its first slide.
    - Open 4:3 PPTX as reference (Magic Resize layout from Canva).
    - Duplicate 16:9 slide 1 as slide 2.
    - For each shape index i on slide 2, read the position/size of shape i on
      the 4:3 slide and map it into the 16:9 coordinate system.
    - Result: slide 2 elements are arranged like the 4:3 layout, but still live
      inside the 16:9 deck (no grayscale/theme issues).
    - Set Morph (if available) or Fade as transition into slide 2.
    - Set timings, save to temp_pptx_path (best-effort).
    """
    print("[PY] Opening base 16:9 presentation...")
    base_pres = app.Presentations.Open(os.path.abspath(input_16_9), WithWindow=False)

    # Keep only first slide in the 16:9 deck
    while base_pres.Slides.Count > 1:
        base_pres.Slides(2).Delete()

    print("[PY] Opening 4:3 reference presentation...")
    ref_pres = app.Presentations.Open(os.path.abspath(input_4_3), WithWindow=False)

    base_slide1 = base_pres.Slides(1)
    ref_slide1 = ref_pres.Slides(1)

    print("[PY] Duplicating base slide 1 to create slide 2...")
    base_slide1.Copy()
    base_pres.Slides.Paste(2)
    slide2 = base_pres.Slides(2)

    const = win32com.client.constants

    # Optional: make slide 2 background black to emphasize "box" feeling
    print("[PY] Setting slide 2 background to black...")
    slide2.FollowMasterBackground = False
    bg = slide2.Background
    bg.Fill.Solid()
    bg.Fill.ForeColor.RGB = 0  # RGB(0,0,0)

    # Coordinate systems
    base_w = base_pres.PageSetup.SlideWidth
    base_h = base_pres.PageSetup.SlideHeight

    ref_w = ref_pres.PageSetup.SlideWidth
    ref_h = ref_pres.PageSetup.SlideHeight

    print(f"[PY] base slide size: {base_w} x {base_h}")
    print(f"[PY] ref  slide size: {ref_w} x {ref_h}")

    # Map 4:3 canvas → 16:9 canvas:
    # - Match height exactly
    # - Pillarbox horizontally
    scale = base_h / float(ref_h) if ref_h else 1.0
    target_w = ref_w * scale
    offset_x = (base_w - target_w) / 2.0

    print("[PY] Using scale =", scale, "offset_x =", offset_x)

    # Map shapes by index (best-effort)
    shape_count = min(slide2.Shapes.Count, ref_slide1.Shapes.Count)
    print(f"[PY] Mapping {shape_count} shapes from 4:3 to 16:9...")

    for i in range(1, shape_count + 1):
        try:
            s2 = slide2.Shapes(i)
            s_ref = ref_slide1.Shapes(i)

            # Skip invisible shapes if any
            if not s2.Visible or not s_ref.Visible:
                continue

            # Reference geometry (in ref slide coords)
            r_left = s_ref.Left
            r_top = s_ref.Top
            r_width = s_ref.Width
            r_height = s_ref.Height

            # Map into base slide coords
            new_left = offset_x + r_left * scale
            new_top = r_top * scale
            new_width = r_width * scale
            new_height = r_height * scale

            s2.Left = new_left
            s2.Top = new_top
            s2.Width = new_width
            s2.Height = new_height

        except Exception as e:
            # Don't let a single weird shape kill the whole layout
            print(f"[PY] Warning: could not map shape {i}: {e}")
            continue

    # Done with 4:3 reference
    ref_pres.Close()

    # Transitions
    slide1 = base_pres.Slides(1)

    print("[PY] Setting transition on slide 2 (Morph if available, else Fade)...")
    try:
        if hasattr(const, "ppEffectMorphByObject"):
            morph_effect = const.ppEffectMorphByObject
        elif hasattr(const, "ppEffectMorph"):
            morph_effect = const.ppEffectMorph
        else:
            morph_effect = 3954  # numeric fallback; may or may not work

        slide2.SlideShowTransition.EntryEffect = morph_effect
        print("[PY] Morph effect applied to slide 2.")
    except Exception as e:
        print("[PY] Could not apply Morph, falling back to Fade. Error:", e)
        try:
            slide2.SlideShowTransition.EntryEffect = const.ppEffectFade
        except Exception:
            print("[PY] Could not apply Fade either, leaving default transition.")

    # Timings
    slide1.SlideShowTransition.AdvanceOnTime = True
    slide1.SlideShowTransition.AdvanceTime = float(slide1_duration)

    slide2.SlideShowTransition.AdvanceOnTime = True
    slide2.SlideShowTransition.AdvanceTime = float(transition_duration)

    # Save combined file (best-effort; ignore "file in use" errors)
    combined_path = os.path.abspath(temp_pptx_path)
    print(f"[PY] Saving combined Morph presentation to {combined_path} (best-effort)...")
    try:
        base_pres.SaveAs(combined_path)
    except Exception as e:
        print("[PY] Warning: could not SaveAs combined presentation:", e)
        print("[PY] Continuing anyway; CreateVideo will still run from memory.")

    return base_pres


def export_video_from_pptx(pres, output_path, width=1920, height=1080,
                           fps=30, quality=100):
    """
    Use PowerPoint's built-in CreateVideo to export video.
    """
    const = win32com.client.constants

    use_timings = True
    default_slide_duration = 5

    print("[PY] Starting CreateVideo export...")
    pres.CreateVideo(
        os.path.abspath(output_path),
        use_timings,
        default_slide_duration,
        height,
        fps,
        quality,
    )

    print("[PY] Waiting for video export to complete...")
    while True:
        status = pres.CreateVideoStatus
        if status == const.ppMediaTaskStatusDone:
            print("[PY] Video export completed.")
            break
        elif status == const.ppMediaTaskStatusFailed:
            raise RuntimeError("PowerPoint reported video export failure.")
        else:
            time.sleep(1)


def main():
    parser = argparse.ArgumentParser(description="Morph 16:9 -> 4:3-like layout into a video")
    parser.add_argument("--input_16_9", required=True, help="Path to the 16:9 PPTX")
    parser.add_argument("--input_4_3", required=True, help="Path to the 4:3 PPTX (layout reference)")
    parser.add_argument("--output", required=True, help="Path to output MP4")
    parser.add_argument(
        "--slide1_duration",
        type=float,
        default=2.0,
        help="Seconds to show the first slide before morph",
    )
    parser.add_argument(
        "--transition_duration",
        type=float,
        default=2.0,
        help="Seconds for the morph transition (and slide 2 display)",
    )
    parser.add_argument(
        "--temp_pptx",
        default="combined_morph.pptx",
        help="Intermediate combined PPTX file",
    )

    args = parser.parse_args()

    input_16_9 = args.input_16_9
    input_4_3 = args.input_4_3
    output_video = args.output
    temp_pptx_path = args.temp_pptx

    print("[PY] Morph pipeline start")
    print("[PY] 16:9 PPTX:", input_16_9)
    print("[PY] 4:3 PPTX (layout reference):", input_4_3)
    print("[PY] Output   :", output_video)

    print("[PY] Launching PowerPoint COM automation...")
    app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
    try:
        app.Visible = True
    except Exception as e:
        print("[PY] Warning: could not set PowerPoint visibility:", e)

    pres_combined = None

    try:
        pres_combined = build_combined_pptx(
            app,
            input_16_9,
            input_4_3,
            temp_pptx_path,
            slide1_duration=args.slide1_duration,
            transition_duration=args.transition_duration,
        )

        export_video_from_pptx(
            pres_combined,
            output_video,
            width=1920,
            height=1080,
            fps=30,
            quality=100,
        )

        print("[PY] Pipeline finished successfully.")
    finally:
        try:
            if pres_combined is not None:
                pres_combined.Close()
        except Exception:
            pass

        try:
            app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
