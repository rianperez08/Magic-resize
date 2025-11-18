import os
import time
import argparse
import win32com.client
from win32com.client import constants


# ----- PowerPoint CreateVideo status codes (common values) -----
PP_MEDIA_TASK_STATUS_NOT_STARTED = 0
PP_MEDIA_TASK_STATUS_IN_PROGRESS = 1
PP_MEDIA_TASK_STATUS_DONE = 3


def get_powerpoint():
    """Start (or re-use) a PowerPoint instance."""
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    ppt.Visible = True  # set False once it's stable, True is nice while debugging
    return ppt


def combine_presentations_with_morph(
    ppt,
    pptx_16x9_path,
    pptx_4x3_path,
    output_pptx_path,
    slide1_duration=2.0,
    transition_duration=1.0,
):
    """
    Create a new presentation with:
      - Slide 1 = slide 1 from 16:9 deck
      - Slide 2 = slide 1 from 4:3 deck (content adapted to base slide size)
    Then:
      - Slide 1 shows for `slide1_duration` seconds, then auto-advances
      - Slide 2 has a Morph transition (for PPT playback) with `transition_duration`
    """

    print("[PPT] Opening base (16:9) presentation:", pptx_16x9_path)
    base_pres = ppt.Presentations.Open(pptx_16x9_path, WithWindow=False)

    # If there are multiple slides, keep only the first as our "16:9 layout"
    while base_pres.Slides.Count > 1:
        base_pres.Slides(2).Delete()

    print("[PPT] Opening 4:3 presentation:", pptx_4x3_path)
    other_pres = ppt.Presentations.Open(pptx_4x3_path, WithWindow=False)

    try:
        # Copy slide 1 from 4:3 deck
        source_slide = other_pres.Slides(1)
        source_slide.Copy()

        # Paste it as slide 2 in the base deck
        base_pres.Slides.Paste(Index=2)

        # Get references to slides
        slide1 = base_pres.Slides(1)
        slide2 = base_pres.Slides(2)

        # --- Slide 1 timing ---
        # Show for slide1_duration seconds, then auto advance
        slide1.SlideShowTransition.AdvanceOnTime = True
        slide1.SlideShowTransition.AdvanceTime = slide1_duration
        # Duration of transition into slide 1: 0 (it's the first slide)
        slide1.SlideShowTransition.Duration = 0.0

        # --- Slide 2 Morph transition ---
        # No auto-advance (last slide)
        slide2.SlideShowTransition.AdvanceOnTime = False
        # Duration of the Morph animation (seconds)
        slide2.SlideShowTransition.Duration = transition_duration

        # Set Morph entry effect if available
        try:
            # This constant is usually present in newer Office:
            #   ppEffectMorph = 40 (but we prefer name if pywin32 exposes it)
            morph_effect = getattr(constants, "ppEffectMorph", 40)
            slide2.SlideShowTransition.EntryEffect = morph_effect
            print("[PPT] Morph effect applied to slide 2.")
        except Exception as e:
            print("[PPT] Could not set Morph entry effect. Falling back to Fade.")
            print("       Error:", e)
            # Fallback to a standard transition
            slide2.SlideShowTransition.EntryEffect = constants.ppEffectFade

        # Save the combined deck
        print("[PPT] Saving combined Morph deck:", output_pptx_path)
        base_pres.SaveAs(output_pptx_path)

    finally:
        # Close 4:3 deck, keep base_pres open (we need it for video export)
        other_pres.Close()

    return base_pres  # return Presentation object so we can call CreateVideo on it


def create_video_from_presentation(
    pres,
    output_video_path,
    use_timings=True,
    default_slide_duration=2.0,
    vert_resolution=1080,
    fps=25,
    quality=85,
    poll_interval=2.0,
):
    """
    Ask PowerPoint to create a video from the given Presentation.

    Arguments roughly map to PowerPoint's CreateVideo:
      - use_timings: if True, uses SlideShowTransition settings (AdvanceTime/Duration)
      - default_slide_duration: used if no timings
      - vert_resolution: e.g. 720, 1080
      - fps: frames per second
      - quality: 1–100, higher = better
    """

    print(f"[PPT] Starting CreateVideo -> {output_video_path}")

    # If file exists, remove it first
    if os.path.exists(output_video_path):
        os.remove(output_video_path)

    # PowerPoint's CreateVideo signature:
    # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration,
    #             VertResolution, FramesPerSecond, Quality)
    pres.CreateVideo(
        output_video_path,
        use_timings,
        default_slide_duration,
        vert_resolution,
        fps,
        quality,
    )

    # Poll CreateVideoStatus until done
    while True:
        status = pres.CreateVideoStatus
        if status == PP_MEDIA_TASK_STATUS_DONE:
            print("[PPT] CreateVideo status: DONE")
            break
        elif status == PP_MEDIA_TASK_STATUS_IN_PROGRESS:
            print("[PPT] CreateVideo status: In progress...")
        elif status == PP_MEDIA_TASK_STATUS_NOT_STARTED:
            print("[PPT] CreateVideo status: NOT STARTED YET")
        else:
            print(f"[PPT] CreateVideo status code: {status}")
        time.sleep(poll_interval)

    print(f"[PPT] Video created at: {output_video_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Combine 16:9 + 4:3 Canva PPTX into a Morph deck and export to video using PowerPoint."
    )
    parser.add_argument("--pptx_16x9", required=True, help="Path to Canva 16:9 PPTX")
    parser.add_argument("--pptx_4x3", required=True, help="Path to Canva 4:3 PPTX")
    parser.add_argument("--out_pptx", required=True, help="Output combined PPTX path")
    parser.add_argument("--out_video", required=True, help="Output video (MP4) path")

    parser.add_argument("--slide1_duration", type=float, default=2.0,
                        help="Seconds to show slide 1 before morph")
    parser.add_argument("--transition_duration", type=float, default=1.0,
                        help="Morph transition duration (seconds)")

    parser.add_argument("--video_fps", type=int, default=25,
                        help="Video frames per second")
    parser.add_argument("--video_resolution", type=int, default=1080,
                        help="Vertical resolution: 720, 1080, etc.")
    parser.add_argument("--video_quality", type=int, default=85,
                        help="Video quality (1–100)")

    args = parser.parse_args()

    ppt = get_powerpoint()

    try:
        # Step 1: Combine presentations and set Morph
        base_pres = combine_presentations_with_morph(
            ppt,
            pptx_16x9_path=args.pptx_16x9,
            pptx_4x3_path=args.pptx_4x3,
            output_pptx_path=args.out_pptx,
            slide1_duration=args.slide1_duration,
            transition_duration=args.transition_duration,
        )

        # Step 2: Export video from the combined presentation
        create_video_from_presentation(
            base_pres,
            output_video_path=args.out_video,
            use_timings=True,
            default_slide_duration=args.slide1_duration,
            vert_resolution=args.video_resolution,
            fps=args.video_fps,
            quality=args.video_quality,
            poll_interval=2.0,
        )

        # Close the base presentation
        base_pres.Close()
        print("[PIPELINE] ✅ Finished successfully.")

    finally:
        # If you want to close PowerPoint automatically:
        # ppt.Quit()
        pass


if __name__ == "__main__":
    main()
