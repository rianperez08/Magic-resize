#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
auto_match_move_to_video.py

Goal:
Make it look like a 16:9 slide is being resized into a 4:3 slide.

Given:
  A = 16:9 slide PNG (original)
  B = 4:3 slide PNG (resized version from Canva)

We do:
  - Video canvas = 16:9 (same size as A).
  - At t = 0: A fills the whole canvas.
  - At t = 1: B appears as 4:3, centered, with black side bars.
  - From t in [0,1]:
      * A "canvas" rectangle shrinks in width from 16:9 → 4:3 width.
      * Inside that rectangle, we crossfade from A → B.
      * Outside the rectangle, black bars are visible.

Usage:
  python auto_match_move_to_video.py \
      --a a.png \
      --b b.png \
      --out output.mp4 \
      --fps 30 \
      --seconds 2.0
"""

import argparse
import os
import sys

import cv2
import numpy as np


def parse_args():
    p = argparse.ArgumentParser(description="Resizing-style morph 16:9 → 4:3.")
    p.add_argument("--a", required=True, help="Path to first slide image (16:9)")
    p.add_argument("--b", required=True, help="Path to second slide image (4:3)")
    p.add_argument("--out", required=True, help="Output MP4 path")
    p.add_argument("--fps", type=int, default=30, help="Frames per second")
    p.add_argument("--seconds", type=float, default=2.0, help="Animation duration (seconds)")
    return p.parse_args()


def load_image(path, label):
    img = cv2.imread(path, cv2.IMREAD_COLOR)
    if img is None:
        print(f"[ERROR] Failed to read image {label}: {path}")
        sys.exit(1)
    return img


def main():
    args = parse_args()

    # Load images
    img_a = load_image(args.a, "A (16:9)")
    img_b = load_image(args.b, "B (4:3)")

    h_a, w_a = img_a.shape[:2]
    h_b, w_b = img_b.shape[:2]
    print(f"[INFO] A size: {w_a}x{h_a}, B size: {w_b}x{h_b}")

    # We'll use A's size as the video size (16:9)
    w_out, h_out = w_a, h_a

    # Scale B so its height matches the output height; width becomes ~4:3
    scale_b = h_out / float(h_b)
    new_w_b = int(round(w_b * scale_b))
    img_b_scaled = cv2.resize(img_b, (new_w_b, h_out), interpolation=cv2.INTER_AREA)
    w_b_scaled = img_b_scaled.shape[1]
    print(f"[INFO] Scaled B to: {w_b_scaled}x{h_out} (should be ~4:3)")

    # We want the "resize" rectangle width to go from full 16:9 width to 4:3 width
    # 16:9 width = w_out
    # 4:3 width at this height = h_out * 4 / 3 (or just use w_b_scaled)
    w_43 = int(round(h_out * 4.0 / 3.0))
    # To be safe, clamp w_43 to <= w_out
    w_43 = min(w_43, w_out)
    print(f"[INFO] Target 4:3 width at this height: {w_43}")

    # We'll also scale A "inside" the rectangle as it changes width.
    # At t=0, rectangle width = w_out → A looks normal (and fills frame).
    # At t=1, rectangle width = w_43 → A is squeezed horizontally, but we crossfade to B.

    fps = args.fps
    seconds = args.seconds
    total_frames = max(2, int(round(fps * seconds)))

    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    writer = cv2.VideoWriter(args.out, fourcc, fps, (w_out, h_out))
    if not writer.isOpened():
        print("[ERROR] Could not open VideoWriter.")
        sys.exit(2)

    print(
        f"[INFO] Rendering video {args.out} at {fps} fps for {seconds} s "
        f"({total_frames} frames), size {w_out}x{h_out}…"
    )

    # Precompute scale for A's height: we assume A already matches h_out
    # If not, scale A to match height
    if h_a != h_out:
        img_a = cv2.resize(img_a, (w_out, h_out), interpolation=cv2.INTER_AREA)
        h_a, w_a = img_a.shape[:2]

    for frame_idx in range(total_frames):
        if total_frames == 1:
            t = 1.0
        else:
            t = frame_idx / float(total_frames - 1)

        # You can use easing if you like:
        # t_eased = 0.5 - 0.5 * np.cos(np.pi * t)
        t_eased = t

        # Current rectangle width (visual "canvas")
        rect_w = int(round((1.0 - t_eased) * w_out + t_eased * w_43))
        rect_h = h_out

        # Center rectangle horizontally
        cx = w_out // 2
        x0 = cx - rect_w // 2
        x1 = x0 + rect_w
        y0 = 0
        y1 = rect_h

        # Clamp just in case
        x0 = max(0, x0)
        x1 = min(w_out, x1)

        # Background: black (so you clearly see side bars at the end)
        frame = np.zeros((h_out, w_out, 3), dtype=np.uint8)

        # Scale A and B to current rectangle size
        # A: scaled to (rect_w, rect_h)
        a_scaled = cv2.resize(img_a, (rect_w, rect_h), interpolation=cv2.INTER_AREA)

        # B: we start from img_b_scaled (already full height h_out, width ~w_43), but
        # to keep the "resizing" feel, we also scale it to (rect_w, rect_h)
        b_scaled = cv2.resize(img_b_scaled, (rect_w, rect_h), interpolation=cv2.INTER_AREA)

        # Crossfade inside the rect: A -> B
        inside = cv2.addWeighted(a_scaled, 1.0 - t_eased, b_scaled, t_eased, 0.0)

        # Place inside-rect onto the frame
        frame[y0:y1, x0:x1] = inside

        # Draw white border to emphasize the "canvas"
        cv2.rectangle(frame, (x0, y0), (x1, y1), (255, 255, 255), thickness=2)

        writer.write(frame)

    writer.release()
    print(f"[OK] Saved video: {args.out}")


if __name__ == "__main__":
    main()
