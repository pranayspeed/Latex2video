#!/usr/bin/env python3
"""Generic Beamer-to-video generator using per-frame \\narration{...} text."""

import argparse
import asyncio
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List, Optional, Tuple

import edge_tts
from moviepy.editor import AudioFileClip, ImageClip, concatenate_videoclips


FRAME_RE = re.compile(
    r"\\begin\{frame\}(?:\[[^\]]*\])?(?:\{(?P<title>[^{}]*)\})?(?:\{[^{}]*\})?(?P<body>.*?)\\end\{frame\}",
    re.DOTALL,
)


class VideoGenerationError(Exception):
    pass


def _require_tool(binary: str) -> None:
    if shutil.which(binary) is None:
        raise VideoGenerationError(f"Required tool '{binary}' not found in PATH.")


def _strip_latex_comments(text: str) -> str:
    return re.sub(r"(?<!\\)%.*", "", text)


def _extract_braced_content(text: str, start: int) -> Tuple[str, int]:
    if start >= len(text) or text[start] != "{":
        return "", start
    depth = 0
    i = start
    while i < len(text):
        ch = text[i]
        if ch == "\\" and i + 1 < len(text):
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start + 1 : i], i + 1
        i += 1
    return text[start + 1 :], len(text)


def _latex_to_plain_text(text: str) -> str:
    text = _strip_latex_comments(text)
    text = re.sub(r"\\\\", "\n", text)
    text = re.sub(r"\\(?:textbf|textit|emph|underline|texttt|textrm|alert|structure)\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\textcolor\{[^{}]*\}\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\[a-zA-Z*]+(?:\[[^\]]*\])?(?:\{[^{}]*\})?", "", text)
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^\[[^\]]+\]\s*", "", text)  # drop prefixes like [20s]
    return text


def extract_frame_narrations(tex_path: Path) -> List[str]:
    source = tex_path.read_text(errors="ignore")
    source = _strip_latex_comments(source)
    narrations: List[str] = []

    for m in FRAME_RE.finditer(source):
        body = m.group("body")
        chunks: List[str] = []
        idx = 0
        while True:
            narr_match = re.search(r"\\narration(?:\[[^\]]*\])?(?:<[^>]*>)?\{", body[idx:])
            if not narr_match:
                break
            narr_start = idx + narr_match.start()
            brace_pos = narr_start + narr_match.group(0).rfind("{")
            content, next_pos = _extract_braced_content(body, brace_pos)
            cleaned = _latex_to_plain_text(content)
            if cleaned:
                chunks.append(cleaned)
            idx = next_pos
        narrations.append("\n\n".join(chunks).strip())

    return narrations


def compile_tex_to_pdf(tex_path: Path, out_pdf: Path) -> None:
    _require_tool("latexmk")
    cmd = [
        "latexmk",
        "-pdf",
        "-interaction=nonstopmode",
        "-halt-on-error",
        tex_path.name,
    ]
    subprocess.run(cmd, check=True, cwd=tex_path.parent, capture_output=True, text=True)
    built_pdf = tex_path.with_suffix(".pdf")
    if not built_pdf.exists():
        raise VideoGenerationError("LaTeX compile completed but PDF was not created.")
    shutil.copy2(built_pdf, out_pdf)


def pdf_to_pngs(pdf_path: Path, out_dir: Path, dpi: int = 220) -> List[Path]:
    _require_tool("pdftoppm")
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = out_dir / "slide"
    cmd = ["pdftoppm", "-r", str(dpi), "-png", str(pdf_path), str(prefix)]
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    images = sorted(out_dir.glob("slide-*.png"))
    if not images:
        raise VideoGenerationError("No slide images generated from PDF.")
    return images


async def _tts_to_file(text: str, output_path: Path, voice: str, rate: str) -> None:
    communicator = edge_tts.Communicate(text, voice=voice, rate=rate)
    await communicator.save(str(output_path))


def generate_audio(text: str, output_path: Path, voice: str, rate: str) -> None:
    asyncio.run(_tts_to_file(text, output_path, voice=voice, rate=rate))


def build_video(
    slide_images: List[Path],
    narrations: List[str],
    output_video: Path,
    voice: str,
    rate: str,
    fps: int,
    default_slide_seconds: float,
    temp_dir: Path,
) -> None:
    audio_dir = temp_dir / "audio"
    audio_dir.mkdir(parents=True, exist_ok=True)
    clips = []

    pair_count = min(len(slide_images), len(narrations))
    if pair_count == 0:
        raise VideoGenerationError("No slides or narrations were found.")

    for idx in range(pair_count):
        slide_path = slide_images[idx]
        narration = (narrations[idx] or "").strip()
        img_clip = ImageClip(str(slide_path))

        if narration:
            audio_path = audio_dir / f"slide_{idx+1}.mp3"
            generate_audio(narration, audio_path, voice=voice, rate=rate)
            audio_clip = AudioFileClip(str(audio_path))
            clip = img_clip.set_duration(audio_clip.duration).set_audio(audio_clip)
        else:
            clip = img_clip.set_duration(default_slide_seconds)

        clips.append(clip)

    final = concatenate_videoclips(clips, method="compose")
    final.write_videofile(str(output_video), fps=fps, codec="libx264", audio_codec="aac")

    final.close()
    for c in clips:
        c.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate MP4 from Beamer .tex and \\narration tags")
    parser.add_argument("--tex", required=True, help="Path to main .tex file")
    parser.add_argument("--output", required=True, help="Output .mp4 path")
    parser.add_argument("--voice", default="en-US-ChristopherNeural", help="edge-tts voice")
    parser.add_argument("--rate", default="+10%", help="edge-tts speaking rate")
    parser.add_argument("--dpi", type=int, default=220, help="Slide rasterization DPI")
    parser.add_argument("--fps", type=int, default=24, help="Output video FPS")
    parser.add_argument("--default-seconds", type=float, default=3.0, help="Duration for slides with empty narration")
    parser.add_argument("--keep-temp", action="store_true", help="Keep temp build artifacts")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    tex_path = Path(args.tex).resolve()
    output_video = Path(args.output).resolve()

    if not tex_path.exists():
        print(f"Error: tex file not found: {tex_path}", file=sys.stderr)
        return 1

    _require_tool("ffmpeg")

    temp_ctx = tempfile.TemporaryDirectory(prefix="beamer_video_")
    temp_root = Path(temp_ctx.name)

    try:
        out_pdf = temp_root / "slides.pdf"
        compile_tex_to_pdf(tex_path, out_pdf)

        slide_images = pdf_to_pngs(out_pdf, temp_root / "slides", dpi=args.dpi)
        narrations = extract_frame_narrations(tex_path)

        if not narrations:
            raise VideoGenerationError("No \\narration{...} entries found in frames.")

        build_video(
            slide_images=slide_images,
            narrations=narrations,
            output_video=output_video,
            voice=args.voice,
            rate=args.rate,
            fps=args.fps,
            default_slide_seconds=args.default_seconds,
            temp_dir=temp_root,
        )

        print(f"Video generated: {output_video}")
        return 0
    except subprocess.CalledProcessError as e:
        stderr = (e.stderr or "").strip()
        print(f"Error: command failed ({e.returncode}): {stderr[:2000]}", file=sys.stderr)
        return 2
    except VideoGenerationError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 3
    finally:
        if args.keep_temp:
            print(f"Kept temp dir: {temp_root}")
        else:
            temp_ctx.cleanup()


if __name__ == "__main__":
    raise SystemExit(main())
