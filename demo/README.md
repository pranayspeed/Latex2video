# Demo Files

## 1) Basic Beamer Demo

- File: `basic_demo.tex`
- Purpose: Minimal 4-slide academic Beamer example with `\\narration{...}` on each slide.
- Includes image: `assets/research_chart.png`
- Ready upload zip: `basic_demo_upload.zip`

## 2) Full Pipeline Demo (LaTeX -> PPTX -> Video)

- File: `latex_to_pptx_to_video_demo.tex`
- Purpose: Demonstrates the exact two-app flow in this repository.
- Includes image: `assets/academic_workflow.png`
- Ready upload zip: `latex_to_pptx_to_video_demo_upload.zip`

### How to test with both apps

1. Open `http://localhost:8000` (`app.py`) and upload one of the ready zip files above.
2. Choose mode: `High-res images to slides (PPTX only + notes)`.
3. Download the generated `.pptx`.
4. Open `http://localhost:8001` (`pptx_video_app.py`) and upload that `.pptx`.
5. Download the narrated `.mp4`.
