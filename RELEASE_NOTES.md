## JapanMoto Release Notes

### What is included
- Windows-only contract generation through original DOGOVIR_6055_template.doc (Word COM) for 1:1 formatting.
- Contract fallback pipeline when Word COM is unavailable:
  - Try LibreOffice/Office CLI to generate legacy .doc from template.
  - If .doc is unavailable, generate template-preserving .docx.
  - Last resort: text-built .docx fallback.
- Contract clause 2.1 now writes amount in words in brackets without currency text, then adds "гривень 00 копійок" outside brackets.
- Improved desktop shortcut creation in build script with fallback to Public Desktop.
- Optional custom app icon support in build script:
  - assets/japanmoto.ico
  - japanmoto.ico
- Autocomplete in empty fields now shows last 4 recent values.
- XLS writing preserves original cell formatting (fonts/styles).
- Automatic UI theme mode based on Windows app theme (light/dark).
- Settings window improvements:
  - Opens centered on screen.
  - Increased height by 5% to avoid clipped bottom actions.
  - Theme switch in settings: Auto / Light / Dark.

### Build info
This release is built automatically by GitHub Actions.
