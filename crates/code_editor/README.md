cosmic_editor
===============

This is a scaffold for a pure‑Rust code editor prototype (keeps the webcore codebase separate).

Goals
- Use `cosmic-text` for text layout and shaping
- Use `tiny-skia` for raster rendering
- Use `ropey` as the text buffer
- Implement an incremental Rust tokenizer/highlighter and folding (no syntect, no non-Rust tooling)

How to proceed
1. Edit `Cargo.toml` and uncomment or pin dependency versions for `cosmic-text`, `tiny-skia`, `winit`, and `ropey`.
2. Implement an event loop in `src/main.rs` using `winit` and render to a `tiny-skia` surface.
3. Integrate `cosmic-text` to layout styled runs and draw them with `tiny-skia`.

If you want, I can now:
- Add concrete dependency versions and implement a first rendering pass that draws a static sample buffer, or
- Implement an incremental tokenizer and basic cursor/selection input handling.

Tell me which to implement next and I will proceed.

What I implemented now
- Added crate dependencies to `Cargo.toml` (cosmic-text, tiny-skia, winit, ropey)
- Added a simple pure-Rust `Editor` in `src/editor.rs` that uses `ropey` as the buffer and
	provides a small single-pass tokenizer suitable for highlighting/folding prototyping.
- `src/main.rs` contains a CLI demo that tokenizes a sample buffer, prints tokens, and performs
	a small edit + re-tokenize to demonstrate the incremental flow.

Next options (pick one):
- Wire `cosmic-text` + `tiny-skia` to render the tokenized styled runs in a `winit` window.
- Improve the tokenizer to be incremental per-line (fast re-highlighting on edits) and add
	folding heuristics (indent/braces).
- Implement cursor/selection input and basic key handling.

Which of these should I implement next? I can start with rendering or the incremental tokenizer.
