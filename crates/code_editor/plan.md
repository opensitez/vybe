# cosmic_editor Implementation Plan

Summary
- Pure‑Rust native code editor prototype.
- GPU rendering via `wgpu` + `wgpu_glyph`.
- `cosmic-text` for shaping/layout where needed.
- `ropey` for the text buffer and incremental edits.

Milestones
1. Workspace & repo
   - Ensure `Cargo.toml` dependencies are finalized.

2. Core buffer
   - Implement `EditorBuffer` using `ropey`.
   - Provide API: insert, delete, slice, undo/redo (simple stack-based history).
   - Expose line/char/byte conversions.

3. Tokenizer / highlighter
   - Design incremental per-line tokenizer API.
   - Implement lightweight language-agnostic tokenizer (comments, strings, numbers, ids, punct).
   - Provide token-change callbacks for renderer to re-highlight minimal ranges.

4. Folding
   - Implement brace-based folding and optional indent-based folding.
   - Provide fold/unfold API and persistence.

5. Shaping & glyphs
   - Integrate `cosmic-text` where high-quality shaping/fallback is required.
   - Or use `ab_glyph`/`wgpu_glyph` path as primary for the first GPU pipeline.
   - Implement system font discovery and a bundled fallback to ensure runnable demos.

6. GPU renderer (wgpu)
   - Initialize `wgpu` + `winit` surface and swapchain.
   - Use `wgpu_glyph` (or a custom glyph-atlas) to render laid-out text.
   - Implement glyph atlas caching and texture updates to minimize uploads.
   - Implement HiDPI scaling.

7. Editor UI wiring
   - Map tokens -> styled glyph runs (colors, font weight, underline).
   - Render caret, selection, line numbers, gutter, folding widgets.
   - Implement scrolling and viewport management.

8. Input & interaction
   - Keyboard: text input, navigation, selection, shortcuts (copy/paste, undo/redo).
   - Mouse: click-to-position, drag-selection, double/triple click, gutter clicks for folds.
   - Clipboard integration.

9. Language features
   - Simple language plugins: provide syntax heuristics via tokenizer.
   - Optional LSP integration later (separate milestone).

10. Tests & perf
   - Unit tests for buffer and tokenizer.
   - Rendering smoke tests.
   - Profile and optimize glyph caching and re-layout hotspots.

11. Docs & examples
   - README with build/run instructions.
   - Example app: `code_editor --gui` running sample buffer.

Decisions & Notes
- Rendering backend: `wgpu` chosen for cross-platform GPU path.
- `wgpu_glyph` can be used to accelerate text rendering; later we can replace with a custom atlas if needed.
- Use system monospace font where available; include a bundled fallback font for demos.

Immediate next steps (this session)
- Add the crate to the workspace `Cargo.toml` at repo root.
- Finalize `cosmic_editor/Cargo.toml` dependency versions.
- Implement `EditorBuffer` basics and unit tests.

If this looks good I will: (A) add `cosmic_editor` to the workspace manifest and (B) implement the `EditorBuffer` skeleton next.
