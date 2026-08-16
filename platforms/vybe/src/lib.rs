//! `vybe_platform_vybe` — the `vybe:*` host surface (currently `vybe:gui`),
//! extracted from `vybe_host`.
//!
//! Holds the GUI implementation: control definitions (`controls`), the
//! `vybe:gui` host functions (`gui`), 2D drawing + canvas (`drawing`,
//! `canvas`), and the widget-backed `GuiState` bridge (`gui_state`).
//!
//! The `gui` feature gates everything that needs `vybe_widgets`. Without it,
//! the widget-free surface still compiles and registers (control `TypeDef`s,
//! `vybe:gui` drawing) — a headless build never links `vybe_widgets`.
//! `gui_state` (the live widget bridge) exists only under the feature.

pub mod builtin_types; // TypeRegistry vtables for the vybe:gui control surface; run in Plugin::finalize
// Canvas is NOT here, and this crate no longer paints for it either.
// `CanvasRenderingContext2D` is WHATWG HTML and lives in `platforms/web`
// (`web:canvas`); the engine behind it is now installed there too, resolving
// through the real `Document`. `canvas_backend_impl` — which resolved through
// `GuiState.form.controls` and so painted only into the capture-only overlay —
// is deleted.
// The control descriptions moved to `vybe_widgets::html`: the tag/CSS of a
// control is what a web engine consumes, so it belongs beside the DOM.
pub mod drawing;
pub mod gui;
#[cfg(feature = "gui")]
pub mod gui_state; // installs vybe_widgets as the `web:canvas` engine

// Input is NOT here. UI events are a web-platform concept and live in
// `platforms/web` (`web:ui-events`), where the queue is owned; SDL reaches
// them through its emitter, so this crate carries no input surface at all.

pub mod plugin;
pub use plugin::Plugin;

pub mod stubs;
#[cfg(feature = "gui")]
pub use plugin::init_platforms_with_gui;
pub use stubs::register_gui_stubs;
