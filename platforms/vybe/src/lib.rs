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
pub mod canvas;
pub mod controls;
pub mod drawing;
pub mod gui;

#[cfg(feature = "gui")]
pub mod gui_state;

pub mod plugin;
pub use plugin::Plugin;


pub mod stubs;
pub use stubs::register_gui_stubs;
