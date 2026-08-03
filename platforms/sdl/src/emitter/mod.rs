//! SDL-side compiler emit adapter.
//!
//! This module is the adapter layer that maps `common:sdl.*` emissions into
//! existing `vybe:gui` host calls. It keeps the C-facing SDL API surface in
//! place while routing runtime behavior through the GUI adapter.

pub mod dispatch;
pub mod sdl;
