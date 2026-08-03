//! SDL adapter platform surface.
//!
//! This crate is intentionally thin: it contributes an `sdl` namespace that
//! aliases SDL-style names to existing Vybe runtime host surfaces (primarily
//! `vybe:gui`). There is no SDL runtime or host-function implementation in this
//! crate.

pub mod emitter;
pub mod plugin;
pub mod tree_register;

pub use plugin::{register, Plugin};
