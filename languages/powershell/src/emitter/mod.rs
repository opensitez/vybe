//! PowerShell-local emit.
//!
//! Deliberately tiny. Everything PowerShell needs that a shared primitive can
//! express is bound through the profile (`[builtins]`, `[value_methods]`,
//! `[builtin_slots.*]`) and lowered by `vybe_compiler::primitives`. This module
//! exists only for the surface where no shared primitive can answer the
//! question, which `documentation/powershellplan.md` §6.5 gates on a proven gap.

pub mod dispatch;
pub mod display;
pub mod json;
pub mod operators;
