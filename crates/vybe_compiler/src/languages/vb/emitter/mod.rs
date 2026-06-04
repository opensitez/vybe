//! VB-specific `common:vb.*` emitter adapters.
//!
//! These helpers keep VB builtins in the frontend/compiler layer and lower
//! directly to portable bytecode plus standard `ecma:*` imports where WASM
//! lacks the required math primitives.

pub mod dispatch;
pub mod financial_adapter;
pub mod misc_adapter;
