//! Python-specific emitter adapters.
//!
//! These adapters keep Python surface semantics in bytecode without
//! introducing Python-only host imports. They are routed via
//! `common:python.*` from the Python profile.

pub mod collections_adapter;
pub mod dispatch;
pub mod float_adapter;
pub mod runtime_adapter;
