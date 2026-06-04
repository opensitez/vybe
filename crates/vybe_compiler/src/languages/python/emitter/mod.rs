//! Python-specific emitter adapters.
//!
//! These adapters keep Python surface semantics in bytecode without
//! introducing Python-only host imports. They are routed via
//! `common:python.*` from the Python profile.

pub mod dispatch;
pub mod collections_adapter;
