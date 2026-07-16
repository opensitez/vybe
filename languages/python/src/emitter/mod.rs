//! Python-specific emitter adapters.
//!
//! These adapters keep Python surface semantics in bytecode without
//! introducing Python-only host imports. They are routed via
//! `common:python.*` from the Python profile.

pub mod array_adapter;
pub mod bisect_adapter;
pub mod collections_adapter;
pub mod datetime_adapter;
pub mod dispatch;
pub mod float_adapter;
pub mod heapq_adapter;
pub mod itertools_adapter;
pub mod repr_adapter;
pub mod runtime_adapter;
pub mod statistics_adapter;
pub mod time_adapter;
