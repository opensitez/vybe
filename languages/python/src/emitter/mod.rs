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
pub mod hash_adapter;
pub mod heapq_adapter;
pub mod itertools_adapter;
pub mod json_adapter;
pub mod math_adapter;
pub mod os_path_adapter;
pub mod repr_adapter;
pub mod re_adapter;
pub mod runtime_adapter;
pub mod sql_adapter;
pub mod statistics_adapter;
pub mod struct_adapter;
pub mod string_adapter;
pub mod time_adapter;
pub mod tree_register;
