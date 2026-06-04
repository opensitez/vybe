//! PHP language hub.
//!
//! This is the central crate-level namespace for PHP-specific code.
//! Frontend and emitter code live here directly; compiler-integrated
//! PHP helpers are sourced from this tree as well.

pub use crate::languages::php as frontend;
pub use crate::languages::php::emitter;
