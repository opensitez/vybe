//! COBOL-specific emitter adapters.
//!
//! These adapters keep COBOL-only execution semantics in inline bytecode
//! and are reached through `common:cobol.*` dispatcher entries from the
//! COBOL profile.

pub mod dispatch;
pub mod arithmetic;
pub mod control;
pub mod data;
pub mod date;
pub mod files;

mod support;
