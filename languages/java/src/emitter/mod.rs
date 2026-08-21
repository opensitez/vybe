//! Java adapter surface — RETIRED except the runtime prelude.
//!
//! Every `common:java.*` emitter moved to `platforms/jvm` and answers as
//! `common:jvm.java.*`: the JDK is platform surface, reached through the
//! common tree resolver, so this crate owns no dispatch and no emit bodies.
//! What remains is the `java.util.Formatter`/`PrintStream` runtime prelude
//! (common-AST construction, prepended by the walker), pending its own move.

pub mod format_runtime;
