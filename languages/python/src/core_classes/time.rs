//! `time` — the one name the concurrency prelude owned.
//!
//! `time.sleep` is rewritten by the walker (`walker.rs:14694`) to the BARE
//! global `sleep`, so that global has to exist whenever a program imports
//! `time`. It lived in `CONCURRENCY_PRELUDE`, whose gate included
//! `source.contains("import time")` — which is why deleting that prelude broke
//! seven `time_module` / `time_extended` tests that have nothing to do with
//! concurrency.
//!
//! It does not sleep. There is one thread and a real sleep would only stall it;
//! `time.sleep(n)` answering immediately is what the prelude did too.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn module_functions() -> Vec<Statement> {
    vec![function(
        "sleep",
        vec![param("seconds", Some(num(0.0)))],
        vec![ret(null())],
    )]
}
