//! `traceback` — the formatting surface.
//!
//! Content that needs the live exception is best-effort: `sys.exc_info()` is
//! not fully populated, so `format_exc` answers the header CPython starts with
//! rather than a real stack. That was true of the prelude too; what changes is
//! that these are real classes and real globals rather than parsed source.

use super::builders::*;
use vybe_ast::Statement;

const HEADER: &str = "Traceback (most recent call last):\n";
const FILE_LINE: &str = "  File \"<unknown>\"\n";

pub(super) fn frame_summary() -> Statement {
    class("FrameSummary", vec![init(any_args(), vec![])])
}

/// `StackSummary` is a `list` subclass in CPython, and the corpus indexes it.
pub(super) fn stack_summary() -> Statement {
    class_extending("StackSummary", &["list"], vec![])
}

pub(super) fn traceback_exception() -> Statement {
    class(
        "TracebackException",
        vec![
            init(any_args(), vec![]),
            stub("format", list_of(vec![str_lit(HEADER)])),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        stub_fn("format_exc", str_lit(HEADER)),
        stub_fn("format_exception", list_of(vec![str_lit(HEADER)])),
        stub_fn("format_exception_only", list_of(vec![str_lit("\n")])),
        stub_fn("format_tb", list_of(vec![str_lit(FILE_LINE)])),
        stub_fn("format_stack", list_of(vec![str_lit(FILE_LINE)])),
        stub_fn("extract_tb", new("StackSummary", vec![])),
        stub_fn("extract_stack", new("StackSummary", vec![])),
        stub_fn("print_exc", null()),
        stub_fn("print_tb", null()),
        stub_fn("print_stack", null()),
        stub_fn("print_exception", null()),
        stub_fn("clear_frames", null()),
        stub_fn("walk_tb", list_of(vec![])),
        stub_fn("walk_stack", list_of(vec![])),
    ]
}
