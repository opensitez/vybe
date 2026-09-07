//! `object()` — a bare instance, useful only for identity.
//!
//! CPython's real `object` carries the default dunder machinery; the corpus
//! only ever constructs one as a unique sentinel (`a is a`, `a is not b`), so
//! an empty class satisfies every observed use without over-building.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn bare_object() -> Statement {
    class("__PyObject", vec![init(vec![], vec![])])
}
