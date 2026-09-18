//! `zoneinfo` surface declarations.

use super::builders::*;
use vybe_ast::{Statement, StmtKind};

pub(super) fn zoneinfo_not_found_error() -> Statement {
    class_extending("ZoneInfoNotFoundError", &["KeyError"], vec![])
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "__py_exc_ZoneInfoNotFoundError",
            vec![param("message", Some(str_lit("ZoneInfo key not found")))],
            vec![ret(new("ZoneInfoNotFoundError", vec![ident("message")]))],
        ),
        function(
            "__py_raise_ZoneInfoNotFoundError",
            vec![param("message", Some(str_lit("ZoneInfo key not found")))],
            vec![Statement::with_span(
                StmtKind::Throw {
                    expr: Some(call_global(
                        "__py_exc_ZoneInfoNotFoundError",
                        vec![ident("message")],
                    )),
                    cause: None,
                },
                span(),
            )],
        ),
    ]
}
