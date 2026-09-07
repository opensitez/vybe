//! `tomllib` surface declarations.
//!
//! Constant `loads` / `load` inputs are normalized by the walker with the Rust
//! TOML parser. This module provides the real exception class and raise helper
//! so `except tomllib.TOMLDecodeError` uses ordinary class/error machinery.

use super::builders::*;
use vybe_ast::{Statement, StmtKind};

pub(super) fn toml_decode_error() -> Statement {
    class_extending("TOMLDecodeError", &["ValueError"], vec![])
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "__py_exc_TOMLDecodeError",
            vec![param("message", Some(str_lit("Invalid TOML")))],
            vec![ret(new("TOMLDecodeError", vec![ident("message")]))],
        ),
        function(
            "__py_raise_TOMLDecodeError",
            vec![param("message", Some(str_lit("Invalid TOML")))],
            vec![Statement::with_span(
                StmtKind::Throw {
                    expr: Some(call_global(
                        "__py_exc_TOMLDecodeError",
                        vec![ident("message")],
                    )),
                    cause: None,
                },
                span(),
            )],
        ),
    ]
}
