//! `string` module classes declared as AST.

use super::builders::*;
use vybe_ast::{Expression, Statement};

pub(super) fn template() -> Statement {
    class(
        "__string_Template",
        vec![
            static_field("delimiter", str_lit("$")),
            init(
                vec![param("template", None)],
                vec![set_this("template", ident("template"))],
            ),
            method(
                "substitute",
                vec![param("mapping", Some(null())), kwargs_param("kws")],
                vec![ret(null())],
            ),
            method(
                "safe_substitute",
                vec![param("mapping", Some(null())), kwargs_param("kws")],
                vec![ret(null())],
            ),
            method("get_identifiers", vec![], vec![ret(list_of(vec![]))]),
            method("is_valid", vec![], vec![ret(Expression::bool(true))]),
        ],
    )
}

pub(super) fn formatter() -> Statement {
    class(
        "__string_Formatter",
        vec![
            init(vec![], vec![]),
            method(
                "format",
                vec![param("fmt", None), rest_param("args"), kwargs_param("kwargs")],
                vec![ret(null())],
            ),
            method(
                "vformat",
                vec![param("fmt", None), param("args", None), param("kwargs", None)],
                vec![ret(null())],
            ),
        ],
    )
}
