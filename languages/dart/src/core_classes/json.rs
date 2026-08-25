//! `dart:convert`'s `JsonDecoder` / `JsonEncoder`, as classes.
//!
//! Both are thin: `convert` is the corresponding builtin — `jsonDecode`
//! (`ecma:json.parse`) and `jsonEncode`/`__dart_json_stringify3`
//! (`ecma:json.stringify`, whose third argument IS the indent, §25.5.2).
//! `JsonEncoder.withIndent(i)` is a walker rewrite to a construction carrying
//! the indent; the plain `JsonEncoder()` ctor defaults it to null.

use super::builders::*;
use vybe_ast::{Argument, ExprKind, Expression, Statement};

fn call_builtin(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

pub(super) fn json_decoder() -> Statement {
    class(
        "JsonDecoder",
        vec![
            constructor(vec![], vec![]),
            method(
                "convert",
                vec![param("source", None, None)],
                Some("dynamic"),
                vec![ret(call_builtin("jsonDecode", vec![ident("source")]))],
            ),
        ],
    )
}

pub(super) fn json_encoder() -> Statement {
    let null_lit = Expression::null();
    class(
        "JsonEncoder",
        vec![
            field("_vybeIndent", "dynamic", Expression::null()),
            constructor(
                vec![param("indent", None, Some(null_lit))],
                vec![set_this("_vybeIndent", ident("indent"))],
            ),
            method(
                "convert",
                vec![param("value", None, None)],
                Some("String"),
                vec![
                    if_stmt(
                        binary(
                            vybe_ast::BinOp::Eq,
                            this_field("_vybeIndent"),
                            Expression::null(),
                        ),
                        vec![ret(call_builtin("jsonEncode", vec![ident("value")]))],
                    ),
                    ret(call_builtin(
                        "__dart_json_stringify3",
                        vec![ident("value"), Expression::null(), this_field("_vybeIndent")],
                    )),
                ],
            ),
        ],
    )
}
