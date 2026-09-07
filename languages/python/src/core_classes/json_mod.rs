//! `json` module classes declared as AST, not source prelude.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn json_decoder() -> Statement {
    class(
        "JSONDecoder",
        vec![
            init(any_args(), vec![]),
            method(
                "decode",
                vec![param("s", None)],
                vec![ret(call_global("__py_json_loads", vec![ident("s")]))],
            ),
            method(
                "raw_decode",
                vec![param("s", None), param("idx", Some(num(0.0)))],
                vec![ret(tuple_of(vec![
                    call_global("__py_json_loads", vec![ident("s")]),
                    call_global("len", vec![ident("s")]),
                ]))],
            ),
        ],
    )
}

pub(super) fn json_encoder() -> Statement {
    class(
        "JSONEncoder",
        vec![
            init(any_args(), vec![]),
            method(
                "default",
                vec![param("o", None)],
                vec![ret(call_global("str", vec![ident("o")]))],
            ),
            method(
                "encode",
                vec![param("o", None)],
                vec![ret(call_global("__py_json_dumps", vec![
                    ident("o"),
                    null(),
                    bool_lit(false),
                    null(),
                    str_lit(", "),
                    str_lit(": "),
                ]))],
            ),
        ],
    )
}
