//! The bytes helper surface — `repr`, the str/bytes conversions, `bytes.join`
//! and the incremental codecs.
//!
//! ⛔ 55 lines of PARSED PYTHON that fired on 1368 corpus tests — more than any
//! other prelude — because the gate is any `b'` or `b"` in the source. Declared
//! here as AST instead: the gate is unchanged, but nothing is parsed.

use super::builders::*;
use vybe_ast::{
    BinOp, BindingPattern, Expression, IntOp, Statement, StmtKind, VarDeclKind, VarDeclarator,
};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

// Compiler-owned accumulators, escape pieces, chr results and hex digits
// are strings. They need the shared concat primitive, not Python's rich +.
// Repr iterates byte payloads, so these operands are integers in 0..255.
// Carry that storage-lane fact into the shared Wasm integer primitives.
fn byte_cmp(operation: vybe_ast::IntOp, l: Expr, r: Expr) -> Expr {
    op(BinOp::Integer(operation, vybe_ast::BitLane::W32), l, r)
}

fn add(l: Expr, r: Expr) -> Expr {
    call_global("__py_known_string_concat", vec![l, r])
}

/// `codecs.getincrementalencoder(...)()` — the pair only has to encode and
/// decode; the incremental state a real codec keeps is never observed here.
pub(super) fn incremental_encoder() -> Statement {
    class(
        "__PyIncrementalEncoder",
        vec![method(
            "encode",
            vec![
                param("value", Some(str_lit(""))),
                param("final", Some(bool_lit(false))),
            ],
            vec![ret(call_global(
                "bytes",
                vec![ident("value"), str_lit("utf-8")],
            ))],
        )],
    )
}

pub(super) fn incremental_decoder() -> Statement {
    class(
        "__PyIncrementalDecoder",
        vec![method(
            "decode",
            vec![
                param("value", Some(null())),
                param("final", Some(bool_lit(false))),
            ],
            vec![ret(call_global(
                "__vybe_bytes_decode",
                vec![ident("value")],
            ))],
        )],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // Escape data is module string data; each byte uses shared string primitives.
        function(
            "__vybe_bytes_repr",
            vec![param("a", Some(null()))],
            repr_body(),
        ),
        function(
            "__vybe_str_encode",
            vec![param("s", Some(str_lit("")))],
            vec![
                assign(ident("out"), list_of(vec![])),
                for_in(
                    "ch",
                    ident("s"),
                    vec![expr_stmt(call(
                        member(ident("out"), "append"),
                        vec![call_global("ord", vec![ident("ch")])],
                    ))],
                ),
                ret(ident("out")),
            ],
        ),
        function(
            "__vybe_bytes_decode",
            vec![param("a", Some(null()))],
            vec![
                assign(ident("r"), str_lit("")),
                for_in(
                    "b",
                    ident("a"),
                    vec![assign(
                        ident("r"),
                        add(ident("r"), call_global("chr", vec![ident("b")])),
                    )],
                ),
                ret(ident("r")),
            ],
        ),
        function(
            "__py_bytes_join",
            vec![param("sep", Some(null())), param("iterable", Some(null()))],
            vec![
                assign(
                    ident("s"),
                    call_global("__vybe_bytes_decode", vec![ident("sep")]),
                ),
                assign(ident("out"), str_lit("")),
                assign(ident("first"), bool_lit(true)),
                for_in(
                    "item",
                    ident("iterable"),
                    vec![
                        if_stmt(
                            unary_not(ident("first")),
                            vec![assign(ident("out"), add(ident("out"), ident("s")))],
                        ),
                        // `str(x)` of a str is that str; of a bytes it is the
                        // repr, so the lengths separate them without a type
                        // test — a spliced class never gets `isinstance`
                        // resolved.
                        if_stmt(
                            op(
                                BinOp::Eq,
                                call_global("len", vec![call_global("str", vec![ident("item")])]),
                                call_global("len", vec![ident("item")]),
                            ),
                            vec![assign(
                                ident("out"),
                                op(BinOp::Add, ident("out"), ident("item")),
                            )],
                        ),
                        if_stmt(
                            op(
                                BinOp::NotEq,
                                call_global("len", vec![call_global("str", vec![ident("item")])]),
                                call_global("len", vec![ident("item")]),
                            ),
                            vec![assign(
                                ident("out"),
                                add(
                                    ident("out"),
                                    call_global("__vybe_bytes_decode", vec![ident("item")]),
                                ),
                            )],
                        ),
                        assign(ident("first"), bool_lit(false)),
                    ],
                ),
                ret(call_global("bytes", vec![ident("out"), str_lit("utf-8")])),
            ],
        ),
    ]
}

fn local(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.into()),
            type_hint: None,
            init: Some(value),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}
fn repr_body() -> Vec<Statement> {
    let mut packed = String::new();
    let mut widths = String::new();
    for byte in 0u16..256 {
        let piece = match byte {
            9 => "\\t".to_string(),
            10 => "\\n".to_string(),
            13 => "\\r".to_string(),
            39 => "\\'".to_string(),
            92 => "\\\\".to_string(),
            32..=126 => char::from(byte as u8).to_string(),
            _ => format!("\\x{byte:02x}"),
        };
        assert!(piece.is_ascii() && (1..=4).contains(&piece.len()));
        widths.push(char::from(b'0' + piece.len() as u8));
        packed.push_str(&piece);
        for _ in piece.len()..4 {
            packed.push(' ');
        }
    }
    assert_eq!(packed.len(), 1024);
    assert_eq!(widths.len(), 256);
    let index = byte_cmp(IntOp::Mul, ident("b"), Expression::int(4));
    let length = byte_cmp(
        IntOp::Sub,
        call_global(
            "__py_known_byte_char_code",
            vec![ident("__byte_repr_widths"), ident("b")],
        ),
        Expression::int(48),
    );
    let end = byte_cmp(IntOp::Add, index.clone(), length);
    let piece = call_global(
        "__py_known_byte_substring",
        vec![ident("__byte_repr_packed"), index, end],
    );
    vec![
        local("__byte_repr_packed", str_lit(&packed)),
        local("__byte_repr_widths", str_lit(&widths)),
        assign(ident("r"), str_lit("b'")),
        for_in(
            "b",
            ident("a"),
            vec![if_else(
                byte_cmp(IntOp::Eq, ident("b"), Expression::int(9)),
                vec![assign(ident("r"), add(ident("r"), str_lit("\\t")))],
                vec![assign(ident("r"), add(ident("r"), piece))],
            )],
        ),
        ret(call_global(
            "__py_known_string_concat",
            vec![ident("r"), str_lit("'")],
        )),
    ]
}
