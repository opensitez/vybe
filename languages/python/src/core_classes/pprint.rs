//! `pprint` — the pretty-printer, declared rather than parsed.
//!
//! ⛔ Every type test goes through `__py_container_kind`, the emitter adapter,
//! NOT `isinstance`. A declared class is never walked and a bare `list` is a
//! call-target profile row rather than a value, so `isinstance(o, list)`
//! answers False for everything. php reached the same conclusion for
//! `print_r`, which is an adapter (`emit_php_print_r`) and not a prelude.

use super::builders::*;
use vybe_ast::{BinOp, Param, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn add(l: Expr, r: Expr) -> Expr {
    op(BinOp::Add, l, r)
}

fn kind_of(e: Expr) -> Expr {
    call_global("__py_container_kind", vec![e])
}

/// Kind codes from `__py_container_kind`. ⛔ Integers: string `==` is a walker
/// rewrite and answers False inside a declaration.
/// ⛔ `str` is kind 1 and is NOT a container here: the original tested
/// `isinstance(o, (list, tuple, dict, set, frozenset))`. Treating any non-zero
/// kind as a container made `"x"` iterate into characters, and the first one
/// tripped the identity-based cycle check.
const LIST_OR_MORE: i64 = 2;

const TUPLE: i64 = 3;
const SET: i64 = 4;
const DICT: i64 = 5;

fn kind_is(e: Expr, k: i64) -> Expr {
    op(BinOp::Eq, kind_of(e), i(k))
}

/// The nine formatting parameters `__pprint_fmt` threads through.
fn fmt_params() -> Vec<Param> {
    vec![
        param("o", Some(null())),
        param("ind", Some(i(1))),
        param("width", Some(i(80))),
        param("depth", Some(null())),
        param("compact", Some(bool_lit(false))),
        param("sort_dicts", Some(bool_lit(true))),
        param("under", Some(bool_lit(false))),
        param("level", Some(i(0))),
        param("seen", Some(null())),
        param("col", Some(i(0))),
        // ⛔⛔ THE TEMPORARIES ARE PARAMETERS. An `assign` to a fresh name
        // inside a spliced function lands in a GLOBAL, not a frame local, so
        // `__pprint_fmt`'s recursive call overwrote the caller's `k`, `parts`
        // and `seen` — tuples came back as lists and sets came back empty.
        // A parameter IS a local, and callers never pass these.
        param("k", Some(i(0))),
        param("parts", Some(null())),
        param("pad", Some(str_lit(""))),
        param("keys", Some(null())),
    ]
}

fn fmt_args(o: Expr, level: Expr, seen: Expr, col: Expr) -> Vec<Expr> {
    vec![
        o,
        ident("ind"),
        ident("width"),
        ident("depth"),
        ident("compact"),
        ident("sort_dicts"),
        ident("under"),
        level,
        seen,
        col,
    ]
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "__pprint_is_container",
            vec![param("o", Some(null()))],
            vec![ret(op(BinOp::GtEq, kind_of(ident("o")), i(LIST_OR_MORE)))],
        ),
        // Identity, not equality: a container that merely COMPARES equal to an
        // ancestor is not a cycle.
        function(
            "__pprint_cycle",
            vec![param("o", Some(null())), param("seen", Some(null()))],
            vec![
                for_in(
                    "s",
                    ident("seen"),
                    vec![if_stmt(
                        call_global("__py_is__", vec![ident("s"), ident("o")]),
                        vec![ret(bool_lit(true))],
                    )],
                ),
                ret(bool_lit(false)),
            ],
        ),
        function(
            "__pprint_has_cycle",
            vec![param("o", Some(null())), param("seen", Some(null()))],
            vec![
                if_stmt(
                    op(BinOp::Lt, kind_of(ident("o")), i(LIST_OR_MORE)),
                    vec![ret(bool_lit(false))],
                ),
                if_stmt(
                    call_global("__pprint_cycle", vec![ident("o"), ident("seen")]),
                    vec![ret(bool_lit(true))],
                ),
                // ⛔ NOT `seen + [o]`. `+` on two lists lowers to numeric add
                // here and traps in `js-number.toF64`; the copy is what the
                // recursion actually needs anyway.
                assign(ident("seen"), call_global("list", vec![ident("seen")])),
                expr_stmt(call(member(ident("seen"), "append"), vec![ident("o")])),
                if_stmt(
                    kind_is(ident("o"), DICT),
                    vec![
                        for_in(
                            "k",
                            ident("o"),
                            vec![if_stmt(
                                call_global(
                                    "__pprint_has_cycle",
                                    vec![index(ident("o"), ident("k")), ident("seen")],
                                ),
                                vec![ret(bool_lit(true))],
                            )],
                        ),
                        ret(bool_lit(false)),
                    ],
                ),
                for_in(
                    "it",
                    ident("o"),
                    vec![if_stmt(
                        call_global("__pprint_has_cycle", vec![ident("it"), ident("seen")]),
                        vec![ret(bool_lit(true))],
                    )],
                ),
                ret(bool_lit(false)),
            ],
        ),
        function(
            "__pprint_kind",
            vec![param("o", Some(null()))],
            vec![ret(kind_of(ident("o")))],
        ),
        // `1234567` → `1_234_567`
        function(
            "__pprint_underscore",
            vec![param("n", Some(i(0)))],
            vec![
                assign(ident("s"), call_global("str", vec![ident("n")])),
                assign(
                    ident("neg"),
                    call(member(ident("s"), "startswith"), vec![str_lit("-")]),
                ),
                if_stmt(
                    ident("neg"),
                    vec![assign(ident("s"), slice_from(ident("s"), i(1)))],
                ),
                assign(ident("out"), str_lit("")),
                assign(ident("c"), i(0)),
                assign(
                    ident("i"),
                    op(BinOp::Sub, call_global("len", vec![ident("s")]), i(1)),
                ),
                while_stmt(
                    op(BinOp::GtEq, ident("i"), i(0)),
                    vec![
                        assign(
                            ident("out"),
                            add(index(ident("s"), ident("i")), ident("out")),
                        ),
                        assign(ident("c"), add(ident("c"), i(1))),
                        if_stmt(
                            binary(
                                BinOp::And,
                                op(BinOp::Eq, op(BinOp::Mod, ident("c"), i(3)), i(0)),
                                op(BinOp::Gt, ident("i"), i(0)),
                            ),
                            vec![assign(ident("out"), add(str_lit("_"), ident("out")))],
                        ),
                        assign(ident("i"), op(BinOp::Sub, ident("i"), i(1))),
                    ],
                ),
                if_stmt(
                    ident("neg"),
                    vec![assign(ident("out"), add(str_lit("-"), ident("out")))],
                ),
                ret(ident("out")),
            ],
        ),
        // `" " * n` — ⛔ NOT `BinOp::Mul` on a string: that is a walker
        // rewrite, and a declaration never gets one.
        function(
            "__pprint_pad",
            vec![param("n", Some(i(0)))],
            vec![
                // ⛔ `__pad_i`, not `k`. A local in a spliced declaration is
                // NOT isolated from its caller's: `__pprint_fmt` held the kind
                // in `k`, called this, and got the pad width back — every
                // container then formatted as a list.
                assign(ident("__pad_out"), str_lit("")),
                assign(ident("__pad_i"), i(0)),
                while_stmt(
                    op(BinOp::Lt, ident("__pad_i"), ident("n")),
                    vec![
                        assign(ident("__pad_out"), add(ident("__pad_out"), str_lit(" "))),
                        assign(ident("__pad_i"), add(ident("__pad_i"), i(1))),
                    ],
                ),
                ret(ident("__pad_out")),
            ],
        ),
        // `sep.join(parts)` — ⛔ a plain loop. `join` on a list binds to a USER
        // method when a program also imports `threading`, and resolves to
        // undefined; `socket` carries the same note.
        function(
            "__pprint_join",
            vec![
                param("parts", Some(null())),
                param("sep", Some(str_lit(", "))),
            ],
            vec![
                assign(ident("out"), str_lit("")),
                assign(ident("first"), bool_lit(true)),
                for_in(
                    "p",
                    ident("parts"),
                    vec![
                        if_stmt(
                            unary_not(ident("first")),
                            vec![assign(ident("out"), add(ident("out"), ident("sep")))],
                        ),
                        assign(ident("out"), add(ident("out"), ident("p"))),
                        assign(ident("first"), bool_lit(false)),
                    ],
                ),
                ret(ident("out")),
            ],
        ),
        function(
            "__pprint_wrap",
            vec![
                param("parts", Some(null())),
                param("open_c", Some(str_lit("["))),
                param("close_c", Some(str_lit("]"))),
                param("width", Some(i(80))),
                param("col", Some(i(0))),
                param("pad", Some(str_lit(""))),
                param("trail_comma", Some(bool_lit(false))),
            ],
            vec![
                assign(ident("tail"), str_lit("")),
                if_stmt(
                    ident("trail_comma"),
                    vec![assign(ident("tail"), str_lit(","))],
                ),
                assign(
                    ident("flat"),
                    add(
                        add(
                            ident("open_c"),
                            call_global("__pprint_join", vec![ident("parts"), str_lit(", ")]),
                        ),
                        add(ident("tail"), ident("close_c")),
                    ),
                ),
                if_stmt(
                    binary(
                        BinOp::Or,
                        op(
                            BinOp::LtEq,
                            add(ident("col"), call_global("len", vec![ident("flat")])),
                            ident("width"),
                        ),
                        op(BinOp::LtEq, call_global("len", vec![ident("parts")]), i(1)),
                    ),
                    vec![ret(ident("flat"))],
                ),
                ret(add(
                    add(
                        ident("open_c"),
                        call_global(
                            "__pprint_join",
                            vec![
                                ident("parts"),
                                add(
                                    add(str_lit(","), call_global("chr", vec![i(10)])),
                                    ident("pad"),
                                ),
                            ],
                        ),
                    ),
                    ident("close_c"),
                )),
            ],
        ),
        function(
            "__pprint_fmt",
            fmt_params(),
            vec![
                if_stmt(
                    is_none(ident("seen")),
                    vec![assign(ident("seen"), list_of(vec![]))],
                ),
                assign(ident("k"), kind_of(ident("o"))),
                if_stmt(
                    op(BinOp::GtEq, ident("k"), i(LIST_OR_MORE)),
                    vec![
                        if_stmt(
                            call_global("__pprint_cycle", vec![ident("o"), ident("seen")]),
                            vec![ret(str_lit("<Recursion on list with id=0>"))],
                        ),
                        // ⛔ NESTED, not `And`: the operands are both
                        // evaluated, so `level >= depth` ran with `depth` None
                        // and trapped in `js-number.toF64`.
                        if_stmt(
                            is_not_none(ident("depth")),
                            vec![if_stmt(
                                op(BinOp::GtEq, ident("level"), ident("depth")),
                                vec![ret(str_lit("..."))],
                            )],
                        ),
                    ],
                ),
                if_stmt(
                    op(BinOp::Lt, ident("k"), i(LIST_OR_MORE)),
                    vec![ret(call_global("repr", vec![ident("o")]))],
                ),
                // ⛔ NOT `seen + [o]`. `+` on two lists lowers to numeric add
                // here and traps in `js-number.toF64`; the copy is what the
                // recursion actually needs anyway.
                assign(ident("seen"), call_global("list", vec![ident("seen")])),
                expr_stmt(call(member(ident("seen"), "append"), vec![ident("o")])),
                assign(
                    ident("pad"),
                    call_global("__pprint_pad", vec![add(ident("col"), ident("ind"))]),
                ),
                assign(ident("parts"), list_of(vec![])),
                if_stmt(
                    op(BinOp::Eq, ident("k"), i(DICT)),
                    vec![
                        assign(
                            ident("keys"),
                            call_global("list", vec![call(member(ident("o"), "keys"), vec![])]),
                        ),
                        if_stmt(
                            ident("sort_dicts"),
                            vec![assign(
                                ident("keys"),
                                call_global("sorted", vec![ident("keys")]),
                            )],
                        ),
                        for_in(
                            "dk",
                            ident("keys"),
                            vec![expr_stmt(call(
                                member(ident("parts"), "append"),
                                vec![add(
                                    add(call_global("repr", vec![ident("dk")]), str_lit(": ")),
                                    call_global(
                                        "__pprint_fmt",
                                        fmt_args(
                                            index(ident("o"), ident("dk")),
                                            add(ident("level"), i(1)),
                                            ident("seen"),
                                            add(ident("col"), ident("ind")),
                                        ),
                                    ),
                                )],
                            ))],
                        ),
                        ret(call_global(
                            "__pprint_wrap",
                            vec![
                                ident("parts"),
                                str_lit("{"),
                                str_lit("}"),
                                ident("width"),
                                ident("col"),
                                ident("pad"),
                                bool_lit(false),
                            ],
                        )),
                    ],
                ),
                if_stmt(
                    op(BinOp::Eq, ident("k"), i(SET)),
                    vec![
                        for_in(
                            "it",
                            // `sorted(set)` yields nothing inside a
                            // declaration; materialising the set first does.
                            call_global("sorted", vec![call_global("list", vec![ident("o")])]),
                            vec![expr_stmt(call(
                                member(ident("parts"), "append"),
                                vec![call_global(
                                    "__pprint_fmt",
                                    fmt_args(
                                        ident("it"),
                                        add(ident("level"), i(1)),
                                        ident("seen"),
                                        add(ident("col"), ident("ind")),
                                    ),
                                )],
                            ))],
                        ),
                        if_stmt(
                            op(BinOp::Eq, call_global("len", vec![ident("parts")]), i(0)),
                            vec![ret(str_lit("set()"))],
                        ),
                        ret(call_global(
                            "__pprint_wrap",
                            vec![
                                ident("parts"),
                                str_lit("{"),
                                str_lit("}"),
                                ident("width"),
                                ident("col"),
                                ident("pad"),
                                bool_lit(false),
                            ],
                        )),
                    ],
                ),
                for_in(
                    "it",
                    ident("o"),
                    vec![expr_stmt(call(
                        member(ident("parts"), "append"),
                        vec![call_global(
                            "__pprint_fmt",
                            fmt_args(
                                ident("it"),
                                add(ident("level"), i(1)),
                                ident("seen"),
                                add(ident("col"), ident("ind")),
                            ),
                        )],
                    ))],
                ),
                if_stmt(
                    op(BinOp::Eq, ident("k"), i(TUPLE)),
                    vec![ret(call_global(
                        "__pprint_wrap",
                        vec![
                            ident("parts"),
                            str_lit("("),
                            str_lit(")"),
                            ident("width"),
                            ident("col"),
                            ident("pad"),
                            op(BinOp::Eq, call_global("len", vec![ident("parts")]), i(1)),
                        ],
                    ))],
                ),
                ret(call_global(
                    "__pprint_wrap",
                    vec![
                        ident("parts"),
                        str_lit("["),
                        str_lit("]"),
                        ident("width"),
                        ident("col"),
                        ident("pad"),
                        bool_lit(false),
                    ],
                )),
            ],
        ),
        function(
            "__pprint_pformat",
            vec![
                param("o", Some(null())),
                param("indent", Some(i(1))),
                param("width", Some(i(80))),
                param("depth", Some(null())),
                param("compact", Some(bool_lit(false))),
                param("sort_dicts", Some(bool_lit(true))),
                param("underscore_numbers", Some(bool_lit(false))),
            ],
            vec![ret(call_global(
                "__pprint_fmt",
                vec![
                    ident("o"),
                    ident("indent"),
                    ident("width"),
                    ident("depth"),
                    ident("compact"),
                    ident("sort_dicts"),
                    ident("underscore_numbers"),
                    i(0),
                    list_of(vec![]),
                    i(0),
                ],
            ))],
        ),
        function(
            "__pprint_pprint",
            vec![
                param("o", Some(null())),
                param("stream", Some(null())),
                param("indent", Some(i(1))),
                param("width", Some(i(80))),
                param("depth", Some(null())),
                param("compact", Some(bool_lit(false))),
                param("sort_dicts", Some(bool_lit(true))),
                param("underscore_numbers", Some(bool_lit(false))),
            ],
            vec![
                assign(
                    ident("text"),
                    call_global(
                        "__pprint_pformat",
                        vec![
                            ident("o"),
                            ident("indent"),
                            ident("width"),
                            ident("depth"),
                            ident("compact"),
                            ident("sort_dicts"),
                            ident("underscore_numbers"),
                        ],
                    ),
                ),
                if_stmt(
                    is_none(ident("stream")),
                    // ⛔ `print` takes (sep, end, *items) — the WALKER inserts
                    // the first two, and a declaration never gets that
                    // rewrite. Calling it with one argument panicked the
                    // adapter on `slots[1]`.
                    vec![expr_stmt(call_global(
                        "print",
                        vec![str_lit(" "), call_global("chr", vec![i(10)]), ident("text")],
                    ))],
                ),
                if_stmt(
                    is_not_none(ident("stream")),
                    vec![expr_stmt(call(
                        member(ident("stream"), "write"),
                        vec![add(ident("text"), call_global("chr", vec![i(10)]))],
                    ))],
                ),
            ],
        ),
        function(
            "__pprint_pp",
            vec![param("o", Some(null())), param("stream", Some(null()))],
            vec![expr_stmt(call_global(
                "__pprint_pprint",
                vec![ident("o"), ident("stream")],
            ))],
        ),
        function(
            "__pprint_saferepr",
            vec![param("o", Some(null()))],
            vec![ret(call_global(
                "__pprint_fmt",
                vec![
                    ident("o"),
                    i(1),
                    i(1000000),
                    null(),
                    bool_lit(false),
                    bool_lit(true),
                    bool_lit(false),
                    i(0),
                    list_of(vec![]),
                    i(0),
                ],
            ))],
        ),
        function(
            "__pprint_isrecursive",
            vec![param("o", Some(null()))],
            vec![ret(call_global(
                "__pprint_has_cycle",
                vec![ident("o"), list_of(vec![])],
            ))],
        ),
        function(
            "__pprint_isreadable",
            vec![param("o", Some(null()))],
            vec![ret(unary_not(call_global(
                "__pprint_has_cycle",
                vec![ident("o"), list_of(vec![])],
            )))],
        ),
    ]
}

pub(super) fn pretty_printer() -> Statement {
    class(
        "__pprint_PrettyPrinter",
        vec![
            init(
                vec![
                    param("indent", Some(i(1))),
                    param("width", Some(i(80))),
                    param("depth", Some(null())),
                    param("stream", Some(null())),
                    param("compact", Some(bool_lit(false))),
                    param("sort_dicts", Some(bool_lit(true))),
                    param("underscore_numbers", Some(bool_lit(false))),
                ],
                vec![
                    set_this("indent", ident("indent")),
                    set_this("width", ident("width")),
                    set_this("depth", ident("depth")),
                    set_this("stream", ident("stream")),
                    set_this("compact", ident("compact")),
                    set_this("sort_dicts", ident("sort_dicts")),
                    set_this("underscore_numbers", ident("underscore_numbers")),
                ],
            ),
            method(
                "pformat",
                vec![param("o", Some(null()))],
                vec![ret(call_global(
                    "__pprint_pformat",
                    vec![
                        ident("o"),
                        this_field("indent"),
                        this_field("width"),
                        this_field("depth"),
                        this_field("compact"),
                        this_field("sort_dicts"),
                        this_field("underscore_numbers"),
                    ],
                ))],
            ),
            method(
                "pprint",
                vec![param("o", Some(null()))],
                vec![expr_stmt(call_global(
                    "__pprint_pprint",
                    vec![
                        ident("o"),
                        this_field("stream"),
                        this_field("indent"),
                        this_field("width"),
                        this_field("depth"),
                        this_field("compact"),
                        this_field("sort_dicts"),
                        this_field("underscore_numbers"),
                    ],
                ))],
            ),
            method(
                "isrecursive",
                vec![param("o", Some(null()))],
                vec![ret(call_global("__pprint_isrecursive", vec![ident("o")]))],
            ),
            method(
                "isreadable",
                vec![param("o", Some(null()))],
                vec![ret(call_global("__pprint_isreadable", vec![ident("o")]))],
            ),
        ],
    )
}
