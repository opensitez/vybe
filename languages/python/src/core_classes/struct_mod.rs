//! `struct` — pack/unpack driven by a PARSE of the format string.
//!
//! ⛔ The adapter matched the WHOLE format against literals (`"i"`, `"@i"`,
//! `"h"`, `"<H"`, `"iii"`), so `">i"` — the ordinary spelling — never matched
//! and returned empty. Format codes cross byte orders, repeat counts and
//! widths; that is a parse, not a lookup table, and enumerating the product was
//! never going to terminate.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

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

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

fn push(target: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(target, "append"), vec![value]))
}

/// `"bB?x"` → 1, `"hH"` → 2, `"iIlLf"` → 4, `"qQd"` → 8.
fn size_fn() -> Statement {
    let mut body = vec![];
    for (codes, size) in [("bB?xcs", 1i64), ("hH", 2), ("iIlLf", 4), ("qQdP", 8)] {
        body.push(if_stmt(
            op(
                BinOp::GtEq,
                call(member(str_lit(codes), "find"), vec![ident("code")]),
                i(0),
            ),
            vec![ret(i(size))],
        ));
    }
    body.push(ret(i(0)));
    function(
        "__py_struct_size",
        vec![param("code", Some(str_lit("")))],
        body,
    )
}

/// Split a format into `[byte_order_is_big, [[code, count], ...]]`.
fn parse_fn() -> Statement {
    function(
        "__py_struct_parse",
        vec![param("fmt", Some(str_lit("")))],
        vec![
            assign(ident("__pr_big"), bool_lit(false)),
            assign(ident("__pr_start"), i(0)),
            assign(ident("__pr_lead"), slice_range(ident("fmt"), i(0), i(1))),
            if_stmt(
                op(
                    BinOp::GtEq,
                    call(member(str_lit("<>!=@"), "find"), vec![ident("__pr_lead")]),
                    i(0),
                ),
                vec![
                    assign(ident("__pr_start"), i(1)),
                    if_stmt(
                        op(
                            BinOp::GtEq,
                            call(member(str_lit(">!"), "find"), vec![ident("__pr_lead")]),
                            i(0),
                        ),
                        vec![assign(ident("__pr_big"), bool_lit(true))],
                    ),
                ],
            ),
            assign(ident("__pr_items"), list_of(vec![])),
            assign(ident("__pr_k"), ident("__pr_start")),
            assign(ident("__pr_num"), str_lit("")),
            while_stmt(
                op(BinOp::Lt, ident("__pr_k"), len_of(ident("fmt"))),
                vec![
                    assign(ident("__pr_ch"), index(ident("fmt"), ident("__pr_k"))),
                    if_stmt(
                        op(
                            BinOp::GtEq,
                            call(
                                member(str_lit("0123456789"), "find"),
                                vec![ident("__pr_ch")],
                            ),
                            i(0),
                        ),
                        vec![assign(
                            ident("__pr_num"),
                            add(ident("__pr_num"), ident("__pr_ch")),
                        )],
                    ),
                    if_stmt(
                        op(
                            BinOp::Lt,
                            call(
                                member(str_lit("0123456789"), "find"),
                                vec![ident("__pr_ch")],
                            ),
                            i(0),
                        ),
                        vec![
                            assign(ident("__pr_cnt"), i(1)),
                            if_stmt(
                                op(BinOp::Gt, len_of(ident("__pr_num")), i(0)),
                                vec![assign(
                                    ident("__pr_cnt"),
                                    call_global("int", vec![ident("__pr_num")]),
                                )],
                            ),
                            push(
                                ident("__pr_items"),
                                list_of(vec![ident("__pr_ch"), ident("__pr_cnt")]),
                            ),
                            assign(ident("__pr_num"), str_lit("")),
                        ],
                    ),
                    assign(ident("__pr_k"), add(ident("__pr_k"), i(1))),
                ],
            ),
            ret(list_of(vec![ident("__pr_big"), ident("__pr_items")])),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        size_fn(),
        parse_fn(),
        // One integer → `size` bytes, two's complement for negatives.
        function(
            "__py_struct_put",
            vec![
                param("value", Some(i(0))),
                param("size", Some(i(1))),
                param("big", Some(bool_lit(false))),
            ],
            vec![
                assign(ident("__put_v"), call_global("int", vec![ident("value")])),
                if_stmt(
                    op(BinOp::Lt, ident("__put_v"), i(0)),
                    vec![assign(
                        ident("__put_v"),
                        add(ident("__put_v"), op(BinOp::Pow, i(256), ident("size"))),
                    )],
                ),
                // ⛔ A LIST OF BYTE VALUES, never text. Going through
                // `chr()` + `bytes(s, "utf-8")` re-encodes anything above
                // U+007F, so `pack("<H", 255)` came back THREE bytes.
                assign(ident("__put_out"), list_of(vec![])),
                assign(ident("__put_n"), i(0)),
                while_stmt(
                    op(BinOp::Lt, ident("__put_n"), ident("size")),
                    vec![
                        assign(
                            ident("__put_byte"),
                            op(BinOp::Mod, ident("__put_v"), i(256)),
                        ),
                        push(ident("__put_out"), ident("__put_byte")),
                        assign(
                            ident("__put_v"),
                            op(BinOp::FloorDiv, ident("__put_v"), i(256)),
                        ),
                        assign(ident("__put_n"), add(ident("__put_n"), i(1))),
                    ],
                ),
                // ⛔ Built LITTLE-ENDIAN by append, then reversed for big.
                // `[b] + out` does not concatenate inside a declaration — `+`
                // on two lists lowers to numeric add and traps, the same trap
                // `pprint`'s `seen + [o]` hit.
                if_stmt(unary_not(ident("__pr_big")), vec![ret(ident("__put_out"))]),
                assign(ident("__put_rev"), list_of(vec![])),
                assign(
                    ident("__put_m"),
                    op(BinOp::Sub, len_of(ident("__put_out")), i(1)),
                ),
                while_stmt(
                    op(BinOp::GtEq, ident("__put_m"), i(0)),
                    vec![
                        push(
                            ident("__put_rev"),
                            index(ident("__put_out"), ident("__put_m")),
                        ),
                        assign(ident("__put_m"), op(BinOp::Sub, ident("__put_m"), i(1))),
                    ],
                ),
                ret(ident("__put_rev")),
            ],
        ),
        // `size` bytes → one integer, signed when the code is lower-case.
        function(
            "__py_struct_get",
            vec![
                param("text", Some(str_lit(""))),
                param("at", Some(i(0))),
                param("size", Some(i(1))),
                param("big", Some(bool_lit(false))),
                param("signed", Some(bool_lit(false))),
            ],
            vec![
                assign(ident("__get_v"), i(0)),
                assign(ident("__get_n"), i(0)),
                while_stmt(
                    op(BinOp::Lt, ident("__get_n"), ident("size")),
                    vec![
                        assign(
                            ident("__get_idx"),
                            ternary(
                                ident("big"),
                                add(ident("at"), ident("__get_n")),
                                add(
                                    ident("at"),
                                    op(
                                        BinOp::Sub,
                                        op(BinOp::Sub, ident("size"), ident("__get_n")),
                                        i(1),
                                    ),
                                ),
                            ),
                        ),
                        assign(
                            ident("__get_v"),
                            add(
                                op(BinOp::Mul, ident("__get_v"), i(256)),
                                index(ident("text"), ident("__get_idx")),
                            ),
                        ),
                        assign(ident("__get_n"), add(ident("__get_n"), i(1))),
                    ],
                ),
                if_stmt(
                    ident("signed"),
                    vec![
                        assign(ident("__get_half"), op(BinOp::Pow, i(256), ident("size"))),
                        if_stmt(
                            op(
                                BinOp::GtEq,
                                ident("__get_v"),
                                op(BinOp::FloorDiv, ident("__get_half"), i(2)),
                            ),
                            vec![assign(
                                ident("__get_v"),
                                op(BinOp::Sub, ident("__get_v"), ident("__get_half")),
                            )],
                        ),
                    ],
                ),
                ret(ident("__get_v")),
            ],
        ),
        function(
            "__py_struct_calcsize",
            vec![param("fmt", Some(str_lit("")))],
            vec![
                assign(
                    ident("__cs_p"),
                    call_global("__py_struct_parse", vec![ident("fmt")]),
                ),
                assign(ident("__cs_total"), i(0)),
                for_in(
                    "__cs_it",
                    index(ident("__cs_p"), i(1)),
                    vec![assign(
                        ident("__cs_total"),
                        add(
                            ident("__cs_total"),
                            op(
                                BinOp::Mul,
                                call_global(
                                    "__py_struct_size",
                                    vec![index(ident("__cs_it"), i(0))],
                                ),
                                index(ident("__cs_it"), i(1)),
                            ),
                        ),
                    )],
                ),
                ret(ident("__cs_total")),
            ],
        ),
        function(
            "__py_struct_pack",
            vec![param("fmt", Some(str_lit(""))), rest_param("values")],
            vec![
                assign(
                    ident("__pk_p"),
                    call_global("__py_struct_parse", vec![ident("fmt")]),
                ),
                assign(ident("__pk_big"), index(ident("__pk_p"), i(0))),
                assign(ident("__pk_out"), list_of(vec![])),
                assign(ident("__pk_vi"), i(0)),
                for_in(
                    "__pk_it",
                    index(ident("__pk_p"), i(1)),
                    vec![
                        assign(ident("__pk_code"), index(ident("__pk_it"), i(0))),
                        assign(ident("__pk_cnt"), index(ident("__pk_it"), i(1))),
                        assign(
                            ident("__pk_sz"),
                            call_global("__py_struct_size", vec![ident("__pk_code")]),
                        ),
                        assign(ident("__pk_j"), i(0)),
                        while_stmt(
                            op(BinOp::Lt, ident("__pk_j"), ident("__pk_cnt")),
                            vec![
                                // `x` is a pad byte and consumes no value.
                                if_stmt(
                                    op(BinOp::Eq, ident("__pk_code"), str_lit("x")),
                                    vec![push(ident("__pk_out"), i(0))],
                                ),
                                // `s` consumes ONE value for the whole field
                                // and pads to the count; `c` is one byte.
                                if_stmt(
                                    op(BinOp::Eq, ident("__pk_code"), str_lit("s")),
                                    vec![
                                        assign(
                                            ident("__pk_raw"),
                                            call_global(
                                                "list",
                                                vec![index(ident("values"), ident("__pk_vi"))],
                                            ),
                                        ),
                                        assign(ident("__pk_vi"), add(ident("__pk_vi"), i(1))),
                                        assign(ident("__pk_q"), i(0)),
                                        while_stmt(
                                            op(BinOp::Lt, ident("__pk_q"), ident("__pk_cnt")),
                                            vec![
                                                assign(ident("__pk_b2"), i(0)),
                                                if_stmt(
                                                    op(
                                                        BinOp::Lt,
                                                        ident("__pk_q"),
                                                        len_of(ident("__pk_raw")),
                                                    ),
                                                    vec![assign(
                                                        ident("__pk_b2"),
                                                        index(ident("__pk_raw"), ident("__pk_q")),
                                                    )],
                                                ),
                                                push(ident("__pk_out"), ident("__pk_b2")),
                                                assign(ident("__pk_q"), add(ident("__pk_q"), i(1))),
                                            ],
                                        ),
                                        assign(ident("__pk_j"), ident("__pk_cnt")),
                                    ],
                                ),
                                if_stmt(
                                    op(BinOp::Eq, ident("__pk_code"), str_lit("c")),
                                    vec![
                                        for_in(
                                            "__pk_cb",
                                            call_global(
                                                "list",
                                                vec![index(ident("values"), ident("__pk_vi"))],
                                            ),
                                            vec![push(ident("__pk_out"), ident("__pk_cb"))],
                                        ),
                                        assign(ident("__pk_vi"), add(ident("__pk_vi"), i(1))),
                                    ],
                                ),
                                if_stmt(
                                    op(
                                        BinOp::Lt,
                                        call(
                                            member(str_lit("xsc"), "find"),
                                            vec![ident("__pk_code")],
                                        ),
                                        i(0),
                                    ),
                                    vec![
                                        for_in(
                                            "__pk_bv",
                                            call_global(
                                                "__py_struct_put",
                                                vec![
                                                    index(ident("values"), ident("__pk_vi")),
                                                    ident("__pk_sz"),
                                                    ident("__pk_big"),
                                                ],
                                            ),
                                            vec![push(ident("__pk_out"), ident("__pk_bv"))],
                                        ),
                                        assign(ident("__pk_vi"), add(ident("__pk_vi"), i(1))),
                                    ],
                                ),
                                assign(ident("__pk_j"), add(ident("__pk_j"), i(1))),
                            ],
                        ),
                    ],
                ),
                ret(call_global("__py_bytes_from_list", vec![ident("__pk_out")])),
            ],
        ),
        function(
            "__py_struct_unpack",
            vec![param("fmt", Some(str_lit(""))), param("data", Some(null()))],
            vec![
                assign(
                    ident("__un_p"),
                    call_global("__py_struct_parse", vec![ident("fmt")]),
                ),
                assign(ident("__un_big"), index(ident("__un_p"), i(0))),
                assign(ident("__un_text"), call_global("list", vec![ident("data")])),
                assign(ident("__un_outv"), list_of(vec![])),
                assign(ident("__un_at"), i(0)),
                for_in(
                    "__un_it",
                    index(ident("__un_p"), i(1)),
                    vec![
                        assign(ident("__un_code"), index(ident("__un_it"), i(0))),
                        assign(ident("__un_cnt"), index(ident("__un_it"), i(1))),
                        assign(
                            ident("__un_sz"),
                            call_global("__py_struct_size", vec![ident("__un_code")]),
                        ),
                        // Lower case is signed, upper case unsigned.
                        assign(
                            ident("__un_signed"),
                            op(
                                BinOp::GtEq,
                                call(member(str_lit("bhilq"), "find"), vec![ident("__un_code")]),
                                i(0),
                            ),
                        ),
                        assign(ident("__un_j"), i(0)),
                        while_stmt(
                            op(BinOp::Lt, ident("__un_j"), ident("__un_cnt")),
                            vec![
                                if_stmt(
                                    op(BinOp::NotEq, ident("__un_code"), str_lit("x")),
                                    vec![push(
                                        ident("__un_outv"),
                                        call_global(
                                            "__py_struct_get",
                                            vec![
                                                ident("__un_text"),
                                                ident("__un_at"),
                                                ident("__un_sz"),
                                                ident("__un_big"),
                                                ident("__un_signed"),
                                            ],
                                        ),
                                    )],
                                ),
                                assign(ident("__un_at"), add(ident("__un_at"), ident("__un_sz"))),
                                assign(ident("__un_j"), add(ident("__un_j"), i(1))),
                            ],
                        ),
                    ],
                ),
                ret(call_global("tuple", vec![ident("__un_outv")])),
            ],
        ),
        // bytes → text, without a type test (see `io`/`tempfile` for the rule).
        function(
            "__py_struct_text",
            vec![param("value", Some(str_lit("")))],
            vec![
                if_stmt(is_none(ident("value")), vec![ret(str_lit(""))]),
                if_stmt(
                    op(
                        BinOp::Eq,
                        len_of(call_global("str", vec![ident("value")])),
                        len_of(ident("value")),
                    ),
                    vec![ret(ident("value"))],
                ),
                assign(ident("__tx_out"), str_lit("")),
                for_in(
                    "__tx_c",
                    ident("value"),
                    vec![assign(
                        ident("__tx_out"),
                        add(ident("__tx_out"), call_global("chr", vec![ident("__tx_c")])),
                    )],
                ),
                ret(ident("__tx_out")),
            ],
        ),
    ]
}
