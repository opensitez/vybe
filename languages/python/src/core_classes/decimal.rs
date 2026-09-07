//! `decimal.Decimal` — exact base-10 arithmetic.
//!
//! A value is a signed integer COEFFICIENT and an EXPONENT, never a float:
//! `Decimal('1.1') + Decimal('2.2')` is `(11,-1) + (22,-1) = (33,-1)`, which
//! prints `3.3`. Going through a float would print `3.3000000000000003` and
//! there would be no point to the class.
//!
//! ⛔ A spliced core class is NEVER WALKED, so nothing here may depend on a
//! walk-time rewrite: no `isinstance`, no `x in s` (use `s.find(x)`), and no
//! `num()` float literals where an integer is meant. All three were measured
//! failures in `fractions.rs` before this file was written.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn mul(l: Expr, r: Expr) -> Expr {
    op(BinOp::Mul, l, r)
}

fn sub(l: Expr, r: Expr) -> Expr {
    op(BinOp::Sub, l, r)
}

fn pow10(e: Expr) -> Expr {
    op(BinOp::Pow, i(10), e)
}

fn c_of(e: Expr) -> Expr {
    read_attr(e, "_c")
}

fn e_of(e: Expr) -> Expr {
    read_attr(e, "_e")
}

fn k_of(e: Expr) -> Expr {
    read_attr(e, "_kind")
}

fn this_c() -> Expr {
    this_field("_c")
}

fn this_e() -> Expr {
    this_field("_e")
}

fn other_dec() -> Statement {
    assign(
        ident("__o"),
        call_global("__py_dec_of", vec![ident("other")]),
    )
}

fn make(c: Expr, e: Expr) -> Expr {
    call_global("__py_dec_make", vec![c, e])
}

/// `self <op> other` through the three-way compare.
fn compare(o: BinOp) -> Vec<Statement> {
    vec![
        other_dec(),
        ret(op(
            o,
            call_global("__py_dec_cmp", vec![ident("self"), ident("__o")]),
            i(0),
        )),
    ]
}

pub(super) fn decimal_tuple() -> Statement {
    class(
        "DecimalTuple",
        vec![init(
            vec![
                param("sign", Some(i(0))),
                param("digits", Some(null())),
                param("exponent", Some(i(0))),
            ],
            vec![
                set_this("sign", ident("sign")),
                set_this("digits", ident("digits")),
                set_this("exponent", ident("exponent")),
            ],
        )],
    )
}

/// The arithmetic context. Only `prec` is consulted, by unary `+`.
pub(super) fn context() -> Statement {
    class(
        "Context",
        vec![
            init(
                vec![param("prec", Some(i(28)))],
                vec![set_this("prec", ident("prec"))],
            ),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method("__exit__", any_args(), vec![ret(null())]),
            method(
                "copy",
                vec![],
                vec![ret(new("Context", vec![this_field("prec")]))],
            ),
        ],
    )
}

pub(super) fn decimal() -> Statement {
    class(
        "Decimal",
        vec![
            init(
                vec![
                    param("value", Some(i(0))),
                    // The rescaling helpers build a Decimal from parts they
                    // have already computed; re-parsing their text would round
                    // twice.
                    param("_coeff", Some(null())),
                    param("_exp", Some(i(0))),
                    param("_special", Some(i(0))),
                ],
                vec![
                    // ⛔ NO early `return` — a constructor that returns makes
                    // `Decimal(...)` evaluate to that value, so every
                    // `__py_dec_make` answered None and all the arithmetic
                    // came back None.
                    if_stmt(
                        is_not_none(ident("_coeff")),
                        vec![
                            set_this("_c", ident("_coeff")),
                            set_this("_e", ident("_exp")),
                            set_this("_kind", ident("_special")),
                        ],
                    ),
                    if_stmt(
                        is_none(ident("_coeff")),
                        vec![
                            assign(
                                ident("__py_dec_parts_tuple"),
                                call_global("__py_dec_parts", vec![ident("value")]),
                            ),
                            set_this("_c", index(ident("__py_dec_parts_tuple"), i(0))),
                            set_this("_e", index(ident("__py_dec_parts_tuple"), i(1))),
                            set_this("_kind", index(ident("__py_dec_parts_tuple"), i(2))),
                        ],
                    ),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![ret(call_global(
                    "__py_dec_text",
                    vec![this_c(), this_e(), this_field("_kind")],
                ))],
            ),
            method(
                "__repr__",
                vec![],
                vec![ret(op(
                    BinOp::Add,
                    op(
                        BinOp::Add,
                        str_lit("Decimal('"),
                        call_global(
                            "__py_dec_text",
                            vec![this_c(), this_e(), this_field("_kind")],
                        ),
                    ),
                    str_lit("')"),
                ))],
            ),
            method(
                "__add__",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    assign(
                        ident("__e"),
                        ternary(
                            op(BinOp::Lt, e_of(ident("__o")), this_e()),
                            e_of(ident("__o")),
                            this_e(),
                        ),
                    ),
                    ret(make(
                        op(
                            BinOp::Add,
                            mul(this_c(), pow10(sub(this_e(), ident("__e")))),
                            mul(
                                c_of(ident("__o")),
                                pow10(sub(e_of(ident("__o")), ident("__e"))),
                            ),
                        ),
                        ident("__e"),
                    )),
                ],
            ),
            method(
                "__sub__",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    assign(
                        ident("__e"),
                        ternary(
                            op(BinOp::Lt, e_of(ident("__o")), this_e()),
                            e_of(ident("__o")),
                            this_e(),
                        ),
                    ),
                    ret(make(
                        sub(
                            mul(this_c(), pow10(sub(this_e(), ident("__e")))),
                            mul(
                                c_of(ident("__o")),
                                pow10(sub(e_of(ident("__o")), ident("__e"))),
                            ),
                        ),
                        ident("__e"),
                    )),
                ],
            ),
            method(
                "__mul__",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    ret(make(
                        mul(this_c(), c_of(ident("__o"))),
                        op(BinOp::Add, this_e(), e_of(ident("__o"))),
                    )),
                ],
            ),
            // Exact division is not always possible, so the coefficient is
            // scaled by a fixed number of digits and the trailing zeros are
            // stripped — `10 / 4` comes back `2.5`, not `2.500000000000000`.
            method(
                "__truediv__",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    assign(
                        ident("__q"),
                        op(
                            BinOp::FloorDiv,
                            mul(this_c(), pow10(i(15))),
                            c_of(ident("__o")),
                        ),
                    ),
                    ret(call_global(
                        "__py_dec_strip",
                        vec![ident("__q"), sub(sub(this_e(), e_of(ident("__o"))), i(15))],
                    )),
                ],
            ),
            method(
                "__floordiv__",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    assign(
                        ident("__e"),
                        ternary(
                            op(BinOp::Lt, e_of(ident("__o")), this_e()),
                            e_of(ident("__o")),
                            this_e(),
                        ),
                    ),
                    ret(make(
                        op(
                            BinOp::FloorDiv,
                            mul(this_c(), pow10(sub(this_e(), ident("__e")))),
                            mul(
                                c_of(ident("__o")),
                                pow10(sub(e_of(ident("__o")), ident("__e"))),
                            ),
                        ),
                        i(0),
                    )),
                ],
            ),
            method(
                "__mod__",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    assign(
                        ident("__e"),
                        ternary(
                            op(BinOp::Lt, e_of(ident("__o")), this_e()),
                            e_of(ident("__o")),
                            this_e(),
                        ),
                    ),
                    ret(make(
                        op(
                            BinOp::Mod,
                            mul(this_c(), pow10(sub(this_e(), ident("__e")))),
                            mul(
                                c_of(ident("__o")),
                                pow10(sub(e_of(ident("__o")), ident("__e"))),
                            ),
                        ),
                        ident("__e"),
                    )),
                ],
            ),
            method(
                "__pow__",
                vec![param("other", Some(null()))],
                vec![
                    assign(
                        ident("__n"),
                        call_global("__py_dec_int_of", vec![ident("other")]),
                    ),
                    ret(make(
                        op(BinOp::Pow, this_c(), ident("__n")),
                        mul(this_e(), ident("__n")),
                    )),
                ],
            ),
            method(
                "__neg__",
                vec![],
                vec![ret(make(sub(i(0), this_c()), this_e()))],
            ),
            // Unary `+` is where the CONTEXT applies: `getcontext().prec = 3`
            // then `+Decimal('1.2345')` is `1.23`.
            method(
                "__pos__",
                vec![],
                vec![ret(call_global(
                    "__py_dec_prec",
                    vec![
                        ident("self"),
                        read_attr(call_global("getcontext", vec![]), "prec"),
                    ],
                ))],
            ),
            method(
                "__abs__",
                vec![],
                vec![ret(make(
                    ternary(op(BinOp::Lt, this_c(), i(0)), sub(i(0), this_c()), this_c()),
                    this_e(),
                ))],
            ),
            method(
                "__float__",
                vec![],
                vec![ret(mul(this_c(), pow10(this_e())))],
            ),
            method(
                "__int__",
                vec![],
                vec![
                    if_stmt(
                        op(BinOp::GtEq, this_e(), i(0)),
                        vec![ret(mul(this_c(), pow10(this_e())))],
                    ),
                    ret(op(BinOp::FloorDiv, this_c(), pow10(sub(i(0), this_e())))),
                ],
            ),
            method("__hash__", vec![], vec![ret(this_c())]),
            method(
                "__eq__",
                vec![param("other", Some(null()))],
                compare(BinOp::Eq),
            ),
            method(
                "__ne__",
                vec![param("other", Some(null()))],
                compare(BinOp::NotEq),
            ),
            method(
                "__lt__",
                vec![param("other", Some(null()))],
                compare(BinOp::Lt),
            ),
            method(
                "__le__",
                vec![param("other", Some(null()))],
                compare(BinOp::LtEq),
            ),
            method(
                "__gt__",
                vec![param("other", Some(null()))],
                compare(BinOp::Gt),
            ),
            method(
                "__ge__",
                vec![param("other", Some(null()))],
                compare(BinOp::GtEq),
            ),
            method(
                "is_nan",
                vec![],
                vec![ret(op(BinOp::Eq, this_field("_kind"), i(2)))],
            ),
            method(
                "is_infinite",
                vec![],
                vec![ret(op(BinOp::Eq, this_field("_kind"), i(1)))],
            ),
            method(
                "is_finite",
                vec![],
                vec![ret(op(BinOp::Eq, this_field("_kind"), i(0)))],
            ),
            method(
                "is_signed",
                vec![],
                vec![ret(op(BinOp::Lt, this_c(), i(0)))],
            ),
            method("is_zero", vec![], vec![ret(op(BinOp::Eq, this_c(), i(0)))]),
            method(
                "copy_abs",
                vec![],
                vec![ret(make(
                    ternary(op(BinOp::Lt, this_c(), i(0)), sub(i(0), this_c()), this_c()),
                    this_e(),
                ))],
            ),
            method(
                "copy_negate",
                vec![],
                vec![ret(make(sub(i(0), this_c()), this_e()))],
            ),
            method(
                "compare",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    ret(new(
                        "Decimal",
                        vec![call_global(
                            "__py_dec_cmp",
                            vec![ident("self"), ident("__o")],
                        )],
                    )),
                ],
            ),
            method(
                "compare_total",
                vec![param("other", Some(null()))],
                vec![
                    other_dec(),
                    ret(new(
                        "Decimal",
                        vec![call_global(
                            "__py_dec_cmp",
                            vec![ident("self"), ident("__o")],
                        )],
                    )),
                ],
            ),
            method(
                "as_tuple",
                vec![],
                vec![ret(call_global("__py_dec_tuple", vec![ident("self")]))],
            ),
            method(
                "sqrt",
                vec![],
                vec![ret(call_global(
                    "__py_dec_of",
                    vec![call_global(
                        "str",
                        vec![op(BinOp::Pow, mul(this_c(), pow10(this_e())), num(0.5))],
                    )],
                ))],
            ),
            // `quantize` rounds to the EXPONENT of its argument, half-up.
            method(
                "quantize",
                vec![
                    param("exp", Some(null())),
                    param("rounding", Some(null())),
                    param("context", Some(null())),
                ],
                vec![ret(call_global(
                    "__py_dec_rescale",
                    vec![
                        ident("self"),
                        e_of(call_global("__py_dec_of", vec![ident("exp")])),
                    ],
                ))],
            ),
            method(
                "normalize",
                vec![],
                vec![ret(call_global("__py_dec_strip", vec![this_c(), this_e()]))],
            ),
            method(
                "to_integral_value",
                vec![],
                vec![ret(call_global(
                    "__py_dec_rescale",
                    vec![ident("self"), i(0)],
                ))],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("__py_dec_ctx", new("Context", vec![])),
        global_assign("ROUND_HALF_UP", str_lit("ROUND_HALF_UP")),
        global_assign("ROUND_HALF_EVEN", str_lit("ROUND_HALF_EVEN")),
        global_assign("ROUND_HALF_DOWN", str_lit("ROUND_HALF_DOWN")),
        global_assign("ROUND_UP", str_lit("ROUND_UP")),
        global_assign("ROUND_DOWN", str_lit("ROUND_DOWN")),
        global_assign("ROUND_CEILING", str_lit("ROUND_CEILING")),
        global_assign("ROUND_FLOOR", str_lit("ROUND_FLOOR")),
        global_assign("ROUND_05UP", str_lit("ROUND_05UP")),
        function("getcontext", vec![], vec![ret(ident("__py_dec_ctx"))]),
        function(
            "setcontext",
            vec![param("ctx", Some(null()))],
            vec![global_assign("__py_dec_ctx", ident("ctx")), ret(null())],
        ),
        // `with localcontext() as ctx:` — a fresh Context that is its own
        // context manager.
        function(
            "localcontext",
            vec![param("ctx", Some(null()))],
            vec![ret(new(
                "Context",
                vec![read_attr(ident("__py_dec_ctx"), "prec")],
            ))],
        ),
        // `[coefficient, exponent, kind]` for whatever the constructor was
        // handed. `kind` is 0 finite, 1 infinite, 2 NaN.
        function(
            "__py_dec_parts",
            vec![param("value", Some(null()))],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("value"), str_lit("_kind")]),
                    vec![ret(list_of(vec![
                        c_of(ident("value")),
                        e_of(ident("value")),
                        k_of(ident("value")),
                    ]))],
                ),
                assign(ident("s"), call_global("str", vec![ident("value")])),
                // The 3-tuple form `Decimal((sign, digits, exponent))`.
                if_stmt(
                    op(
                        BinOp::Eq,
                        call(member(ident("s"), "find"), vec![str_lit("(")]),
                        i(0),
                    ),
                    vec![
                        assign(ident("c"), i(0)),
                        for_in(
                            "dg",
                            index(ident("value"), i(1)),
                            vec![assign(
                                ident("c"),
                                op(BinOp::Add, mul(ident("c"), i(10)), ident("dg")),
                            )],
                        ),
                        if_stmt(
                            op(BinOp::Eq, index(ident("value"), i(0)), i(1)),
                            vec![assign(ident("c"), sub(i(0), ident("c")))],
                        ),
                        ret(list_of(vec![ident("c"), index(ident("value"), i(2)), i(0)])),
                    ],
                ),
                if_stmt(
                    op(
                        BinOp::GtEq,
                        call(member(ident("s"), "find"), vec![str_lit("N")]),
                        i(0),
                    ),
                    vec![ret(list_of(vec![i(0), i(0), i(2)]))],
                ),
                if_stmt(
                    op(
                        BinOp::GtEq,
                        call(member(ident("s"), "find"), vec![str_lit("nfinity")]),
                        i(0),
                    ),
                    vec![ret(list_of(vec![
                        ternary(
                            op(
                                BinOp::Eq,
                                call(member(ident("s"), "find"), vec![str_lit("-")]),
                                i(0),
                            ),
                            i(-1),
                            i(1),
                        ),
                        i(0),
                        i(1),
                    ]))],
                ),
                assign(
                    ident("dot"),
                    call(member(ident("s"), "find"), vec![str_lit(".")]),
                ),
                if_stmt(
                    op(BinOp::Lt, ident("dot"), i(0)),
                    vec![ret(list_of(vec![
                        call_global("int", vec![ident("s")]),
                        i(0),
                        i(0),
                    ]))],
                ),
                assign(ident("whole"), slice_range(ident("s"), i(0), ident("dot"))),
                assign(
                    ident("frac"),
                    slice_from(ident("s"), op(BinOp::Add, ident("dot"), i(1))),
                ),
                assign(
                    ident("scale"),
                    pow10(call_global("len", vec![ident("frac")])),
                ),
                assign(ident("digits"), call_global("int", vec![ident("frac")])),
                if_stmt(
                    op(
                        BinOp::Eq,
                        call(member(ident("whole"), "find"), vec![str_lit("-")]),
                        i(0),
                    ),
                    vec![assign(ident("digits"), sub(i(0), ident("digits")))],
                ),
                ret(list_of(vec![
                    op(
                        BinOp::Add,
                        mul(call_global("int", vec![ident("whole")]), ident("scale")),
                        ident("digits"),
                    ),
                    sub(i(0), call_global("len", vec![ident("frac")])),
                    i(0),
                ])),
            ],
        ),
        function(
            "__py_dec_make",
            vec![param("c", Some(i(0))), param("e", Some(i(0)))],
            vec![ret(new(
                "Decimal",
                vec![i(0), ident("c"), ident("e"), i(0)],
            ))],
        ),
        function(
            "__py_dec_of",
            vec![param("value", Some(null()))],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("value"), str_lit("_kind")]),
                    vec![ret(ident("value"))],
                ),
                ret(new("Decimal", vec![ident("value")])),
            ],
        ),
        function(
            "__py_dec_int_of",
            vec![param("value", Some(null()))],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("value"), str_lit("_kind")]),
                    vec![ret(mul(c_of(ident("value")), pow10(e_of(ident("value")))))],
                ),
                ret(ident("value")),
            ],
        ),
        // The printed form: the coefficient with a point inserted `-e` digits
        // from the right.
        function(
            "__py_dec_text",
            vec![
                param("c", Some(i(0))),
                param("e", Some(i(0))),
                param("kind", Some(i(0))),
            ],
            vec![
                if_stmt(
                    op(BinOp::Eq, ident("kind"), i(2)),
                    vec![ret(str_lit("NaN"))],
                ),
                if_stmt(
                    op(BinOp::Eq, ident("kind"), i(1)),
                    vec![ret(ternary(
                        op(BinOp::Lt, ident("c"), i(0)),
                        str_lit("-Infinity"),
                        str_lit("Infinity"),
                    ))],
                ),
                assign(ident("neg"), op(BinOp::Lt, ident("c"), i(0))),
                if_stmt(
                    ident("neg"),
                    vec![assign(ident("c"), sub(i(0), ident("c")))],
                ),
                assign(ident("s"), call_global("str", vec![ident("c")])),
                if_stmt(
                    op(BinOp::Gt, ident("e"), i(0)),
                    vec![
                        assign(ident("k"), ident("e")),
                        while_stmt(
                            op(BinOp::Gt, ident("k"), i(0)),
                            vec![
                                assign(ident("s"), op(BinOp::Add, ident("s"), str_lit("0"))),
                                assign(ident("k"), sub(ident("k"), i(1))),
                            ],
                        ),
                    ],
                ),
                if_stmt(
                    op(BinOp::Lt, ident("e"), i(0)),
                    vec![
                        assign(ident("k"), sub(i(0), ident("e"))),
                        while_stmt(
                            op(
                                BinOp::LtEq,
                                call_global("len", vec![ident("s")]),
                                ident("k"),
                            ),
                            vec![assign(ident("s"), op(BinOp::Add, str_lit("0"), ident("s")))],
                        ),
                        assign(
                            ident("cut"),
                            sub(call_global("len", vec![ident("s")]), ident("k")),
                        ),
                        assign(
                            ident("s"),
                            op(
                                BinOp::Add,
                                op(
                                    BinOp::Add,
                                    slice_range(ident("s"), i(0), ident("cut")),
                                    str_lit("."),
                                ),
                                slice_from(ident("s"), ident("cut")),
                            ),
                        ),
                    ],
                ),
                if_stmt(
                    ident("neg"),
                    vec![assign(ident("s"), op(BinOp::Add, str_lit("-"), ident("s")))],
                ),
                ret(ident("s")),
            ],
        ),
        // Three-way compare. An infinity outranks every finite value, which is
        // what `Decimal('Infinity') > Decimal('1000')` is asking.
        function(
            "__py_dec_cmp",
            vec![param("a", Some(null())), param("b", Some(null()))],
            vec![
                assign(ident("ra"), i(0)),
                if_stmt(
                    op(BinOp::Eq, k_of(ident("a")), i(1)),
                    vec![assign(
                        ident("ra"),
                        ternary(op(BinOp::Lt, c_of(ident("a")), i(0)), i(-1), i(1)),
                    )],
                ),
                assign(ident("rb"), i(0)),
                if_stmt(
                    op(BinOp::Eq, k_of(ident("b")), i(1)),
                    vec![assign(
                        ident("rb"),
                        ternary(op(BinOp::Lt, c_of(ident("b")), i(0)), i(-1), i(1)),
                    )],
                ),
                if_stmt(
                    binary(
                        BinOp::Or,
                        op(BinOp::NotEq, ident("ra"), i(0)),
                        op(BinOp::NotEq, ident("rb"), i(0)),
                    ),
                    vec![
                        if_stmt(op(BinOp::Lt, ident("ra"), ident("rb")), vec![ret(i(-1))]),
                        if_stmt(op(BinOp::Gt, ident("ra"), ident("rb")), vec![ret(i(1))]),
                        ret(i(0)),
                    ],
                ),
                assign(
                    ident("e"),
                    ternary(
                        op(BinOp::Lt, e_of(ident("b")), e_of(ident("a"))),
                        e_of(ident("b")),
                        e_of(ident("a")),
                    ),
                ),
                assign(
                    ident("ca"),
                    mul(c_of(ident("a")), pow10(sub(e_of(ident("a")), ident("e")))),
                ),
                assign(
                    ident("cb"),
                    mul(c_of(ident("b")), pow10(sub(e_of(ident("b")), ident("e")))),
                ),
                if_stmt(op(BinOp::Lt, ident("ca"), ident("cb")), vec![ret(i(-1))]),
                if_stmt(op(BinOp::Gt, ident("ca"), ident("cb")), vec![ret(i(1))]),
                ret(i(0)),
            ],
        ),
        // Drop trailing zeros from the coefficient, raising the exponent.
        function(
            "__py_dec_strip",
            vec![param("c", Some(i(0))), param("e", Some(i(0)))],
            vec![
                while_stmt(
                    binary(
                        BinOp::And,
                        op(BinOp::Lt, ident("e"), i(0)),
                        op(BinOp::Eq, op(BinOp::Mod, ident("c"), i(10)), i(0)),
                    ),
                    vec![
                        assign(ident("c"), op(BinOp::FloorDiv, ident("c"), i(10))),
                        assign(ident("e"), op(BinOp::Add, ident("e"), i(1))),
                    ],
                ),
                ret(call_global("__py_dec_make", vec![ident("c"), ident("e")])),
            ],
        ),
        // Round to a target exponent, half away from zero.
        function(
            "__py_dec_rescale",
            vec![param("d", Some(null())), param("target", Some(i(0)))],
            vec![
                assign(ident("c"), c_of(ident("d"))),
                assign(ident("e"), e_of(ident("d"))),
                if_stmt(
                    op(BinOp::GtEq, ident("e"), ident("target")),
                    vec![ret(call_global(
                        "__py_dec_make",
                        vec![
                            mul(ident("c"), pow10(sub(ident("e"), ident("target")))),
                            ident("target"),
                        ],
                    ))],
                ),
                assign(ident("p"), pow10(sub(ident("target"), ident("e")))),
                assign(ident("neg"), op(BinOp::Lt, ident("c"), i(0))),
                if_stmt(
                    ident("neg"),
                    vec![assign(ident("c"), sub(i(0), ident("c")))],
                ),
                assign(ident("q"), op(BinOp::FloorDiv, ident("c"), ident("p"))),
                assign(ident("r"), sub(ident("c"), mul(ident("q"), ident("p")))),
                if_stmt(
                    op(BinOp::GtEq, mul(ident("r"), i(2)), ident("p")),
                    vec![assign(ident("q"), op(BinOp::Add, ident("q"), i(1)))],
                ),
                if_stmt(
                    ident("neg"),
                    vec![assign(ident("q"), sub(i(0), ident("q")))],
                ),
                ret(call_global(
                    "__py_dec_make",
                    vec![ident("q"), ident("target")],
                )),
            ],
        ),
        // Round to `prec` SIGNIFICANT digits — what unary `+` applies.
        function(
            "__py_dec_prec",
            vec![param("d", Some(null())), param("prec", Some(i(28)))],
            vec![
                assign(ident("c"), c_of(ident("d"))),
                if_stmt(
                    op(BinOp::Lt, ident("c"), i(0)),
                    vec![assign(ident("c"), sub(i(0), ident("c")))],
                ),
                assign(
                    ident("n"),
                    call_global("len", vec![call_global("str", vec![ident("c")])]),
                ),
                if_stmt(
                    op(BinOp::LtEq, ident("n"), ident("prec")),
                    vec![ret(call_global(
                        "__py_dec_make",
                        vec![c_of(ident("d")), e_of(ident("d"))],
                    ))],
                ),
                ret(call_global(
                    "__py_dec_rescale",
                    vec![
                        ident("d"),
                        op(BinOp::Add, e_of(ident("d")), sub(ident("n"), ident("prec"))),
                    ],
                )),
            ],
        ),
        function(
            "__py_dec_tuple",
            vec![param("d", Some(null()))],
            vec![
                assign(ident("c"), c_of(ident("d"))),
                assign(
                    ident("sign"),
                    ternary(op(BinOp::Lt, ident("c"), i(0)), i(1), i(0)),
                ),
                if_stmt(
                    op(BinOp::Lt, ident("c"), i(0)),
                    vec![assign(ident("c"), sub(i(0), ident("c")))],
                ),
                assign(ident("digits"), list_of(vec![])),
                for_in(
                    "ch",
                    call_global("str", vec![ident("c")]),
                    vec![expr_stmt(call(
                        member(ident("digits"), "append"),
                        vec![call_global("int", vec![ident("ch")])],
                    ))],
                ),
                ret(new(
                    "DecimalTuple",
                    vec![ident("sign"), ident("digits"), e_of(ident("d"))],
                )),
            ],
        ),
    ]
}

/// `decimal`'s exception tree — `class X(Y): pass`, one row each.
pub(super) const EXCEPTIONS: &[(&str, &str)] = &[
    ("DecimalException", "ArithmeticError"),
    ("Clamped", "DecimalException"),
    ("InvalidOperation", "DecimalException"),
    ("DivisionByZero", "DecimalException"),
    ("Inexact", "DecimalException"),
    ("Rounded", "DecimalException"),
    ("Subnormal", "DecimalException"),
    ("Overflow", "DecimalException"),
    ("Underflow", "DecimalException"),
    ("FloatOperation", "DecimalException"),
];

pub(super) fn exception(name: &str, parent: &str) -> Statement {
    class_extending(name, &[parent], vec![])
}
