//! `fractions.Fraction` — an exact rational, normalised on construction.
//!
//! Numerator and denominator are ordinary FIELDS, not properties: they are
//! stored by the constructor rather than computed, so an accessor would buy
//! nothing and would carry the receiver-binding constraint that applies inside
//! one. The arithmetic is the ordinary dunder surface, so `+`, `<` and `abs()`
//! reach it through the same protocol slots a user class binds.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

/// An INTEGER literal. `num()` builds `Literal::Float`, and a float index
/// does not subscript a list — `parts[0.0]` read nothing.
fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn frac(n: Expr, d: Expr) -> Expr {
    new("Fraction", vec![n, d])
}

type Expr = vybe_ast::Expression;

fn n_of(e: Expr) -> Expr {
    read_attr(e, "numerator")
}

fn d_of(e: Expr) -> Expr {
    read_attr(e, "denominator")
}

fn this_n() -> Expr {
    this_field("numerator")
}

fn this_d() -> Expr {
    this_field("denominator")
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

/// `Fraction(other)` for an operand that may be a plain int — the dunders
/// accept `Fraction + 2` the way CPython's do.
fn coerced(name: &str) -> Expr {
    call_global("__py_frac_of", vec![ident(name)])
}

/// `self <op> other` as a cross-multiplied comparison: `a/b ? c/d` is
/// `a*d ? c*b`, and both denominators are positive by construction.
fn compare(o: BinOp) -> Vec<Statement> {
    vec![
        assign(ident("__o"), coerced("other")),
        ret(op(
            o,
            mul(this_n(), d_of(ident("__o"))),
            mul(n_of(ident("__o")), this_d()),
        )),
    ]
}

pub(super) fn fraction() -> Statement {
    class(
        "Fraction",
        vec![
            init(
                vec![
                    param("numerator", Some(i(0))),
                    // ⛔ `Some(null())`, not `None`: `None` is NO DEFAULT, so a
                    // one-argument `Fraction(0.5)` left this unbound, `d`
                    // became NaN and `__py_frac_gcd` spun forever.
                    param("denominator", Some(null())),
                ],
                vec![
                    assign(
                        ident("__parts"),
                        call_global(
                            "__py_frac_parts",
                            vec![ident("numerator"), ident("denominator")],
                        ),
                    ),
                    set_this("numerator", index(ident("__parts"), i(0))),
                    set_this("denominator", index(ident("__parts"), i(1))),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![
                    if_stmt(
                        op(BinOp::Eq, this_d(), i(1)),
                        vec![ret(call_global("str", vec![this_n()]))],
                    ),
                    ret(op(
                        BinOp::Add,
                        op(BinOp::Add, call_global("str", vec![this_n()]), str_lit("/")),
                        call_global("str", vec![this_d()]),
                    )),
                ],
            ),
            method(
                "__repr__",
                vec![],
                vec![ret(op(
                    BinOp::Add,
                    op(
                        BinOp::Add,
                        op(
                            BinOp::Add,
                            str_lit("Fraction("),
                            call_global("str", vec![this_n()]),
                        ),
                        op(
                            BinOp::Add,
                            str_lit(", "),
                            call_global("str", vec![this_d()]),
                        ),
                    ),
                    str_lit(")"),
                ))],
            ),
            method(
                "__add__",
                vec![param("other", None.into())],
                vec![
                    assign(ident("__o"), coerced("other")),
                    ret(frac(
                        op(
                            BinOp::Add,
                            mul(this_n(), d_of(ident("__o"))),
                            mul(n_of(ident("__o")), this_d()),
                        ),
                        mul(this_d(), d_of(ident("__o"))),
                    )),
                ],
            ),
            method(
                "__sub__",
                vec![param("other", None.into())],
                vec![
                    assign(ident("__o"), coerced("other")),
                    ret(frac(
                        sub(
                            mul(this_n(), d_of(ident("__o"))),
                            mul(n_of(ident("__o")), this_d()),
                        ),
                        mul(this_d(), d_of(ident("__o"))),
                    )),
                ],
            ),
            method(
                "__mul__",
                vec![param("other", None.into())],
                vec![
                    assign(ident("__o"), coerced("other")),
                    ret(frac(
                        mul(this_n(), n_of(ident("__o"))),
                        mul(this_d(), d_of(ident("__o"))),
                    )),
                ],
            ),
            method(
                "__truediv__",
                vec![param("other", None.into())],
                vec![
                    assign(ident("__o"), coerced("other")),
                    ret(frac(
                        mul(this_n(), d_of(ident("__o"))),
                        mul(this_d(), n_of(ident("__o"))),
                    )),
                ],
            ),
            // `//` on two rationals is an INTEGER in CPython, not a Fraction.
            method(
                "__floordiv__",
                vec![param("other", None.into())],
                vec![
                    assign(ident("__o"), coerced("other")),
                    ret(op(
                        BinOp::FloorDiv,
                        mul(this_n(), d_of(ident("__o"))),
                        mul(this_d(), n_of(ident("__o"))),
                    )),
                ],
            ),
            method(
                "__mod__",
                vec![param("other", None.into())],
                vec![
                    assign(ident("__o"), coerced("other")),
                    assign(
                        ident("__q"),
                        op(
                            BinOp::FloorDiv,
                            mul(this_n(), d_of(ident("__o"))),
                            mul(this_d(), n_of(ident("__o"))),
                        ),
                    ),
                    ret(frac(
                        sub(
                            mul(this_n(), d_of(ident("__o"))),
                            mul(ident("__q"), mul(this_d(), n_of(ident("__o")))),
                        ),
                        mul(this_d(), d_of(ident("__o"))),
                    )),
                ],
            ),
            method(
                "__pow__",
                vec![param("other", None.into())],
                vec![
                    if_stmt(
                        op(BinOp::Lt, ident("other"), i(0)),
                        vec![ret(frac(
                            op(BinOp::Pow, this_d(), op(BinOp::Sub, i(0), ident("other"))),
                            op(BinOp::Pow, this_n(), op(BinOp::Sub, i(0), ident("other"))),
                        ))],
                    ),
                    ret(frac(
                        op(BinOp::Pow, this_n(), ident("other")),
                        op(BinOp::Pow, this_d(), ident("other")),
                    )),
                ],
            ),
            method(
                "__neg__",
                vec![],
                vec![ret(frac(sub(i(0), this_n()), this_d()))],
            ),
            method("__pos__", vec![], vec![ret(frac(this_n(), this_d()))]),
            method(
                "__abs__",
                vec![],
                vec![ret(frac(
                    ternary(op(BinOp::Lt, this_n(), i(0)), sub(i(0), this_n()), this_n()),
                    this_d(),
                ))],
            ),
            method(
                "__float__",
                vec![],
                vec![ret(op(BinOp::Div, this_n(), this_d()))],
            ),
            method(
                "__int__",
                vec![],
                vec![ret(op(BinOp::FloorDiv, this_n(), this_d()))],
            ),
            method(
                "__hash__",
                vec![],
                vec![ret(op(BinOp::Add, mul(this_n(), i(1000003)), this_d()))],
            ),
            method(
                "__eq__",
                vec![param("other", None.into())],
                compare(BinOp::Eq),
            ),
            method(
                "__ne__",
                vec![param("other", None.into())],
                compare(BinOp::NotEq),
            ),
            method(
                "__lt__",
                vec![param("other", None.into())],
                compare(BinOp::Lt),
            ),
            method(
                "__le__",
                vec![param("other", None.into())],
                compare(BinOp::LtEq),
            ),
            method(
                "__gt__",
                vec![param("other", None.into())],
                compare(BinOp::Gt),
            ),
            method(
                "__ge__",
                vec![param("other", None.into())],
                compare(BinOp::GtEq),
            ),
            // The Stern–Brocot walk CPython uses: the best rational whose
            // denominator does not exceed the bound.
            method(
                "limit_denominator",
                vec![param("max_denominator", Some(i(1000000)))],
                vec![
                    if_stmt(
                        op(BinOp::LtEq, this_d(), ident("max_denominator")),
                        vec![ret(frac(this_n(), this_d()))],
                    ),
                    assign(ident("__p0"), i(0)),
                    assign(ident("__q0"), i(1)),
                    assign(ident("__p1"), i(1)),
                    assign(ident("__q1"), i(0)),
                    assign(ident("__n"), this_n()),
                    assign(ident("__d"), this_d()),
                    while_stmt(
                        op(BinOp::NotEq, ident("__d"), i(0)),
                        vec![
                            assign(
                                ident("__a"),
                                op(BinOp::FloorDiv, ident("__n"), ident("__d")),
                            ),
                            assign(
                                ident("__q2"),
                                op(BinOp::Add, ident("__q0"), mul(ident("__a"), ident("__q1"))),
                            ),
                            if_stmt(
                                op(BinOp::Gt, ident("__q2"), ident("max_denominator")),
                                vec![Statement::with_span(
                                    vybe_ast::StmtKind::Break(vybe_ast::BreakTarget::Implicit),
                                    vybe_ast::Span::default(),
                                )],
                            ),
                            assign(ident("__t"), ident("__p1")),
                            assign(
                                ident("__p1"),
                                op(BinOp::Add, ident("__p0"), mul(ident("__a"), ident("__p1"))),
                            ),
                            assign(ident("__p0"), ident("__t")),
                            assign(ident("__t"), ident("__q1")),
                            assign(ident("__q1"), ident("__q2")),
                            assign(ident("__q0"), ident("__t")),
                            assign(ident("__t"), ident("__d")),
                            assign(
                                ident("__d"),
                                sub(ident("__n"), mul(ident("__a"), ident("__d"))),
                            ),
                            assign(ident("__n"), ident("__t")),
                        ],
                    ),
                    assign(
                        ident("__k"),
                        op(
                            BinOp::FloorDiv,
                            sub(ident("max_denominator"), ident("__q0")),
                            ident("__q1"),
                        ),
                    ),
                    assign(
                        ident("__b1"),
                        frac(
                            op(BinOp::Add, ident("__p0"), mul(ident("__k"), ident("__p1"))),
                            op(BinOp::Add, ident("__q0"), mul(ident("__k"), ident("__q1"))),
                        ),
                    ),
                    assign(ident("__b2"), frac(ident("__p1"), ident("__q1"))),
                    // |b2 - self| <= |b1 - self|, cross-multiplied so no
                    // Fraction-valued abs is needed.
                    assign(
                        ident("__e2"),
                        call_global("__py_frac_dist", vec![ident("__b2"), ident("self")]),
                    ),
                    assign(
                        ident("__e1"),
                        call_global("__py_frac_dist", vec![ident("__b1"), ident("self")]),
                    ),
                    if_stmt(
                        op(BinOp::LtEq, ident("__e2"), ident("__e1")),
                        vec![ret(ident("__b2"))],
                    ),
                    ret(ident("__b1")),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // gcd, sign-normalised: the denominator is kept positive so every
        // comparison can cross-multiply without flipping.
        function(
            "__py_frac_gcd",
            vec![param("a", None.into()), param("b", None.into())],
            vec![
                if_stmt(
                    op(BinOp::Lt, ident("a"), i(0)),
                    vec![assign(ident("a"), sub(i(0), ident("a")))],
                ),
                if_stmt(
                    op(BinOp::Lt, ident("b"), i(0)),
                    vec![assign(ident("b"), sub(i(0), ident("b")))],
                ),
                while_stmt(
                    op(BinOp::NotEq, ident("b"), i(0)),
                    vec![
                        assign(ident("t"), op(BinOp::Mod, ident("a"), ident("b"))),
                        assign(ident("a"), ident("b")),
                        assign(ident("b"), ident("t")),
                    ],
                ),
                ret(ident("a")),
            ],
        ),
        // `[numerator, denominator]` for whatever the constructor was handed.
        //
        // ⛔ NO `isinstance`. A spliced core class is never walked, so
        // `isinstance(value, float)` leaves `float` an unresolved identifier
        // and the test is always False — all three numeric legs missed and
        // `Fraction(0.5)` answered `0/1`. `hasattr` is no probe either: it
        // answers False for every primitive. The text of the value is the one
        // thing that needs no walk-time rewrite.
        function(
            "__py_frac_parts",
            vec![
                param("value", Some(null())),
                param("denominator", Some(null())),
            ],
            vec![
                assign(ident("n"), i(0)),
                assign(ident("d"), i(1)),
                if_stmt(
                    call_global("hasattr", vec![ident("value"), str_lit("denominator")]),
                    vec![
                        assign(ident("n"), read_attr(ident("value"), "numerator")),
                        assign(ident("d"), read_attr(ident("value"), "denominator")),
                    ],
                ),
                if_stmt(
                    unary_not(call_global(
                        "hasattr",
                        vec![ident("value"), str_lit("denominator")],
                    )),
                    vec![
                        assign(
                            ident("p"),
                            call_global("__py_frac_parse", vec![ident("value")]),
                        ),
                        assign(ident("n"), index(ident("p"), i(0))),
                        assign(ident("d"), index(ident("p"), i(1))),
                    ],
                ),
                if_stmt(
                    is_not_none(ident("denominator")),
                    vec![assign(
                        ident("d"),
                        mul(
                            ident("d"),
                            call_global("__py_frac_int", vec![ident("denominator")]),
                        ),
                    )],
                ),
                if_stmt(
                    op(BinOp::Lt, ident("d"), i(0)),
                    vec![
                        assign(ident("n"), sub(i(0), ident("n"))),
                        assign(ident("d"), sub(i(0), ident("d"))),
                    ],
                ),
                assign(
                    ident("g"),
                    call_global("__py_frac_gcd", vec![ident("n"), ident("d")]),
                ),
                if_stmt(
                    op(BinOp::Gt, ident("g"), i(1)),
                    vec![
                        assign(ident("n"), op(BinOp::FloorDiv, ident("n"), ident("g"))),
                        assign(ident("d"), op(BinOp::FloorDiv, ident("d"), ident("g"))),
                    ],
                ),
                ret(list_of(vec![ident("n"), ident("d")])),
            ],
        ),
        // A denominator argument is itself a Rational in CPython; the corpus
        // only ever passes an int, and this keeps that honest.
        function(
            "__py_frac_int",
            vec![param("value", None.into())],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("value"), str_lit("denominator")]),
                    vec![ret(read_attr(ident("value"), "numerator"))],
                ),
                ret(ident("value")),
            ],
        ),
        // `"3/4"`, `"3.14159"` and `"12"` — the three spellings the
        // constructor accepts, plus whatever `str()` makes of a number.
        //
        // ⛔ `s.find(x)`, NOT `x in s`. `BinOp::In` on a string is a WALKER
        // rewrite, and a spliced class is never walked, so the membership test
        // answered False and every value fell through to `int(s)` — `"0.5"`
        // became `0`. Same failure shape as `isinstance` above.
        function(
            "__py_frac_parse",
            vec![param("value", Some(null()))],
            vec![
                assign(ident("s"), call_global("str", vec![ident("value")])),
                assign(
                    ident("bar"),
                    call(member(ident("s"), "find"), vec![str_lit("/")]),
                ),
                if_stmt(
                    op(BinOp::GtEq, ident("bar"), i(0)),
                    vec![ret(list_of(vec![
                        call_global("int", vec![slice_range(ident("s"), i(0), ident("bar"))]),
                        call_global(
                            "int",
                            vec![slice_from(ident("s"), op(BinOp::Add, ident("bar"), i(1)))],
                        ),
                    ]))],
                ),
                assign(
                    ident("dot"),
                    call(member(ident("s"), "find"), vec![str_lit(".")]),
                ),
                if_stmt(
                    op(BinOp::GtEq, ident("dot"), i(0)),
                    vec![
                        assign(ident("whole"), slice_range(ident("s"), i(0), ident("dot"))),
                        assign(
                            ident("frac"),
                            slice_from(ident("s"), op(BinOp::Add, ident("dot"), i(1))),
                        ),
                        assign(
                            ident("scale"),
                            op(BinOp::Pow, i(10), call_global("len", vec![ident("frac")])),
                        ),
                        assign(ident("digits"), call_global("int", vec![ident("frac")])),
                        if_stmt(
                            op(
                                BinOp::GtEq,
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
                            ident("scale"),
                        ])),
                    ],
                ),
                ret(list_of(vec![call_global("int", vec![ident("s")]), i(1)])),
            ],
        ),
        // A float is a binary rational: doubling until the value is integral
        // reproduces it EXACTLY, which is what `Fraction(0.5)` promises.
        function(
            "__py_frac_of_float",
            vec![param("x", None.into())],
            vec![
                assign(ident("d"), i(1)),
                while_stmt(
                    binary(
                        BinOp::And,
                        op(
                            BinOp::NotEq,
                            ident("x"),
                            call_global("int", vec![ident("x")]),
                        ),
                        op(BinOp::Lt, ident("d"), i(1000000000000000000)),
                    ),
                    vec![
                        assign(ident("x"), mul(ident("x"), i(2))),
                        assign(ident("d"), mul(ident("d"), i(2))),
                    ],
                ),
                ret(list_of(vec![
                    call_global("int", vec![ident("x")]),
                    ident("d"),
                ])),
            ],
        ),
        // |a - b| as a plain float — enough to ORDER two candidates.
        function(
            "__py_frac_dist",
            vec![param("a", None.into()), param("b", None.into())],
            vec![
                assign(
                    ident("v"),
                    sub(
                        op(
                            BinOp::Div,
                            read_attr(ident("a"), "numerator"),
                            read_attr(ident("a"), "denominator"),
                        ),
                        op(
                            BinOp::Div,
                            read_attr(ident("b"), "numerator"),
                            read_attr(ident("b"), "denominator"),
                        ),
                    ),
                ),
                if_stmt(
                    op(BinOp::Lt, ident("v"), i(0)),
                    vec![ret(sub(i(0), ident("v")))],
                ),
                ret(ident("v")),
            ],
        ),
        // Any operand of an arithmetic dunder, as a Fraction.
        function(
            "__py_frac_of",
            vec![param("value", None.into())],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("value"), str_lit("denominator")]),
                    vec![ret(ident("value"))],
                ),
                ret(new("Fraction", vec![ident("value")])),
            ],
        ),
    ]
}
