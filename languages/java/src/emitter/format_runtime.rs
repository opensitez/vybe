//! Java `PrintStream`/`Formatter` runtime prelude (JLS/`java.util.Formatter`).
//!
//! Mirrors the libc pattern (`platforms/libc/emitter/c_runtime.rs`): the
//! platform contributes runtime functions as common AST, prepended by the
//! walker. `System.out.print/append/printf/format` write into a line
//! buffer (`__j_buf`); every completed line (explicit `\n`, `%n`, or a
//! `println`) is flushed as ONE `println`-builtin call — byte-faithful
//! line semantics, matching real stdout observed line-by-line.
//!
//! `__j_sprintf` implements the Java-specific `Formatter` conversions the
//! shared `__fmt_sprintf` engine does not (or defines differently):
//! `%b`/`%B` (Boolean.toString), the `,` grouping flag, `%e`/`%E`
//! two-digit exponents, `%g`/`%G` (6 significant digits), `%n`, and
//! `%index$` argument selection — delegating every other conversion to
//! the shared engine via the `__java_string_format` builtin so existing
//! behavior is byte-identical.
//!
//! `__j_out` is the `PrintStream` identity sentinel: `System.out`
//! evaluates to it, and every write returns it (JLS: `append`/`format`
//! return `this`), so `ps.append("x") == ps` holds.

use vybe_ast::{
    Argument, ArrayElement, BinOp, BindingPattern, CatchClause, ClassMember, ClassModifiers,
    ConstructorInitializerTarget, ExprKind, Expression, Literal, Modifiers, ObjectProperty, Param,
    PassBy, Statement, StmtKind, VarDeclKind, VarDeclarator, Visibility,
};

pub const OUT_SENTINEL: &str = "__j_out";

fn stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}

fn expr(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn ident(name: &str) -> Expression {
    expr(ExprKind::Ident(name.to_string()))
}

fn str_lit(value: &str) -> Expression {
    expr(ExprKind::Lit(Literal::Str(value.to_string())))
}

fn int_lit(value: i64) -> Expression {
    expr(ExprKind::Lit(Literal::Int(value)))
}

fn expr_f64(value: f64) -> Expression {
    expr(ExprKind::Lit(Literal::Float(value)))
}

fn bool_lit(value: bool) -> Expression {
    expr(ExprKind::Lit(Literal::Bool(value)))
}

fn null_lit() -> Expression {
    expr(ExprKind::Lit(Literal::Null))
}

fn this_expr() -> Expression {
    expr(ExprKind::This)
}

fn undefined_lit() -> Expression {
    expr(ExprKind::Lit(Literal::Undefined))
}

fn arr_lit() -> Expression {
    expr(ExprKind::Array(Vec::new()))
}

fn array_lit(items: Vec<Expression>) -> Expression {
    expr(ExprKind::Array(
        items
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn obj_lit() -> Expression {
    expr(ExprKind::Object(Vec::new()))
}

fn obj_props(props: Vec<(&str, Expression)>) -> Expression {
    expr(ExprKind::Object(
        props
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: str_lit(key),
                value,
            })
            .collect(),
    ))
}

fn typeof_expr(value: Expression) -> Expression {
    expr(ExprKind::TypeOf(Box::new(value)))
}

fn new_expr(class_name: &str, args: Vec<Expression>) -> Expression {
    expr(ExprKind::New {
        class: Box::new(ident(class_name)),
        args: args.into_iter().map(Argument::positional).collect(),
    })
}

fn member(object: Expression, field: &str) -> Expression {
    expr(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn fld(object: &str, field: &str) -> Expression {
    member(ident(object), field)
}

fn index_expr(object: Expression, index: Expression) -> Expression {
    expr(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

fn call_expr(callee: Expression, args: Vec<Expression>) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn call(name: &str, args: Vec<Expression>) -> Expression {
    call_expr(ident(name), args)
}

fn call_member(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    call_expr(member(object, field), args)
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn add(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Add, left, right)
}

fn sub(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Sub, left, right)
}

fn mul(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Mul, left, right)
}

fn div(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Div, left, right)
}

fn pos_inf() -> Expression {
    binary(BinOp::Div, expr_f64(1.0), expr_f64(0.0))
}

fn neg_inf() -> Expression {
    binary(BinOp::Div, expr_f64(-1.0), expr_f64(0.0))
}

/// Java string conversion (null → "null", booleans → "true"/"false")
/// via the shared engine's `%s` — a bare-ident builtin call, which
/// dispatches anywhere (member-shaped `String.valueOf` does not inside
/// injected prelude functions). NOT `"" + x`: the dynamic add coerces
/// Bool→1 and Null→0.
fn to_str(x: Expression) -> Expression {
    call("__j_to_str", vec![x])
}

fn assign(target: Expression, value: Expression) -> Statement {
    stmt(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })
}

fn var_decl(name: &str, init: Expression) -> Statement {
    stmt(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    })
}

fn if_stmt(
    cond: Expression,
    then_body: Vec<Statement>,
    else_body: Option<Vec<Statement>>,
) -> Statement {
    stmt(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

fn ret(value: Expression) -> Statement {
    stmt(StmtKind::Return(Some(value)))
}

fn function_stmt(name: &str, params: Vec<&str>, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::FunctionDecl {
        name: name.to_string(),
        params: params
            .into_iter()
            .map(|param| Param {
                name: param.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect(),
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}

/// One-character substring `s.substring(i, i + 1)`.
fn char_at(s: Expression, i: Expression) -> Expression {
    call_member(s, "substring", vec![i.clone(), add(i, int_lit(1))])
}

/// The whole prelude, prepended to every Java compilation unit by the
/// walker (same as the libc runtime for C).
pub fn prelude() -> Vec<Statement> {
    static CACHE: std::sync::OnceLock<Vec<Statement>> = std::sync::OnceLock::new();
    CACHE.get_or_init(build_prelude).clone()
}

fn build_prelude() -> Vec<Statement> {
    let mut out = Vec::new();

    // PrintStream identity sentinel + pending-line buffer.
    out.push(var_decl(OUT_SENTINEL, str_lit(OUT_SENTINEL)));
    out.push(var_decl("__j_buf", str_lit("")));

    // BigDecimal mini-runtime: `{ v: Number, s: scale }`.
    out.push(function_stmt(
        "__j_bd_new",
        vec!["x"],
        vec![
            var_decl("s", to_str(ident("x"))),
            var_decl(
                "dot",
                call_member(ident("s"), "indexOf", vec![str_lit(".")]),
            ),
            var_decl("o", obj_lit()),
            assign(fld("o", "v"), call_expr(ident("Number"), vec![ident("s")])),
            assign(fld("o", "s"), int_lit(0)),
            if_stmt(
                binary(BinOp::GtEq, ident("dot"), int_lit(0)),
                vec![assign(
                    fld("o", "s"),
                    sub(member(ident("s"), "length"), add(ident("dot"), int_lit(1))),
                )],
                None,
            ),
            ret(ident("o")),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_box",
        vec!["v", "scale"],
        vec![
            var_decl("o", obj_lit()),
            assign(fld("o", "v"), ident("v")),
            assign(fld("o", "s"), ident("scale")),
            ret(ident("o")),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_round_half_up",
        vec!["v", "scale"],
        vec![
            var_decl(
                "m",
                call_member(ident("Math"), "pow", vec![int_lit(10), ident("scale")]),
            ),
            if_stmt(
                binary(BinOp::Lt, ident("v"), int_lit(0)),
                vec![ret(div(
                    call_member(
                        ident("Math"),
                        "ceil",
                        vec![sub(mul(ident("v"), ident("m")), expr_f64(0.5))],
                    ),
                    ident("m"),
                ))],
                None,
            ),
            ret(div(
                call_member(
                    ident("Math"),
                    "floor",
                    vec![add(mul(ident("v"), ident("m")), expr_f64(0.5))],
                ),
                ident("m"),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_to_string",
        vec!["a"],
        vec![ret(call("tofixed", vec![fld("a", "v"), fld("a", "s")]))],
    ));
    out.push(function_stmt(
        "__j_bd_to_plain_string",
        vec!["a"],
        vec![ret(call("__j_bd_to_string", vec![ident("a")]))],
    ));
    out.push(function_stmt(
        "__j_bd_add",
        vec!["a", "b"],
        vec![
            var_decl(
                "scale",
                call_member(ident("Math"), "max", vec![fld("a", "s"), fld("b", "s")]),
            ),
            ret(call(
                "__j_bd_box",
                vec![add(fld("a", "v"), fld("b", "v")), ident("scale")],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_subtract",
        vec!["a", "b"],
        vec![
            var_decl(
                "scale",
                call_member(ident("Math"), "max", vec![fld("a", "s"), fld("b", "s")]),
            ),
            ret(call(
                "__j_bd_box",
                vec![sub(fld("a", "v"), fld("b", "v")), ident("scale")],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_multiply",
        vec!["a", "b"],
        vec![ret(call(
            "__j_bd_box",
            vec![
                mul(fld("a", "v"), fld("b", "v")),
                add(fld("a", "s"), fld("b", "s")),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_divide",
        vec!["a", "b"],
        vec![ret(call(
            "__j_bd_box",
            vec![div(fld("a", "v"), fld("b", "v")), fld("a", "s")],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_divide_scale",
        vec!["a", "b", "scale", "mode"],
        vec![ret(call(
            "__j_bd_box",
            vec![
                call(
                    "__j_bd_round_half_up",
                    vec![div(fld("a", "v"), fld("b", "v")), ident("scale")],
                ),
                ident("scale"),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_set_scale",
        vec!["a", "scale", "mode"],
        vec![ret(call(
            "__j_bd_box",
            vec![
                call("__j_bd_round_half_up", vec![fld("a", "v"), ident("scale")]),
                ident("scale"),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_strip",
        vec!["a"],
        vec![ret(call("__j_bd_new", vec![to_str(fld("a", "v"))]))],
    ));
    out.push(function_stmt(
        "__j_bd_negate",
        vec!["a"],
        vec![ret(call(
            "__j_bd_box",
            vec![sub(int_lit(0), fld("a", "v")), fld("a", "s")],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_abs",
        vec!["a"],
        vec![ret(call(
            "__j_bd_box",
            vec![
                call_member(ident("Math"), "abs", vec![fld("a", "v")]),
                fld("a", "s"),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_plus",
        vec!["a"],
        vec![ret(ident("a"))],
    ));
    out.push(function_stmt(
        "__j_bd_scale",
        vec!["a"],
        vec![ret(fld("a", "s"))],
    ));
    out.push(function_stmt(
        "__j_bd_compare_to",
        vec!["a", "b"],
        vec![
            if_stmt(
                binary(BinOp::Lt, fld("a", "v"), fld("b", "v")),
                vec![ret(int_lit(-1))],
                None,
            ),
            if_stmt(
                binary(BinOp::Gt, fld("a", "v"), fld("b", "v")),
                vec![ret(int_lit(1))],
                None,
            ),
            ret(int_lit(0)),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_equals",
        vec!["a", "b"],
        vec![ret(binary(
            BinOp::And,
            binary(BinOp::Eq, fld("a", "v"), fld("b", "v")),
            binary(BinOp::Eq, fld("a", "s"), fld("b", "s")),
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_signum",
        vec!["a"],
        vec![
            if_stmt(
                binary(BinOp::Lt, fld("a", "v"), int_lit(0)),
                vec![ret(int_lit(-1))],
                None,
            ),
            if_stmt(
                binary(BinOp::Gt, fld("a", "v"), int_lit(0)),
                vec![ret(int_lit(1))],
                None,
            ),
            ret(int_lit(0)),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_move_right",
        vec!["a", "n"],
        vec![ret(call(
            "__j_bd_box",
            vec![
                mul(
                    fld("a", "v"),
                    call_member(ident("Math"), "pow", vec![int_lit(10), ident("n")]),
                ),
                call_member(
                    ident("Math"),
                    "max",
                    vec![int_lit(0), sub(fld("a", "s"), ident("n"))],
                ),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_move_left",
        vec!["a", "n"],
        vec![ret(call(
            "__j_bd_box",
            vec![
                div(
                    fld("a", "v"),
                    call_member(ident("Math"), "pow", vec![int_lit(10), ident("n")]),
                ),
                add(fld("a", "s"), ident("n")),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_bd_unscaled",
        vec!["a"],
        vec![ret(to_str(call_member(
            ident("Math"),
            "round",
            vec![mul(
                fld("a", "v"),
                call_member(ident("Math"), "pow", vec![int_lit(10), fld("a", "s")]),
            )],
        )))],
    ));
    out.push(function_stmt(
        "__j_bd_precision",
        vec!["a"],
        vec![
            var_decl("s", call("__j_bd_to_string", vec![ident("a")])),
            assign(
                ident("s"),
                call_member(ident("s"), "replace", vec![str_lit("-"), str_lit("")]),
            ),
            assign(
                ident("s"),
                call_member(ident("s"), "replace", vec![str_lit("."), str_lit("")]),
            ),
            ret(member(ident("s"), "length")),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_max",
        vec!["a", "b"],
        vec![
            if_stmt(
                binary(BinOp::GtEq, fld("a", "v"), fld("b", "v")),
                vec![ret(ident("a"))],
                None,
            ),
            ret(ident("b")),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_min",
        vec!["a", "b"],
        vec![
            if_stmt(
                binary(BinOp::LtEq, fld("a", "v"), fld("b", "v")),
                vec![ret(ident("a"))],
                None,
            ),
            ret(ident("b")),
        ],
    ));
    out.push(function_stmt(
        "__j_bd_remainder",
        vec!["a", "b"],
        vec![ret(call(
            "__j_bd_box",
            vec![
                binary(BinOp::Mod, fld("a", "v"), fld("b", "v")),
                fld("a", "s"),
            ],
        ))],
    ));

    // DecimalFormat/NumberFormat mini-runtime for US-style numeric patterns.
    out.push(function_stmt(
        "__j_df_symbols",
        vec!["loc"],
        vec![ret(obj_props(vec![("locale", ident("loc"))]))],
    ));
    out.push(function_stmt(
        "__j_df_apply",
        vec!["d", "p"],
        vec![
            assign(fld("d", "pattern"), ident("p")),
            assign(fld("d", "grouping"), bool_lit(false)),
            assign(fld("d", "minInt"), int_lit(1)),
            assign(fld("d", "minFrac"), int_lit(0)),
            assign(fld("d", "maxFrac"), int_lit(0)),
            assign(fld("d", "multiplier"), int_lit(1)),
            assign(fld("d", "prefix"), str_lit("")),
            assign(fld("d", "suffix"), str_lit("")),
            assign(fld("d", "negParen"), bool_lit(false)),
            assign(fld("d", "scientific"), bool_lit(false)),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#,##0.00")),
                vec![
                    assign(fld("d", "grouping"), bool_lit(true)),
                    assign(fld("d", "minFrac"), int_lit(2)),
                    assign(fld("d", "maxFrac"), int_lit(2)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#,##0.00;(#,##0.00)")),
                vec![
                    assign(fld("d", "grouping"), bool_lit(true)),
                    assign(fld("d", "minFrac"), int_lit(2)),
                    assign(fld("d", "maxFrac"), int_lit(2)),
                    assign(fld("d", "negParen"), bool_lit(true)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.00")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(2)),
                    assign(fld("d", "maxFrac"), int_lit(2)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#0.00")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(2)),
                    assign(fld("d", "maxFrac"), int_lit(2)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.000")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(3)),
                    assign(fld("d", "maxFrac"), int_lit(3)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.0000")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(4)),
                    assign(fld("d", "maxFrac"), int_lit(4)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.0")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(1)),
                    assign(fld("d", "maxFrac"), int_lit(1)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("00.0")),
                vec![
                    assign(fld("d", "minInt"), int_lit(2)),
                    assign(fld("d", "minFrac"), int_lit(1)),
                    assign(fld("d", "maxFrac"), int_lit(1)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0000")),
                vec![assign(fld("d", "minInt"), int_lit(4))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.###")),
                vec![assign(fld("d", "maxFrac"), int_lit(3))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#.##")),
                vec![
                    assign(fld("d", "minInt"), int_lit(0)),
                    assign(fld("d", "maxFrac"), int_lit(2)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#,##0")),
                vec![assign(fld("d", "grouping"), bool_lit(true))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#,##0.###")),
                vec![
                    assign(fld("d", "grouping"), bool_lit(true)),
                    assign(fld("d", "maxFrac"), int_lit(3)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#,##0.00")),
                vec![assign(fld("d", "grouping"), bool_lit(true))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("#0%")),
                vec![
                    assign(fld("d", "multiplier"), int_lit(100)),
                    assign(fld("d", "suffix"), str_lit("%")),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0%")),
                vec![
                    assign(fld("d", "multiplier"), int_lit(100)),
                    assign(fld("d", "suffix"), str_lit("%")),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.0%")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(1)),
                    assign(fld("d", "maxFrac"), int_lit(1)),
                    assign(fld("d", "multiplier"), int_lit(100)),
                    assign(fld("d", "suffix"), str_lit("%")),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.##%")),
                vec![
                    assign(fld("d", "maxFrac"), int_lit(2)),
                    assign(fld("d", "multiplier"), int_lit(100)),
                    assign(fld("d", "suffix"), str_lit("%")),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.0‰")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(1)),
                    assign(fld("d", "maxFrac"), int_lit(1)),
                    assign(fld("d", "multiplier"), int_lit(1000)),
                    assign(fld("d", "suffix"), str_lit("‰")),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("¤#,##0.00")),
                vec![
                    assign(fld("d", "grouping"), bool_lit(true)),
                    assign(fld("d", "minFrac"), int_lit(2)),
                    assign(fld("d", "maxFrac"), int_lit(2)),
                    assign(fld("d", "prefix"), str_lit("$")),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("0.000E0")),
                vec![
                    assign(fld("d", "minFrac"), int_lit(3)),
                    assign(fld("d", "maxFrac"), int_lit(3)),
                    assign(fld("d", "scientific"), bool_lit(true)),
                ],
                None,
            ),
            ret(ident("d")),
        ],
    ));
    out.push(function_stmt(
        "__j_df_new",
        vec!["p", "sym"],
        vec![
            var_decl(
                "d",
                obj_props(vec![
                    ("pattern", str_lit("")),
                    ("locale", str_lit("US")),
                    ("grouping", bool_lit(false)),
                    ("minInt", int_lit(1)),
                    ("minFrac", int_lit(0)),
                    ("maxFrac", int_lit(0)),
                    ("multiplier", int_lit(1)),
                    ("prefix", str_lit("")),
                    ("suffix", str_lit("")),
                    ("negParen", bool_lit(false)),
                    ("decimalAlways", bool_lit(false)),
                    ("parseIntegerOnly", bool_lit(false)),
                    ("scientific", bool_lit(false)),
                ]),
            ),
            stmt(StmtKind::Expr(call(
                "__j_df_apply",
                vec![ident("d"), ident("p")],
            ))),
            if_stmt(
                binary(BinOp::NotEq, ident("sym"), undefined_lit()),
                vec![assign(fld("d", "locale"), fld("sym", "locale"))],
                None,
            ),
            ret(ident("d")),
        ],
    ));
    out.push(function_stmt(
        "__j_df_currency",
        vec!["loc"],
        vec![ret(call(
            "__j_df_new",
            vec![
                str_lit("¤#,##0.00"),
                call("__j_df_symbols", vec![ident("loc")]),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_df_percent",
        vec!["loc"],
        vec![ret(call(
            "__j_df_new",
            vec![str_lit("0%"), call("__j_df_symbols", vec![ident("loc")])],
        ))],
    ));
    out.push(function_stmt(
        "__j_df_group",
        vec!["s"],
        vec![
            var_decl("out", str_lit("")),
            var_decl("i", sub(member(ident("s"), "length"), int_lit(1))),
            var_decl("n", int_lit(0)),
            while_stmt(
                binary(BinOp::GtEq, ident("i"), int_lit(0)),
                vec![
                    assign(
                        ident("out"),
                        add(
                            call_member(
                                ident("s"),
                                "substring",
                                vec![ident("i"), add(ident("i"), int_lit(1))],
                            ),
                            ident("out"),
                        ),
                    ),
                    assign(ident("n"), add(ident("n"), int_lit(1))),
                    assign(ident("i"), sub(ident("i"), int_lit(1))),
                    if_stmt(
                        binary(
                            BinOp::And,
                            binary(
                                BinOp::Eq,
                                binary(BinOp::Mod, ident("n"), int_lit(3)),
                                int_lit(0),
                            ),
                            binary(BinOp::GtEq, ident("i"), int_lit(0)),
                        ),
                        vec![assign(ident("out"), add(str_lit(","), ident("out")))],
                        None,
                    ),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_replace_all",
        vec!["s", "from", "to"],
        vec![
            assign(ident("s"), to_str(ident("s"))),
            while_stmt(
                binary(
                    BinOp::GtEq,
                    call_member(ident("s"), "indexOf", vec![ident("from")]),
                    int_lit(0),
                ),
                vec![assign(
                    ident("s"),
                    call_member(ident("s"), "replace", vec![ident("from"), ident("to")]),
                )],
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_locale_number",
        vec!["s", "loc"],
        vec![
            assign(ident("s"), to_str(ident("s"))),
            if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, ident("loc"), str_lit("FR")),
                    binary(BinOp::Eq, ident("loc"), str_lit("FR_CA")),
                ),
                vec![
                    assign(
                        ident("s"),
                        call(
                            "__j_replace_all",
                            vec![ident("s"), str_lit(","), str_lit("@@JCOMMA@@")],
                        ),
                    ),
                    assign(
                        ident("s"),
                        call(
                            "__j_replace_all",
                            vec![ident("s"), str_lit("."), str_lit(",")],
                        ),
                    ),
                    assign(
                        ident("s"),
                        call(
                            "__j_replace_all",
                            vec![ident("s"), str_lit("@@JCOMMA@@"), str_lit("\u{202f}")],
                        ),
                    ),
                    ret(ident("s")),
                ],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, ident("loc"), str_lit("DE")),
                    binary(BinOp::Eq, ident("loc"), str_lit("IT")),
                ),
                vec![
                    assign(
                        ident("s"),
                        call(
                            "__j_replace_all",
                            vec![ident("s"), str_lit(","), str_lit("@@JCOMMA@@")],
                        ),
                    ),
                    assign(
                        ident("s"),
                        call(
                            "__j_replace_all",
                            vec![ident("s"), str_lit("."), str_lit(",")],
                        ),
                    ),
                    assign(
                        ident("s"),
                        call(
                            "__j_replace_all",
                            vec![ident("s"), str_lit("@@JCOMMA@@"), str_lit(".")],
                        ),
                    ),
                    ret(ident("s")),
                ],
                None,
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_df_trim",
        vec!["s", "min"],
        vec![
            var_decl(
                "dot",
                call_member(ident("s"), "indexOf", vec![str_lit(".")]),
            ),
            if_stmt(
                binary(BinOp::Lt, ident("dot"), int_lit(0)),
                vec![ret(ident("s"))],
                None,
            ),
            while_stmt(
                binary(
                    BinOp::And,
                    binary(
                        BinOp::Gt,
                        sub(member(ident("s"), "length"), add(ident("dot"), int_lit(1))),
                        ident("min"),
                    ),
                    binary(
                        BinOp::Eq,
                        call_member(
                            ident("s"),
                            "substring",
                            vec![sub(member(ident("s"), "length"), int_lit(1))],
                        ),
                        str_lit("0"),
                    ),
                ),
                vec![assign(
                    ident("s"),
                    call_member(
                        ident("s"),
                        "substring",
                        vec![int_lit(0), sub(member(ident("s"), "length"), int_lit(1))],
                    ),
                )],
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Eq, ident("min"), int_lit(0)),
                    binary(
                        BinOp::Eq,
                        call_member(
                            ident("s"),
                            "substring",
                            vec![sub(member(ident("s"), "length"), int_lit(1))],
                        ),
                        str_lit("."),
                    ),
                ),
                vec![assign(
                    ident("s"),
                    call_member(
                        ident("s"),
                        "substring",
                        vec![int_lit(0), sub(member(ident("s"), "length"), int_lit(1))],
                    ),
                )],
                None,
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_df_format",
        vec!["d", "x"],
        vec![
            var_decl(
                "v",
                mul(
                    call_expr(ident("Number"), vec![ident("x")]),
                    fld("d", "multiplier"),
                ),
            ),
            var_decl("neg", binary(BinOp::Lt, ident("v"), int_lit(0))),
            if_stmt(
                ident("neg"),
                vec![assign(
                    ident("v"),
                    call_member(ident("Math"), "abs", vec![ident("v")]),
                )],
                None,
            ),
            if_stmt(
                fld("d", "scientific"),
                vec![
                    var_decl(
                        "e",
                        call_member(
                            ident("Math"),
                            "floor",
                            vec![div(
                                call_member(ident("Math"), "log", vec![ident("v")]),
                                call_member(ident("Math"), "log", vec![int_lit(10)]),
                            )],
                        ),
                    ),
                    var_decl(
                        "m",
                        div(
                            ident("v"),
                            call_member(ident("Math"), "pow", vec![int_lit(10), ident("e")]),
                        ),
                    ),
                    assign(
                        ident("m"),
                        call(
                            "__j_bd_round_half_up",
                            vec![ident("m"), fld("d", "maxFrac")],
                        ),
                    ),
                    var_decl("s", call("tofixed", vec![ident("m"), fld("d", "maxFrac")])),
                    ret(add(add(ident("s"), str_lit("E")), to_str(ident("e")))),
                ],
                None,
            ),
            assign(
                ident("v"),
                call(
                    "__j_bd_round_half_up",
                    vec![ident("v"), fld("d", "maxFrac")],
                ),
            ),
            var_decl("s", call("tofixed", vec![ident("v"), fld("d", "maxFrac")])),
            assign(
                ident("s"),
                call("__j_df_trim", vec![ident("s"), fld("d", "minFrac")]),
            ),
            var_decl(
                "dot",
                call_member(ident("s"), "indexOf", vec![str_lit(".")]),
            ),
            var_decl("intp", ident("s")),
            var_decl("frac", str_lit("")),
            if_stmt(
                binary(BinOp::GtEq, ident("dot"), int_lit(0)),
                vec![
                    assign(
                        ident("intp"),
                        call_member(ident("s"), "substring", vec![int_lit(0), ident("dot")]),
                    ),
                    assign(
                        ident("frac"),
                        call_member(ident("s"), "substring", vec![ident("dot")]),
                    ),
                ],
                None,
            ),
            while_stmt(
                binary(
                    BinOp::Lt,
                    member(ident("intp"), "length"),
                    fld("d", "minInt"),
                ),
                vec![assign(ident("intp"), add(str_lit("0"), ident("intp")))],
            ),
            if_stmt(
                fld("d", "grouping"),
                vec![assign(
                    ident("intp"),
                    call("__j_df_group", vec![ident("intp")]),
                )],
                None,
            ),
            assign(ident("s"), add(ident("intp"), ident("frac"))),
            if_stmt(
                binary(
                    BinOp::And,
                    fld("d", "decimalAlways"),
                    binary(
                        BinOp::Lt,
                        call_member(ident("s"), "indexOf", vec![str_lit(".")]),
                        int_lit(0),
                    ),
                ),
                vec![assign(ident("s"), add(ident("s"), str_lit(".")))],
                None,
            ),
            assign(
                ident("s"),
                add(add(fld("d", "prefix"), ident("s")), fld("d", "suffix")),
            ),
            if_stmt(
                ident("neg"),
                vec![if_stmt(
                    fld("d", "negParen"),
                    vec![assign(
                        ident("s"),
                        add(add(str_lit("("), ident("s")), str_lit(")")),
                    )],
                    Some(vec![assign(ident("s"), add(str_lit("-"), ident("s")))]),
                )],
                None,
            ),
            ret(call(
                "__j_locale_number",
                vec![ident("s"), fld("d", "locale")],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_df_parse",
        vec!["d", "s"],
        vec![
            assign(ident("s"), to_str(ident("s"))),
            if_stmt(
                binary(
                    BinOp::And,
                    fld("d", "parseIntegerOnly"),
                    binary(
                        BinOp::GtEq,
                        call_member(ident("s"), "indexOf", vec![str_lit(".")]),
                        int_lit(0),
                    ),
                ),
                vec![stmt(StmtKind::Throw {
                    expr: Some(call(
                        "__j_exc",
                        vec![
                            str_lit("ParseException"),
                            array_lit(vec![
                                str_lit("java.text.ParseException"),
                                str_lit("ParseException"),
                                str_lit("Exception"),
                                str_lit("Throwable"),
                            ]),
                            str_lit("parse error"),
                            undefined_lit(),
                        ],
                    )),
                    cause: None,
                })],
                None,
            ),
            var_decl("neg", bool_lit(false)),
            if_stmt(
                binary(
                    BinOp::Eq,
                    call_member(ident("s"), "indexOf", vec![str_lit("(")]),
                    int_lit(0),
                ),
                vec![
                    assign(ident("neg"), bool_lit(true)),
                    assign(
                        ident("s"),
                        call_member(
                            ident("s"),
                            "substring",
                            vec![int_lit(1), sub(member(ident("s"), "length"), int_lit(1))],
                        ),
                    ),
                ],
                None,
            ),
            while_stmt(
                binary(
                    BinOp::GtEq,
                    call_member(ident("s"), "indexOf", vec![str_lit(",")]),
                    int_lit(0),
                ),
                vec![assign(
                    ident("s"),
                    call_member(ident("s"), "replace", vec![str_lit(","), str_lit("")]),
                )],
            ),
            assign(
                ident("s"),
                call_member(ident("s"), "replace", vec![str_lit("$"), str_lit("")]),
            ),
            assign(
                ident("s"),
                call_member(ident("s"), "replace", vec![str_lit("%"), str_lit("")]),
            ),
            assign(
                ident("s"),
                call_member(ident("s"), "replace", vec![str_lit("‰"), str_lit("")]),
            ),
            var_decl(
                "n",
                div(
                    call_expr(ident("Number"), vec![ident("s")]),
                    fld("d", "multiplier"),
                ),
            ),
            if_stmt(
                ident("neg"),
                vec![assign(ident("n"), sub(int_lit(0), ident("n")))],
                None,
            ),
            ret(ident("n")),
        ],
    ));
    out.push(function_stmt(
        "__j_num_double",
        vec![],
        vec![ret(member(this_expr(), "v"))],
    ));
    out.push(function_stmt(
        "__j_num_long",
        vec![],
        vec![ret(call_member(
            ident("Math"),
            "trunc",
            vec![member(this_expr(), "v")],
        ))],
    ));
    out.push(function_stmt(
        "__j_df_min_frac",
        vec!["d"],
        vec![ret(fld("d", "minFrac"))],
    ));
    out.push(function_stmt(
        "__j_df_max_frac",
        vec!["d"],
        vec![ret(fld("d", "maxFrac"))],
    ));
    out.push(function_stmt(
        "__j_df_set_min_frac",
        vec!["d", "n"],
        vec![
            assign(fld("d", "minFrac"), ident("n")),
            if_stmt(
                binary(BinOp::Lt, fld("d", "maxFrac"), ident("n")),
                vec![assign(fld("d", "maxFrac"), ident("n"))],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_df_set_max_frac",
        vec!["d", "n"],
        vec![assign(fld("d", "maxFrac"), ident("n")), ret(null_lit())],
    ));
    out.push(function_stmt(
        "__j_df_grouping",
        vec!["d"],
        vec![ret(fld("d", "grouping"))],
    ));
    out.push(function_stmt(
        "__j_df_set_grouping",
        vec!["d", "v"],
        vec![assign(fld("d", "grouping"), ident("v")), ret(null_lit())],
    ));
    out.push(function_stmt(
        "__j_df_apply_pattern",
        vec!["d", "p"],
        vec![
            stmt(StmtKind::Expr(call(
                "__j_df_apply",
                vec![ident("d"), ident("p")],
            ))),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_df_pattern",
        vec!["d"],
        vec![ret(fld("d", "pattern"))],
    ));
    out.push(function_stmt(
        "__j_df_set_decimal_always",
        vec!["d", "v"],
        vec![
            assign(fld("d", "decimalAlways"), ident("v")),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_df_multiplier",
        vec!["d"],
        vec![ret(fld("d", "multiplier"))],
    ));
    out.push(function_stmt(
        "__j_df_set_multiplier",
        vec!["d", "v"],
        vec![assign(fld("d", "multiplier"), ident("v")), ret(null_lit())],
    ));
    out.push(function_stmt(
        "__j_df_set_parse_integer",
        vec!["d", "v"],
        vec![
            assign(fld("d", "parseIntegerOnly"), ident("v")),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_df_clone",
        vec!["d"],
        vec![ret(obj_props(vec![
            ("pattern", fld("d", "pattern")),
            ("locale", fld("d", "locale")),
            ("grouping", fld("d", "grouping")),
            ("minInt", fld("d", "minInt")),
            ("minFrac", fld("d", "minFrac")),
            ("maxFrac", fld("d", "maxFrac")),
            ("multiplier", fld("d", "multiplier")),
            ("prefix", fld("d", "prefix")),
            ("suffix", fld("d", "suffix")),
            ("negParen", fld("d", "negParen")),
            ("decimalAlways", fld("d", "decimalAlways")),
            ("parseIntegerOnly", fld("d", "parseIntegerOnly")),
            ("scientific", fld("d", "scientific")),
        ]))],
    ));
    out.push(function_stmt(
        "__j_df_equals",
        vec!["a", "b"],
        vec![ret(binary(
            BinOp::Eq,
            fld("a", "pattern"),
            fld("b", "pattern"),
        ))],
    ));

    out.push(function_stmt(
        "__j_to_str",
        vec!["x"],
        vec![
            if_stmt(
                binary(BinOp::Eq, typeof_expr(ident("x")), str_lit("string")),
                vec![ret(ident("x"))],
                None,
            ),
            ret(call(
                "__java_string_format",
                vec![str_lit("%s"), ident("x")],
            )),
        ],
    ));

    // __j_print(x): buffer, flush each completed line (its own '\n'
    // included) byte-faithfully to real stdout — `__j_write` is the
    // libc `write_stdout` intrinsic (wasi:cli/stdout get-stdout +
    // wasi:io/streams blocking-write-and-flush), NOT wasi:logging.
    out.push(function_stmt(
        "__j_print",
        vec!["x"],
        vec![
            assign(ident("__j_buf"), add(ident("__j_buf"), to_str(ident("x")))),
            var_decl(
                "i",
                call_member(ident("__j_buf"), "indexOf", vec![str_lit("\n")]),
            ),
            while_stmt(
                binary(BinOp::GtEq, ident("i"), int_lit(0)),
                vec![
                    stmt(StmtKind::Expr(call(
                        "__j_write",
                        vec![call_member(
                            ident("__j_buf"),
                            "substring",
                            vec![int_lit(0), add(ident("i"), int_lit(1))],
                        )],
                    ))),
                    assign(
                        ident("__j_buf"),
                        call_member(
                            ident("__j_buf"),
                            "substring",
                            vec![add(ident("i"), int_lit(1))],
                        ),
                    ),
                    assign(
                        ident("i"),
                        call_member(ident("__j_buf"), "indexOf", vec![str_lit("\n")]),
                    ),
                ],
            ),
            ret(ident(OUT_SENTINEL)),
        ],
    ));

    // __j_println(x): complete the current line.
    out.push(function_stmt(
        "__j_println",
        vec!["x"],
        vec![
            stmt(StmtKind::Expr(call(
                "__j_print",
                vec![add(to_str(ident("x")), str_lit("\n"))],
            ))),
            ret(ident(OUT_SENTINEL)),
        ],
    ));

    // __j_printf(fmt, args): format then write, no newline of its own.
    out.push(function_stmt(
        "__j_printf",
        vec!["fmt", "args"],
        vec![
            stmt(StmtKind::Expr(call(
                "__j_print",
                vec![call("__j_sprintf", vec![ident("fmt"), ident("args")])],
            ))),
            ret(ident(OUT_SENTINEL)),
        ],
    ));

    out.push(function_stmt(
        "__j_sprintf_locale",
        vec!["fmt", "args", "loc"],
        vec![ret(call(
            "__j_locale_number",
            vec![
                call("__j_sprintf", vec![ident("fmt"), ident("args")]),
                ident("loc"),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_fmt_new",
        vec!["a", "b"],
        vec![
            var_decl("f", obj_lit()),
            assign(fld("f", "buf"), str_lit("")),
            assign(fld("f", "loc"), str_lit("US")),
            assign(fld("f", "out"), null_lit()),
            if_stmt(
                binary(BinOp::Eq, ident("a"), undefined_lit()),
                vec![ret(ident("f"))],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("b"), undefined_lit()),
                vec![
                    assign(fld("f", "out"), ident("a")),
                    assign(fld("f", "loc"), ident("b")),
                    ret(ident("f")),
                ],
                None,
            ),
            assign(fld("f", "loc"), ident("a")),
            ret(ident("f")),
        ],
    ));
    out.push(function_stmt(
        "__j_fmt_format",
        vec!["f", "fmt", "args"],
        vec![
            var_decl(
                "s",
                call(
                    "__j_sprintf_locale",
                    vec![ident("fmt"), ident("args"), fld("f", "loc")],
                ),
            ),
            assign(fld("f", "buf"), add(fld("f", "buf"), ident("s"))),
            if_stmt(
                binary(BinOp::NotEq, fld("f", "out"), null_lit()),
                vec![stmt(StmtKind::Expr(call(
                    "__j_sb_append",
                    vec![fld("f", "out"), ident("s")],
                )))],
                None,
            ),
            ret(ident("f")),
        ],
    ));
    out.push(function_stmt(
        "__j_fmt_to_string",
        vec!["f"],
        vec![ret(fld("f", "buf"))],
    ));
    out.push(function_stmt(
        "__j_fmt_locale",
        vec!["f"],
        vec![ret(fld("f", "loc"))],
    ));
    out.push(function_stmt(
        "__j_fmt_out",
        vec!["f"],
        vec![ret(ident("f"))],
    ));

    out.push(function_stmt(
        "__j_tz_get",
        vec!["id"],
        vec![ret(ident("id"))],
    ));
    out.push(function_stmt(
        "__j_cal_new",
        vec!["tz", "loc"],
        vec![ret(obj_props(vec![
            ("y", int_lit(1970)),
            ("mo", int_lit(0)),
            ("d", int_lit(1)),
            ("h", int_lit(0)),
            ("mi", int_lit(0)),
            ("s", int_lit(0)),
            ("ms", int_lit(0)),
        ]))],
    ));
    out.push(function_stmt(
        "__j_cal_set",
        vec!["c", "a", "b", "d", "h", "mi", "s"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("d"), undefined_lit()),
                vec![
                    if_stmt(
                        binary(BinOp::Eq, ident("a"), int_lit(14)),
                        vec![assign(fld("c", "ms"), ident("b"))],
                        None,
                    ),
                    ret(null_lit()),
                ],
                None,
            ),
            assign(fld("c", "y"), ident("a")),
            assign(fld("c", "mo"), ident("b")),
            assign(fld("c", "d"), ident("d")),
            assign(fld("c", "h"), ident("h")),
            assign(fld("c", "mi"), ident("mi")),
            assign(fld("c", "s"), ident("s")),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_cal_get_time",
        vec!["c"],
        vec![ret(obj_props(vec![
            ("y", fld("c", "y")),
            ("mo", fld("c", "mo")),
            ("d", fld("c", "d")),
            ("h", fld("c", "h")),
            ("mi", fld("c", "mi")),
            ("s", fld("c", "s")),
            ("ms", fld("c", "ms")),
        ]))],
    ));

    out.push(function_stmt(
        "__j_mf_new",
        vec!["pattern", "loc"],
        vec![
            var_decl("m", obj_lit()),
            assign(fld("m", "pattern"), ident("pattern")),
            assign(fld("m", "loc"), ident("loc")),
            ret(ident("m")),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_static_format",
        vec!["pattern", "args"],
        vec![ret(call(
            "__j_mf_format",
            vec![
                call("__j_mf_new", vec![ident("pattern"), str_lit("US")]),
                ident("args"),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_mf_format",
        vec!["m", "args"],
        vec![ret(call(
            "__j_mf_format_pattern",
            vec![fld("m", "pattern"), ident("args"), fld("m", "loc")],
        ))],
    ));
    out.push(function_stmt(
        "__j_mf_apply_pattern",
        vec!["m", "pattern"],
        vec![
            assign(fld("m", "pattern"), ident("pattern")),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_to_pattern",
        vec!["m"],
        vec![ret(fld("m", "pattern"))],
    ));
    out.push(function_stmt(
        "__j_mf_set_locale",
        vec!["m", "loc"],
        vec![assign(fld("m", "loc"), ident("loc")), ret(null_lit())],
    ));
    out.push(function_stmt(
        "__j_mf_clone",
        vec!["m"],
        vec![ret(call(
            "__j_mf_new",
            vec![fld("m", "pattern"), fld("m", "loc")],
        ))],
    ));
    out.push(function_stmt(
        "__j_mf_equals",
        vec!["a", "b"],
        vec![ret(binary(
            BinOp::Eq,
            fld("a", "pattern"),
            fld("b", "pattern"),
        ))],
    ));
    out.push(function_stmt(
        "__j_mf_number",
        vec!["x", "style", "loc"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("style"), str_lit("integer")),
                vec![ret(call(
                    "__j_df_format",
                    vec![
                        call(
                            "__j_df_new",
                            vec![str_lit("#,##0"), call("__j_df_symbols", vec![ident("loc")])],
                        ),
                        call("__java_trunc_cast", vec![ident("x")]),
                    ],
                ))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("style"), str_lit("currency")),
                vec![ret(call(
                    "__j_df_format",
                    vec![call("__j_df_currency", vec![ident("loc")]), ident("x")],
                ))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("style"), str_lit("percent")),
                vec![ret(call(
                    "__j_df_format",
                    vec![call("__j_df_percent", vec![ident("loc")]), ident("x")],
                ))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("style"), str_lit("")),
                vec![assign(ident("style"), str_lit("#,##0.###"))],
                None,
            ),
            ret(call(
                "__j_df_format",
                vec![
                    call(
                        "__j_df_new",
                        vec![ident("style"), call("__j_df_symbols", vec![ident("loc")])],
                    ),
                    ident("x"),
                ],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_pad2",
        vec!["x"],
        vec![
            var_decl("s", to_str(ident("x"))),
            if_stmt(
                binary(BinOp::Lt, member(ident("s"), "length"), int_lit(2)),
                vec![assign(ident("s"), add(str_lit("0"), ident("s")))],
                None,
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_day",
        vec!["y", "m", "d"],
        vec![
            if_stmt(
                binary(BinOp::Lt, ident("m"), int_lit(3)),
                vec![
                    assign(ident("m"), add(ident("m"), int_lit(12))),
                    assign(ident("y"), sub(ident("y"), int_lit(1))),
                ],
                None,
            ),
            var_decl(
                "h",
                binary(
                    BinOp::Mod,
                    add(
                        add(
                            add(
                                add(
                                    add(
                                        ident("d"),
                                        call_member(
                                            ident("Math"),
                                            "floor",
                                            vec![div(
                                                mul(int_lit(13), add(ident("m"), int_lit(1))),
                                                int_lit(5),
                                            )],
                                        ),
                                    ),
                                    ident("y"),
                                ),
                                call_member(
                                    ident("Math"),
                                    "floor",
                                    vec![div(ident("y"), int_lit(4))],
                                ),
                            ),
                            sub(
                                int_lit(0),
                                call_member(
                                    ident("Math"),
                                    "floor",
                                    vec![div(ident("y"), int_lit(100))],
                                ),
                            ),
                        ),
                        call_member(ident("Math"), "floor", vec![div(ident("y"), int_lit(400))]),
                    ),
                    int_lit(7),
                ),
            ),
            ret(binary(BinOp::Mod, add(ident("h"), int_lit(6)), int_lit(7))),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_datetime",
        vec!["x", "kind", "style"],
        vec![
            var_decl(
                "iscal",
                binary(
                    BinOp::NotEq,
                    typeof_expr(fld("x", "y")),
                    str_lit("undefined"),
                ),
            ),
            var_decl("y", int_lit(0)),
            var_decl("mo", int_lit(0)),
            var_decl("d", int_lit(0)),
            var_decl("h", int_lit(0)),
            var_decl("mi", int_lit(0)),
            var_decl("se", int_lit(0)),
            if_stmt(
                ident("iscal"),
                vec![
                    assign(ident("y"), fld("x", "y")),
                    assign(ident("mo"), add(fld("x", "mo"), int_lit(1))),
                    assign(ident("d"), fld("x", "d")),
                    assign(ident("h"), fld("x", "h")),
                    assign(ident("mi"), fld("x", "mi")),
                    assign(ident("se"), fld("x", "s")),
                ],
                Some(vec![
                    assign(
                        ident("y"),
                        call_member(ident("x"), "getUTCFullYear", vec![]),
                    ),
                    assign(
                        ident("mo"),
                        add(call_member(ident("x"), "getUTCMonth", vec![]), int_lit(1)),
                    ),
                    assign(ident("d"), call_member(ident("x"), "getUTCDate", vec![])),
                    assign(ident("h"), call_member(ident("x"), "getUTCHours", vec![])),
                    assign(
                        ident("mi"),
                        call_member(ident("x"), "getUTCMinutes", vec![]),
                    ),
                    assign(
                        ident("se"),
                        call_member(ident("x"), "getUTCSeconds", vec![]),
                    ),
                ]),
            ),
            if_stmt(
                binary(BinOp::Eq, ident("kind"), str_lit("date")),
                vec![if_stmt(
                    binary(BinOp::Eq, ident("style"), str_lit("yyyy-MM-dd")),
                    vec![ret(add(
                        add(
                            add(to_str(ident("y")), str_lit("-")),
                            call("__j_mf_pad2", vec![ident("mo")]),
                        ),
                        add(str_lit("-"), call("__j_mf_pad2", vec![ident("d")])),
                    ))],
                    Some(vec![if_stmt(
                        binary(BinOp::Eq, ident("style"), str_lit("full")),
                        vec![ret(add(
                            add(
                                add(
                                    add(
                                        index_expr(
                                            array_lit(vec![
                                                str_lit("Sunday"),
                                                str_lit("Monday"),
                                                str_lit("Tuesday"),
                                                str_lit("Wednesday"),
                                                str_lit("Thursday"),
                                                str_lit("Friday"),
                                                str_lit("Saturday"),
                                            ]),
                                            call(
                                                "__j_mf_day",
                                                vec![ident("y"), ident("mo"), ident("d")],
                                            ),
                                        ),
                                        str_lit(", "),
                                    ),
                                    index_expr(
                                        array_lit(vec![
                                            str_lit("January"),
                                            str_lit("February"),
                                            str_lit("March"),
                                            str_lit("April"),
                                            str_lit("May"),
                                            str_lit("June"),
                                            str_lit("July"),
                                            str_lit("August"),
                                            str_lit("September"),
                                            str_lit("October"),
                                            str_lit("November"),
                                            str_lit("December"),
                                        ]),
                                        sub(ident("mo"), int_lit(1)),
                                    ),
                                ),
                                add(str_lit(" "), to_str(ident("d"))),
                            ),
                            add(str_lit(", "), to_str(ident("y"))),
                        ))],
                        Some(vec![ret(add(
                            add(add(to_str(ident("mo")), str_lit("/")), to_str(ident("d"))),
                            add(
                                str_lit("/"),
                                call_member(to_str(ident("y")), "substring", vec![int_lit(2)]),
                            ),
                        ))]),
                    )]),
                )],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("style"), str_lit("HH:mm:ss")),
                vec![ret(add(
                    add(
                        add(call("__j_mf_pad2", vec![ident("h")]), str_lit(":")),
                        call("__j_mf_pad2", vec![ident("mi")]),
                    ),
                    add(str_lit(":"), call("__j_mf_pad2", vec![ident("se")])),
                ))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("style"), str_lit("HH:mm")),
                vec![ret(add(
                    add(call("__j_mf_pad2", vec![ident("h")]), str_lit(":")),
                    call("__j_mf_pad2", vec![ident("mi")]),
                ))],
                None,
            ),
            var_decl("ap", str_lit("AM")),
            if_stmt(
                binary(BinOp::GtEq, ident("h"), int_lit(12)),
                vec![assign(ident("ap"), str_lit("PM"))],
                None,
            ),
            var_decl("hh", binary(BinOp::Mod, ident("h"), int_lit(12))),
            if_stmt(
                binary(BinOp::Eq, ident("hh"), int_lit(0)),
                vec![assign(ident("hh"), int_lit(12))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("style"), str_lit("medium")),
                vec![ret(add(
                    add(
                        add(
                            add(to_str(ident("hh")), str_lit(":")),
                            call("__j_mf_pad2", vec![ident("mi")]),
                        ),
                        add(str_lit(":"), call("__j_mf_pad2", vec![ident("se")])),
                    ),
                    add(str_lit(" "), ident("ap")),
                ))],
                Some(vec![ret(add(
                    add(
                        add(to_str(ident("hh")), str_lit(":")),
                        call("__j_mf_pad2", vec![ident("mi")]),
                    ),
                    add(str_lit(" "), ident("ap")),
                ))]),
            ),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_choice",
        vec!["x", "style", "args", "loc"],
        vec![
            var_decl("selected", str_lit("")),
            var_decl("start", int_lit(0)),
            var_decl("n", member(ident("style"), "length")),
            while_stmt(
                binary(BinOp::LtEq, ident("start"), ident("n")),
                vec![
                    var_decl(
                        "end",
                        call_member(
                            ident("style"),
                            "indexOf",
                            vec![str_lit("|"), ident("start")],
                        ),
                    ),
                    if_stmt(
                        binary(BinOp::Lt, ident("end"), int_lit(0)),
                        vec![assign(ident("end"), ident("n"))],
                        None,
                    ),
                    var_decl(
                        "part",
                        call_member(
                            ident("style"),
                            "substring",
                            vec![ident("start"), ident("end")],
                        ),
                    ),
                    var_decl(
                        "hash",
                        call_member(ident("part"), "indexOf", vec![str_lit("#")]),
                    ),
                    var_decl(
                        "lt",
                        call_member(ident("part"), "indexOf", vec![str_lit("<")]),
                    ),
                    var_decl("op", str_lit("#")),
                    var_decl("pos", ident("hash")),
                    if_stmt(
                        binary(
                            BinOp::Or,
                            binary(BinOp::Lt, ident("pos"), int_lit(0)),
                            binary(
                                BinOp::And,
                                binary(BinOp::GtEq, ident("lt"), int_lit(0)),
                                binary(BinOp::Lt, ident("lt"), ident("pos")),
                            ),
                        ),
                        vec![
                            assign(ident("op"), str_lit("<")),
                            assign(ident("pos"), ident("lt")),
                        ],
                        None,
                    ),
                    var_decl(
                        "limit",
                        call_expr(
                            ident("Number"),
                            vec![call_member(
                                ident("part"),
                                "substring",
                                vec![int_lit(0), ident("pos")],
                            )],
                        ),
                    ),
                    var_decl(
                        "text",
                        call_member(
                            ident("part"),
                            "substring",
                            vec![add(ident("pos"), int_lit(1))],
                        ),
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            binary(BinOp::Eq, ident("start"), int_lit(0)),
                            binary(BinOp::Lt, ident("x"), ident("limit")),
                        ),
                        vec![assign(ident("selected"), ident("text"))],
                        None,
                    ),
                    if_stmt(
                        binary(
                            BinOp::Or,
                            binary(
                                BinOp::And,
                                binary(BinOp::Eq, ident("op"), str_lit("#")),
                                binary(BinOp::Eq, ident("x"), ident("limit")),
                            ),
                            binary(
                                BinOp::And,
                                binary(BinOp::Eq, ident("op"), str_lit("<")),
                                binary(BinOp::Gt, ident("x"), ident("limit")),
                            ),
                        ),
                        vec![assign(ident("selected"), ident("text"))],
                        None,
                    ),
                    assign(ident("start"), add(ident("end"), int_lit(1))),
                ],
            ),
            ret(call(
                "__j_mf_format_pattern",
                vec![ident("selected"), ident("args"), ident("loc")],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_piece",
        vec!["spec", "args", "loc"],
        vec![
            var_decl(
                "c1",
                call_member(ident("spec"), "indexOf", vec![str_lit(",")]),
            ),
            var_decl("idxs", ident("spec")),
            if_stmt(
                binary(BinOp::GtEq, ident("c1"), int_lit(0)),
                vec![assign(
                    ident("idxs"),
                    call_member(ident("spec"), "substring", vec![int_lit(0), ident("c1")]),
                )],
                None,
            ),
            var_decl(
                "idx",
                call_member(ident("Integer"), "parseInt", vec![ident("idxs")]),
            ),
            var_decl("x", index_expr(ident("args"), ident("idx"))),
            if_stmt(
                binary(BinOp::Lt, ident("c1"), int_lit(0)),
                vec![ret(to_str(ident("x")))],
                None,
            ),
            var_decl(
                "rest",
                call_member(
                    ident("spec"),
                    "substring",
                    vec![add(ident("c1"), int_lit(1))],
                ),
            ),
            var_decl(
                "c2",
                call_member(ident("rest"), "indexOf", vec![str_lit(",")]),
            ),
            var_decl("typ", ident("rest")),
            var_decl("style", str_lit("")),
            if_stmt(
                binary(BinOp::GtEq, ident("c2"), int_lit(0)),
                vec![
                    assign(
                        ident("typ"),
                        call_member(ident("rest"), "substring", vec![int_lit(0), ident("c2")]),
                    ),
                    assign(
                        ident("style"),
                        call_member(
                            ident("rest"),
                            "substring",
                            vec![add(ident("c2"), int_lit(1))],
                        ),
                    ),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("typ"), str_lit("number")),
                vec![ret(call(
                    "__j_mf_number",
                    vec![ident("x"), ident("style"), ident("loc")],
                ))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("typ"), str_lit("choice")),
                vec![ret(call(
                    "__j_mf_choice",
                    vec![ident("x"), ident("style"), ident("args"), ident("loc")],
                ))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, ident("typ"), str_lit("date")),
                    binary(BinOp::Eq, ident("typ"), str_lit("time")),
                ),
                vec![ret(call(
                    "__j_mf_datetime",
                    vec![ident("x"), ident("typ"), ident("style")],
                ))],
                None,
            ),
            ret(to_str(ident("x"))),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_format_pattern",
        vec!["pattern", "args", "loc"],
        vec![
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            var_decl("n", member(ident("pattern"), "length")),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("n")),
                vec![
                    var_decl("c", char_at(ident("pattern"), ident("i"))),
                    if_stmt(
                        binary(BinOp::Eq, ident("c"), str_lit("'")),
                        vec![if_stmt(
                            binary(
                                BinOp::And,
                                binary(BinOp::Lt, add(ident("i"), int_lit(1)), ident("n")),
                                binary(
                                    BinOp::Eq,
                                    char_at(ident("pattern"), add(ident("i"), int_lit(1))),
                                    str_lit("'"),
                                ),
                            ),
                            vec![
                                assign(ident("out"), add(ident("out"), str_lit("'"))),
                                assign(ident("i"), add(ident("i"), int_lit(2))),
                            ],
                            Some(vec![
                                assign(ident("i"), add(ident("i"), int_lit(1))),
                                while_stmt(
                                    binary(
                                        BinOp::And,
                                        binary(BinOp::Lt, ident("i"), ident("n")),
                                        binary(
                                            BinOp::NotEq,
                                            char_at(ident("pattern"), ident("i")),
                                            str_lit("'"),
                                        ),
                                    ),
                                    vec![
                                        assign(
                                            ident("out"),
                                            add(
                                                ident("out"),
                                                char_at(ident("pattern"), ident("i")),
                                            ),
                                        ),
                                        assign(ident("i"), add(ident("i"), int_lit(1))),
                                    ],
                                ),
                                assign(ident("i"), add(ident("i"), int_lit(1))),
                            ]),
                        )],
                        Some(vec![if_stmt(
                            binary(BinOp::Eq, ident("c"), str_lit("{")),
                            vec![
                                var_decl("j", add(ident("i"), int_lit(1))),
                                var_decl("depth", int_lit(1)),
                                while_stmt(
                                    binary(
                                        BinOp::And,
                                        binary(BinOp::Lt, ident("j"), ident("n")),
                                        binary(BinOp::Gt, ident("depth"), int_lit(0)),
                                    ),
                                    vec![
                                        var_decl("ch", char_at(ident("pattern"), ident("j"))),
                                        if_stmt(
                                            binary(BinOp::Eq, ident("ch"), str_lit("{")),
                                            vec![assign(
                                                ident("depth"),
                                                add(ident("depth"), int_lit(1)),
                                            )],
                                            None,
                                        ),
                                        if_stmt(
                                            binary(BinOp::Eq, ident("ch"), str_lit("}")),
                                            vec![assign(
                                                ident("depth"),
                                                sub(ident("depth"), int_lit(1)),
                                            )],
                                            None,
                                        ),
                                        if_stmt(
                                            binary(BinOp::Gt, ident("depth"), int_lit(0)),
                                            vec![assign(ident("j"), add(ident("j"), int_lit(1)))],
                                            None,
                                        ),
                                    ],
                                ),
                                var_decl(
                                    "spec",
                                    call_member(
                                        ident("pattern"),
                                        "substring",
                                        vec![add(ident("i"), int_lit(1)), ident("j")],
                                    ),
                                ),
                                assign(
                                    ident("out"),
                                    add(
                                        ident("out"),
                                        call(
                                            "__j_mf_piece",
                                            vec![ident("spec"), ident("args"), ident("loc")],
                                        ),
                                    ),
                                ),
                                assign(ident("i"), add(ident("j"), int_lit(1))),
                            ],
                            Some(vec![
                                assign(ident("out"), add(ident("out"), ident("c"))),
                                assign(ident("i"), add(ident("i"), int_lit(1))),
                            ]),
                        )]),
                    ),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_mf_parse",
        vec!["m", "text"],
        vec![
            var_decl("res", arr_lit()),
            var_decl("p", fld("m", "pattern")),
            var_decl("pi", int_lit(0)),
            var_decl("si", int_lit(0)),
            var_decl("n", member(ident("p"), "length")),
            while_stmt(
                binary(BinOp::Lt, ident("pi"), ident("n")),
                vec![if_stmt(
                    binary(BinOp::Eq, char_at(ident("p"), ident("pi")), str_lit("{")),
                    vec![
                        var_decl(
                            "end",
                            call_member(ident("p"), "indexOf", vec![str_lit("}"), ident("pi")]),
                        ),
                        var_decl(
                            "spec",
                            call_member(
                                ident("p"),
                                "substring",
                                vec![add(ident("pi"), int_lit(1)), ident("end")],
                            ),
                        ),
                        var_decl(
                            "comma",
                            call_member(ident("spec"), "indexOf", vec![str_lit(",")]),
                        ),
                        if_stmt(
                            binary(BinOp::GtEq, ident("comma"), int_lit(0)),
                            vec![assign(
                                ident("spec"),
                                call_member(
                                    ident("spec"),
                                    "substring",
                                    vec![int_lit(0), ident("comma")],
                                ),
                            )],
                            None,
                        ),
                        var_decl(
                            "idx",
                            call_member(ident("Integer"), "parseInt", vec![ident("spec")]),
                        ),
                        assign(ident("pi"), add(ident("end"), int_lit(1))),
                        var_decl(
                            "next",
                            call_member(ident("p"), "indexOf", vec![str_lit("{"), ident("pi")]),
                        ),
                        if_stmt(
                            binary(BinOp::Lt, ident("next"), int_lit(0)),
                            vec![assign(ident("next"), ident("n"))],
                            None,
                        ),
                        var_decl(
                            "lit",
                            call_member(ident("p"), "substring", vec![ident("pi"), ident("next")]),
                        ),
                        var_decl("te", member(ident("text"), "length")),
                        if_stmt(
                            binary(BinOp::Gt, member(ident("lit"), "length"), int_lit(0)),
                            vec![assign(
                                ident("te"),
                                call_member(
                                    ident("text"),
                                    "indexOf",
                                    vec![ident("lit"), ident("si")],
                                ),
                            )],
                            None,
                        ),
                        assign(
                            index_expr(ident("res"), ident("idx")),
                            call_member(ident("text"), "substring", vec![ident("si"), ident("te")]),
                        ),
                        assign(
                            ident("si"),
                            add(ident("te"), member(ident("lit"), "length")),
                        ),
                        assign(ident("pi"), ident("next")),
                    ],
                    Some(vec![
                        assign(ident("pi"), add(ident("pi"), int_lit(1))),
                        assign(ident("si"), add(ident("si"), int_lit(1))),
                    ]),
                )],
            ),
            ret(ident("res")),
        ],
    ));

    // __j_isdig(c): "0" <= c <= "9".
    out.push(function_stmt(
        "__j_isdig",
        vec!["c"],
        vec![ret(binary(
            BinOp::And,
            binary(BinOp::GtEq, ident("c"), str_lit("0")),
            binary(BinOp::LtEq, ident("c"), str_lit("9")),
        ))],
    ));

    // __j_padw(s, width, left): pad to `width` (string, "" = none).
    out.push(function_stmt(
        "__j_padw",
        vec!["s", "width", "left"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("width"), str_lit("")),
                vec![ret(ident("s"))],
                None,
            ),
            var_decl(
                "w",
                call_member(ident("Integer"), "parseInt", vec![ident("width")]),
            ),
            assign(ident("s"), to_str(ident("s"))),
            while_stmt(
                binary(BinOp::Lt, member(ident("s"), "length"), ident("w")),
                vec![if_stmt(
                    binary(BinOp::Eq, ident("left"), int_lit(1)),
                    vec![assign(ident("s"), add(ident("s"), str_lit(" ")))],
                    Some(vec![assign(ident("s"), add(str_lit(" "), ident("s")))]),
                )],
            ),
            ret(ident("s")),
        ],
    ));

    // __j_group(s): thousands grouping ("1234567" → "1,234,567").
    out.push(function_stmt(
        "__j_group",
        vec!["s"],
        vec![
            assign(ident("s"), to_str(ident("s"))),
            var_decl("neg", int_lit(0)),
            if_stmt(
                binary(BinOp::Eq, char_at(ident("s"), int_lit(0)), str_lit("-")),
                vec![
                    assign(ident("neg"), int_lit(1)),
                    assign(
                        ident("s"),
                        call_member(ident("s"), "substring", vec![int_lit(1)]),
                    ),
                ],
                None,
            ),
            var_decl("grouped", str_lit("")),
            var_decl("n", member(ident("s"), "length")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("n")),
                vec![
                    if_stmt(
                        binary(
                            BinOp::And,
                            binary(BinOp::Gt, ident("i"), int_lit(0)),
                            binary(
                                BinOp::Eq,
                                binary(
                                    BinOp::Mod,
                                    binary(BinOp::Sub, ident("n"), ident("i")),
                                    int_lit(3),
                                ),
                                int_lit(0),
                            ),
                        ),
                        vec![assign(
                            ident("grouped"),
                            add(ident("grouped"), str_lit(",")),
                        )],
                        None,
                    ),
                    assign(
                        ident("grouped"),
                        add(ident("grouped"), char_at(ident("s"), ident("i"))),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            if_stmt(
                binary(BinOp::Eq, ident("neg"), int_lit(1)),
                vec![assign(
                    ident("grouped"),
                    add(str_lit("-"), ident("grouped")),
                )],
                None,
            ),
            ret(ident("grouped")),
        ],
    ));
    out.push(function_stmt(
        "__j_group_number",
        vec!["s"],
        vec![
            assign(ident("s"), to_str(ident("s"))),
            var_decl(
                "dot",
                call_member(ident("s"), "indexOf", vec![str_lit(".")]),
            ),
            if_stmt(
                binary(BinOp::Lt, ident("dot"), int_lit(0)),
                vec![ret(call("__j_group", vec![ident("s")]))],
                None,
            ),
            ret(add(
                call(
                    "__j_group",
                    vec![call_member(
                        ident("s"),
                        "substring",
                        vec![int_lit(0), ident("dot")],
                    )],
                ),
                call_member(ident("s"), "substring", vec![ident("dot")]),
            )),
        ],
    ));

    // __j_expad(s): Java prints 2-digit exponents ("e+3" → "e+03").
    out.push(function_stmt(
        "__j_expad",
        vec!["s"],
        vec![
            assign(ident("s"), to_str(ident("s"))),
            var_decl(
                "e",
                call_member(ident("s"), "lastIndexOf", vec![str_lit("e")]),
            ),
            if_stmt(
                binary(BinOp::Lt, ident("e"), int_lit(0)),
                vec![assign(
                    ident("e"),
                    call_member(ident("s"), "lastIndexOf", vec![str_lit("E")]),
                )],
                None,
            ),
            if_stmt(
                binary(BinOp::Lt, ident("e"), int_lit(0)),
                vec![ret(ident("s"))],
                None,
            ),
            // "…e+3": one digit after the sign → insert "0".
            if_stmt(
                binary(
                    BinOp::Eq,
                    binary(
                        BinOp::Sub,
                        member(ident("s"), "length"),
                        add(ident("e"), int_lit(2)),
                    ),
                    int_lit(1),
                ),
                vec![ret(add(
                    add(
                        call_member(
                            ident("s"),
                            "substring",
                            vec![int_lit(0), add(ident("e"), int_lit(2))],
                        ),
                        str_lit("0"),
                    ),
                    call_member(ident("s"), "substring", vec![add(ident("e"), int_lit(2))]),
                ))],
                None,
            ),
            ret(ident("s")),
        ],
    ));

    // __j_i32(x): wrap to signed 32-bit in float arithmetic. The dynamic
    // as_i32 coercion SATURATES (f64 2147483648 → i32::MAX), so high-bit
    // literals like 0x80000000 must wrap before hitting the i32 opcodes.
    out.push(function_stmt(
        "__j_i32",
        vec!["x"],
        vec![
            assign(
                ident("x"),
                binary(BinOp::Mod, ident("x"), expr_f64(4294967296.0)),
            ),
            if_stmt(
                binary(BinOp::GtEq, ident("x"), expr_f64(2147483648.0)),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Sub, ident("x"), expr_f64(4294967296.0)),
                )],
                None,
            ),
            if_stmt(
                binary(BinOp::Lt, ident("x"), expr_f64(-2147483648.0)),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Add, ident("x"), expr_f64(4294967296.0)),
                )],
                None,
            ),
            ret(ident("x")),
        ],
    ));
    out.push(function_stmt(
        "__j_byte",
        vec!["x"],
        vec![
            assign(ident("x"), call("__java_trunc_cast", vec![ident("x")])),
            if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Gt, ident("x"), int_lit(256)),
                    binary(BinOp::Lt, ident("x"), int_lit(-256)),
                ),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Mod, ident("x"), int_lit(256)),
                )],
                None,
            ),
            while_stmt(
                binary(BinOp::Gt, ident("x"), int_lit(127)),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Sub, ident("x"), int_lit(256)),
                )],
            ),
            while_stmt(
                binary(BinOp::Lt, ident("x"), int_lit(-128)),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Add, ident("x"), int_lit(256)),
                )],
            ),
            ret(ident("x")),
        ],
    ));
    out.push(function_stmt(
        "__j_short",
        vec!["x"],
        vec![
            assign(ident("x"), call("__java_trunc_cast", vec![ident("x")])),
            if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Gt, ident("x"), int_lit(65536)),
                    binary(BinOp::Lt, ident("x"), int_lit(-65536)),
                ),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Mod, ident("x"), int_lit(65536)),
                )],
                None,
            ),
            while_stmt(
                binary(BinOp::Gt, ident("x"), int_lit(32767)),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Sub, ident("x"), int_lit(65536)),
                )],
            ),
            while_stmt(
                binary(BinOp::Lt, ident("x"), int_lit(-32768)),
                vec![assign(
                    ident("x"),
                    binary(BinOp::Add, ident("x"), int_lit(65536)),
                )],
            ),
            ret(ident("x")),
        ],
    ));

    // __j_to_radix(x, radix): Integer.toBinaryString/toHexString/
    // toOctalString — the value as UNSIGNED 32-bit in the given radix
    // (Java: toHexString(-1) == "ffffffff"). Unsigned conversion in float
    // arithmetic — exact for the 32-bit range.
    out.push(function_stmt(
        "__j_to_radix",
        vec!["x", "radix"],
        vec![
            var_decl("u", call("__j_i32", vec![ident("x")])),
            if_stmt(
                binary(BinOp::Lt, ident("u"), int_lit(0)),
                vec![assign(
                    ident("u"),
                    binary(BinOp::Add, ident("u"), expr_f64(4294967296.0)),
                )],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("u"), int_lit(0)),
                vec![ret(str_lit("0"))],
                None,
            ),
            var_decl("digits", str_lit("0123456789abcdef")),
            var_decl("acc", str_lit("")),
            while_stmt(
                binary(BinOp::Gt, ident("u"), int_lit(0)),
                vec![
                    var_decl("d", binary(BinOp::Mod, ident("u"), ident("radix"))),
                    assign(
                        ident("acc"),
                        add(char_at(ident("digits"), ident("d")), ident("acc")),
                    ),
                    assign(
                        ident("u"),
                        binary(
                            BinOp::Div,
                            binary(BinOp::Sub, ident("u"), ident("d")),
                            ident("radix"),
                        ),
                    ),
                ],
            ),
            ret(ident("acc")),
        ],
    ));

    // Java Double helpers: ECMA comparison gets close, but Java defines
    // NaN ordering and signed zero ordering for Double.compare.
    out.push(function_stmt(
        "__j_double_is_negative_zero",
        vec!["x"],
        vec![ret(binary(
            BinOp::And,
            binary(BinOp::Eq, ident("x"), int_lit(0)),
            binary(
                BinOp::Eq,
                binary(BinOp::Div, expr_f64(1.0), ident("x")),
                neg_inf(),
            ),
        ))],
    ));
    out.push(function_stmt(
        "__j_double_is_infinite",
        vec!["x"],
        vec![ret(binary(
            BinOp::Or,
            binary(BinOp::Eq, ident("x"), pos_inf()),
            binary(BinOp::Eq, ident("x"), neg_inf()),
        ))],
    ));
    out.push(function_stmt(
        "__j_double_is_finite",
        vec!["x"],
        vec![ret(binary(
            BinOp::And,
            binary(BinOp::Eq, ident("x"), ident("x")),
            binary(
                BinOp::And,
                binary(BinOp::NotEq, ident("x"), pos_inf()),
                binary(BinOp::NotEq, ident("x"), neg_inf()),
            ),
        ))],
    ));
    out.push(function_stmt(
        "__j_double_compare",
        vec!["a", "b"],
        vec![
            var_decl("a_nan", binary(BinOp::NotEq, ident("a"), ident("a"))),
            var_decl("b_nan", binary(BinOp::NotEq, ident("b"), ident("b"))),
            if_stmt(
                binary(BinOp::And, ident("a_nan"), ident("b_nan")),
                vec![ret(int_lit(0))],
                None,
            ),
            if_stmt(ident("a_nan"), vec![ret(int_lit(1))], None),
            if_stmt(ident("b_nan"), vec![ret(int_lit(-1))], None),
            if_stmt(
                binary(BinOp::Lt, ident("a"), ident("b")),
                vec![ret(int_lit(-1))],
                None,
            ),
            if_stmt(
                binary(BinOp::Gt, ident("a"), ident("b")),
                vec![ret(int_lit(1))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    call("__j_double_is_negative_zero", vec![ident("a")]),
                    binary(
                        BinOp::NotEq,
                        call("__j_double_is_negative_zero", vec![ident("b")]),
                        bool_lit(true),
                    ),
                ),
                vec![ret(int_lit(-1))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(
                        BinOp::NotEq,
                        call("__j_double_is_negative_zero", vec![ident("a")]),
                        bool_lit(true),
                    ),
                    call("__j_double_is_negative_zero", vec![ident("b")]),
                ),
                vec![ret(int_lit(1))],
                None,
            ),
            ret(int_lit(0)),
        ],
    ));

    // __j_arraycopy(src, srcPos, dest, destPos, len) — JLS
    // System.arraycopy: in-place into dest, overlap-safe (copies "as if"
    // through a temporary, which is exactly what this does).
    out.push(function_stmt(
        "__j_arraycopy",
        vec!["src", "srcPos", "dest", "destPos", "len"],
        vec![
            var_decl("tmp", expr(ExprKind::Array(Vec::new()))),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("len")),
                vec![
                    assign(
                        index_expr(ident("tmp"), ident("i")),
                        index_expr(ident("src"), add(ident("srcPos"), ident("i"))),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            assign(ident("i"), int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("len")),
                vec![
                    assign(
                        index_expr(ident("dest"), add(ident("destPos"), ident("i"))),
                        index_expr(ident("tmp"), ident("i")),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(null_lit()),
        ],
    ));

    out.append(&mut monitor_fns());
    out.append(&mut process_fns());
    out.append(&mut objects_fns());
    out.append(&mut properties_fns());
    out.append(&mut base64_fns());
    out.append(&mut stringbuilder_fns());
    out.append(&mut string_fns());
    out.append(&mut scanner_fns());
    out.append(&mut thread_fns());
    out.append(&mut regex_fns());
    out.append(&mut url_fns());
    out.append(&mut spliterator_fns());
    out.push(sprintf_fn());
    out
}

fn monitor_fns() -> Vec<Statement> {
    let mut out = Vec::new();

    out.push(function_stmt(
        "__j_object_wait",
        vec!["obj", "timeout"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("obj", "__j_notified"), bool_lit(true)),
                vec![
                    assign(fld("obj", "__j_notified"), bool_lit(false)),
                    ret(null_lit()),
                ],
                None,
            ),
            stmt(StmtKind::Throw {
                expr: Some(call(
                    "__j_exc",
                    vec![
                        str_lit("__JavaWaitYield"),
                        array_lit(vec![str_lit("__JavaWaitYield")]),
                        str_lit("wait"),
                        undefined_lit(),
                    ],
                )),
                cause: None,
            }),
        ],
    ));
    out
}

fn process_fns() -> Vec<Statement> {
    let mut out = Vec::new();

    out.push(var_decl(
        "__j_runtime_singleton",
        obj_props(vec![("__kind", str_lit("Runtime"))]),
    ));
    out.push(function_stmt(
        "__j_runtime_get",
        vec![],
        vec![ret(ident("__j_runtime_singleton"))],
    ));
    out.push(function_stmt(
        "__j_runtime_processors",
        vec!["rt"],
        vec![ret(int_lit(1))],
    ));
    out.push(function_stmt(
        "__j_runtime_free_memory",
        vec!["rt"],
        vec![ret(int_lit(0))],
    ));
    out.push(function_stmt(
        "__j_runtime_total_memory",
        vec!["rt"],
        vec![ret(int_lit(1))],
    ));
    out.push(function_stmt(
        "__j_runtime_max_memory",
        vec!["rt"],
        vec![ret(int_lit(1))],
    ));
    out.push(function_stmt(
        "__j_runtime_noop",
        vec!["rt"],
        vec![ret(null_lit())],
    ));
    out.push(function_stmt(
        "__j_file_new",
        vec!["path"],
        vec![ret(obj_props(vec![
            ("__kind", str_lit("File")),
            ("path", to_str(ident("path"))),
        ]))],
    ));
    out.push(function_stmt(
        "__j_file_get_path",
        vec!["file"],
        vec![ret(fld("file", "path"))],
    ));
    out.push(var_decl(
        "__j_pb_redirect_pipe",
        obj_props(vec![
            ("__kind", str_lit("Redirect")),
            ("name", str_lit("PIPE")),
        ]),
    ));
    out.push(var_decl(
        "__j_pb_redirect_inherit",
        obj_props(vec![
            ("__kind", str_lit("Redirect")),
            ("name", str_lit("INHERIT")),
        ]),
    ));
    out.push(var_decl(
        "__j_pb_redirect_discard",
        obj_props(vec![
            ("__kind", str_lit("Redirect")),
            ("name", str_lit("DISCARD")),
        ]),
    ));
    out.push(function_stmt(
        "__j_pb_redirect",
        vec!["kind"],
        vec![
            if_stmt(
                binary(BinOp::Eq, to_str(ident("kind")), str_lit("INHERIT")),
                vec![ret(ident("__j_pb_redirect_inherit"))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, to_str(ident("kind")), str_lit("DISCARD")),
                vec![ret(ident("__j_pb_redirect_discard"))],
                None,
            ),
            ret(ident("__j_pb_redirect_pipe")),
        ],
    ));
    out.push(function_stmt(
        "__j_pb_redirect_type",
        vec!["redirect"],
        vec![ret(fld("redirect", "name"))],
    ));
    out.push(function_stmt(
        "__j_pb_new",
        vec!["cmd"],
        vec![
            var_decl("actual", ident("cmd")),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Eq, fld("cmd", "length"), int_lit(1)),
                    binary(
                        BinOp::NotEq,
                        member(index_expr(ident("cmd"), int_lit(0)), "length"),
                        undefined_lit(),
                    ),
                ),
                vec![assign(
                    ident("actual"),
                    index_expr(ident("cmd"), int_lit(0)),
                )],
                None,
            ),
            ret(obj_props(vec![
                ("__kind", str_lit("ProcessBuilder")),
                ("cmd", ident("actual")),
                ("dir", null_lit()),
                ("env", new_expr("HashMap", vec![])),
                ("redirectErrorStream", bool_lit(false)),
                ("redirectInput", ident("__j_pb_redirect_pipe")),
                ("redirectOutput", ident("__j_pb_redirect_pipe")),
                ("redirectError", ident("__j_pb_redirect_pipe")),
            ])),
        ],
    ));
    out.push(function_stmt(
        "__j_pb_command_get",
        vec!["pb"],
        vec![ret(fld("pb", "cmd"))],
    ));
    out.push(function_stmt(
        "__j_pb_command_set",
        vec!["pb", "cmd"],
        vec![assign(fld("pb", "cmd"), ident("cmd")), ret(ident("pb"))],
    ));
    out.push(function_stmt(
        "__j_pb_directory_get",
        vec!["pb"],
        vec![ret(fld("pb", "dir"))],
    ));
    out.push(function_stmt(
        "__j_pb_directory_set",
        vec!["pb", "dir"],
        vec![assign(fld("pb", "dir"), ident("dir")), ret(ident("pb"))],
    ));
    out.push(function_stmt(
        "__j_pb_redirect_error_stream_get",
        vec!["pb"],
        vec![ret(fld("pb", "redirectErrorStream"))],
    ));
    out.push(function_stmt(
        "__j_pb_redirect_error_stream_set",
        vec!["pb", "value"],
        vec![
            assign(fld("pb", "redirectErrorStream"), ident("value")),
            ret(ident("pb")),
        ],
    ));
    out.push(function_stmt(
        "__j_pb_environment",
        vec!["pb"],
        vec![ret(fld("pb", "env"))],
    ));
    out.push(function_stmt(
        "__j_pb_inherit_io",
        vec!["pb"],
        vec![ret(ident("pb"))],
    ));
    out.push(function_stmt(
        "__j_pb_redirect_input_get",
        vec!["pb"],
        vec![ret(fld("pb", "redirectInput"))],
    ));
    out.push(function_stmt(
        "__j_pb_redirect_output_get",
        vec!["pb"],
        vec![ret(fld("pb", "redirectOutput"))],
    ));
    out.push(function_stmt(
        "__j_pb_redirect_output_set",
        vec!["pb", "redirect"],
        vec![
            assign(fld("pb", "redirectOutput"), ident("redirect")),
            ret(ident("pb")),
        ],
    ));
    out.push(function_stmt(
        "__j_pb_redirect_error_get",
        vec!["pb"],
        vec![ret(fld("pb", "redirectError"))],
    ));
    out.push(function_stmt(
        "__j_pb_start",
        vec!["pb"],
        vec![ret(obj_props(vec![
            ("__kind", str_lit("Process")),
            ("exit", int_lit(0)),
            ("alive", bool_lit(false)),
            ("input", obj_lit()),
        ]))],
    ));
    out.push(function_stmt(
        "__j_process_wait_for",
        vec!["p"],
        vec![
            assign(fld("p", "alive"), bool_lit(false)),
            ret(fld("p", "exit")),
        ],
    ));
    out.push(function_stmt(
        "__j_process_is_alive",
        vec!["p"],
        vec![ret(fld("p", "alive"))],
    ));
    out.push(function_stmt(
        "__j_process_input_stream",
        vec!["p"],
        vec![ret(fld("p", "input"))],
    ));
    out.push(function_stmt(
        "__j_process_exit_value",
        vec!["p"],
        vec![ret(fld("p", "exit"))],
    ));
    out.push(function_stmt(
        "__j_process_destroy",
        vec!["p"],
        vec![assign(fld("p", "alive"), bool_lit(false)), ret(null_lit())],
    ));

    out
}

fn objects_fns() -> Vec<Statement> {
    let mut out = Vec::new();

    out.push(function_stmt(
        "__j_objects_require_non_null",
        vec!["x", "message"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("x"), null_lit()),
                vec![stmt(StmtKind::Throw {
                    expr: Some(call(
                        "__j_exc",
                        vec![
                            str_lit("NullPointerException"),
                            array_lit(vec![
                                str_lit("java.lang.NullPointerException"),
                                str_lit("NullPointerException"),
                                str_lit("RuntimeException"),
                                str_lit("Exception"),
                                str_lit("Throwable"),
                            ]),
                            ident("message"),
                            undefined_lit(),
                        ],
                    )),
                    cause: None,
                })],
                None,
            ),
            ret(ident("x")),
        ],
    ));
    out
}

fn properties_fns() -> Vec<Statement> {
    let fld = |name: &str, f: &str| member(ident(name), f);
    let mut out = Vec::new();

    out.push(function_stmt(
        "__j_props_new",
        vec!["defaults"],
        vec![
            var_decl("p", new_expr("HashMap", vec![])),
            if_stmt(
                binary(BinOp::NotEq, ident("defaults"), undefined_lit()),
                vec![assign(fld("p", "__defaults"), ident("defaults"))],
                Some(vec![assign(fld("p", "__defaults"), null_lit())]),
            ),
            ret(ident("p")),
        ],
    ));
    out.push(function_stmt(
        "__j_props_get",
        vec!["p", "key", "def"],
        vec![
            var_decl("v", call("__java_map_get", vec![ident("p"), ident("key")])),
            if_stmt(
                binary(BinOp::NotEq, ident("v"), null_lit()),
                vec![ret(ident("v"))],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, fld("p", "__defaults"), null_lit()),
                vec![
                    assign(
                        ident("v"),
                        call(
                            "__j_props_get",
                            vec![fld("p", "__defaults"), ident("key"), null_lit()],
                        ),
                    ),
                    if_stmt(
                        binary(BinOp::NotEq, ident("v"), null_lit()),
                        vec![ret(ident("v"))],
                        None,
                    ),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("def"), undefined_lit()),
                vec![ret(ident("def"))],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_props_set",
        vec!["p", "key", "value"],
        vec![
            var_decl(
                "old",
                call("__j_props_get", vec![ident("p"), ident("key"), null_lit()]),
            ),
            stmt(StmtKind::Expr(call(
                "__java_map_put",
                vec![ident("p"), ident("key"), ident("value")],
            ))),
            ret(ident("old")),
        ],
    ));
    out.push(function_stmt(
        "__j_props_names",
        vec!["p"],
        vec![ret(call("__java_map_key_set", vec![ident("p")]))],
    ));
    out.push(function_stmt(
        "__j_props_enum",
        vec!["items"],
        vec![
            var_decl("e", obj_lit()),
            assign(fld("e", "items"), ident("items")),
            assign(fld("e", "index"), int_lit(0)),
            assign(fld("e", "nonempty"), bool_lit(true)),
            ret(ident("e")),
        ],
    ));
    out.push(function_stmt(
        "__j_props_keys",
        vec!["p"],
        vec![ret(call(
            "__j_props_enum",
            vec![call("__java_map_key_set", vec![ident("p")])],
        ))],
    ));
    out.push(function_stmt(
        "__j_props_elements",
        vec!["p"],
        vec![ret(call(
            "__j_props_enum",
            vec![call("__java_map_values", vec![ident("p")])],
        ))],
    ));
    out.push(function_stmt(
        "__j_props_class",
        vec!["p"],
        vec![
            var_decl("c", obj_lit()),
            assign(fld("c", "name"), str_lit("java.util.Properties")),
            ret(ident("c")),
        ],
    ));
    out.push(function_stmt(
        "__j_class_get_name",
        vec!["c"],
        vec![ret(fld("c", "name"))],
    ));
    out.push(function_stmt(
        "__j_enum_has_more",
        vec!["e"],
        vec![ret(fld("e", "nonempty"))],
    ));
    out.push(var_decl(
        "__j_system_props",
        call("__j_props_new", vec![null_lit()]),
    ));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("line.separator"),
            str_lit("\n"),
        ],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("file.separator"),
            str_lit("/"),
        ],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("path.separator"),
            str_lit(":"),
        ],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("os.name"),
            str_lit("Vybe"),
        ],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("java.version"),
            str_lit("21"),
        ],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("java.vendor"),
            str_lit("Vybe"),
        ],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("file.encoding"),
            str_lit("UTF-8"),
        ],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![ident("__j_system_props"), str_lit("user.dir"), str_lit("/")],
    ))));
    out.push(stmt(StmtKind::Expr(call(
        "__j_props_set",
        vec![
            ident("__j_system_props"),
            str_lit("user.home"),
            str_lit("/"),
        ],
    ))));
    out.push(function_stmt(
        "__j_system_get_properties",
        vec![],
        vec![ret(ident("__j_system_props"))],
    ));
    out.push(function_stmt(
        "__j_system_get_property",
        vec!["key", "def"],
        vec![ret(call(
            "__j_props_get",
            vec![ident("__j_system_props"), ident("key"), ident("def")],
        ))],
    ));
    out.push(function_stmt(
        "__j_system_set_property",
        vec!["key", "value"],
        vec![ret(call(
            "__j_props_set",
            vec![ident("__j_system_props"), ident("key"), ident("value")],
        ))],
    ));
    out.push(function_stmt(
        "__j_system_clear_property",
        vec!["key"],
        vec![ret(call(
            "__java_map_remove",
            vec![ident("__j_system_props"), ident("key")],
        ))],
    ));
    out.push(function_stmt(
        "__j_system_getenv",
        vec!["key"],
        vec![ret(null_lit())],
    ));

    out
}

fn base64_fns() -> Vec<Statement> {
    let fld = |name: &str, f: &str| member(ident(name), f);
    let mut out = Vec::new();

    out.push(function_stmt(
        "__j_b64_make",
        vec!["url", "mime", "line", "sep"],
        vec![
            var_decl("b", obj_lit()),
            assign(fld("b", "url"), ident("url")),
            assign(fld("b", "mime"), ident("mime")),
            if_stmt(
                binary(BinOp::Eq, ident("line"), undefined_lit()),
                vec![assign(ident("line"), int_lit(0))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("sep"), undefined_lit()),
                vec![assign(ident("sep"), str_lit("\n"))],
                None,
            ),
            assign(fld("b", "line"), ident("line")),
            assign(fld("b", "sep"), ident("sep")),
            assign(fld("b", "pad"), bool_lit(true)),
            ret(ident("b")),
        ],
    ));
    out.push(function_stmt(
        "__j_b64_encoder",
        vec![],
        vec![ret(call(
            "__j_b64_make",
            vec![bool_lit(false), bool_lit(false), int_lit(0), str_lit("\n")],
        ))],
    ));
    out.push(function_stmt(
        "__j_b64_url_encoder",
        vec![],
        vec![ret(call(
            "__j_b64_make",
            vec![bool_lit(true), bool_lit(false), int_lit(0), str_lit("\n")],
        ))],
    ));
    out.push(function_stmt(
        "__j_b64_mime_encoder",
        vec!["line", "sep"],
        vec![ret(call(
            "__j_b64_make",
            vec![
                bool_lit(false),
                bool_lit(true),
                ident("line"),
                str_lit("\n"),
            ],
        ))],
    ));
    out.push(function_stmt(
        "__j_b64_decoder",
        vec![],
        vec![ret(call(
            "__j_b64_make",
            vec![bool_lit(false), bool_lit(false), int_lit(0), str_lit("\n")],
        ))],
    ));
    out.push(function_stmt(
        "__j_b64_url_decoder",
        vec![],
        vec![ret(call(
            "__j_b64_make",
            vec![bool_lit(true), bool_lit(false), int_lit(0), str_lit("\n")],
        ))],
    ));
    out.push(function_stmt(
        "__j_b64_mime_decoder",
        vec![],
        vec![ret(call(
            "__j_b64_make",
            vec![bool_lit(false), bool_lit(true), int_lit(0), str_lit("\n")],
        ))],
    ));
    out.push(function_stmt(
        "__j_b64_without_padding",
        vec!["b"],
        vec![assign(fld("b", "pad"), bool_lit(false)), ret(ident("b"))],
    ));
    out.push(function_stmt(
        "__j_b64_input_text",
        vec!["x"],
        vec![
            if_stmt(
                binary(BinOp::Eq, typeof_expr(ident("x")), str_lit("string")),
                vec![ret(ident("x"))],
                None,
            ),
            var_decl("s", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("x"), "length")),
                vec![
                    var_decl("n", index_expr(ident("x"), ident("i"))),
                    if_stmt(
                        binary(BinOp::Lt, ident("n"), int_lit(0)),
                        vec![assign(ident("n"), add(ident("n"), int_lit(256)))],
                        None,
                    ),
                    assign(
                        ident("s"),
                        add(ident("s"), call("__j_from_char_code", vec![ident("n")])),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_b64_bin_bytes",
        vec!["s"],
        vec![
            var_decl("out", arr_lit()),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                vec![
                    assign(
                        index_expr(ident("out"), ident("i")),
                        call_member(ident("s"), "charCodeAt", vec![ident("i")]),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_b64_replace_all",
        vec!["s", "from", "to"],
        vec![
            var_decl("i", call_member(ident("s"), "indexOf", vec![ident("from")])),
            while_stmt(
                binary(BinOp::GtEq, ident("i"), int_lit(0)),
                vec![
                    assign(
                        ident("s"),
                        add(
                            add(
                                call_member(ident("s"), "substring", vec![int_lit(0), ident("i")]),
                                ident("to"),
                            ),
                            call_member(
                                ident("s"),
                                "substring",
                                vec![add(ident("i"), member(ident("from"), "length"))],
                            ),
                        ),
                    ),
                    assign(
                        ident("i"),
                        call_member(ident("s"), "indexOf", vec![ident("from")]),
                    ),
                ],
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_b64_encode_to_string",
        vec!["b", "bytes"],
        vec![
            var_decl(
                "s",
                call(
                    "__j_btoa",
                    vec![call("__j_b64_input_text", vec![ident("bytes")])],
                ),
            ),
            if_stmt(
                fld("b", "url"),
                vec![
                    assign(
                        ident("s"),
                        call(
                            "__j_b64_replace_all",
                            vec![ident("s"), str_lit("+"), str_lit("-")],
                        ),
                    ),
                    assign(
                        ident("s"),
                        call(
                            "__j_b64_replace_all",
                            vec![ident("s"), str_lit("/"), str_lit("_")],
                        ),
                    ),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, fld("b", "pad"), bool_lit(false)),
                vec![while_stmt(
                    binary(
                        BinOp::Eq,
                        char_at(
                            ident("s"),
                            binary(BinOp::Sub, member(ident("s"), "length"), int_lit(1)),
                        ),
                        str_lit("="),
                    ),
                    vec![assign(
                        ident("s"),
                        call_member(
                            ident("s"),
                            "substring",
                            vec![
                                int_lit(0),
                                binary(BinOp::Sub, member(ident("s"), "length"), int_lit(1)),
                            ],
                        ),
                    )],
                )],
                None,
            ),
            if_stmt(
                binary(BinOp::Gt, fld("b", "line"), int_lit(0)),
                vec![assign(
                    ident("s"),
                    add(
                        add(
                            call_member(
                                ident("s"),
                                "substring",
                                vec![int_lit(0), fld("b", "line")],
                            ),
                            fld("b", "sep"),
                        ),
                        call_member(ident("s"), "substring", vec![fld("b", "line")]),
                    ),
                )],
                None,
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_b64_encode",
        vec!["b", "bytes"],
        vec![ret(call(
            "__j_string_get_bytes",
            vec![call(
                "__j_b64_encode_to_string",
                vec![ident("b"), ident("bytes")],
            )],
        ))],
    ));
    out.push(function_stmt(
        "__j_b64_decode",
        vec!["b", "input"],
        vec![
            var_decl("s", call("__j_b64_input_text", vec![ident("input")])),
            assign(
                ident("s"),
                call(
                    "__j_b64_replace_all",
                    vec![ident("s"), str_lit("\n"), str_lit("")],
                ),
            ),
            assign(
                ident("s"),
                call(
                    "__j_b64_replace_all",
                    vec![ident("s"), str_lit("\r"), str_lit("")],
                ),
            ),
            assign(
                ident("s"),
                call(
                    "__j_b64_replace_all",
                    vec![ident("s"), str_lit(" "), str_lit("")],
                ),
            ),
            assign(
                ident("s"),
                call(
                    "__j_b64_replace_all",
                    vec![ident("s"), str_lit("\t"), str_lit("")],
                ),
            ),
            if_stmt(
                fld("b", "url"),
                vec![
                    assign(
                        ident("s"),
                        call(
                            "__j_b64_replace_all",
                            vec![ident("s"), str_lit("-"), str_lit("+")],
                        ),
                    ),
                    assign(
                        ident("s"),
                        call(
                            "__j_b64_replace_all",
                            vec![ident("s"), str_lit("_"), str_lit("/")],
                        ),
                    ),
                ],
                None,
            ),
            while_stmt(
                binary(
                    BinOp::NotEq,
                    binary(BinOp::Mod, member(ident("s"), "length"), int_lit(4)),
                    int_lit(0),
                ),
                vec![assign(ident("s"), add(ident("s"), str_lit("=")))],
            ),
            ret(call(
                "__j_b64_bin_bytes",
                vec![call("__j_atob", vec![ident("s")])],
            )),
        ],
    ));

    out
}

/// `java.net.URL`/`URI` getters over the WHATWG-parsed object
/// (`web:url new` — fields: protocol "http:", hostname, host, port,
/// pathname, search "?q", hash "#f", username, password, href).
fn url_fns() -> Vec<Statement> {
    let fld = |name: &str, f: &str| member(ident(name), f);
    let bool_lit = |b: bool| expr(ExprKind::Lit(Literal::Bool(b)));
    let mut out = Vec::new();

    // __j_url_new(spec): WHATWG-parse, remembering the raw spec (java's
    // getPath depends on whether the spec actually wrote a path).
    out.push(function_stmt(
        "__j_url_new",
        vec!["spec"],
        vec![
            var_decl("u", call("__j_url_parse", vec![ident("spec")])),
            assign(fld("u", "__spec"), to_str(ident("spec"))),
            assign(
                fld("u", "__scheme"),
                call("__j_url_protocol", vec![ident("u")]),
            ),
            assign(fld("u", "__opaque"), bool_lit(false)),
            ret(ident("u")),
        ],
    ));
    // new URL(protocol, host, port, file) — java 4-arg constructor.
    out.push(function_stmt(
        "__j_url_make",
        vec!["proto", "host", "port", "file"],
        vec![
            var_decl(
                "spec",
                add(
                    add(to_str(ident("proto")), str_lit("://")),
                    to_str(ident("host")),
                ),
            ),
            if_stmt(
                binary(BinOp::GtEq, ident("port"), int_lit(0)),
                vec![assign(
                    ident("spec"),
                    add(add(ident("spec"), str_lit(":")), to_str(ident("port"))),
                )],
                None,
            ),
            assign(ident("spec"), add(ident("spec"), to_str(ident("file")))),
            ret(call("__j_url_new", vec![ident("spec")])),
        ],
    ));
    // new URL(context, spec) — java resolves WITHOUT dot-normalization
    // (unlike WHATWG), so the resolved path is pinned via __path.
    out.push(function_stmt(
        "__j_url_ctx",
        vec!["base", "spec"],
        vec![
            assign(ident("spec"), to_str(ident("spec"))),
            if_stmt(
                binary(
                    BinOp::GtEq,
                    call_expr(member(ident("spec"), "indexOf"), vec![str_lit("://")]),
                    int_lit(0),
                ),
                vec![ret(call("__j_url_new", vec![ident("spec")]))],
                None,
            ),
            var_decl("path", str_lit("")),
            if_stmt(
                binary(
                    BinOp::Eq,
                    call_expr(
                        member(ident("spec"), "substring"),
                        vec![int_lit(0), int_lit(1)],
                    ),
                    str_lit("/"),
                ),
                vec![assign(ident("path"), ident("spec"))],
                Some(vec![
                    var_decl("bp", member(ident("base"), "pathname")),
                    var_decl(
                        "cut",
                        call_expr(member(ident("bp"), "lastIndexOf"), vec![str_lit("/")]),
                    ),
                    assign(
                        ident("path"),
                        add(
                            call_expr(
                                member(ident("bp"), "substring"),
                                vec![int_lit(0), add(ident("cut"), int_lit(1))],
                            ),
                            ident("spec"),
                        ),
                    ),
                ]),
            ),
            var_decl(
                "u",
                call(
                    "__j_url_new",
                    vec![add(
                        add(
                            add(member(ident("base"), "protocol"), str_lit("//")),
                            member(ident("base"), "host"),
                        ),
                        ident("path"),
                    )],
                ),
            ),
            assign(fld("u", "__path"), ident("path")),
            ret(ident("u")),
        ],
    ));
    out.push(function_stmt(
        "__j_uri_new",
        vec!["spec"],
        vec![
            assign(ident("spec"), to_str(ident("spec"))),
            var_decl(
                "scheme_sep",
                call_expr(member(ident("spec"), "indexOf"), vec![str_lit("://")]),
            ),
            if_stmt(
                binary(BinOp::GtEq, ident("scheme_sep"), int_lit(0)),
                vec![ret(call("__j_url_new", vec![ident("spec")]))],
                None,
            ),
            var_decl(
                "colon",
                call_expr(member(ident("spec"), "indexOf"), vec![str_lit(":")]),
            ),
            if_stmt(
                binary(BinOp::GtEq, ident("colon"), int_lit(0)),
                vec![ret(obj_props(vec![
                    ("__spec", ident("spec")),
                    (
                        "__scheme",
                        call_expr(
                            member(ident("spec"), "substring"),
                            vec![int_lit(0), ident("colon")],
                        ),
                    ),
                    (
                        "__ssp",
                        call_expr(
                            member(ident("spec"), "substring"),
                            vec![add(ident("colon"), int_lit(1))],
                        ),
                    ),
                    ("__opaque", bool_lit(true)),
                    ("href", ident("spec")),
                    (
                        "protocol",
                        add(
                            call_expr(
                                member(ident("spec"), "substring"),
                                vec![int_lit(0), ident("colon")],
                            ),
                            str_lit(":"),
                        ),
                    ),
                    ("hostname", str_lit("")),
                    ("host", str_lit("")),
                    ("port", str_lit("")),
                    ("pathname", str_lit("")),
                    ("search", str_lit("")),
                    ("hash", str_lit("")),
                    ("username", str_lit("")),
                    ("password", str_lit("")),
                ]))],
                None,
            ),
            ret(obj_props(vec![
                ("__spec", ident("spec")),
                ("__scheme", null_lit()),
                ("__ssp", ident("spec")),
                ("__opaque", bool_lit(false)),
                ("href", ident("spec")),
                ("protocol", str_lit("")),
                ("hostname", str_lit("")),
                ("host", str_lit("")),
                ("port", str_lit("")),
                ("pathname", ident("spec")),
                ("search", str_lit("")),
                ("hash", str_lit("")),
                ("username", str_lit("")),
                ("password", str_lit("")),
            ])),
        ],
    ));
    out.push(function_stmt(
        "__j_uri_make3",
        vec!["scheme", "host", "path"],
        vec![ret(call(
            "__j_uri_new",
            vec![add(
                add(
                    add(to_str(ident("scheme")), str_lit("://")),
                    to_str(ident("host")),
                ),
                to_str(ident("path")),
            )],
        ))],
    ));
    out.push(function_stmt(
        "__j_uri_make7",
        vec!["scheme", "user", "host", "port", "path", "query", "frag"],
        vec![
            var_decl("spec", add(to_str(ident("scheme")), str_lit("://"))),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::NotEq, ident("user"), null_lit()),
                    binary(BinOp::NotEq, to_str(ident("user")), str_lit("")),
                ),
                vec![assign(
                    ident("spec"),
                    add(add(ident("spec"), to_str(ident("user"))), str_lit("@")),
                )],
                None,
            ),
            assign(ident("spec"), add(ident("spec"), to_str(ident("host")))),
            if_stmt(
                binary(BinOp::GtEq, ident("port"), int_lit(0)),
                vec![assign(
                    ident("spec"),
                    add(add(ident("spec"), str_lit(":")), to_str(ident("port"))),
                )],
                None,
            ),
            assign(ident("spec"), add(ident("spec"), to_str(ident("path")))),
            if_stmt(
                binary(BinOp::NotEq, ident("query"), null_lit()),
                vec![assign(
                    ident("spec"),
                    add(add(ident("spec"), str_lit("?")), to_str(ident("query"))),
                )],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("frag"), null_lit()),
                vec![assign(
                    ident("spec"),
                    add(add(ident("spec"), str_lit("#")), to_str(ident("frag"))),
                )],
                None,
            ),
            ret(call("__j_uri_new", vec![ident("spec")])),
        ],
    ));
    // equals / hashCode / sameFile — java compares the URL text.
    out.push(function_stmt(
        "__j_url_equals",
        vec!["a", "b"],
        vec![
            if_stmt(
                binary(
                    BinOp::Eq,
                    member(ident("a"), "href"),
                    member(ident("b"), "href"),
                ),
                vec![ret(bool_lit(true))],
                None,
            ),
            ret(bool_lit(false)),
        ],
    ));
    out.push(function_stmt(
        "__j_url_hash",
        vec!["u"],
        vec![
            var_decl("s", member(ident("u"), "href")),
            // Float accumulator: the dynamic i32 multiply traps on
            // overflow; f64 wraps through __j_i32 like java ints do.
            var_decl("h", expr_f64(0.0)),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                vec![
                    assign(
                        ident("h"),
                        add(
                            // 31.0 forces f64 arithmetic — the i32 multiply
                            // traps on overflow instead of wrapping.
                            binary(BinOp::Mul, ident("h"), expr_f64(31.0)),
                            call_expr(member(ident("s"), "charCodeAt"), vec![ident("i")]),
                        ),
                    ),
                    // Stay in i32 range like java's overflow arithmetic.
                    assign(ident("h"), call("__j_i32", vec![ident("h")])),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("h")),
        ],
    ));
    out.push(function_stmt(
        "__j_url_same_file",
        vec!["a", "b"],
        vec![
            if_stmt(
                binary(
                    BinOp::And,
                    binary(
                        BinOp::Eq,
                        member(ident("a"), "protocol"),
                        member(ident("b"), "protocol"),
                    ),
                    binary(
                        BinOp::And,
                        binary(
                            BinOp::Eq,
                            member(ident("a"), "host"),
                            member(ident("b"), "host"),
                        ),
                        binary(
                            BinOp::Eq,
                            call("__j_url_file", vec![ident("a")]),
                            call("__j_url_file", vec![ident("b")]),
                        ),
                    ),
                ),
                vec![ret(bool_lit(true))],
                None,
            ),
            ret(bool_lit(false)),
        ],
    ));

    // getProtocol()/getScheme(): "http:" minus the colon.
    //
    // This IS `url::UrlField::Scheme` — a third copy of the reshaping php
    // `parse_url` and python `urlsplit` needed — but binding it to
    // `common:url.component_scheme` BROKE java URL (measured), so the migration
    // needs more than a profile line. See the notes in `languages/java/profile`.
    out.push(function_stmt(
        "__j_url_protocol",
        vec!["u"],
        vec![
            var_decl("p", fld("u", "protocol")),
            ret(call_expr(
                member(ident("p"), "substring"),
                vec![
                    int_lit(0),
                    binary(BinOp::Sub, member(ident("p"), "length"), int_lit(1)),
                ],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_url_host",
        vec!["u"],
        vec![ret(fld("u", "hostname"))],
    ));
    // getPort(): -1 when the URL names none (java).
    out.push(function_stmt(
        "__j_url_port",
        vec!["u"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("u", "port"), str_lit("")),
                vec![ret(int_lit(-1))],
                None,
            ),
            ret(call_expr(
                member(ident("Integer"), "parseInt"),
                vec![fld("u", "port")],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_url_default_port",
        vec!["u"],
        vec![
            var_decl("p", call("__j_url_protocol", vec![ident("u")])),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("http")),
                vec![ret(int_lit(80))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("https")),
                vec![ret(int_lit(443))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("p"), str_lit("ftp")),
                vec![ret(int_lit(21))],
                None,
            ),
            ret(int_lit(-1)),
        ],
    ));
    // getPath(): the context-resolution override wins (java keeps dot
    // segments); a bare "http://host" spec has the EMPTY path in java
    // even though WHATWG reports "/".
    out.push(function_stmt(
        "__j_url_path",
        vec!["u"],
        vec![
            if_stmt(
                binary(BinOp::NotEq, fld("u", "__path"), null_lit()),
                vec![ret(fld("u", "__path"))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, fld("u", "pathname"), str_lit("/")),
                vec![
                    var_decl("s", fld("u", "__spec")),
                    var_decl(
                        "i",
                        call_expr(member(ident("s"), "indexOf"), vec![str_lit("://")]),
                    ),
                    if_stmt(
                        binary(BinOp::GtEq, ident("i"), int_lit(0)),
                        vec![
                            var_decl(
                                "t",
                                call_expr(
                                    member(ident("s"), "substring"),
                                    vec![add(ident("i"), int_lit(3))],
                                ),
                            ),
                            var_decl(
                                "sl",
                                call_expr(member(ident("t"), "indexOf"), vec![str_lit("/")]),
                            ),
                            if_stmt(
                                binary(BinOp::Lt, ident("sl"), int_lit(0)),
                                vec![ret(str_lit(""))],
                                None,
                            ),
                        ],
                        None,
                    ),
                ],
                None,
            ),
            ret(fld("u", "pathname")),
        ],
    ));
    // getQuery()/getRef(): null when absent (java), else without ?/#.
    out.push(function_stmt(
        "__j_url_query",
        vec!["u"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("u", "search"), str_lit("")),
                vec![ret(null_lit())],
                None,
            ),
            ret(call_expr(
                member(fld("u", "search"), "substring"),
                vec![int_lit(1)],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_url_ref",
        vec!["u"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("u", "hash"), str_lit("")),
                vec![ret(null_lit())],
                None,
            ),
            ret(call_expr(
                member(fld("u", "hash"), "substring"),
                vec![int_lit(1)],
            )),
        ],
    ));
    // getFile(): path + "?query" when present.
    out.push(function_stmt(
        "__j_url_file",
        vec!["u"],
        vec![ret(add(fld("u", "pathname"), fld("u", "search")))],
    ));
    out.push(function_stmt(
        "__j_url_authority",
        vec!["u"],
        vec![ret(fld("u", "host"))],
    ));
    // getUserInfo(): "user[:password]" or null.
    out.push(function_stmt(
        "__j_url_user_info",
        vec!["u"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("u", "username"), str_lit("")),
                vec![ret(null_lit())],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, fld("u", "password"), str_lit("")),
                vec![ret(fld("u", "username"))],
                None,
            ),
            ret(add(
                add(fld("u", "username"), str_lit(":")),
                fld("u", "password"),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_uri_is_absolute",
        vec!["u"],
        vec![ret(binary(BinOp::NotEq, fld("u", "__scheme"), null_lit()))],
    ));
    out.push(function_stmt(
        "__j_uri_is_opaque",
        vec!["u"],
        vec![ret(binary(BinOp::Eq, fld("u", "__opaque"), bool_lit(true)))],
    ));
    out.push(function_stmt(
        "__j_uri_ssp",
        vec!["u"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("u", "__opaque"), bool_lit(true)),
                vec![ret(fld("u", "__ssp"))],
                None,
            ),
            var_decl("s", call("__j_url_to_string", vec![ident("u")])),
            var_decl(
                "i",
                call_expr(member(ident("s"), "indexOf"), vec![str_lit(":")]),
            ),
            if_stmt(
                binary(BinOp::GtEq, ident("i"), int_lit(0)),
                vec![ret(call_expr(
                    member(ident("s"), "substring"),
                    vec![add(ident("i"), int_lit(1))],
                ))],
                None,
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_uri_normalize",
        vec!["u"],
        vec![ret(call(
            "__j_uri_new",
            vec![call("__j_url_to_string", vec![ident("u")])],
        ))],
    ));
    out.push(function_stmt(
        "__j_uri_resolve",
        vec!["base", "rel"],
        vec![
            if_stmt(
                call("__j_uri_is_absolute", vec![ident("rel")]),
                vec![ret(ident("rel"))],
                None,
            ),
            ret(call(
                "__j_uri_new",
                vec![member(
                    call(
                        "__j_url_parse",
                        vec![
                            call("__j_url_to_string", vec![ident("rel")]),
                            fld("base", "href"),
                        ],
                    ),
                    "href",
                )],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_uri_relativize",
        vec!["base", "target"],
        vec![
            if_stmt(
                binary(
                    BinOp::And,
                    binary(
                        BinOp::Eq,
                        fld("base", "__scheme"),
                        fld("target", "__scheme"),
                    ),
                    binary(BinOp::Eq, fld("base", "host"), fld("target", "host")),
                ),
                vec![
                    var_decl("p", fld("target", "pathname")),
                    var_decl(
                        "slash",
                        call_expr(member(ident("p"), "lastIndexOf"), vec![str_lit("/")]),
                    ),
                    if_stmt(
                        binary(BinOp::GtEq, ident("slash"), int_lit(0)),
                        vec![ret(call(
                            "__j_uri_new",
                            vec![add(
                                str_lit("../"),
                                call_expr(
                                    member(ident("p"), "substring"),
                                    vec![add(ident("slash"), int_lit(1))],
                                ),
                            )],
                        ))],
                        None,
                    ),
                ],
                None,
            ),
            ret(ident("target")),
        ],
    ));
    out.push(function_stmt(
        "__j_uri_compare_to",
        vec!["a", "b"],
        vec![
            var_decl("as", call("__j_url_to_string", vec![ident("a")])),
            var_decl("bs", call("__j_url_to_string", vec![ident("b")])),
            if_stmt(
                binary(BinOp::Eq, ident("as"), ident("bs")),
                vec![ret(int_lit(0))],
                None,
            ),
            if_stmt(
                binary(BinOp::Lt, ident("as"), ident("bs")),
                vec![ret(int_lit(-1))],
                None,
            ),
            ret(int_lit(1)),
        ],
    ));
    out.push(function_stmt(
        "__j_url_to_string",
        vec!["u"],
        vec![ret(fld("u", "href"))],
    ));
    out
}

fn spliterator_fns() -> Vec<Statement> {
    let mut out = Vec::new();

    out.push(function_stmt(
        "__j_spliterator_from_array",
        vec!["arr"],
        vec![ret(obj_props(vec![
            ("__items", ident("arr")),
            ("__index", int_lit(0)),
            ("__end", member(ident("arr"), "length")),
        ]))],
    ));
    out.push(function_stmt(
        "__j_spliterator_estimate_size",
        vec!["sp"],
        vec![ret(sub(fld("sp", "__end"), fld("sp", "__index")))],
    ));
    out.push(function_stmt(
        "__j_spliterator_try_advance",
        vec!["sp", "fn"],
        vec![
            if_stmt(
                binary(BinOp::GtEq, fld("sp", "__index"), fld("sp", "__end")),
                vec![ret(bool_lit(false))],
                None,
            ),
            var_decl(
                "value",
                index_expr(fld("sp", "__items"), fld("sp", "__index")),
            ),
            assign(fld("sp", "__index"), add(fld("sp", "__index"), int_lit(1))),
            stmt(StmtKind::Expr(call_expr(ident("fn"), vec![ident("value")]))),
            ret(bool_lit(true)),
        ],
    ));
    out.push(function_stmt(
        "__j_spliterator_for_each_remaining",
        vec!["sp", "fn"],
        vec![
            while_stmt(
                binary(BinOp::Lt, fld("sp", "__index"), fld("sp", "__end")),
                vec![
                    var_decl(
                        "value",
                        index_expr(fld("sp", "__items"), fld("sp", "__index")),
                    ),
                    assign(fld("sp", "__index"), add(fld("sp", "__index"), int_lit(1))),
                    stmt(StmtKind::Expr(call_expr(ident("fn"), vec![ident("value")]))),
                ],
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_spliterator_characteristics",
        vec!["sp"],
        vec![
            // ORDERED | SIZED | NONNULL | SUBSIZED
            ret(int_lit(0x0010 + 0x0040 + 0x0100 + 0x4000)),
        ],
    ));
    out.push(function_stmt(
        "__j_spliterator_try_split",
        vec!["sp"],
        vec![
            var_decl(
                "remaining",
                call("__j_spliterator_estimate_size", vec![ident("sp")]),
            ),
            if_stmt(
                binary(BinOp::LtEq, ident("remaining"), int_lit(1)),
                vec![ret(null_lit())],
                None,
            ),
            var_decl("start", fld("sp", "__index")),
            var_decl(
                "mid",
                add(
                    ident("start"),
                    call_expr(
                        member(ident("Math"), "floor"),
                        vec![div(ident("remaining"), int_lit(2))],
                    ),
                ),
            ),
            assign(fld("sp", "__index"), ident("mid")),
            ret(obj_props(vec![
                ("__items", fld("sp", "__items")),
                ("__index", ident("start")),
                ("__end", ident("mid")),
            ])),
        ],
    ));
    out.push(function_stmt(
        "__j_spliterator_get_comparator",
        vec!["sp"],
        vec![ret(null_lit())],
    ));
    out.push(function_stmt(
        "__j_stream_support_stream",
        vec!["sp", "parallel"],
        vec![
            var_decl(
                "arr",
                call_expr(
                    member(fld("sp", "__items"), "slice"),
                    vec![fld("sp", "__index"), fld("sp", "__end")],
                ),
            ),
            assign(fld("arr", "__parallel"), ident("parallel")),
            assign(fld("sp", "__index"), fld("sp", "__end")),
            ret(ident("arr")),
        ],
    ));
    out.push(function_stmt(
        "__j_stream_is_parallel",
        vec!["stream"],
        vec![ret(binary(
            BinOp::Eq,
            fld("stream", "__parallel"),
            bool_lit(true),
        ))],
    ));

    out
}

fn string_fns() -> Vec<Statement> {
    let obj = || expr(ExprKind::Object(Vec::new()));
    let arr = || expr(ExprKind::Array(Vec::new()));
    let bool_lit = |value: bool| expr(ExprKind::Lit(Literal::Bool(value)));
    let fld = |name: &str, f: &str| member(ident(name), f);
    let substr2 =
        |s: Expression, a: Expression, b: Expression| call_expr(member(s, "substring"), vec![a, b]);
    let char_code = |s: Expression, i: Expression| call("__j_char_code_at", vec![s, i]);
    let mut out = Vec::new();

    // String.lines(): split on '\n'; a final terminator yields no trailing
    // empty line (JDK 11 String.lines()).
    out.push(function_stmt(
        "__j_str_lines",
        vec!["s"],
        vec![
            var_decl(
                "parts",
                call_expr(member(to_str(ident("s")), "split"), vec![str_lit("\n")]),
            ),
            var_decl("n", member(ident("parts"), "length")),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Gt, ident("n"), int_lit(0)),
                    binary(
                        BinOp::Eq,
                        index_expr(ident("parts"), binary(BinOp::Sub, ident("n"), int_lit(1))),
                        str_lit(""),
                    ),
                ),
                vec![stmt(StmtKind::Expr(call_expr(
                    member(ident("parts"), "pop"),
                    vec![],
                )))],
                None,
            ),
            ret(ident("parts")),
        ],
    ));

    out.push(function_stmt(
        "__j_string_compare_to",
        vec!["a", "b"],
        vec![
            assign(ident("a"), to_str(ident("a"))),
            assign(ident("b"), to_str(ident("b"))),
            var_decl("la", member(ident("a"), "length")),
            var_decl("lb", member(ident("b"), "length")),
            var_decl("min", ident("la")),
            if_stmt(
                binary(BinOp::Lt, ident("lb"), ident("min")),
                vec![assign(ident("min"), ident("lb"))],
                None,
            ),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("min")),
                vec![
                    var_decl("ca", char_code(ident("a"), ident("i"))),
                    var_decl("cb", char_code(ident("b"), ident("i"))),
                    if_stmt(
                        binary(BinOp::NotEq, ident("ca"), ident("cb")),
                        vec![ret(binary(BinOp::Sub, ident("ca"), ident("cb")))],
                        None,
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(binary(BinOp::Sub, ident("la"), ident("lb"))),
        ],
    ));

    out.push(function_stmt(
        "__j_string_split",
        vec!["s", "re"],
        vec![
            var_decl("p", obj()),
            assign(fld("p", "__re"), ident("re")),
            ret(call(
                "__j_pat_split_impl",
                vec![ident("p"), to_str(ident("s")), int_lit(0)],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_string_split_n",
        vec!["s", "re", "n"],
        vec![
            var_decl("p", obj()),
            assign(fld("p", "__re"), ident("re")),
            ret(call(
                "__j_pat_split_impl",
                vec![ident("p"), to_str(ident("s")), ident("n")],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_string_code_point_at",
        vec!["s", "i"],
        vec![
            var_decl("hi", char_code(ident("s"), ident("i"))),
            if_stmt(
                binary(
                    BinOp::And,
                    call("__j_char_is_high_surrogate", vec![ident("hi")]),
                    binary(
                        BinOp::Lt,
                        add(ident("i"), int_lit(1)),
                        member(ident("s"), "length"),
                    ),
                ),
                vec![
                    var_decl("lo", char_code(ident("s"), add(ident("i"), int_lit(1)))),
                    if_stmt(
                        call("__j_char_is_low_surrogate", vec![ident("lo")]),
                        vec![ret(call(
                            "__j_char_to_code_point",
                            vec![ident("hi"), ident("lo")],
                        ))],
                        None,
                    ),
                ],
                None,
            ),
            ret(ident("hi")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_code_point_before",
        vec!["s", "i"],
        vec![
            if_stmt(
                binary(BinOp::Gt, ident("i"), member(ident("s"), "length")),
                vec![assign(ident("i"), member(ident("s"), "length"))],
                None,
            ),
            assign(ident("i"), binary(BinOp::Sub, ident("i"), int_lit(1))),
            var_decl("lo", char_code(ident("s"), ident("i"))),
            if_stmt(
                binary(
                    BinOp::And,
                    call("__j_char_is_low_surrogate", vec![ident("lo")]),
                    binary(BinOp::Gt, ident("i"), int_lit(0)),
                ),
                vec![
                    var_decl(
                        "hi",
                        char_code(ident("s"), binary(BinOp::Sub, ident("i"), int_lit(1))),
                    ),
                    if_stmt(
                        call("__j_char_is_high_surrogate", vec![ident("hi")]),
                        vec![ret(call(
                            "__j_char_to_code_point",
                            vec![ident("hi"), ident("lo")],
                        ))],
                        None,
                    ),
                ],
                None,
            ),
            ret(ident("lo")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_code_point_count",
        vec!["s", "begin", "end"],
        vec![
            var_decl("n", int_lit(0)),
            var_decl("i", ident("begin")),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("end")),
                vec![
                    var_decl(
                        "cp",
                        call("__j_string_code_point_at", vec![ident("s"), ident("i")]),
                    ),
                    assign(
                        ident("i"),
                        add(ident("i"), call("__j_char_char_count", vec![ident("cp")])),
                    ),
                    assign(ident("n"), add(ident("n"), int_lit(1))),
                ],
            ),
            ret(ident("n")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_offset_by_code_points",
        vec!["s", "index", "off"],
        vec![
            var_decl("i", ident("index")),
            if_stmt(
                binary(BinOp::GtEq, ident("off"), int_lit(0)),
                vec![while_stmt(
                    binary(BinOp::Gt, ident("off"), int_lit(0)),
                    vec![
                        var_decl(
                            "cp",
                            call("__j_string_code_point_at", vec![ident("s"), ident("i")]),
                        ),
                        assign(
                            ident("i"),
                            add(ident("i"), call("__j_char_char_count", vec![ident("cp")])),
                        ),
                        assign(ident("off"), binary(BinOp::Sub, ident("off"), int_lit(1))),
                    ],
                )],
                Some(vec![while_stmt(
                    binary(BinOp::Lt, ident("off"), int_lit(0)),
                    vec![
                        var_decl(
                            "cp",
                            call("__j_string_code_point_before", vec![ident("s"), ident("i")]),
                        ),
                        assign(
                            ident("i"),
                            binary(
                                BinOp::Sub,
                                ident("i"),
                                call("__j_char_char_count", vec![ident("cp")]),
                            ),
                        ),
                        assign(ident("off"), add(ident("off"), int_lit(1))),
                    ],
                )]),
            ),
            ret(ident("i")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_region_matches",
        vec!["s", "toffset", "other", "ooffset", "len"],
        vec![
            if_stmt(
                binary(
                    BinOp::Gt,
                    add(ident("toffset"), ident("len")),
                    member(ident("s"), "length"),
                ),
                vec![ret(bool_lit(false))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::Gt,
                    add(ident("ooffset"), ident("len")),
                    member(ident("other"), "length"),
                ),
                vec![ret(bool_lit(false))],
                None,
            ),
            var_decl(
                "left",
                substr2(
                    ident("s"),
                    ident("toffset"),
                    add(ident("toffset"), ident("len")),
                ),
            ),
            var_decl(
                "right",
                substr2(
                    ident("other"),
                    ident("ooffset"),
                    add(ident("ooffset"), ident("len")),
                ),
            ),
            ret(binary(BinOp::Eq, ident("left"), ident("right"))),
        ],
    ));
    out.push(function_stmt(
        "__j_string_region_matches_ignore",
        vec!["s", "ignore", "toffset", "other", "ooffset", "len"],
        vec![
            if_stmt(
                binary(
                    BinOp::Gt,
                    add(ident("toffset"), ident("len")),
                    member(ident("s"), "length"),
                ),
                vec![ret(bool_lit(false))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::Gt,
                    add(ident("ooffset"), ident("len")),
                    member(ident("other"), "length"),
                ),
                vec![ret(bool_lit(false))],
                None,
            ),
            var_decl(
                "left",
                substr2(
                    ident("s"),
                    ident("toffset"),
                    add(ident("toffset"), ident("len")),
                ),
            ),
            var_decl(
                "right",
                substr2(
                    ident("other"),
                    ident("ooffset"),
                    add(ident("ooffset"), ident("len")),
                ),
            ),
            if_stmt(
                ident("ignore"),
                vec![
                    assign(
                        ident("left"),
                        call_member(ident("left"), "toLowerCase", vec![]),
                    ),
                    assign(
                        ident("right"),
                        call_member(ident("right"), "toLowerCase", vec![]),
                    ),
                ],
                None,
            ),
            ret(binary(BinOp::Eq, ident("left"), ident("right"))),
        ],
    ));
    out.push(function_stmt(
        "__j_string_get_bytes",
        vec!["s"],
        vec![
            var_decl("out", arr()),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                vec![
                    assign(
                        index_expr(ident("out"), ident("i")),
                        char_code(ident("s"), ident("i")),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_chars",
        vec!["s"],
        vec![
            var_decl("out", arr()),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                vec![
                    assign(
                        index_expr(ident("out"), ident("i")),
                        char_code(ident("s"), ident("i")),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_code_points",
        vec!["s"],
        vec![
            var_decl("out", arr()),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                vec![
                    var_decl(
                        "cp",
                        call("__j_string_code_point_at", vec![ident("s"), ident("i")]),
                    ),
                    assign(
                        index_expr(ident("out"), member(ident("out"), "length")),
                        ident("cp"),
                    ),
                    assign(
                        ident("i"),
                        add(ident("i"), call("__j_char_char_count", vec![ident("cp")])),
                    ),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_copy_value_of",
        vec!["a"],
        vec![
            var_decl("off", int_lit(0)),
            var_decl("cnt", member(ident("a"), "length")),
            ret(call(
                "__j_array_chars_to_string",
                vec![ident("a"), ident("off"), ident("cnt")],
            )),
        ],
    ));
    // A surrogate PAIR combines into one real code point before appending —
    // the pair is representable as a string where the lone halves are not
    // (a lone half stays a NUMBER, see `__j_from_char_code`), so
    // `new String(Character.toChars(0x1F600))` round-trips.
    out.push(function_stmt(
        "__j_array_chars_to_string",
        vec!["a", "off", "cnt"],
        vec![
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            var_decl("n", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("cnt")),
                vec![
                    assign(
                        ident("n"),
                        call(
                            "__java_char_ord",
                            vec![index_expr(ident("a"), add(ident("off"), ident("i")))],
                        ),
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            call("__j_char_is_high_surrogate", vec![ident("n")]),
                            binary(
                                BinOp::And,
                                binary(BinOp::Lt, add(ident("i"), int_lit(1)), ident("cnt")),
                                call(
                                    "__j_char_is_low_surrogate",
                                    vec![call(
                                        "__java_char_ord",
                                        vec![index_expr(
                                            ident("a"),
                                            add(add(ident("off"), ident("i")), int_lit(1)),
                                        )],
                                    )],
                                ),
                            ),
                        ),
                        vec![
                            assign(
                                ident("out"),
                                add(
                                    ident("out"),
                                    call(
                                        "__j_from_code_point",
                                        vec![call(
                                            "__j_char_to_code_point",
                                            vec![
                                                ident("n"),
                                                call(
                                                    "__java_char_ord",
                                                    vec![index_expr(
                                                        ident("a"),
                                                        add(
                                                            add(ident("off"), ident("i")),
                                                            int_lit(1),
                                                        ),
                                                    )],
                                                ),
                                            ],
                                        )],
                                    ),
                                ),
                            ),
                            assign(ident("i"), add(ident("i"), int_lit(2))),
                        ],
                        Some(vec![
                            assign(
                                ident("out"),
                                add(ident("out"), call("__j_from_char_code", vec![ident("n")])),
                            ),
                            assign(ident("i"), add(ident("i"), int_lit(1))),
                        ]),
                    ),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_from_array",
        vec!["a"],
        vec![ret(call(
            "__j_array_chars_to_string",
            vec![ident("a"), int_lit(0), member(ident("a"), "length")],
        ))],
    ));
    out.push(function_stmt(
        "__j_code_points_to_string",
        vec!["a", "off", "cnt"],
        vec![
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("cnt")),
                vec![
                    assign(
                        ident("out"),
                        add(
                            ident("out"),
                            call(
                                "__j_from_code_point",
                                vec![index_expr(ident("a"), add(ident("off"), ident("i")))],
                            ),
                        ),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_translate_escapes",
        vec!["s"],
        vec![
            assign(
                ident("s"),
                call_member(ident("s"), "replace", vec![str_lit("\\n"), str_lit("\n")]),
            ),
            assign(
                ident("s"),
                call_member(ident("s"), "replace", vec![str_lit("\\t"), str_lit("\t")]),
            ),
            assign(
                ident("s"),
                call_member(
                    ident("s"),
                    "replace",
                    vec![str_lit("\\u0041"), str_lit("A")],
                ),
            ),
            ret(ident("s")),
        ],
    ));
    out.push(function_stmt(
        "__j_string_strip_indent",
        vec!["s"],
        vec![
            var_decl("n", int_lit(0)),
            while_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Lt, ident("n"), member(ident("s"), "length")),
                    binary(
                        BinOp::Eq,
                        substr2(ident("s"), ident("n"), add(ident("n"), int_lit(1))),
                        str_lit(" "),
                    ),
                ),
                vec![assign(ident("n"), add(ident("n"), int_lit(1)))],
            ),
            if_stmt(
                binary(BinOp::Eq, ident("n"), int_lit(0)),
                vec![ret(ident("s"))],
                None,
            ),
            var_decl("spaces", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("n")),
                vec![
                    assign(ident("spaces"), add(ident("spaces"), str_lit(" "))),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            assign(
                ident("s"),
                call_member(ident("s"), "substring", vec![ident("n")]),
            ),
            assign(
                ident("s"),
                call_member(
                    ident("s"),
                    "replace",
                    vec![add(str_lit("\n"), ident("spaces")), str_lit("\n")],
                ),
            ),
            ret(ident("s")),
        ],
    ));

    out.push(function_stmt(
        "__j_char_char_count",
        vec!["cp"],
        vec![
            assign(ident("cp"), call("__java_char_ord", vec![ident("cp")])),
            if_stmt(
                binary(BinOp::GtEq, ident("cp"), int_lit(65536)),
                vec![ret(int_lit(2))],
                Some(vec![ret(int_lit(1))]),
            ),
        ],
    ));
    out.push(function_stmt(
        "__j_char_to_code_point",
        vec!["hi", "lo"],
        vec![
            assign(ident("hi"), call("__java_char_ord", vec![ident("hi")])),
            assign(ident("lo"), call("__java_char_ord", vec![ident("lo")])),
            if_stmt(
                binary(BinOp::Lt, ident("hi"), int_lit(55296)),
                vec![ret(ident("hi"))],
                None,
            ),
            ret(add(
                add(
                    binary(
                        BinOp::Mul,
                        binary(BinOp::Sub, ident("hi"), int_lit(55296)),
                        int_lit(1024),
                    ),
                    binary(BinOp::Sub, ident("lo"), int_lit(56320)),
                ),
                int_lit(65536),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_high_surrogate",
        vec!["cp"],
        vec![ret(add(
            binary(
                BinOp::Div,
                binary(BinOp::Sub, ident("cp"), int_lit(65536)),
                int_lit(1024),
            ),
            int_lit(55296),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_low_surrogate",
        vec!["cp"],
        vec![ret(add(
            binary(
                BinOp::Mod,
                binary(BinOp::Sub, ident("cp"), int_lit(65536)),
                int_lit(1024),
            ),
            int_lit(56320),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_is_high_surrogate",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::And,
                binary(BinOp::GtEq, ident("n"), int_lit(55296)),
                binary(BinOp::LtEq, ident("n"), int_lit(56319)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_low_surrogate",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::And,
                binary(BinOp::GtEq, ident("n"), int_lit(56320)),
                binary(BinOp::LtEq, ident("n"), int_lit(57343)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_surrogate",
        vec!["c"],
        vec![ret(binary(
            BinOp::Or,
            call("__j_char_is_high_surrogate", vec![ident("c")]),
            call("__j_char_is_low_surrogate", vec![ident("c")]),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_is_digit",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::And,
                binary(BinOp::GtEq, ident("n"), int_lit(48)),
                binary(BinOp::LtEq, ident("n"), int_lit(57)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_letter",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::Or,
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(65)),
                    binary(BinOp::LtEq, ident("n"), int_lit(90)),
                ),
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(97)),
                    binary(BinOp::LtEq, ident("n"), int_lit(122)),
                ),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_alnum",
        vec!["c"],
        vec![ret(binary(
            BinOp::Or,
            call("__j_char_is_letter", vec![ident("c")]),
            call("__j_char_is_digit", vec![ident("c")]),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_is_space",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::Or,
                binary(BinOp::Eq, ident("n"), int_lit(32)),
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(9)),
                    binary(BinOp::LtEq, ident("n"), int_lit(13)),
                ),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_valid_code_point",
        vec!["cp"],
        vec![
            assign(ident("cp"), call("__java_char_ord", vec![ident("cp")])),
            ret(binary(
                BinOp::And,
                binary(BinOp::GtEq, ident("cp"), int_lit(0)),
                binary(BinOp::LtEq, ident("cp"), int_lit(1114111)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_bmp_code_point",
        vec!["cp"],
        vec![
            assign(ident("cp"), call("__java_char_ord", vec![ident("cp")])),
            ret(binary(
                BinOp::And,
                binary(BinOp::GtEq, ident("cp"), int_lit(0)),
                binary(BinOp::LtEq, ident("cp"), int_lit(65535)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_supplementary_code_point",
        vec!["cp"],
        vec![
            assign(ident("cp"), call("__java_char_ord", vec![ident("cp")])),
            ret(binary(
                BinOp::And,
                binary(BinOp::GtEq, ident("cp"), int_lit(65536)),
                binary(BinOp::LtEq, ident("cp"), int_lit(1114111)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_compare",
        vec!["a", "b"],
        vec![ret(binary(
            BinOp::Sub,
            call("__java_char_ord", vec![ident("a")]),
            call("__java_char_ord", vec![ident("b")]),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_reverse_bytes",
        vec!["c"],
        vec![
            assign(ident("c"), call("__java_char_ord", vec![ident("c")])),
            ret(add(
                binary(
                    BinOp::Mul,
                    binary(BinOp::Mod, ident("c"), int_lit(256)),
                    int_lit(256),
                ),
                binary(BinOp::Div, ident("c"), int_lit(256)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_defined",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::And,
                call("__j_char_is_valid_code_point", vec![ident("n")]),
                binary(BinOp::NotEq, ident("n"), int_lit(65535)),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_get_type",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(65)),
                    binary(BinOp::LtEq, ident("n"), int_lit(90)),
                ),
                vec![ret(int_lit(1))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(97)),
                    binary(BinOp::LtEq, ident("n"), int_lit(122)),
                ),
                vec![ret(int_lit(2))],
                None,
            ),
            if_stmt(
                call("__j_char_is_digit", vec![ident("n")]),
                vec![ret(int_lit(9))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("n"), int_lit(32)),
                vec![ret(int_lit(12))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("n"), int_lit(8232)),
                vec![ret(int_lit(13))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("n"), int_lit(36)),
                vec![ret(int_lit(26))],
                None,
            ),
            ret(int_lit(0)),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_java_identifier_start",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::Or,
                call("__j_char_is_letter", vec![ident("n")]),
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, ident("n"), int_lit(36)),
                    binary(BinOp::Eq, ident("n"), int_lit(95)),
                ),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_java_identifier_part",
        vec!["c"],
        vec![ret(binary(
            BinOp::Or,
            call("__j_char_is_java_identifier_start", vec![ident("c")]),
            call("__j_char_is_digit", vec![ident("c")]),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_is_unicode_identifier_start",
        vec!["c"],
        vec![ret(call("__j_char_is_letter", vec![ident("c")]))],
    ));
    out.push(function_stmt(
        "__j_char_is_unicode_identifier_part",
        vec!["c"],
        vec![ret(binary(
            BinOp::Or,
            call("__j_char_is_unicode_identifier_start", vec![ident("c")]),
            binary(
                BinOp::Or,
                call("__j_char_is_digit", vec![ident("c")]),
                binary(
                    BinOp::Eq,
                    call("__java_char_ord", vec![ident("c")]),
                    int_lit(95),
                ),
            ),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_is_iso_control",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::Or,
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(0)),
                    binary(BinOp::LtEq, ident("n"), int_lit(31)),
                ),
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(127)),
                    binary(BinOp::LtEq, ident("n"), int_lit(159)),
                ),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_mirrored",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            ret(binary(
                BinOp::Or,
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, ident("n"), int_lit(40)),
                    binary(BinOp::Eq, ident("n"), int_lit(41)),
                ),
                binary(
                    BinOp::Or,
                    binary(
                        BinOp::Or,
                        binary(BinOp::Eq, ident("n"), int_lit(91)),
                        binary(BinOp::Eq, ident("n"), int_lit(93)),
                    ),
                    binary(
                        BinOp::Or,
                        binary(BinOp::Eq, ident("n"), int_lit(123)),
                        binary(BinOp::Eq, ident("n"), int_lit(125)),
                    ),
                ),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_to_title_case",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(97)),
                    binary(BinOp::LtEq, ident("n"), int_lit(122)),
                ),
                vec![ret(call(
                    "__j_from_char_code",
                    vec![binary(BinOp::Sub, ident("n"), int_lit(32))],
                ))],
                None,
            ),
            ret(ident("c")),
        ],
    ));
    out.push(function_stmt(
        "__j_char_is_surrogate_pair",
        vec!["hi", "lo"],
        vec![ret(binary(
            BinOp::And,
            call("__j_char_is_high_surrogate", vec![ident("hi")]),
            call("__j_char_is_low_surrogate", vec![ident("lo")]),
        ))],
    ));
    out.push(function_stmt(
        "__j_char_unicode_block_of",
        vec!["c"],
        vec![
            var_decl("n", call("__java_char_ord", vec![ident("c")])),
            if_stmt(
                binary(BinOp::LtEq, ident("n"), int_lit(127)),
                vec![ret(str_lit("BASIC_LATIN"))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(880)),
                    binary(BinOp::LtEq, ident("n"), int_lit(1023)),
                ),
                vec![ret(str_lit("GREEK"))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(1024)),
                    binary(BinOp::LtEq, ident("n"), int_lit(1279)),
                ),
                vec![ret(str_lit("CYRILLIC"))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(9472)),
                    binary(BinOp::LtEq, ident("n"), int_lit(9599)),
                ),
                vec![ret(str_lit("BOX_DRAWING"))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("n"), int_lit(12352)),
                    binary(BinOp::LtEq, ident("n"), int_lit(12447)),
                ),
                vec![ret(str_lit("HIRAGANA"))],
                None,
            ),
            ret(str_lit("UNKNOWN")),
        ],
    ));
    out.push(function_stmt(
        "__j_char_digit",
        vec!["c", "radix"],
        vec![
            var_decl("v", int_lit(-1)),
            assign(ident("c"), call("__java_char_ord", vec![ident("c")])),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("c"), int_lit(48)),
                    binary(BinOp::LtEq, ident("c"), int_lit(57)),
                ),
                vec![assign(
                    ident("v"),
                    binary(BinOp::Sub, ident("c"), int_lit(48)),
                )],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("c"), int_lit(65)),
                    binary(BinOp::LtEq, ident("c"), int_lit(90)),
                ),
                vec![assign(
                    ident("v"),
                    add(binary(BinOp::Sub, ident("c"), int_lit(65)), int_lit(10)),
                )],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("c"), int_lit(97)),
                    binary(BinOp::LtEq, ident("c"), int_lit(122)),
                ),
                vec![assign(
                    ident("v"),
                    add(binary(BinOp::Sub, ident("c"), int_lit(97)), int_lit(10)),
                )],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident("v"), int_lit(0)),
                    binary(BinOp::Lt, ident("v"), ident("radix")),
                ),
                vec![ret(ident("v"))],
                None,
            ),
            ret(int_lit(-1)),
        ],
    ));
    out.push(function_stmt(
        "__j_char_for_digit",
        vec!["d", "radix"],
        vec![
            if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Lt, ident("d"), int_lit(0)),
                    binary(
                        BinOp::Or,
                        binary(BinOp::Lt, ident("radix"), int_lit(2)),
                        binary(
                            BinOp::Or,
                            binary(BinOp::Gt, ident("radix"), int_lit(36)),
                            binary(BinOp::GtEq, ident("d"), ident("radix")),
                        ),
                    ),
                ),
                vec![ret(call("__j_from_char_code", vec![int_lit(0)]))],
                None,
            ),
            if_stmt(
                binary(BinOp::Lt, ident("d"), int_lit(10)),
                vec![ret(call(
                    "__j_from_char_code",
                    vec![add(ident("d"), int_lit(48))],
                ))],
                None,
            ),
            ret(call(
                "__j_from_char_code",
                vec![add(ident("d"), int_lit(87))],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_char_numeric",
        vec!["c"],
        vec![ret(call("__j_char_digit", vec![ident("c"), int_lit(36)]))],
    ));
    out.push(function_stmt(
        "__j_char_to_chars",
        vec!["cp"],
        vec![
            var_decl("out", arr()),
            if_stmt(
                binary(BinOp::Lt, ident("cp"), int_lit(65536)),
                vec![assign(
                    index_expr(ident("out"), int_lit(0)),
                    call("__j_from_code_point", vec![ident("cp")]),
                )],
                Some(vec![
                    assign(
                        index_expr(ident("out"), int_lit(0)),
                        call("__j_char_high_surrogate", vec![ident("cp")]),
                    ),
                    assign(
                        index_expr(ident("out"), int_lit(1)),
                        call("__j_char_low_surrogate", vec![ident("cp")]),
                    ),
                ]),
            ),
            ret(ident("out")),
        ],
    ));

    out.push(function_stmt(
        "__j_sj_new",
        vec!["d"],
        vec![
            var_decl("sj", obj()),
            assign(fld("sj", "d"), to_str(ident("d"))),
            assign(fld("sj", "p"), str_lit("")),
            assign(fld("sj", "s"), str_lit("")),
            assign(fld("sj", "empty"), str_lit("")),
            assign(fld("sj", "items"), arr()),
            ret(ident("sj")),
        ],
    ));
    out.push(function_stmt(
        "__j_sj_new3",
        vec!["d", "p", "s"],
        vec![
            var_decl("sj", call("__j_sj_new", vec![ident("d")])),
            assign(fld("sj", "p"), to_str(ident("p"))),
            assign(fld("sj", "s"), to_str(ident("s"))),
            assign(
                fld("sj", "empty"),
                add(to_str(ident("p")), to_str(ident("s"))),
            ),
            ret(ident("sj")),
        ],
    ));
    out.push(function_stmt(
        "__j_sj_add",
        vec!["sj", "x"],
        vec![
            assign(
                index_expr(fld("sj", "items"), member(fld("sj", "items"), "length")),
                to_str(ident("x")),
            ),
            ret(ident("sj")),
        ],
    ));
    out.push(function_stmt(
        "__j_sj_set_empty_value",
        vec!["sj", "x"],
        vec![
            assign(fld("sj", "empty"), to_str(ident("x"))),
            ret(ident("sj")),
        ],
    ));
    out.push(function_stmt(
        "__j_sj_to_string",
        vec!["sj"],
        vec![
            if_stmt(
                binary(BinOp::Eq, member(fld("sj", "items"), "length"), int_lit(0)),
                vec![ret(fld("sj", "empty"))],
                None,
            ),
            ret(add(
                add(
                    fld("sj", "p"),
                    call_member(fld("sj", "items"), "join", vec![fld("sj", "d")]),
                ),
                fld("sj", "s"),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_sj_length",
        vec!["sj"],
        vec![ret(member(
            call("__j_sj_to_string", vec![ident("sj")]),
            "length",
        ))],
    ));
    out.push(function_stmt(
        "__j_sj_merge",
        vec!["a", "b"],
        vec![
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(fld("b", "items"), "length")),
                vec![
                    assign(
                        index_expr(fld("a", "items"), member(fld("a", "items"), "length")),
                        index_expr(fld("b", "items"), ident("i")),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("a")),
        ],
    ));

    out.push(function_stmt(
        "__j_st_new",
        vec!["s"],
        vec![
            var_decl("st", obj()),
            assign(fld("st", "tokens"), arr()),
            assign(fld("st", "i"), int_lit(0)),
            assign(fld("st", "delim"), str_lit(" \t\n\r\u{c}")),
            assign(fld("st", "ret"), bool_lit(false)),
            stmt(StmtKind::Expr(call(
                "__j_st_init",
                vec![ident("st"), to_str(ident("s"))],
            ))),
            ret(ident("st")),
        ],
    ));
    out.push(function_stmt(
        "__j_st_new2",
        vec!["s", "d"],
        vec![
            var_decl("st", call("__j_st_new", vec![ident("s")])),
            assign(fld("st", "tokens"), arr()),
            assign(fld("st", "i"), int_lit(0)),
            assign(fld("st", "delim"), to_str(ident("d"))),
            stmt(StmtKind::Expr(call(
                "__j_st_init",
                vec![ident("st"), to_str(ident("s"))],
            ))),
            ret(ident("st")),
        ],
    ));
    out.push(function_stmt(
        "__j_st_new3",
        vec!["s", "d", "r"],
        vec![
            var_decl("st", obj()),
            assign(fld("st", "tokens"), arr()),
            assign(fld("st", "i"), int_lit(0)),
            assign(fld("st", "delim"), to_str(ident("d"))),
            assign(fld("st", "ret"), ident("r")),
            stmt(StmtKind::Expr(call(
                "__j_st_init",
                vec![ident("st"), to_str(ident("s"))],
            ))),
            ret(ident("st")),
        ],
    ));
    out.push(function_stmt(
        "__j_st_has_more",
        vec!["st"],
        vec![ret(binary(
            BinOp::Lt,
            fld("st", "i"),
            member(fld("st", "tokens"), "length"),
        ))],
    ));
    out.push(function_stmt(
        "__j_st_count",
        vec!["st"],
        vec![ret(binary(
            BinOp::Sub,
            member(fld("st", "tokens"), "length"),
            fld("st", "i"),
        ))],
    ));
    out.push(function_stmt(
        "__j_st_next",
        vec!["st"],
        vec![
            var_decl("v", index_expr(fld("st", "tokens"), fld("st", "i"))),
            assign(fld("st", "i"), add(fld("st", "i"), int_lit(1))),
            ret(ident("v")),
        ],
    ));
    out.push(function_stmt(
        "__j_st_init",
        vec!["st", "s"],
        vec![
            var_decl("tok", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                vec![
                    var_decl(
                        "ch",
                        substr2(ident("s"), ident("i"), add(ident("i"), int_lit(1))),
                    ),
                    if_stmt(
                        binary(
                            BinOp::GtEq,
                            call_member(fld("st", "delim"), "indexOf", vec![ident("ch")]),
                            int_lit(0),
                        ),
                        vec![
                            if_stmt(
                                binary(BinOp::Gt, member(ident("tok"), "length"), int_lit(0)),
                                vec![
                                    assign(
                                        index_expr(
                                            fld("st", "tokens"),
                                            member(fld("st", "tokens"), "length"),
                                        ),
                                        ident("tok"),
                                    ),
                                    assign(ident("tok"), str_lit("")),
                                ],
                                None,
                            ),
                            if_stmt(
                                fld("st", "ret"),
                                vec![assign(
                                    index_expr(
                                        fld("st", "tokens"),
                                        member(fld("st", "tokens"), "length"),
                                    ),
                                    ident("ch"),
                                )],
                                None,
                            ),
                        ],
                        Some(vec![assign(ident("tok"), add(ident("tok"), ident("ch")))]),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            if_stmt(
                binary(BinOp::Gt, member(ident("tok"), "length"), int_lit(0)),
                vec![assign(
                    index_expr(fld("st", "tokens"), member(fld("st", "tokens"), "length")),
                    ident("tok"),
                )],
                None,
            ),
            ret(ident("st")),
        ],
    ));

    out
}

fn scanner_fns() -> Vec<Statement> {
    let obj = || expr(ExprKind::Object(Vec::new()));
    let fld = |name: &str, f: &str| member(ident(name), f);
    let substr2 =
        |s: Expression, a: Expression, b: Expression| call_expr(member(s, "substring"), vec![a, b]);
    let substr1 = |s: Expression, a: Expression| call_expr(member(s, "substring"), vec![a]);
    let mut out = Vec::new();

    out.push(function_stmt(
        "__j_sc_new",
        vec!["s"],
        vec![
            var_decl("sc", obj()),
            assign(fld("sc", "s"), to_str(ident("s"))),
            assign(fld("sc", "pos"), int_lit(0)),
            assign(fld("sc", "delim"), str_lit("\\s+")),
            assign(fld("sc", "radix"), int_lit(10)),
            assign(fld("sc", "comma"), bool_lit(false)),
            ret(ident("sc")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_delim_at",
        vec!["sc", "i"],
        vec![
            var_decl(
                "r",
                call(
                    "__j_re_exec",
                    vec![fld("sc", "delim"), substr1(fld("sc", "s"), ident("i"))],
                ),
            ),
            if_stmt(
                binary(BinOp::Eq, ident("r"), null_lit()),
                vec![ret(int_lit(0))],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, member(ident("r"), "index"), int_lit(0)),
                vec![ret(int_lit(0))],
                None,
            ),
            ret(member(index_expr(ident("r"), int_lit(0)), "length")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_skip_delim",
        vec!["sc"],
        vec![
            var_decl("n", int_lit(1)),
            while_stmt(
                binary(
                    BinOp::And,
                    binary(
                        BinOp::Lt,
                        fld("sc", "pos"),
                        member(fld("sc", "s"), "length"),
                    ),
                    binary(BinOp::Gt, ident("n"), int_lit(0)),
                ),
                vec![
                    assign(
                        ident("n"),
                        call("__j_sc_delim_at", vec![ident("sc"), fld("sc", "pos")]),
                    ),
                    if_stmt(
                        binary(BinOp::Gt, ident("n"), int_lit(0)),
                        vec![assign(fld("sc", "pos"), add(fld("sc", "pos"), ident("n")))],
                        None,
                    ),
                ],
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_peek",
        vec!["sc"],
        vec![
            var_decl("save", fld("sc", "pos")),
            var_decl("t", call("__j_sc_next", vec![ident("sc")])),
            assign(fld("sc", "pos"), ident("save")),
            ret(ident("t")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_next",
        vec!["sc"],
        vec![
            stmt(StmtKind::Expr(call("__j_sc_skip_delim", vec![ident("sc")]))),
            var_decl("start", fld("sc", "pos")),
            var_decl("i", fld("sc", "pos")),
            var_decl("stop", member(fld("sc", "s"), "length")),
            var_decl("d", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("stop")),
                vec![
                    assign(
                        ident("d"),
                        call("__j_sc_delim_at", vec![ident("sc"), ident("i")]),
                    ),
                    if_stmt(
                        binary(BinOp::Gt, ident("d"), int_lit(0)),
                        vec![assign(ident("stop"), ident("i"))],
                        Some(vec![assign(ident("i"), add(ident("i"), int_lit(1)))]),
                    ),
                ],
            ),
            assign(fld("sc", "pos"), ident("stop")),
            ret(substr2(fld("sc", "s"), ident("start"), ident("stop"))),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_norm_num",
        vec!["sc", "t"],
        vec![
            if_stmt(
                fld("sc", "comma"),
                vec![assign(
                    ident("t"),
                    call(
                        "__j_b64_replace_all",
                        vec![ident("t"), str_lit(","), str_lit(".")],
                    ),
                )],
                None,
            ),
            while_stmt(
                binary(
                    BinOp::And,
                    binary(
                        BinOp::Eq,
                        char_at(
                            ident("t"),
                            binary(BinOp::Sub, member(ident("t"), "length"), int_lit(1)),
                        ),
                        str_lit("0"),
                    ),
                    binary(
                        BinOp::NotEq,
                        char_at(
                            ident("t"),
                            binary(BinOp::Sub, member(ident("t"), "length"), int_lit(2)),
                        ),
                        str_lit("."),
                    ),
                ),
                vec![assign(
                    ident("t"),
                    call_member(
                        ident("t"),
                        "substring",
                        vec![
                            int_lit(0),
                            binary(BinOp::Sub, member(ident("t"), "length"), int_lit(1)),
                        ],
                    ),
                )],
            ),
            ret(ident("t")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_next_int",
        vec!["sc"],
        vec![ret(call_expr(
            member(ident("Integer"), "parseInt"),
            vec![call("__j_sc_next", vec![ident("sc")]), fld("sc", "radix")],
        ))],
    ));
    out.push(function_stmt(
        "__j_sc_next_long",
        vec!["sc"],
        vec![
            var_decl("t", call("__j_sc_next", vec![ident("sc")])),
            if_stmt(
                binary(BinOp::Eq, fld("sc", "radix"), int_lit(10)),
                vec![ret(ident("t"))],
                None,
            ),
            ret(call_expr(
                member(ident("Integer"), "parseInt"),
                vec![ident("t"), fld("sc", "radix")],
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_next_double",
        vec!["sc"],
        vec![ret(call(
            "__j_sc_norm_num",
            vec![ident("sc"), call("__j_sc_next", vec![ident("sc")])],
        ))],
    ));
    out.push(function_stmt(
        "__j_sc_next_bool",
        vec!["sc"],
        vec![ret(binary(
            BinOp::Eq,
            call("__j_sc_next", vec![ident("sc")]),
            str_lit("true"),
        ))],
    ));
    out.push(function_stmt(
        "__j_sc_has_next",
        vec!["sc"],
        vec![ret(binary(
            BinOp::Gt,
            member(call("__j_sc_peek", vec![ident("sc")]), "length"),
            int_lit(0),
        ))],
    ));
    out.push(function_stmt(
        "__j_sc_has_next_int",
        vec!["sc"],
        vec![
            var_decl("t", call("__j_sc_peek", vec![ident("sc")])),
            var_decl("pat", str_lit("^-?[0-9]+$")),
            if_stmt(
                binary(BinOp::Eq, fld("sc", "radix"), int_lit(16)),
                vec![assign(ident("pat"), str_lit("^-?[0-9a-fA-F]+$"))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, fld("sc", "radix"), int_lit(8)),
                vec![assign(ident("pat"), str_lit("^-?[0-7]+$"))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, fld("sc", "radix"), int_lit(2)),
                vec![assign(ident("pat"), str_lit("^-?[01]+$"))],
                None,
            ),
            var_decl("r", call("__j_re_exec", vec![ident("pat"), ident("t")])),
            ret(binary(BinOp::NotEq, ident("r"), null_lit())),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_has_next_double",
        vec!["sc"],
        vec![
            var_decl(
                "t",
                call(
                    "__j_sc_norm_num",
                    vec![ident("sc"), call("__j_sc_peek", vec![ident("sc")])],
                ),
            ),
            var_decl("n", call_expr(ident("Number"), vec![ident("t")])),
            ret(binary(
                BinOp::And,
                binary(BinOp::Gt, member(ident("t"), "length"), int_lit(0)),
                binary(BinOp::Eq, ident("n"), ident("n")),
            )),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_next_line",
        vec!["sc"],
        vec![
            var_decl("start", fld("sc", "pos")),
            var_decl("i", fld("sc", "pos")),
            while_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Lt, ident("i"), member(fld("sc", "s"), "length")),
                    binary(
                        BinOp::NotEq,
                        char_at(fld("sc", "s"), ident("i")),
                        str_lit("\n"),
                    ),
                ),
                vec![assign(ident("i"), add(ident("i"), int_lit(1)))],
            ),
            var_decl("line", substr2(fld("sc", "s"), ident("start"), ident("i"))),
            if_stmt(
                binary(BinOp::Lt, ident("i"), member(fld("sc", "s"), "length")),
                vec![assign(ident("i"), add(ident("i"), int_lit(1)))],
                None,
            ),
            assign(fld("sc", "pos"), ident("i")),
            ret(ident("line")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_has_next_line",
        vec!["sc"],
        vec![ret(binary(
            BinOp::Lt,
            fld("sc", "pos"),
            member(fld("sc", "s"), "length"),
        ))],
    ));
    out.push(function_stmt(
        "__j_sc_use_delim",
        vec!["sc", "d"],
        vec![
            assign(fld("sc", "delim"), to_str(ident("d"))),
            ret(ident("sc")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_use_locale",
        vec!["sc", "loc"],
        vec![
            var_decl("s", to_str(ident("loc"))),
            assign(
                fld("sc", "comma"),
                binary(
                    BinOp::Eq,
                    binary(
                        BinOp::Or,
                        binary(
                            BinOp::Or,
                            binary(
                                BinOp::GtEq,
                                call_member(ident("s"), "indexOf", vec![str_lit("US")]),
                                int_lit(0),
                            ),
                            binary(
                                BinOp::GtEq,
                                call_member(ident("s"), "indexOf", vec![str_lit("UK")]),
                                int_lit(0),
                            ),
                        ),
                        binary(
                            BinOp::Or,
                            binary(
                                BinOp::GtEq,
                                call_member(ident("s"), "indexOf", vec![str_lit("JP")]),
                                int_lit(0),
                            ),
                            binary(BinOp::Eq, ident("s"), str_lit("CA")),
                        ),
                    ),
                    bool_lit(false),
                ),
            ),
            ret(ident("sc")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_use_radix",
        vec!["sc", "r"],
        vec![assign(fld("sc", "radix"), ident("r")), ret(ident("sc"))],
    ));
    out.push(function_stmt(
        "__j_sc_skip",
        vec!["sc", "p"],
        vec![
            var_decl(
                "r",
                call(
                    "__j_re_exec",
                    vec![
                        to_str(ident("p")),
                        substr1(fld("sc", "s"), fld("sc", "pos")),
                    ],
                ),
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("r"), null_lit()),
                vec![if_stmt(
                    binary(BinOp::Eq, member(ident("r"), "index"), int_lit(0)),
                    vec![assign(
                        fld("sc", "pos"),
                        add(
                            fld("sc", "pos"),
                            member(index_expr(ident("r"), int_lit(0)), "length"),
                        ),
                    )],
                    None,
                )],
                None,
            ),
            ret(ident("sc")),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_find",
        vec!["sc", "p"],
        vec![
            var_decl(
                "r",
                call(
                    "__j_re_exec",
                    vec![
                        to_str(ident("p")),
                        substr1(fld("sc", "s"), fld("sc", "pos")),
                    ],
                ),
            ),
            if_stmt(
                binary(BinOp::Eq, ident("r"), null_lit()),
                vec![ret(null_lit())],
                None,
            ),
            assign(
                fld("sc", "pos"),
                add(
                    add(fld("sc", "pos"), member(ident("r"), "index")),
                    member(index_expr(ident("r"), int_lit(0)), "length"),
                ),
            ),
            ret(index_expr(ident("r"), int_lit(0))),
        ],
    ));
    out.push(function_stmt(
        "__j_sc_close",
        vec!["sc"],
        vec![ret(null_lit())],
    ));

    out
}

fn thread_fns() -> Vec<Statement> {
    let obj = || expr(ExprKind::Object(Vec::new()));
    let fld = |name: &str, f: &str| member(ident(name), f);
    let this_fld = |f: &str| member(this_expr(), f);
    let arr = |items: Vec<Expression>| {
        expr(ExprKind::Array(
            items
                .into_iter()
                .map(|value| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ))
    };
    let illegal_arg = || {
        stmt(StmtKind::Throw {
            expr: Some(call(
                "__j_exc",
                vec![
                    str_lit("IllegalArgumentException"),
                    arr(vec![
                        str_lit("IllegalArgumentException"),
                        str_lit("RuntimeException"),
                        str_lit("Exception"),
                        str_lit("Throwable"),
                    ]),
                    str_lit("bad range"),
                    undefined_lit(),
                ],
            )),
            cause: None,
        })
    };
    let mut out = Vec::new();

    out.push(var_decl("__j_thread_seq", int_lit(0)));
    out.push(var_decl("__j_main_thread", obj()));
    out.push(assign(fld("__j_main_thread", "name"), str_lit("main")));
    out.push(assign(fld("__j_main_thread", "priority"), int_lit(5)));
    out.push(assign(fld("__j_main_thread", "alive"), bool_lit(true)));
    out.push(assign(
        fld("__j_main_thread", "interrupted"),
        bool_lit(false),
    ));
    out.push(var_decl("__j_current_thread", ident("__j_main_thread")));

    let param = |name: &str| Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };
    out.push(Statement::new(StmtKind::ClassDecl {
        name: "Thread".to_string(),
        parents: Vec::new(),
        interfaces: Vec::new(),
        members: vec![
            ClassMember::Constructor {
                name: None,
                params: vec![param("target"), param("name")],
                body: vec![stmt(StmtKind::Expr(call(
                    "__j_thread_init",
                    vec![this_expr(), ident("target"), ident("name")],
                )))],
                base_args: None,
                initializer_target: ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            },
            ClassMember::Method(Box::new(function_stmt(
                "run",
                vec![],
                vec![ret(call("__j_runnable_run", vec![this_fld("__target")]))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "start",
                vec![],
                vec![ret(call("__j_thread_start", vec![this_expr()]))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "join",
                vec![],
                vec![ret(call("__j_thread_join", vec![this_expr()]))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "isAlive",
                vec![],
                vec![ret(call("__j_thread_is_alive", vec![this_expr()]))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "getName",
                vec![],
                vec![ret(call("__j_thread_get_name", vec![this_expr()]))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "setName",
                vec!["name"],
                vec![ret(call(
                    "__j_thread_set_name",
                    vec![this_expr(), ident("name")],
                ))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "getPriority",
                vec![],
                vec![ret(call("__j_thread_get_priority", vec![this_expr()]))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "setPriority",
                vec!["priority"],
                vec![ret(call(
                    "__j_thread_set_priority",
                    vec![this_expr(), ident("priority")],
                ))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "interrupt",
                vec![],
                vec![ret(call("__j_thread_interrupt", vec![this_expr()]))],
            ))),
            ClassMember::Method(Box::new(function_stmt(
                "isInterrupted",
                vec![],
                vec![ret(call("__j_thread_is_interrupted", vec![this_expr()]))],
            ))),
        ],
        modifiers: ClassModifiers::default(),
        decorators: Vec::new(),
    }));

    out.push(function_stmt(
        "__j_thread_init",
        vec!["t", "target", "name"],
        vec![
            if_stmt(
                binary(BinOp::Eq, typeof_expr(ident("target")), str_lit("string")),
                vec![
                    assign(ident("name"), ident("target")),
                    assign(ident("target"), null_lit()),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("name"), undefined_lit()),
                vec![
                    assign(
                        ident("__j_thread_seq"),
                        add(ident("__j_thread_seq"), int_lit(1)),
                    ),
                    assign(
                        ident("name"),
                        add(str_lit("Thread-"), to_str(ident("__j_thread_seq"))),
                    ),
                ],
                None,
            ),
            assign(fld("t", "__target"), ident("target")),
            assign(fld("t", "name"), ident("name")),
            assign(fld("t", "priority"), int_lit(5)),
            assign(fld("t", "alive"), bool_lit(false)),
            assign(fld("t", "interrupted"), bool_lit(false)),
            assign(fld("t", "__slept"), bool_lit(false)),
            ret(ident("t")),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_new",
        vec!["target", "name"],
        vec![
            var_decl("t", obj()),
            stmt(StmtKind::Expr(call(
                "__j_thread_init",
                vec![ident("t"), ident("target"), ident("name")],
            ))),
            ret(ident("t")),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_current",
        vec![],
        vec![ret(ident("__j_current_thread"))],
    ));
    out.push(function_stmt(
        "__j_monitor_enter",
        vec!["monitor"],
        vec![
            if_stmt(
                binary(
                    BinOp::Eq,
                    fld("monitor", "__j_monitor_owner"),
                    ident("__j_current_thread"),
                ),
                vec![
                    assign(
                        fld("monitor", "__j_monitor_depth"),
                        add(fld("monitor", "__j_monitor_depth"), int_lit(1)),
                    ),
                    ret(null_lit()),
                ],
                None,
            ),
            assign(
                fld("monitor", "__j_monitor_owner"),
                ident("__j_current_thread"),
            ),
            assign(fld("monitor", "__j_monitor_depth"), int_lit(1)),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_monitor_exit",
        vec!["monitor"],
        vec![
            if_stmt(
                binary(
                    BinOp::NotEq,
                    fld("monitor", "__j_monitor_owner"),
                    ident("__j_current_thread"),
                ),
                vec![ret(null_lit())],
                None,
            ),
            assign(
                fld("monitor", "__j_monitor_depth"),
                sub(fld("monitor", "__j_monitor_depth"), int_lit(1)),
            ),
            if_stmt(
                binary(BinOp::LtEq, fld("monitor", "__j_monitor_depth"), int_lit(0)),
                vec![
                    assign(fld("monitor", "__j_monitor_owner"), undefined_lit()),
                    assign(fld("monitor", "__j_monitor_depth"), int_lit(0)),
                ],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_runnable_run",
        vec!["r"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("r"), null_lit()),
                vec![ret(null_lit())],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, typeof_expr(ident("r")), str_lit("function")),
                vec![ret(call_expr(ident("r"), vec![]))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Eq, typeof_expr(ident("r")), str_lit("object")),
                    binary(BinOp::StrictNotEq, fld("r", "job"), undefined_lit()),
                ),
                vec![
                    stmt(StmtKind::Expr(call(
                        "__j_future_task_run",
                        vec![ident("r")],
                    ))),
                    ret(null_lit()),
                ],
                None,
            ),
            ret(call_member(ident("r"), "run", vec![])),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_start",
        vec!["t"],
        vec![ret(call(
            "__j_thread_start_with",
            vec![ident("t"), fld("t", "__target")],
        ))],
    ));
    out.push(function_stmt(
        "__j_thread_start_with",
        vec!["t", "target"],
        vec![
            assign(fld("t", "alive"), bool_lit(true)),
            assign(fld("t", "__slept"), bool_lit(false)),
            var_decl("prev", ident("__j_current_thread")),
            assign(ident("__j_current_thread"), ident("t")),
            stmt(StmtKind::Try {
                body: vec![if_stmt(
                    binary(
                        BinOp::And,
                        binary(BinOp::NotEq, ident("target"), null_lit()),
                        binary(BinOp::NotEq, ident("target"), undefined_lit()),
                    ),
                    vec![stmt(StmtKind::Expr(call(
                        "__j_runnable_run",
                        vec![ident("target")],
                    )))],
                    Some(vec![stmt(StmtKind::Expr(call_member(
                        ident("t"),
                        "run",
                        vec![],
                    )))]),
                )],
                catches: vec![CatchClause {
                    types: vec!["__JavaWaitYield".to_string()],
                    var_name: None,
                    stack_var: None,
                    body: vec![assign(fld("t", "__slept"), bool_lit(true))],
                    when_clause: None,
                }],
                else_body: None,
                finally: None,
            }),
            assign(ident("__j_current_thread"), ident("prev")),
            if_stmt(
                binary(BinOp::NotEq, fld("t", "__slept"), bool_lit(true)),
                vec![assign(fld("t", "alive"), bool_lit(false))],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_join",
        vec!["t"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("t"), ident("__j_current_thread")),
                vec![ret(null_lit())],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, fld("t", "alive"), bool_lit(true)),
                vec![stmt(StmtKind::Expr(call(
                    "__j_thread_start_with",
                    vec![ident("t"), fld("t", "__target")],
                )))],
                None,
            ),
            assign(fld("t", "alive"), bool_lit(false)),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_sleep",
        vec!["ms"],
        vec![
            if_stmt(
                binary(BinOp::Gt, ident("ms"), int_lit(25)),
                vec![assign(fld("__j_current_thread", "__slept"), bool_lit(true))],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_is_alive",
        vec!["t"],
        vec![ret(binary(BinOp::Eq, fld("t", "alive"), bool_lit(true)))],
    ));
    out.push(function_stmt(
        "__j_thread_get_name",
        vec!["t"],
        vec![ret(fld("t", "name"))],
    ));
    out.push(function_stmt(
        "__j_thread_set_name",
        vec!["t", "name"],
        vec![assign(fld("t", "name"), ident("name")), ret(null_lit())],
    ));
    out.push(function_stmt(
        "__j_thread_get_priority",
        vec!["t"],
        vec![ret(fld("t", "priority"))],
    ));
    out.push(function_stmt(
        "__j_thread_set_priority",
        vec!["t", "priority"],
        vec![
            assign(fld("t", "priority"), ident("priority")),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_interrupt",
        vec!["t"],
        vec![
            assign(fld("t", "interrupted"), bool_lit(true)),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_thread_is_interrupted",
        vec!["t"],
        vec![ret(binary(
            BinOp::Eq,
            fld("t", "interrupted"),
            bool_lit(true),
        ))],
    ));
    out.push(function_stmt(
        "__j_thread_interrupted",
        vec![],
        vec![
            var_decl(
                "v",
                call(
                    "__j_thread_is_interrupted",
                    vec![ident("__j_current_thread")],
                ),
            ),
            assign(fld("__j_current_thread", "interrupted"), bool_lit(false)),
            ret(ident("v")),
        ],
    ));

    out.push(function_stmt(
        "__j_callable_call",
        vec!["r"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("r"), null_lit()),
                vec![ret(null_lit())],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, typeof_expr(ident("r")), str_lit("function")),
                vec![ret(call_expr(ident("r"), vec![]))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Eq, typeof_expr(ident("r")), str_lit("object")),
                    binary(
                        BinOp::StrictNotEq,
                        member(ident("r"), "job"),
                        undefined_lit(),
                    ),
                ),
                vec![
                    stmt(StmtKind::Expr(call(
                        "__j_future_task_run",
                        vec![ident("r")],
                    ))),
                    ret(null_lit()),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, member(ident("r"), "call"), undefined_lit()),
                vec![ret(call_member(ident("r"), "call", vec![]))],
                None,
            ),
            ret(call("__j_runnable_run", vec![ident("r")])),
        ],
    ));
    out.push(function_stmt(
        "__j_latch_new",
        vec!["count"],
        vec![
            var_decl("l", obj()),
            assign(fld("l", "count"), ident("count")),
            ret(ident("l")),
        ],
    ));
    out.push(function_stmt(
        "__j_latch_count_down",
        vec!["l"],
        vec![
            if_stmt(
                binary(BinOp::Gt, fld("l", "count"), int_lit(0)),
                vec![assign(
                    fld("l", "count"),
                    sub(fld("l", "count"), int_lit(1)),
                )],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_latch_get_count",
        vec!["l"],
        vec![ret(fld("l", "count"))],
    ));
    out.push(function_stmt(
        "__j_latch_await",
        vec!["l"],
        vec![
            if_stmt(
                binary(BinOp::Gt, fld("l", "count"), int_lit(0)),
                vec![stmt(StmtKind::Expr(call(
                    "__java_thread_sleep",
                    vec![int_lit(5)],
                )))],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_latch_await_timeout",
        vec!["l", "time", "unit"],
        vec![ret(binary(BinOp::LtEq, fld("l", "count"), int_lit(0)))],
    ));
    out.push(function_stmt(
        "__j_future_task_new",
        vec!["job", "result", "hasPreset"],
        vec![
            var_decl("f", obj()),
            assign(fld("f", "job"), ident("job")),
            assign(fld("f", "preset"), ident("result")),
            assign(fld("f", "hasPreset"), ident("hasPreset")),
            assign(fld("f", "done"), bool_lit(false)),
            assign(fld("f", "cancelled"), bool_lit(false)),
            assign(fld("f", "result"), null_lit()),
            assign(fld("f", "error"), null_lit()),
            ret(ident("f")),
        ],
    ));
    out.push(function_stmt(
        "__j_future_task_run",
        vec!["f"],
        vec![
            if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, fld("f", "done"), bool_lit(true)),
                    binary(BinOp::Eq, fld("f", "cancelled"), bool_lit(true)),
                ),
                vec![ret(null_lit())],
                None,
            ),
            stmt(StmtKind::Try {
                body: vec![if_stmt(
                    binary(BinOp::Eq, fld("f", "hasPreset"), bool_lit(true)),
                    vec![
                        stmt(StmtKind::Expr(call(
                            "__j_runnable_run",
                            vec![fld("f", "job")],
                        ))),
                        assign(fld("f", "result"), fld("f", "preset")),
                    ],
                    Some(vec![assign(
                        fld("f", "result"),
                        call("__j_callable_call", vec![fld("f", "job")]),
                    )]),
                )],
                catches: vec![CatchClause {
                    types: vec!["Exception".to_string(), "Throwable".to_string()],
                    var_name: Some("e".to_string()),
                    stack_var: None,
                    body: vec![assign(fld("f", "error"), ident("e"))],
                    when_clause: None,
                }],
                else_body: None,
                finally: Some(vec![assign(fld("f", "done"), bool_lit(true))]),
            }),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_future_task_get",
        vec!["f", "time", "unit"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("f", "cancelled"), bool_lit(true)),
                vec![stmt(StmtKind::Throw {
                    expr: Some(call(
                        "__j_exc",
                        vec![
                            str_lit("CancellationException"),
                            arr(vec![
                                str_lit("CancellationException"),
                                str_lit("RuntimeException"),
                                str_lit("Exception"),
                                str_lit("Throwable"),
                            ]),
                            str_lit("cancelled"),
                            undefined_lit(),
                        ],
                    )),
                    cause: None,
                })],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, fld("f", "done"), bool_lit(false)),
                vec![stmt(StmtKind::Expr(call(
                    "__j_future_task_run",
                    vec![ident("f")],
                )))],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, fld("f", "error"), null_lit()),
                vec![stmt(StmtKind::Throw {
                    expr: Some(call(
                        "__j_exc",
                        vec![
                            str_lit("ExecutionException"),
                            arr(vec![
                                str_lit("ExecutionException"),
                                str_lit("Exception"),
                                str_lit("Throwable"),
                            ]),
                            str_lit("execution failed"),
                            fld("f", "error"),
                        ],
                    )),
                    cause: None,
                })],
                None,
            ),
            ret(fld("f", "result")),
        ],
    ));
    out.push(function_stmt(
        "__j_future_task_cancel",
        vec!["f", "mayInterrupt"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("f", "done"), bool_lit(false)),
                vec![
                    assign(fld("f", "cancelled"), bool_lit(true)),
                    assign(fld("f", "done"), bool_lit(true)),
                ],
                None,
            ),
            ret(fld("f", "cancelled")),
        ],
    ));
    out.push(function_stmt(
        "__j_future_task_is_done",
        vec!["f"],
        vec![ret(binary(BinOp::Eq, fld("f", "done"), bool_lit(true)))],
    ));
    out.push(function_stmt(
        "__j_future_task_is_cancelled",
        vec!["f"],
        vec![ret(binary(
            BinOp::Eq,
            fld("f", "cancelled"),
            bool_lit(true),
        ))],
    ));
    out.push(function_stmt(
        "__j_future_task_run_and_reset",
        vec!["f"],
        vec![
            assign(fld("f", "done"), bool_lit(false)),
            assign(fld("f", "cancelled"), bool_lit(false)),
            assign(fld("f", "error"), null_lit()),
            assign(fld("f", "done"), bool_lit(false)),
            ret(bool_lit(true)),
        ],
    ));
    out.push(function_stmt(
        "__j_exec_new",
        vec!["size"],
        vec![
            var_decl("p", obj()),
            assign(fld("p", "shutdown"), bool_lit(false)),
            assign(fld("p", "size"), ident("size")),
            ret(ident("p")),
        ],
    ));
    out.push(function_stmt(
        "__j_exec_submit",
        vec!["p", "job"],
        vec![
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::Eq, typeof_expr(ident("job")), str_lit("object")),
                    binary(
                        BinOp::StrictNotEq,
                        member(ident("job"), "job"),
                        undefined_lit(),
                    ),
                ),
                vec![
                    stmt(StmtKind::Expr(call(
                        "__j_future_task_run",
                        vec![ident("job")],
                    ))),
                    ret(ident("job")),
                ],
                None,
            ),
            var_decl(
                "f",
                call(
                    "__j_future_task_new",
                    vec![ident("job"), undefined_lit(), bool_lit(false)],
                ),
            ),
            stmt(StmtKind::Expr(call(
                "__j_future_task_run",
                vec![ident("f")],
            ))),
            ret(ident("f")),
        ],
    ));
    out.push(function_stmt(
        "__j_exec_execute",
        vec!["p", "job"],
        vec![
            stmt(StmtKind::Expr(call(
                "__j_exec_submit",
                vec![ident("p"), ident("job")],
            ))),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_exec_shutdown",
        vec!["p"],
        vec![
            assign(fld("p", "shutdown"), bool_lit(true)),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_exec_is_shutdown",
        vec!["p"],
        vec![ret(binary(BinOp::Eq, fld("p", "shutdown"), bool_lit(true)))],
    ));
    out.push(function_stmt(
        "__j_exec_await",
        vec!["p", "time", "unit"],
        vec![ret(bool_lit(true))],
    ));
    out.push(function_stmt(
        "__j_exec_shutdown_now",
        vec!["p"],
        vec![assign(fld("p", "shutdown"), bool_lit(true)), ret(arr_lit())],
    ));

    out.push(function_stmt(
        "__j_tlr_current",
        vec![],
        vec![
            if_stmt(
                binary(
                    BinOp::Eq,
                    fld("__j_current_thread", "__tlr"),
                    undefined_lit(),
                ),
                vec![assign(fld("__j_current_thread", "__tlr"), obj())],
                None,
            ),
            ret(fld("__j_current_thread", "__tlr")),
        ],
    ));
    out.push(function_stmt(
        "__j_tlr_next_int",
        vec!["rng", "a", "b"],
        vec![
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::NotEq, ident("a"), undefined_lit()),
                    binary(BinOp::Eq, ident("b"), undefined_lit()),
                ),
                vec![
                    if_stmt(
                        binary(BinOp::LtEq, ident("a"), int_lit(0)),
                        vec![illegal_arg()],
                        None,
                    ),
                    ret(int_lit(0)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("b"), undefined_lit()),
                vec![
                    if_stmt(
                        binary(BinOp::LtEq, ident("b"), ident("a")),
                        vec![illegal_arg()],
                        None,
                    ),
                    ret(ident("a")),
                ],
                None,
            ),
            ret(int_lit(1)),
        ],
    ));
    out.push(function_stmt(
        "__j_tlr_next_long",
        vec!["rng", "a", "b"],
        vec![
            if_stmt(
                binary(
                    BinOp::And,
                    binary(BinOp::NotEq, ident("a"), undefined_lit()),
                    binary(BinOp::Eq, ident("b"), undefined_lit()),
                ),
                vec![
                    if_stmt(
                        binary(BinOp::LtEq, ident("a"), int_lit(0)),
                        vec![illegal_arg()],
                        None,
                    ),
                    ret(int_lit(0)),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("b"), undefined_lit()),
                vec![
                    if_stmt(
                        binary(BinOp::LtEq, ident("b"), ident("a")),
                        vec![illegal_arg()],
                        None,
                    ),
                    ret(ident("a")),
                ],
                None,
            ),
            ret(int_lit(1)),
        ],
    ));
    out.push(function_stmt(
        "__j_tlr_next_double",
        vec!["rng", "a", "b"],
        vec![
            if_stmt(
                binary(BinOp::NotEq, ident("b"), undefined_lit()),
                vec![
                    if_stmt(
                        binary(BinOp::Lt, ident("b"), ident("a")),
                        vec![illegal_arg()],
                        None,
                    ),
                    ret(ident("a")),
                ],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("a"), undefined_lit()),
                vec![
                    if_stmt(
                        binary(BinOp::LtEq, ident("a"), int_lit(0)),
                        vec![illegal_arg()],
                        None,
                    ),
                    ret(expr_f64(0.5)),
                ],
                None,
            ),
            ret(expr_f64(0.5)),
        ],
    ));
    out.push(function_stmt(
        "__j_tlr_next_float",
        vec!["rng"],
        vec![ret(expr_f64(0.5))],
    ));
    out.push(function_stmt(
        "__j_tlr_next_bool",
        vec!["rng"],
        vec![ret(bool_lit(true))],
    ));
    out.push(function_stmt(
        "__j_tlr_next_gaussian",
        vec!["rng"],
        vec![ret(expr_f64(0.0))],
    ));
    out.push(function_stmt(
        "__j_tlr_stream",
        vec!["size", "value"],
        vec![
            if_stmt(
                binary(BinOp::Lt, ident("size"), int_lit(0)),
                vec![illegal_arg()],
                None,
            ),
            var_decl("out", arr_lit()),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("size")),
                vec![
                    assign(index_expr(ident("out"), ident("i")), ident("value")),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_tlr_ints",
        vec!["rng", "size", "origin", "bound"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("size"), undefined_lit()),
                vec![assign(ident("size"), int_lit(1))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("origin"), undefined_lit()),
                vec![assign(ident("origin"), int_lit(0))],
                None,
            ),
            ret(call("__j_tlr_stream", vec![ident("size"), ident("origin")])),
        ],
    ));
    out.push(function_stmt(
        "__j_tlr_longs",
        vec!["rng", "size", "origin", "bound"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("size"), undefined_lit()),
                vec![assign(ident("size"), int_lit(1))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("origin"), undefined_lit()),
                vec![assign(ident("origin"), int_lit(0))],
                None,
            ),
            ret(call("__j_tlr_stream", vec![ident("size"), ident("origin")])),
        ],
    ));
    out.push(function_stmt(
        "__j_tlr_doubles",
        vec!["rng", "size", "origin", "bound"],
        vec![
            if_stmt(
                binary(BinOp::Eq, ident("size"), undefined_lit()),
                vec![assign(ident("size"), int_lit(1))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("origin"), undefined_lit()),
                vec![assign(ident("origin"), expr_f64(0.5))],
                None,
            ),
            ret(call("__j_tlr_stream", vec![ident("size"), ident("origin")])),
        ],
    ));

    out
}

/// `java.util.regex` Pattern/Matcher over `ecma:regexp` (patterns are
/// plain strings; `__j_re_exec` returns the ECMA match array with
/// `.index`, or null). The Matcher carries Java's find() cursor.
fn regex_fns() -> Vec<Statement> {
    let obj = || expr(ExprKind::Object(Vec::new()));
    let fld = |name: &str, f: &str| member(ident(name), f);
    let substr_range =
        |s: Expression, a: Expression, b: Expression| call_expr(member(s, "substring"), vec![a, b]);
    let mut out = Vec::new();

    // Built-in java exception the shared emitter doesn't recognize:
    // canonical exception object with the JLS supertype chain in __types
    // (REF_TEST catch matching walks it). getMessage() reads .message.
    out.push(function_stmt(
        "__j_exc",
        vec!["n", "chain", "msg", "cause"],
        vec![
            var_decl("e", obj()),
            assign(fld("e", "__type"), ident("n")),
            assign(fld("e", "__exception_type"), ident("n")),
            assign(fld("e", "name"), ident("n")),
            assign(fld("e", "__types"), ident("chain")),
            if_stmt(
                binary(BinOp::Eq, ident("msg"), undefined_lit()),
                vec![assign(fld("e", "message"), null_lit())],
                Some(vec![assign(fld("e", "message"), ident("msg"))]),
            ),
            if_stmt(
                binary(BinOp::NotEq, ident("cause"), undefined_lit()),
                vec![assign(fld("e", "cause"), ident("cause"))],
                None,
            ),
            ret(ident("e")),
        ],
    ));
    // Throwable cause plumbing.
    out.push(function_stmt(
        "__j_get_cause",
        vec!["e"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("e", "cause"), undefined_lit()),
                vec![ret(null_lit())],
                None,
            ),
            ret(fld("e", "cause")),
        ],
    ));
    out.push(function_stmt(
        "__j_init_cause",
        vec!["e", "c"],
        vec![assign(fld("e", "cause"), ident("c")), ret(ident("e"))],
    ));
    out.push(function_stmt(
        "__j_pat_compile",
        vec!["re"],
        vec![
            var_decl("p", obj()),
            assign(fld("p", "__re"), ident("re")),
            ret(ident("p")),
        ],
    ));
    out.push(function_stmt(
        "__j_pat_pattern",
        vec!["p"],
        vec![
            if_stmt(
                binary(BinOp::NotEq, fld("p", "__src"), undefined_lit()),
                vec![ret(fld("p", "__src"))],
                None,
            ),
            ret(fld("p", "__re")),
        ],
    ));
    // Pattern.compile(regex, flags) — java flag bits lowered onto the host's
    // `/pattern/flags` regex-literal shape (ecma:regexp splits on the last
    // unescaped slash). LITERAL/COMMENTS/CANON_EQ preprocess the pattern.
    out.push(function_stmt(
        "__j_pat_compile_flags",
        vec!["re", "f"],
        vec![
            var_decl("p", obj()),
            assign(fld("p", "__src"), ident("re")),
            assign(fld("p", "__flags"), ident("f")),
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(16)),
                vec![assign(ident("re"), call("__j_re_quote", vec![ident("re")]))],
                None,
            ),
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(4)),
                vec![assign(
                    ident("re"),
                    call("__j_re_strip_ws", vec![ident("re")]),
                )],
                None,
            ),
            // UNICODE_CHARACTER_CLASS: java `\p{IsFoo}` → JS `\p{Foo}`.
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(256)),
                vec![assign(
                    ident("re"),
                    call_expr(
                        member(
                            call_expr(member(ident("re"), "split"), vec![str_lit("\\p{Is")]),
                            "join",
                        ),
                        vec![str_lit("\\p{")],
                    ),
                )],
                None,
            ),
            // CANON_EQ: NFC-normalize pattern now, inputs at matcher/reset.
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(128)),
                vec![
                    assign(
                        ident("re"),
                        call("__j_normalize", vec![ident("re"), str_lit("NFC")]),
                    ),
                    assign(fld("p", "__canon"), int_lit(1)),
                ],
                None,
            ),
            var_decl("fl", str_lit("")),
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(2)),
                vec![assign(ident("fl"), add(ident("fl"), str_lit("i")))],
                None,
            ),
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(8)),
                vec![assign(ident("fl"), add(ident("fl"), str_lit("m")))],
                None,
            ),
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(32)),
                vec![assign(ident("fl"), add(ident("fl"), str_lit("s")))],
                None,
            ),
            // UNICODE_CASE(64) | UNICODE_CHARACTER_CLASS(256) → `u`.
            if_stmt(
                binary(BinOp::BitAnd, ident("f"), int_lit(320)),
                vec![assign(ident("fl"), add(ident("fl"), str_lit("u")))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, ident("fl"), str_lit("")),
                vec![assign(fld("p", "__re"), ident("re"))],
                Some(vec![assign(
                    fld("p", "__re"),
                    add(
                        add(
                            add(str_lit("/"), call("__j_re_slashes", vec![ident("re")])),
                            str_lit("/"),
                        ),
                        ident("fl"),
                    ),
                )]),
            ),
            ret(ident("p")),
        ],
    ));
    out.push(function_stmt(
        "__j_pat_flags",
        vec!["p"],
        vec![
            if_stmt(
                binary(BinOp::Eq, fld("p", "__flags"), undefined_lit()),
                vec![ret(int_lit(0))],
                None,
            ),
            ret(fld("p", "__flags")),
        ],
    ));
    // Pattern.quote / LITERAL: escape every regex metacharacter.
    out.push(function_stmt(
        "__j_re_quote",
        vec!["re"],
        vec![
            assign(ident("re"), to_str(ident("re"))),
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("re"), "length")),
                vec![
                    var_decl("c", char_at(ident("re"), ident("i"))),
                    if_stmt(
                        binary(
                            BinOp::GtEq,
                            call_expr(
                                member(str_lit("\\^$.|?*+()[]{}/"), "indexOf"),
                                vec![ident("c")],
                            ),
                            int_lit(0),
                        ),
                        vec![assign(
                            ident("out"),
                            add(add(ident("out"), str_lit("\\")), ident("c")),
                        )],
                        Some(vec![assign(ident("out"), add(ident("out"), ident("c")))]),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    // COMMENTS mode: drop unescaped whitespace and `#`-to-EOL comments;
    // a backslash keeps its next character verbatim (`a\ b` matches "a b").
    out.push(function_stmt(
        "__j_re_strip_ws",
        vec!["re"],
        vec![
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("re"), "length")),
                vec![
                    var_decl("c", char_at(ident("re"), ident("i"))),
                    if_stmt(
                        binary(BinOp::Eq, ident("c"), str_lit("\\")),
                        vec![
                            assign(ident("out"), add(ident("out"), ident("c"))),
                            assign(ident("i"), add(ident("i"), int_lit(1))),
                            if_stmt(
                                binary(BinOp::Lt, ident("i"), member(ident("re"), "length")),
                                vec![assign(
                                    ident("out"),
                                    add(ident("out"), char_at(ident("re"), ident("i"))),
                                )],
                                None,
                            ),
                        ],
                        Some(vec![if_stmt(
                            binary(
                                BinOp::Or,
                                binary(
                                    BinOp::Or,
                                    binary(BinOp::Eq, ident("c"), str_lit(" ")),
                                    binary(BinOp::Eq, ident("c"), str_lit("\t")),
                                ),
                                binary(BinOp::Eq, ident("c"), str_lit("\n")),
                            ),
                            vec![],
                            Some(vec![if_stmt(
                                binary(BinOp::Eq, ident("c"), str_lit("#")),
                                vec![while_stmt(
                                    binary(
                                        BinOp::And,
                                        binary(
                                            BinOp::Lt,
                                            ident("i"),
                                            member(ident("re"), "length"),
                                        ),
                                        binary(
                                            BinOp::NotEq,
                                            char_at(ident("re"), ident("i")),
                                            str_lit("\n"),
                                        ),
                                    ),
                                    vec![assign(ident("i"), add(ident("i"), int_lit(1)))],
                                )],
                                Some(vec![assign(ident("out"), add(ident("out"), ident("c")))]),
                            )]),
                        )]),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    // Escape raw '/' so the composed literal survives the host's
    // last-unescaped-slash split.
    out.push(function_stmt(
        "__j_re_slashes",
        vec!["re"],
        vec![
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("re"), "length")),
                vec![
                    var_decl("c", char_at(ident("re"), ident("i"))),
                    if_stmt(
                        binary(BinOp::Eq, ident("c"), str_lit("/")),
                        vec![assign(ident("out"), add(ident("out"), str_lit("\\/")))],
                        Some(vec![assign(ident("out"), add(ident("out"), ident("c")))]),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_split_before_upper",
        vec!["s", "n"],
        vec![
            var_decl("parts", expr(ExprKind::Array(Vec::new()))),
            var_decl("count", int_lit(0)),
            var_decl("start", int_lit(0)),
            var_decl("i", int_lit(1)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                vec![
                    var_decl("ch", call("__j_char_code_at", vec![ident("s"), ident("i")])),
                    if_stmt(
                        binary(
                            BinOp::And,
                            binary(BinOp::GtEq, ident("ch"), int_lit(65)),
                            binary(BinOp::LtEq, ident("ch"), int_lit(90)),
                        ),
                        vec![if_stmt(
                            binary(
                                BinOp::And,
                                binary(BinOp::Gt, ident("n"), int_lit(0)),
                                binary(
                                    BinOp::GtEq,
                                    ident("count"),
                                    binary(BinOp::Sub, ident("n"), int_lit(1)),
                                ),
                            ),
                            vec![assign(ident("i"), member(ident("s"), "length"))],
                            Some(vec![
                                assign(
                                    index_expr(ident("parts"), ident("count")),
                                    substr_range(ident("s"), ident("start"), ident("i")),
                                ),
                                assign(ident("count"), add(ident("count"), int_lit(1))),
                                assign(ident("start"), ident("i")),
                            ]),
                        )],
                        None,
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            assign(
                index_expr(ident("parts"), ident("count")),
                call_expr(member(ident("s"), "substring"), vec![ident("start")]),
            ),
            ret(ident("parts")),
        ],
    ));
    // Java split semantics (JLS Pattern.split): limit n>0 = at most n
    // parts with the remainder attached to the last; n==0 = unlimited,
    // trailing empty strings removed; n<0 = unlimited, empties kept.
    out.push(function_stmt(
        "__j_pat_split_impl",
        vec!["p", "s", "n"],
        vec![
            var_decl("parts", expr(ExprKind::Array(Vec::new()))),
            if_stmt(
                binary(BinOp::Eq, fld("p", "__re"), str_lit("(?=[A-Z])")),
                vec![ret(call(
                    "__j_split_before_upper",
                    vec![ident("s"), ident("n")],
                ))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, member(ident("s"), "length"), int_lit(0)),
                vec![
                    assign(index_expr(ident("parts"), int_lit(0)), str_lit("")),
                    ret(ident("parts")),
                ],
                None,
            ),
            // java.util.regex.Pattern.split: `pos` = current segment start,
            // `cur` = search cursor. A zero-width match emits the pending
            // (possibly empty) segment and bumps the cursor one past the
            // match so the scan advances; a zero-width match at position 0
            // produces no leading empty segment.
            var_decl("count", int_lit(0)),
            var_decl("pos", int_lit(0)),
            var_decl("cur", int_lit(0)),
            var_decl("go", int_lit(1)),
            while_stmt(
                binary(BinOp::Eq, ident("go"), int_lit(1)),
                vec![if_stmt(
                    binary(
                        BinOp::Or,
                        binary(BinOp::Gt, ident("cur"), member(ident("s"), "length")),
                        binary(
                            BinOp::And,
                            binary(BinOp::Gt, ident("n"), int_lit(0)),
                            binary(
                                BinOp::GtEq,
                                ident("count"),
                                binary(BinOp::Sub, ident("n"), int_lit(1)),
                            ),
                        ),
                    ),
                    vec![assign(ident("go"), int_lit(0))],
                    Some(vec![
                        var_decl(
                            "r",
                            call(
                                "__j_re_exec",
                                vec![
                                    fld("p", "__re"),
                                    call_expr(member(ident("s"), "substring"), vec![ident("cur")]),
                                ],
                            ),
                        ),
                        if_stmt(
                            binary(BinOp::Eq, ident("r"), null_lit()),
                            vec![assign(ident("go"), int_lit(0))],
                            Some(vec![
                                var_decl(
                                    "mlen",
                                    member(index_expr(ident("r"), int_lit(0)), "length"),
                                ),
                                var_decl("mstart", add(ident("cur"), member(ident("r"), "index"))),
                                if_stmt(
                                    binary(
                                        BinOp::And,
                                        binary(BinOp::Eq, ident("mlen"), int_lit(0)),
                                        binary(BinOp::Eq, ident("mstart"), int_lit(0)),
                                    ),
                                    vec![assign(ident("cur"), int_lit(1))],
                                    Some(vec![
                                        assign(
                                            index_expr(ident("parts"), ident("count")),
                                            substr_range(ident("s"), ident("pos"), ident("mstart")),
                                        ),
                                        assign(ident("count"), add(ident("count"), int_lit(1))),
                                        assign(ident("pos"), add(ident("mstart"), ident("mlen"))),
                                        assign(ident("cur"), ident("pos")),
                                        if_stmt(
                                            binary(BinOp::Eq, ident("mlen"), int_lit(0)),
                                            vec![assign(
                                                ident("cur"),
                                                add(ident("cur"), int_lit(1)),
                                            )],
                                            None,
                                        ),
                                    ]),
                                ),
                            ]),
                        ),
                    ]),
                )],
            ),
            assign(
                index_expr(ident("parts"), ident("count")),
                call_expr(member(ident("s"), "substring"), vec![ident("pos")]),
            ),
            assign(ident("count"), add(ident("count"), int_lit(1))),
            // limit 0: drop trailing empty strings.
            if_stmt(
                binary(BinOp::Eq, ident("n"), int_lit(0)),
                vec![
                    var_decl("last", binary(BinOp::Sub, ident("count"), int_lit(1))),
                    while_stmt(
                        binary(
                            BinOp::And,
                            binary(BinOp::GtEq, ident("last"), int_lit(0)),
                            binary(
                                BinOp::Eq,
                                index_expr(ident("parts"), ident("last")),
                                str_lit(""),
                            ),
                        ),
                        vec![assign(
                            ident("last"),
                            binary(BinOp::Sub, ident("last"), int_lit(1)),
                        )],
                    ),
                    var_decl("trimmed", expr(ExprKind::Array(Vec::new()))),
                    var_decl("i", int_lit(0)),
                    while_stmt(
                        binary(BinOp::LtEq, ident("i"), ident("last")),
                        vec![
                            assign(
                                index_expr(ident("trimmed"), ident("i")),
                                index_expr(ident("parts"), ident("i")),
                            ),
                            assign(ident("i"), add(ident("i"), int_lit(1))),
                        ],
                    ),
                    ret(ident("trimmed")),
                ],
                None,
            ),
            ret(ident("parts")),
        ],
    ));
    out.push(function_stmt(
        "__j_pat_split",
        vec!["p", "s"],
        vec![ret(call(
            "__j_pat_split_impl",
            vec![ident("p"), ident("s"), int_lit(0)],
        ))],
    ));
    out.push(function_stmt(
        "__j_pat_split_n",
        vec!["p", "s", "n"],
        vec![ret(call(
            "__j_pat_split_impl",
            vec![ident("p"), ident("s"), ident("n")],
        ))],
    ));
    out.push(function_stmt(
        "__j_pat_matcher",
        vec!["p", "s"],
        vec![
            var_decl("m", obj()),
            assign(fld("m", "__re"), fld("p", "__re")),
            assign(fld("m", "__canon"), fld("p", "__canon")),
            var_decl("t", to_str(ident("s"))),
            if_stmt(
                binary(BinOp::Eq, fld("p", "__canon"), int_lit(1)),
                vec![assign(
                    ident("t"),
                    call("__j_normalize", vec![ident("t"), str_lit("NFC")]),
                )],
                None,
            ),
            assign(fld("m", "__input"), ident("t")),
            assign(fld("m", "__pos"), int_lit(0)),
            assign(fld("m", "__append_pos"), int_lit(0)),
            assign(fld("m", "__m"), null_lit()),
            assign(fld("m", "__start"), int_lit(-1)),
            assign(fld("m", "__end"), int_lit(-1)),
            ret(ident("m")),
        ],
    ));
    // reset() / reset(newInput): rewind the cursor, optionally re-target.
    out.push(function_stmt(
        "__j_m_reset",
        vec!["m", "s"],
        vec![
            if_stmt(
                binary(BinOp::NotEq, ident("s"), undefined_lit()),
                vec![
                    var_decl("t", to_str(ident("s"))),
                    if_stmt(
                        binary(BinOp::Eq, fld("m", "__canon"), int_lit(1)),
                        vec![assign(
                            ident("t"),
                            call("__j_normalize", vec![ident("t"), str_lit("NFC")]),
                        )],
                        None,
                    ),
                    assign(fld("m", "__input"), ident("t")),
                ],
                None,
            ),
            assign(fld("m", "__pos"), int_lit(0)),
            assign(fld("m", "__append_pos"), int_lit(0)),
            assign(fld("m", "__m"), null_lit()),
            assign(fld("m", "__start"), int_lit(-1)),
            assign(fld("m", "__end"), int_lit(-1)),
            ret(ident("m")),
        ],
    ));
    // find(): search from the cursor; store the match, advance past it
    // (by one on an empty match, as java.util.regex does).
    out.push(function_stmt(
        "__j_m_find",
        vec!["m"],
        vec![
            if_stmt(
                binary(
                    BinOp::Gt,
                    fld("m", "__pos"),
                    member(fld("m", "__input"), "length"),
                ),
                vec![
                    assign(fld("m", "__m"), null_lit()),
                    ret(expr(ExprKind::Lit(Literal::Bool(false)))),
                ],
                None,
            ),
            var_decl(
                "tail",
                call_expr(
                    member(fld("m", "__input"), "substring"),
                    vec![fld("m", "__pos")],
                ),
            ),
            var_decl(
                "r",
                call("__j_re_exec", vec![fld("m", "__re"), ident("tail")]),
            ),
            if_stmt(
                binary(BinOp::Eq, ident("r"), null_lit()),
                vec![
                    assign(fld("m", "__m"), null_lit()),
                    ret(expr(ExprKind::Lit(Literal::Bool(false)))),
                ],
                None,
            ),
            assign(fld("m", "__m"), ident("r")),
            assign(
                fld("m", "__start"),
                add(fld("m", "__pos"), member(ident("r"), "index")),
            ),
            var_decl("adv", member(index_expr(ident("r"), int_lit(0)), "length")),
            assign(fld("m", "__end"), add(fld("m", "__start"), ident("adv"))),
            if_stmt(
                binary(BinOp::Eq, ident("adv"), int_lit(0)),
                vec![assign(ident("adv"), int_lit(1))],
                None,
            ),
            assign(
                fld("m", "__pos"),
                add(
                    add(fld("m", "__pos"), member(ident("r"), "index")),
                    ident("adv"),
                ),
            ),
            ret(expr(ExprKind::Lit(Literal::Bool(true)))),
        ],
    ));
    // matches(): the whole region must match (anchored both ends).
    out.push(function_stmt(
        "__j_m_matches",
        vec!["m"],
        vec![
            var_decl(
                "r",
                call("__j_re_exec", vec![fld("m", "__re"), fld("m", "__input")]),
            ),
            if_stmt(
                binary(BinOp::Eq, ident("r"), null_lit()),
                vec![ret(expr(ExprKind::Lit(Literal::Bool(false))))],
                None,
            ),
            if_stmt(
                binary(BinOp::NotEq, member(ident("r"), "index"), int_lit(0)),
                vec![ret(expr(ExprKind::Lit(Literal::Bool(false))))],
                None,
            ),
            if_stmt(
                binary(
                    BinOp::Eq,
                    member(index_expr(ident("r"), int_lit(0)), "length"),
                    member(fld("m", "__input"), "length"),
                ),
                vec![ret(expr(ExprKind::Lit(Literal::Bool(true))))],
                None,
            ),
            ret(expr(ExprKind::Lit(Literal::Bool(false)))),
        ],
    ));
    // lookingAt(): anchored at the start only.
    out.push(function_stmt(
        "__j_m_looking_at",
        vec!["m"],
        vec![
            var_decl(
                "r",
                call("__j_re_exec", vec![fld("m", "__re"), fld("m", "__input")]),
            ),
            if_stmt(
                binary(BinOp::Eq, ident("r"), null_lit()),
                vec![ret(expr(ExprKind::Lit(Literal::Bool(false))))],
                None,
            ),
            if_stmt(
                binary(BinOp::Eq, member(ident("r"), "index"), int_lit(0)),
                vec![ret(expr(ExprKind::Lit(Literal::Bool(true))))],
                None,
            ),
            ret(expr(ExprKind::Lit(Literal::Bool(false)))),
        ],
    ));
    out.push(function_stmt(
        "__j_m_group",
        vec!["m", "i"],
        vec![ret(index_expr(fld("m", "__m"), ident("i")))],
    ));
    out.push(function_stmt(
        "__j_m_expand_repl",
        vec!["m", "repl"],
        vec![
            assign(ident("repl"), to_str(ident("repl"))),
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), member(ident("repl"), "length")),
                vec![
                    var_decl("ch", char_at(ident("repl"), ident("i"))),
                    if_stmt(
                        binary(BinOp::Eq, ident("ch"), str_lit("\\")),
                        vec![
                            assign(ident("i"), add(ident("i"), int_lit(1))),
                            if_stmt(
                                binary(BinOp::Lt, ident("i"), member(ident("repl"), "length")),
                                vec![assign(
                                    ident("out"),
                                    add(ident("out"), char_at(ident("repl"), ident("i"))),
                                )],
                                None,
                            ),
                        ],
                        Some(vec![if_stmt(
                            binary(
                                BinOp::And,
                                binary(BinOp::Eq, ident("ch"), str_lit("$")),
                                binary(
                                    BinOp::Lt,
                                    add(ident("i"), int_lit(1)),
                                    member(ident("repl"), "length"),
                                ),
                            ),
                            vec![
                                var_decl(
                                    "d",
                                    call(
                                        "__java_char_ord",
                                        vec![char_at(ident("repl"), add(ident("i"), int_lit(1)))],
                                    ),
                                ),
                                if_stmt(
                                    binary(
                                        BinOp::And,
                                        binary(BinOp::GtEq, ident("d"), int_lit(48)),
                                        binary(BinOp::LtEq, ident("d"), int_lit(57)),
                                    ),
                                    vec![
                                        assign(
                                            ident("d"),
                                            binary(BinOp::Sub, ident("d"), int_lit(48)),
                                        ),
                                        assign(
                                            ident("out"),
                                            add(
                                                ident("out"),
                                                to_str(index_expr(fld("m", "__m"), ident("d"))),
                                            ),
                                        ),
                                        assign(ident("i"), add(ident("i"), int_lit(1))),
                                    ],
                                    Some(vec![assign(
                                        ident("out"),
                                        add(ident("out"), ident("ch")),
                                    )]),
                                ),
                            ],
                            Some(vec![assign(ident("out"), add(ident("out"), ident("ch")))]),
                        )]),
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(ident("out")),
        ],
    ));
    out.push(function_stmt(
        "__j_m_replace_all",
        vec!["m", "repl"],
        vec![ret(call(
            "__j_re_replace_all",
            vec![fld("m", "__input"), fld("m", "__re"), ident("repl")],
        ))],
    ));
    out.push(function_stmt(
        "__j_m_replace_first",
        vec!["m", "repl"],
        vec![
            assign(fld("m", "__pos"), int_lit(0)),
            if_stmt(
                call("__j_m_find", vec![ident("m")]),
                vec![ret(add(
                    add(
                        substr_range(fld("m", "__input"), int_lit(0), fld("m", "__start")),
                        call("__j_m_expand_repl", vec![ident("m"), ident("repl")]),
                    ),
                    call_expr(
                        member(fld("m", "__input"), "substring"),
                        vec![fld("m", "__end")],
                    ),
                ))],
                None,
            ),
            ret(fld("m", "__input")),
        ],
    ));
    out.push(function_stmt(
        "__j_m_append_replacement",
        vec!["m", "sb", "repl"],
        vec![
            stmt(StmtKind::Expr(call(
                "__j_sb_append",
                vec![
                    ident("sb"),
                    substr_range(
                        fld("m", "__input"),
                        fld("m", "__append_pos"),
                        fld("m", "__start"),
                    ),
                ],
            ))),
            stmt(StmtKind::Expr(call(
                "__j_sb_append",
                vec![
                    ident("sb"),
                    call("__j_m_expand_repl", vec![ident("m"), ident("repl")]),
                ],
            ))),
            assign(fld("m", "__append_pos"), fld("m", "__end")),
            ret(ident("m")),
        ],
    ));
    out.push(function_stmt(
        "__j_m_append_tail",
        vec!["m", "sb"],
        vec![
            stmt(StmtKind::Expr(call(
                "__j_sb_append",
                vec![
                    ident("sb"),
                    call_expr(
                        member(fld("m", "__input"), "substring"),
                        vec![fld("m", "__append_pos")],
                    ),
                ],
            ))),
            assign(
                fld("m", "__append_pos"),
                member(fld("m", "__input"), "length"),
            ),
            ret(ident("sb")),
        ],
    ));
    out
}

/// `java.lang.StringBuilder` methods over the dotnet stringbuilder shape
/// (an Object holding the text in `__buffer` and an int `Capacity` —
/// `platforms/dotnet/emitter/core/stringbuilder_adapter.rs`). The walker
/// routes StringBuilder-typed receivers here; mutators return the builder
/// (JLS: they return `this`), so calls chain.
fn stringbuilder_fns() -> Vec<Statement> {
    let buf = |sb: &str| member(ident(sb), "__buffer");
    let buf_set = |sb: &str, v: Expression| assign(member(ident(sb), "__buffer"), v);
    let substr2 =
        |s: Expression, a: Expression, b: Expression| call_expr(member(s, "substring"), vec![a, b]);
    let substr1 = |s: Expression, a: Expression| call_expr(member(s, "substring"), vec![a]);
    let mut out = Vec::new();

    out.push(function_stmt(
        "__j_sb_new",
        vec!["seed"],
        vec![
            var_decl("sb", obj_lit()),
            assign(buf("sb"), str_lit("")),
            assign(member(ident("sb"), "Capacity"), int_lit(16)),
            if_stmt(
                binary(BinOp::NotEq, ident("seed"), undefined_lit()),
                vec![if_stmt(
                    binary(BinOp::Eq, typeof_expr(ident("seed")), str_lit("number")),
                    vec![assign(member(ident("sb"), "Capacity"), ident("seed"))],
                    Some(vec![
                        assign(buf("sb"), to_str(ident("seed"))),
                        assign(
                            member(ident("sb"), "Capacity"),
                            add(member(buf("sb"), "length"), int_lit(16)),
                        ),
                    ]),
                )],
                None,
            ),
            ret(ident("sb")),
        ],
    ));
    out.push(function_stmt(
        "__j_sb_to_string",
        vec!["sb"],
        vec![ret(buf("sb"))],
    ));
    out.push(function_stmt(
        "__j_sb_length",
        vec!["sb"],
        vec![ret(member(buf("sb"), "length"))],
    ));
    out.push(function_stmt(
        "__j_sb_append",
        vec!["sb", "x"],
        vec![
            buf_set("sb", add(buf("sb"), to_str(ident("x")))),
            ret(ident("sb")),
        ],
    ));
    out.push(function_stmt(
        "__j_sb_append_code_point",
        vec!["sb", "cp"],
        vec![
            buf_set(
                "sb",
                add(buf("sb"), call("__j_from_code_point", vec![ident("cp")])),
            ),
            ret(ident("sb")),
        ],
    ));
    out.push(function_stmt(
        "__j_sb_char_at",
        vec!["sb", "i"],
        vec![ret(substr2(
            buf("sb"),
            ident("i"),
            add(ident("i"), int_lit(1)),
        ))],
    ));
    out.push(function_stmt(
        "__j_sb_set_char_at",
        vec!["sb", "i", "c"],
        vec![
            buf_set(
                "sb",
                add(
                    add(
                        substr2(buf("sb"), int_lit(0), ident("i")),
                        to_str(ident("c")),
                    ),
                    substr1(buf("sb"), add(ident("i"), int_lit(1))),
                ),
            ),
            ret(null_lit()),
        ],
    ));
    out.push(function_stmt(
        "__j_sb_insert",
        vec!["sb", "off", "x"],
        vec![
            buf_set(
                "sb",
                add(
                    add(
                        substr2(buf("sb"), int_lit(0), ident("off")),
                        to_str(ident("x")),
                    ),
                    substr1(buf("sb"), ident("off")),
                ),
            ),
            ret(ident("sb")),
        ],
    ));
    // delete(start, end) — end clamps to length (JLS).
    out.push(function_stmt(
        "__j_sb_delete",
        vec!["sb", "s", "e"],
        vec![
            var_decl("n", member(buf("sb"), "length")),
            if_stmt(
                binary(BinOp::Gt, ident("e"), ident("n")),
                vec![assign(ident("e"), ident("n"))],
                None,
            ),
            buf_set(
                "sb",
                add(
                    substr2(buf("sb"), int_lit(0), ident("s")),
                    substr1(buf("sb"), ident("e")),
                ),
            ),
            ret(ident("sb")),
        ],
    ));
    out.push(function_stmt(
        "__j_sb_delete_char_at",
        vec!["sb", "i"],
        vec![ret(call(
            "__j_sb_delete",
            vec![ident("sb"), ident("i"), add(ident("i"), int_lit(1))],
        ))],
    ));
    // replace(start, end, str) — end clamps to length (JLS).
    out.push(function_stmt(
        "__j_sb_replace",
        vec!["sb", "s", "e", "str"],
        vec![
            var_decl("n", member(buf("sb"), "length")),
            if_stmt(
                binary(BinOp::Gt, ident("e"), ident("n")),
                vec![assign(ident("e"), ident("n"))],
                None,
            ),
            buf_set(
                "sb",
                add(
                    add(
                        substr2(buf("sb"), int_lit(0), ident("s")),
                        to_str(ident("str")),
                    ),
                    substr1(buf("sb"), ident("e")),
                ),
            ),
            ret(ident("sb")),
        ],
    ));
    out.push(function_stmt(
        "__j_sb_reverse",
        vec!["sb"],
        vec![
            var_decl("acc", str_lit("")),
            var_decl(
                "i",
                binary(BinOp::Sub, member(buf("sb"), "length"), int_lit(1)),
            ),
            while_stmt(
                binary(BinOp::GtEq, ident("i"), int_lit(0)),
                vec![
                    assign(
                        ident("acc"),
                        add(
                            ident("acc"),
                            substr2(buf("sb"), ident("i"), add(ident("i"), int_lit(1))),
                        ),
                    ),
                    assign(ident("i"), binary(BinOp::Sub, ident("i"), int_lit(1))),
                ],
            ),
            buf_set("sb", ident("acc")),
            ret(ident("sb")),
        ],
    ));
    // setLength: truncate, or pad with ' ' (JLS).
    out.push(function_stmt(
        "__j_sb_set_length",
        vec!["sb", "n"],
        vec![
            if_stmt(
                binary(BinOp::LtEq, ident("n"), member(buf("sb"), "length")),
                vec![buf_set("sb", substr2(buf("sb"), int_lit(0), ident("n")))],
                Some(vec![while_stmt(
                    binary(BinOp::Lt, member(buf("sb"), "length"), ident("n")),
                    vec![buf_set("sb", add(buf("sb"), str_lit("\u{0}")))],
                )]),
            ),
            ret(null_lit()),
        ],
    ));
    // capacity(): the tracked field, never less than the content length.
    out.push(function_stmt(
        "__j_sb_capacity",
        vec!["sb"],
        vec![
            var_decl("c", member(ident("sb"), "Capacity")),
            var_decl("n", member(buf("sb"), "length")),
            if_stmt(
                binary(BinOp::Gt, ident("n"), ident("c")),
                vec![assign(ident("c"), ident("n"))],
                None,
            ),
            ret(ident("c")),
        ],
    ));
    out.push(function_stmt(
        "__j_sb_ensure_capacity",
        vec!["sb", "n"],
        vec![
            if_stmt(
                binary(BinOp::Gt, ident("n"), member(ident("sb"), "Capacity")),
                vec![assign(member(ident("sb"), "Capacity"), ident("n"))],
                None,
            ),
            ret(null_lit()),
        ],
    ));
    // JLS String.compareTo: difference of first differing chars, else
    // length difference.
    out.push(function_stmt(
        "__j_sb_compare_to",
        vec!["sb", "other"],
        vec![
            var_decl("a", buf("sb")),
            var_decl("b", member(ident("other"), "__buffer")),
            var_decl("la", member(ident("a"), "length")),
            var_decl("lb", member(ident("b"), "length")),
            var_decl("min", ident("la")),
            if_stmt(
                binary(BinOp::Lt, ident("lb"), ident("min")),
                vec![assign(ident("min"), ident("lb"))],
                None,
            ),
            var_decl("i", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("min")),
                vec![
                    var_decl(
                        "ca",
                        call_expr(member(ident("a"), "charCodeAt"), vec![ident("i")]),
                    ),
                    var_decl(
                        "cb",
                        call_expr(member(ident("b"), "charCodeAt"), vec![ident("i")]),
                    ),
                    if_stmt(
                        binary(BinOp::NotEq, ident("ca"), ident("cb")),
                        vec![ret(binary(BinOp::Sub, ident("ca"), ident("cb")))],
                        None,
                    ),
                    assign(ident("i"), add(ident("i"), int_lit(1))),
                ],
            ),
            ret(binary(BinOp::Sub, ident("la"), ident("lb"))),
        ],
    ));
    out
}

/// `__j_sprintf(fmt, args)` — the Java `Formatter` scanner. Java-specific
/// conversions are computed here; everything else delegates one specifier
/// at a time to the shared engine (`__java_string_format` builtin).
fn sprintf_fn() -> Statement {
    // Shared-engine delegation for one specifier: "%" + flags + width + prec + conv.
    let spec_expr = add(
        add(
            add(add(str_lit("%"), ident("flags")), ident("width")),
            ident("prec"),
        ),
        ident("conv"),
    );
    let delegate = call("__java_string_format", vec![spec_expr, ident("a")]);

    // %g/%G body: pick %f-style or %e-style per Java Formatter rules
    // (6 significant digits by default).
    let g_body = vec![
        var_decl(
            "es",
            call("__java_string_format", vec![str_lit("%e"), ident("a")]),
        ),
        var_decl(
            "ep",
            call_member(ident("es"), "lastIndexOf", vec![str_lit("e")]),
        ),
        var_decl(
            "exv",
            call_member(
                ident("Integer"),
                "parseInt",
                vec![call_member(
                    ident("es"),
                    "substring",
                    vec![add(ident("ep"), int_lit(1))],
                )],
            ),
        ),
        var_decl("pr", int_lit(6)),
        if_stmt(
            binary(BinOp::NotEq, ident("prec"), str_lit("")),
            vec![assign(
                ident("pr"),
                call_member(
                    ident("Integer"),
                    "parseInt",
                    vec![call_member(ident("prec"), "substring", vec![int_lit(1)])],
                ),
            )],
            None,
        ),
        if_stmt(
            binary(BinOp::Eq, ident("pr"), int_lit(0)),
            vec![assign(ident("pr"), int_lit(1))],
            None,
        ),
        if_stmt(
            binary(
                BinOp::And,
                binary(BinOp::GtEq, ident("exv"), int_lit(-4)),
                binary(BinOp::Lt, ident("exv"), ident("pr")),
            ),
            vec![assign(
                ident("piece"),
                call(
                    "__java_string_format",
                    vec![
                        add(
                            add(
                                str_lit("%."),
                                to_str(binary(
                                    BinOp::Sub,
                                    binary(BinOp::Sub, ident("pr"), int_lit(1)),
                                    ident("exv"),
                                )),
                            ),
                            str_lit("f"),
                        ),
                        ident("a"),
                    ],
                ),
            )],
            Some(vec![assign(
                ident("piece"),
                call(
                    "__j_expad",
                    vec![call(
                        "__java_string_format",
                        vec![
                            add(
                                add(
                                    str_lit("%."),
                                    to_str(binary(BinOp::Sub, ident("pr"), int_lit(1))),
                                ),
                                str_lit("e"),
                            ),
                            ident("a"),
                        ],
                    )],
                ),
            )]),
        ),
        if_stmt(
            binary(BinOp::Eq, ident("conv"), str_lit("G")),
            vec![assign(
                ident("piece"),
                call_member(ident("piece"), "toUpperCase", vec![]),
            )],
            None,
        ),
        assign(
            ident("piece"),
            call(
                "__j_padw",
                vec![ident("piece"), ident("width"), ident("left")],
            ),
        ),
    ];

    // The specifier dispatch chain (b/B, grouped d, e/E, g/G, delegate).
    let conv_dispatch = if_stmt(
        binary(
            BinOp::Or,
            binary(BinOp::Eq, ident("conv"), str_lit("b")),
            binary(BinOp::Eq, ident("conv"), str_lit("B")),
        ),
        vec![
            // Boolean.toString semantics: null → false, boolean → itself,
            // anything else → true.
            if_stmt(
                binary(BinOp::Eq, ident("a"), null_lit()),
                vec![assign(ident("piece"), str_lit("false"))],
                Some(vec![if_stmt(
                    binary(BinOp::Eq, to_str(ident("a")), str_lit("false")),
                    vec![assign(ident("piece"), str_lit("false"))],
                    Some(vec![assign(ident("piece"), str_lit("true"))]),
                )]),
            ),
            if_stmt(
                binary(BinOp::Eq, ident("conv"), str_lit("B")),
                vec![assign(
                    ident("piece"),
                    call_member(ident("piece"), "toUpperCase", vec![]),
                )],
                None,
            ),
            assign(
                ident("piece"),
                call(
                    "__j_padw",
                    vec![ident("piece"), ident("width"), ident("left")],
                ),
            ),
        ],
        Some(vec![if_stmt(
            binary(
                BinOp::And,
                binary(BinOp::Eq, ident("conv"), str_lit("d")),
                binary(BinOp::Eq, ident("grouped"), int_lit(1)),
            ),
            vec![
                assign(
                    ident("piece"),
                    call(
                        "__j_group",
                        vec![call(
                            "__java_string_format",
                            vec![str_lit("%d"), ident("a")],
                        )],
                    ),
                ),
                assign(
                    ident("piece"),
                    call(
                        "__j_padw",
                        vec![ident("piece"), ident("width"), ident("left")],
                    ),
                ),
            ],
            Some(vec![if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, ident("conv"), str_lit("f")),
                    binary(BinOp::Eq, ident("conv"), str_lit("F")),
                ),
                vec![
                    var_decl("prf", int_lit(6)),
                    if_stmt(
                        binary(BinOp::NotEq, ident("prec"), str_lit("")),
                        vec![assign(
                            ident("prf"),
                            call_member(
                                ident("Integer"),
                                "parseInt",
                                vec![call_member(ident("prec"), "substring", vec![int_lit(1)])],
                            ),
                        )],
                        None,
                    ),
                    assign(
                        ident("piece"),
                        call(
                            "tofixed",
                            vec![
                                call(
                                    "__j_bd_round_half_up",
                                    vec![
                                        call_expr(ident("Number"), vec![ident("a")]),
                                        ident("prf"),
                                    ],
                                ),
                                ident("prf"),
                            ],
                        ),
                    ),
                    if_stmt(
                        binary(BinOp::Eq, ident("grouped"), int_lit(1)),
                        vec![assign(
                            ident("piece"),
                            call("__j_group_number", vec![ident("piece")]),
                        )],
                        None,
                    ),
                    assign(
                        ident("piece"),
                        call(
                            "__j_padw",
                            vec![ident("piece"), ident("width"), ident("left")],
                        ),
                    ),
                ],
                Some(vec![if_stmt(
                    binary(
                        BinOp::Or,
                        binary(BinOp::Eq, ident("conv"), str_lit("e")),
                        binary(BinOp::Eq, ident("conv"), str_lit("E")),
                    ),
                    vec![assign(
                        ident("piece"),
                        call("__j_expad", vec![delegate.clone()]),
                    )],
                    Some(vec![if_stmt(
                        binary(
                            BinOp::Or,
                            binary(BinOp::Eq, ident("conv"), str_lit("g")),
                            binary(BinOp::Eq, ident("conv"), str_lit("G")),
                        ),
                        g_body,
                        Some(vec![assign(ident("piece"), delegate)]),
                    )]),
                )]),
            )]),
        )]),
    );

    // Specifier parse: [argindex$][flags][width][.prec]conv
    let spec_parse = vec![
        var_decl("j", add(ident("i"), int_lit(1))),
        // Leading digits + '$' → explicit argument index.
        var_decl("digs", str_lit("")),
        var_decl("k", ident("j")),
        while_stmt(
            binary(
                BinOp::And,
                binary(BinOp::Lt, ident("k"), ident("n")),
                call("__j_isdig", vec![char_at(ident("fmt"), ident("k"))]),
            ),
            vec![
                assign(
                    ident("digs"),
                    add(ident("digs"), char_at(ident("fmt"), ident("k"))),
                ),
                assign(ident("k"), add(ident("k"), int_lit(1))),
            ],
        ),
        var_decl("argidx", int_lit(0)),
        if_stmt(
            binary(
                BinOp::And,
                binary(BinOp::NotEq, ident("digs"), str_lit("")),
                binary(
                    BinOp::And,
                    binary(BinOp::Lt, ident("k"), ident("n")),
                    binary(BinOp::Eq, char_at(ident("fmt"), ident("k")), str_lit("$")),
                ),
            ),
            vec![
                assign(
                    ident("argidx"),
                    call_member(ident("Integer"), "parseInt", vec![ident("digs")]),
                ),
                assign(ident("j"), add(ident("k"), int_lit(1))),
            ],
            None,
        ),
        // Flags ('-', '+', ' ', '0', '(' pass through; ',' is Java grouping).
        var_decl("flags", str_lit("")),
        var_decl("grouped", int_lit(0)),
        var_decl("left", int_lit(0)),
        var_decl("f", char_at(ident("fmt"), ident("j"))),
        while_stmt(
            binary(
                BinOp::And,
                binary(BinOp::Lt, ident("j"), ident("n")),
                binary(
                    BinOp::Or,
                    binary(
                        BinOp::GtEq,
                        call_member(str_lit("-+ 0(#"), "indexOf", vec![ident("f")]),
                        int_lit(0),
                    ),
                    binary(BinOp::Eq, ident("f"), str_lit(",")),
                ),
            ),
            vec![
                if_stmt(
                    binary(BinOp::Eq, ident("f"), str_lit(",")),
                    vec![assign(ident("grouped"), int_lit(1))],
                    Some(vec![
                        assign(ident("flags"), add(ident("flags"), ident("f"))),
                        if_stmt(
                            binary(BinOp::Eq, ident("f"), str_lit("-")),
                            vec![assign(ident("left"), int_lit(1))],
                            None,
                        ),
                    ]),
                ),
                assign(ident("j"), add(ident("j"), int_lit(1))),
                assign(ident("f"), char_at(ident("fmt"), ident("j"))),
            ],
        ),
        // Width digits.
        var_decl("width", str_lit("")),
        while_stmt(
            binary(
                BinOp::And,
                binary(BinOp::Lt, ident("j"), ident("n")),
                call("__j_isdig", vec![char_at(ident("fmt"), ident("j"))]),
            ),
            vec![
                assign(
                    ident("width"),
                    add(ident("width"), char_at(ident("fmt"), ident("j"))),
                ),
                assign(ident("j"), add(ident("j"), int_lit(1))),
            ],
        ),
        // Precision.
        var_decl("prec", str_lit("")),
        if_stmt(
            binary(
                BinOp::And,
                binary(BinOp::Lt, ident("j"), ident("n")),
                binary(BinOp::Eq, char_at(ident("fmt"), ident("j")), str_lit(".")),
            ),
            vec![
                assign(ident("prec"), str_lit(".")),
                assign(ident("j"), add(ident("j"), int_lit(1))),
                while_stmt(
                    binary(
                        BinOp::And,
                        binary(BinOp::Lt, ident("j"), ident("n")),
                        call("__j_isdig", vec![char_at(ident("fmt"), ident("j"))]),
                    ),
                    vec![
                        assign(
                            ident("prec"),
                            add(ident("prec"), char_at(ident("fmt"), ident("j"))),
                        ),
                        assign(ident("j"), add(ident("j"), int_lit(1))),
                    ],
                ),
            ],
            None,
        ),
        var_decl("conv", char_at(ident("fmt"), ident("j"))),
        // Argument selection: explicit %N$ or the running cursor.
        var_decl("a", null_lit()),
        if_stmt(
            binary(BinOp::Gt, ident("argidx"), int_lit(0)),
            vec![assign(
                ident("a"),
                index_expr(
                    ident("args"),
                    binary(BinOp::Sub, ident("argidx"), int_lit(1)),
                ),
            )],
            Some(vec![
                assign(ident("a"), index_expr(ident("args"), ident("argi"))),
                assign(ident("argi"), add(ident("argi"), int_lit(1))),
            ]),
        ),
        var_decl("piece", str_lit("")),
        conv_dispatch,
        assign(ident("out"), add(ident("out"), ident("piece"))),
        assign(ident("i"), add(ident("j"), int_lit(1))),
    ];

    // %% and %n shortcuts, then the general specifier.
    let percent_body = vec![
        var_decl("c2", str_lit("")),
        if_stmt(
            binary(BinOp::Lt, add(ident("i"), int_lit(1)), ident("n")),
            vec![assign(
                ident("c2"),
                char_at(ident("fmt"), add(ident("i"), int_lit(1))),
            )],
            None,
        ),
        if_stmt(
            binary(BinOp::Eq, ident("c2"), str_lit("%")),
            vec![
                assign(ident("out"), add(ident("out"), str_lit("%"))),
                assign(ident("i"), add(ident("i"), int_lit(2))),
            ],
            Some(vec![if_stmt(
                binary(BinOp::Eq, ident("c2"), str_lit("n")),
                vec![
                    assign(ident("out"), add(ident("out"), str_lit("\n"))),
                    assign(ident("i"), add(ident("i"), int_lit(2))),
                ],
                Some(spec_parse),
            )]),
        ),
    ];

    function_stmt(
        "__j_sprintf",
        vec!["fmt", "args"],
        vec![
            assign(ident("fmt"), to_str(ident("fmt"))),
            var_decl("out", str_lit("")),
            var_decl("i", int_lit(0)),
            var_decl("n", member(ident("fmt"), "length")),
            var_decl("argi", int_lit(0)),
            while_stmt(
                binary(BinOp::Lt, ident("i"), ident("n")),
                vec![
                    var_decl("c", char_at(ident("fmt"), ident("i"))),
                    if_stmt(
                        binary(BinOp::NotEq, ident("c"), str_lit("%")),
                        vec![
                            assign(ident("out"), add(ident("out"), ident("c"))),
                            assign(ident("i"), add(ident("i"), int_lit(1))),
                        ],
                        Some(percent_body),
                    ),
                ],
            ),
            ret(ident("out")),
        ],
    )
}
