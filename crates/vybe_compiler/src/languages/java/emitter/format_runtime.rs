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
    Argument, BinOp, BindingPattern, ExprKind, Expression, Literal, Modifiers, Param, PassBy,
    Statement, StmtKind, VarDeclKind, VarDeclarator,
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

fn null_lit() -> Expression {
    expr(ExprKind::Lit(Literal::Null))
}

fn member(object: Expression, field: &str) -> Expression {
    expr(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
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

/// Java string conversion (null → "null", booleans → "true"/"false")
/// via the shared engine's `%s` — a bare-ident builtin call, which
/// dispatches anywhere (member-shaped `String.valueOf` does not inside
/// injected prelude functions). NOT `"" + x`: the dynamic add coerces
/// Bool→1 and Null→0.
fn to_str(x: Expression) -> Expression {
    call("__java_string_format", vec![str_lit("%s"), x])
}

fn assign(target: Expression, value: Expression) -> Statement {
    stmt(StmtKind::Assign {
        targets: vec![target],
        value,
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
    let mut out = Vec::new();

    // PrintStream identity sentinel + pending-line buffer.
    out.push(var_decl(OUT_SENTINEL, str_lit(OUT_SENTINEL)));
    out.push(var_decl("__j_buf", str_lit("")));

    // __j_print(x): buffer, flush completed lines one println each.
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
                        "println",
                        vec![call_member(
                            ident("__j_buf"),
                            "substring",
                            vec![int_lit(0), ident("i")],
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
            var_decl("w", call_member(ident("Integer"), "parseInt", vec![ident("width")])),
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
                binary(
                    BinOp::Eq,
                    char_at(ident("s"), int_lit(0)),
                    str_lit("-"),
                ),
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
                        vec![assign(ident("grouped"), add(ident("grouped"), str_lit(",")))],
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
                vec![assign(ident("grouped"), add(str_lit("-"), ident("grouped")))],
                None,
            ),
            ret(ident("grouped")),
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

    out.append(&mut stringbuilder_fns());
    out.append(&mut regex_fns());
    out.append(&mut url_fns());
    out.push(sprintf_fn());
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
                add(add(to_str(ident("proto")), str_lit("://")), to_str(ident("host"))),
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
    // equals / hashCode / sameFile — java compares the URL text.
    out.push(function_stmt(
        "__j_url_equals",
        vec!["a", "b"],
        vec![
            if_stmt(
                binary(BinOp::Eq, member(ident("a"), "href"), member(ident("b"), "href")),
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
                        binary(BinOp::Eq, member(ident("a"), "host"), member(ident("b"), "host")),
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
        "__j_url_to_string",
        vec!["u"],
        vec![ret(fld("u", "href"))],
    ));
    out
}

/// `java.util.regex` Pattern/Matcher over `ecma:regexp` (patterns are
/// plain strings; `__j_re_exec` returns the ECMA match array with
/// `.index`, or null). The Matcher carries Java's find() cursor.
fn regex_fns() -> Vec<Statement> {
    let obj = || expr(ExprKind::Object(Vec::new()));
    let fld = |name: &str, f: &str| member(ident(name), f);
    let substr_range = |s: Expression, a: Expression, b: Expression| {
        call_expr(member(s, "substring"), vec![a, b])
    };
    let mut out = Vec::new();

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
        vec![ret(fld("p", "__re"))],
    ));
    // Java split semantics (JLS Pattern.split): limit n>0 = at most n
    // parts with the remainder attached to the last; n==0 = unlimited,
    // trailing empty strings removed; n<0 = unlimited, empties kept.
    out.push(function_stmt(
        "__j_pat_split_impl",
        vec!["p", "s", "n"],
        vec![
            var_decl("parts", expr(ExprKind::Array(Vec::new()))),
            var_decl("count", int_lit(0)),
            var_decl("pos", int_lit(0)),
            var_decl("go", int_lit(1)),
            while_stmt(
                binary(BinOp::Eq, ident("go"), int_lit(1)),
                vec![
                    if_stmt(
                        binary(
                            BinOp::And,
                            binary(BinOp::Gt, ident("n"), int_lit(0)),
                            binary(
                                BinOp::GtEq,
                                ident("count"),
                                binary(BinOp::Sub, ident("n"), int_lit(1)),
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
                                        call_expr(
                                            member(ident("s"), "substring"),
                                            vec![ident("pos")],
                                        ),
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
                                    if_stmt(
                                        binary(BinOp::Eq, ident("mlen"), int_lit(0)),
                                        vec![assign(ident("go"), int_lit(0))],
                                        Some(vec![
                                            assign(
                                                index_expr(ident("parts"), ident("count")),
                                                substr_range(
                                                    ident("s"),
                                                    ident("pos"),
                                                    add(ident("pos"), member(ident("r"), "index")),
                                                ),
                                            ),
                                            assign(ident("count"), add(ident("count"), int_lit(1))),
                                            assign(
                                                ident("pos"),
                                                add(
                                                    add(ident("pos"), member(ident("r"), "index")),
                                                    ident("mlen"),
                                                ),
                                            ),
                                        ]),
                                    ),
                                ]),
                            ),
                        ]),
                    ),
                ],
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
                        vec![assign(ident("last"), binary(BinOp::Sub, ident("last"), int_lit(1)))],
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
            assign(fld("m", "__input"), to_str(ident("s"))),
            assign(fld("m", "__pos"), int_lit(0)),
            assign(fld("m", "__m"), null_lit()),
            assign(fld("m", "__start"), int_lit(-1)),
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
                vec![assign(fld("m", "__m"), null_lit()), ret(expr(ExprKind::Lit(Literal::Bool(false))))],
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
            var_decl(
                "adv",
                member(index_expr(ident("r"), int_lit(0)), "length"),
            ),
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
        "__j_m_replace_all",
        vec!["m", "repl"],
        vec![ret(call(
            "__j_re_replace_all",
            vec![fld("m", "__input"), fld("m", "__re"), ident("repl")],
        ))],
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
    let substr2 = |s: Expression, a: Expression, b: Expression| {
        call_expr(member(s, "substring"), vec![a, b])
    };
    let substr1 = |s: Expression, a: Expression| call_expr(member(s, "substring"), vec![a]);
    let mut out = Vec::new();

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
        "__j_sb_char_at",
        vec!["sb", "i"],
        vec![ret(substr2(buf("sb"), ident("i"), add(ident("i"), int_lit(1))))],
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
            var_decl("i", binary(BinOp::Sub, member(buf("sb"), "length"), int_lit(1))),
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
                            add(str_lit("%."), to_str(binary(
                                BinOp::Sub,
                                binary(BinOp::Sub, ident("pr"), int_lit(1)),
                                ident("exv"),
                            ))),
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
                                add(str_lit("%."), to_str(binary(BinOp::Sub, ident("pr"), int_lit(1)))),
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
            call("__j_padw", vec![ident("piece"), ident("width"), ident("left")]),
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
                call("__j_padw", vec![ident("piece"), ident("width"), ident("left")]),
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
                    call("__j_padw", vec![ident("piece"), ident("width"), ident("left")]),
                ),
            ],
            Some(vec![if_stmt(
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, ident("conv"), str_lit("e")),
                    binary(BinOp::Eq, ident("conv"), str_lit("E")),
                ),
                vec![assign(ident("piece"), call("__j_expad", vec![delegate.clone()]))],
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
                assign(ident("digs"), add(ident("digs"), char_at(ident("fmt"), ident("k")))),
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
                assign(ident("width"), add(ident("width"), char_at(ident("fmt"), ident("j")))),
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
                        assign(ident("prec"), add(ident("prec"), char_at(ident("fmt"), ident("j")))),
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
                index_expr(ident("args"), binary(BinOp::Sub, ident("argidx"), int_lit(1))),
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
            vec![assign(ident("c2"), char_at(ident("fmt"), add(ident("i"), int_lit(1))))],
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
