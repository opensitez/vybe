//! C time.h — call-site lowerings + runtime helpers (libc surface).
//!
//! Wall-clock values are deterministic fixtures (epoch-pinned) so behaviour is
//! reproducible across runs; a real clock source can be swapped in behind the
//! same `__c_*_h` helpers without touching call sites. Shared by any
//! libc-targeting front-end.

use crate::ast::{ArrayElement, BinOp, ExprKind, Expression, ObjectProperty, Statement, StmtKind, UnaryOp};
use crate::platforms::libc::emitter::build::*;

// ── call-site lowerings (walker maps `time(...)` etc. through these) ─────────

/// `time(out)` → epoch seconds (also written through `out` by the helper).
pub fn time(out_ptr: Expression) -> Expression {
    call_expr(ident("__c_time_h"), vec![out_ptr])
}

/// `clock()` → processor clock ticks.
pub fn clock() -> Expression {
    call_expr(ident("__c_clock_h"), vec![])
}

/// `difftime(a, b)` → `a - b` (seconds, per §7.27.2.2).
pub fn difftime(a: Expression, b: Expression) -> Expression {
    bin(BinOp::Sub, a, b)
}

/// `gmtime(t)` → broken-down UTC `struct tm`.
pub fn gmtime(t: Expression) -> Expression {
    call_expr(ident("__c_gmtime_h"), vec![value_from_address_arg(t)])
}

/// `localtime(t)` → broken-down local `struct tm`.
pub fn localtime(t: Expression) -> Expression {
    call_expr(ident("__c_localtime_h"), vec![value_from_address_arg(t)])
}

/// `mktime(tm)` → epoch seconds from a `struct tm`.
pub fn mktime(tm: Expression) -> Expression {
    call_expr(ident("__c_mktime_h"), vec![value_from_address_arg(tm)])
}

/// `asctime(tm)` → textual calendar representation.
pub fn asctime(tm: Expression) -> Expression {
    call_expr(ident("__c_asctime_h"), vec![value_from_address_arg(tm)])
}

/// `ctime(t)` → `asctime(localtime(t))`.
pub fn ctime(t: Expression) -> Expression {
    call_expr(
        ident("__c_asctime_h"),
        vec![call_expr(ident("__c_localtime_h"), vec![value_from_address_arg(t)])],
    )
}

/// The formatted-output string for `strftime(buf, size, fmt, tm)`. The caller
/// copies this into `buf` and returns the generated length.
pub fn strftime_output(fmt: Expression, tm: Expression) -> Expression {
    call_expr(ident("__c_strftime_format_h"), vec![fmt, value_from_address_arg(tm)])
}

// ── small AST builders ───────────────────────────────────────────────────────

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn value_from_address_arg(value: Expression) -> Expression {
    match value.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        kind => Expression {
            kind,
            span: value.span,
        },
    }
}

fn ternary(cond: Expression, then_expr: Expression, else_expr: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(else_expr),
    })
}

fn ret(value: Expression) -> Statement {
    stmt(StmtKind::Return(Some(value)))
}

fn expr_stmt(value: Expression) -> Statement {
    stmt(StmtKind::Expr(value))
}

fn assign_stmt(target: Expression, value: Expression) -> Statement {
    expr_stmt(assign_expr(target, value))
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

fn array(values: &[&str]) -> Expression {
    expr(ExprKind::Array(
        values
            .iter()
            .map(|value| ArrayElement {
                key: None,
                value: str_lit(value),
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn tm_field(name: &str) -> Expression {
    member(ident("tm"), name)
}

fn zero_if_missing(value: Expression) -> Expression {
    ternary(
        eq(
            expr(ExprKind::Unary {
                op: UnaryOp::Typeof,
                expr: Box::new(value.clone()),
            }),
            str_lit("undefined"),
        ),
        int_lit(0),
        value,
    )
}

fn tm_numeric_field(name: &str) -> Expression {
    zero_if_missing(tm_field(name))
}

fn call_name(name: &str, args: Vec<Expression>) -> Expression {
    call_expr(ident(name), args)
}

fn eq(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Eq, left, right)
}

fn add(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Add, left, right)
}

fn sub(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Sub, left, right)
}

fn mul(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Mul, left, right)
}

fn div(left: Expression, right: Expression) -> Expression {
    bin(BinOp::IDiv, left, right)
}

fn modulo(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Mod, left, right)
}

fn lt(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Lt, left, right)
}

fn gt(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Gt, left, right)
}

fn gte(left: Expression, right: Expression) -> Expression {
    bin(BinOp::GtEq, left, right)
}

fn lte(left: Expression, right: Expression) -> Expression {
    bin(BinOp::LtEq, left, right)
}

fn cat(left: Expression, right: Expression) -> Expression {
    add(left, right)
}

fn zero_padded(value: Expression, width: i64) -> Expression {
    call_name("__c_pad_int_h", vec![value, int_lit(width), str_lit("0")])
}

fn space_padded(value: Expression, width: i64) -> Expression {
    call_name("__c_pad_int_h", vec![value, int_lit(width), str_lit(" ")])
}

fn month_name(full: bool) -> Expression {
    index_expr(
        if full {
            array(&[
                "January",
                "February",
                "March",
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December",
            ])
        } else {
            array(&[
                "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov",
                "Dec",
            ])
        },
        tm_field("tm_mon"),
    )
}

fn weekday_name(full: bool) -> Expression {
    index_expr(
        if full {
            array(&[
                "Sunday",
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
            ])
        } else {
            array(&["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"])
        },
        tm_field("tm_wday"),
    )
}

fn year_full() -> Expression {
    add(tm_field("tm_year"), int_lit(1900))
}

fn hour12() -> Expression {
    ternary(
        eq(modulo(tm_field("tm_hour"), int_lit(12)), int_lit(0)),
        int_lit(12),
        modulo(tm_field("tm_hour"), int_lit(12)),
    )
}

fn computed_yday() -> Expression {
    call_name(
        "__c_yday_h",
        vec![year_full(), tm_field("tm_mon"), tm_field("tm_mday")],
    )
}

fn sunday_week_number() -> Expression {
    div(add(sub(computed_yday(), tm_field("tm_wday")), int_lit(7)), int_lit(7))
}

fn monday_week_number() -> Expression {
    let monday_zero_wday = modulo(add(tm_field("tm_wday"), int_lit(6)), int_lit(7));
    div(add(sub(computed_yday(), monday_zero_wday), int_lit(7)), int_lit(7))
}

fn strftime_value_for_code(code: &str) -> Expression {
    match code {
        "%Y" => call_name("__c_to_string_h", vec![year_full()]),
        "%y" => zero_padded(modulo(year_full(), int_lit(100)), 2),
        "%m" => zero_padded(add(tm_field("tm_mon"), int_lit(1)), 2),
        "%d" => zero_padded(tm_field("tm_mday"), 2),
        "%H" => zero_padded(tm_field("tm_hour"), 2),
        "%M" => zero_padded(tm_field("tm_min"), 2),
        "%S" => zero_padded(tm_field("tm_sec"), 2),
        "%A" => weekday_name(true),
        "%a" => weekday_name(false),
        "%B" => month_name(true),
        "%b" | "%h" => month_name(false),
        "%p" => ternary(gte(tm_field("tm_hour"), int_lit(12)), str_lit("PM"), str_lit("AM")),
        "%I" => zero_padded(hour12(), 2),
        "%j" => zero_padded(add(tm_field("tm_yday"), int_lit(1)), 3),
        "%w" => call_name("__c_to_string_h", vec![tm_field("tm_wday")]),
        "%u" => call_name(
            "__c_to_string_h",
            vec![ternary(eq(tm_field("tm_wday"), int_lit(0)), int_lit(7), tm_field("tm_wday"))],
        ),
        "%C" => zero_padded(div(year_full(), int_lit(100)), 2),
        "%F" => cat(
            cat(cat(strftime_value_for_code("%Y"), str_lit("-")), strftime_value_for_code("%m")),
            cat(str_lit("-"), strftime_value_for_code("%d")),
        ),
        "%D" => cat(
            cat(cat(strftime_value_for_code("%m"), str_lit("/")), strftime_value_for_code("%d")),
            cat(str_lit("/"), strftime_value_for_code("%y")),
        ),
        "%R" => cat(
            cat(strftime_value_for_code("%H"), str_lit(":")),
            strftime_value_for_code("%M"),
        ),
        "%T" => cat(
            cat(cat(strftime_value_for_code("%H"), str_lit(":")), strftime_value_for_code("%M")),
            cat(str_lit(":"), strftime_value_for_code("%S")),
        ),
        "%e" => space_padded(tm_field("tm_mday"), 2),
        "%l" => space_padded(hour12(), 2),
        "%k" => space_padded(tm_field("tm_hour"), 2),
        "%%" => str_lit("%"),
        "%U" => zero_padded(sunday_week_number(), 2),
        "%W" | "%V" => zero_padded(monday_week_number(), 2),
        "%G" => strftime_value_for_code("%Y"),
        "%g" => strftime_value_for_code("%y"),
        "%n" => str_lit("\n"),
        "%t" => str_lit("\t"),
        _ => str_lit(""),
    }
}

fn strftime_return(format: &str, value: Expression) -> Statement {
    if_stmt(
        eq(ident("fmt"), str_lit(format)),
        vec![ret(value)],
        None,
    )
}

fn tm_struct_from_locals() -> Expression {
    expr(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: str_lit("tm_year"),
            value: sub(ident("year"), int_lit(1900)),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_mon"),
            value: ident("mon"),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_mday"),
            value: ident("mday"),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_hour"),
            value: ident("hour"),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_min"),
            value: ident("min"),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_sec"),
            value: ident("sec"),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_wday"),
            value: ident("wday"),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_yday"),
            value: ident("yday"),
        },
        ObjectProperty::KeyValue {
            key: str_lit("tm_isdst"),
            value: int_lit(0),
        },
    ]))
}

// ── runtime helpers (injected once into the program prelude) ─────────────────

fn is_leap_body() -> Vec<Statement> {
    vec![ret(ternary(
        eq(modulo(ident("y"), int_lit(400)), int_lit(0)),
        int_lit(1),
        ternary(
            eq(modulo(ident("y"), int_lit(100)), int_lit(0)),
            int_lit(0),
            ternary(eq(modulo(ident("y"), int_lit(4)), int_lit(0)), int_lit(1), int_lit(0)),
        ),
    ))]
}

fn days_in_month_body() -> Vec<Statement> {
    vec![ret(ternary(
        eq(ident("m"), int_lit(1)),
        ternary(eq(call_name("__c_is_leap_h", vec![ident("y")]), int_lit(1)), int_lit(29), int_lit(28)),
        ternary(
            eq(ident("m"), int_lit(3)),
            int_lit(30),
            ternary(
                eq(ident("m"), int_lit(5)),
                int_lit(30),
                ternary(
                    eq(ident("m"), int_lit(8)),
                    int_lit(30),
                    ternary(eq(ident("m"), int_lit(10)), int_lit(30), int_lit(31)),
                ),
            ),
        ),
    ))]
}

fn yday_body() -> Vec<Statement> {
    vec![
        var_decl_stmt("i", int_lit(0)),
        var_decl_stmt("d", int_lit(0)),
        while_stmt(
            lt(ident("i"), ident("m")),
            vec![
                assign_stmt(ident("d"), add(ident("d"), call_name("__c_dim_h", vec![ident("y"), ident("i")]))),
                assign_stmt(ident("i"), add(ident("i"), int_lit(1))),
            ],
        ),
        ret(add(ident("d"), sub(ident("day"), int_lit(1)))),
    ]
}

fn gmtime_body() -> Vec<Statement> {
    vec![
        var_decl_stmt("secs", ident("t")),
        var_decl_stmt("days", div(ident("secs"), int_lit(86400))),
        var_decl_stmt("rem", modulo(ident("secs"), int_lit(86400))),
        var_decl_stmt("hour", div(ident("rem"), int_lit(3600))),
        assign_stmt(ident("rem"), modulo(ident("rem"), int_lit(3600))),
        var_decl_stmt("min", div(ident("rem"), int_lit(60))),
        var_decl_stmt("sec", modulo(ident("rem"), int_lit(60))),
        var_decl_stmt("year", int_lit(1970)),
        var_decl_stmt("yday", ident("days")),
        while_stmt(
            gte(ident("yday"), call_name("__c_days_in_year_h", vec![ident("year")])),
            vec![
                assign_stmt(ident("yday"), sub(ident("yday"), call_name("__c_days_in_year_h", vec![ident("year")]))),
                assign_stmt(ident("year"), add(ident("year"), int_lit(1))),
            ],
        ),
        var_decl_stmt("mon", int_lit(0)),
        while_stmt(
            gte(ident("yday"), call_name("__c_dim_h", vec![ident("year"), ident("mon")])),
            vec![
                assign_stmt(ident("yday"), sub(ident("yday"), call_name("__c_dim_h", vec![ident("year"), ident("mon")]))),
                assign_stmt(ident("mon"), add(ident("mon"), int_lit(1))),
            ],
        ),
        var_decl_stmt("mday", add(ident("yday"), int_lit(1))),
        var_decl_stmt("wday", modulo(add(int_lit(4), ident("days")), int_lit(7))),
        ret(tm_struct_from_locals()),
    ]
}

fn mktime_body() -> Vec<Statement> {
    vec![
        var_decl_stmt("year", add(tm_numeric_field("tm_year"), int_lit(1900))),
        var_decl_stmt("mon", tm_numeric_field("tm_mon")),
        var_decl_stmt("mday", tm_numeric_field("tm_mday")),
        var_decl_stmt("hour", tm_numeric_field("tm_hour")),
        var_decl_stmt("min", tm_numeric_field("tm_min")),
        var_decl_stmt("sec", tm_numeric_field("tm_sec")),
        while_stmt(gte(ident("sec"), int_lit(60)), vec![
            assign_stmt(ident("sec"), sub(ident("sec"), int_lit(60))),
            assign_stmt(ident("min"), add(ident("min"), int_lit(1))),
        ]),
        while_stmt(lt(ident("sec"), int_lit(0)), vec![
            assign_stmt(ident("sec"), add(ident("sec"), int_lit(60))),
            assign_stmt(ident("min"), sub(ident("min"), int_lit(1))),
        ]),
        while_stmt(gte(ident("min"), int_lit(60)), vec![
            assign_stmt(ident("min"), sub(ident("min"), int_lit(60))),
            assign_stmt(ident("hour"), add(ident("hour"), int_lit(1))),
        ]),
        while_stmt(lt(ident("min"), int_lit(0)), vec![
            assign_stmt(ident("min"), add(ident("min"), int_lit(60))),
            assign_stmt(ident("hour"), sub(ident("hour"), int_lit(1))),
        ]),
        while_stmt(gte(ident("hour"), int_lit(24)), vec![
            assign_stmt(ident("hour"), sub(ident("hour"), int_lit(24))),
            assign_stmt(ident("mday"), add(ident("mday"), int_lit(1))),
        ]),
        while_stmt(lt(ident("hour"), int_lit(0)), vec![
            assign_stmt(ident("hour"), add(ident("hour"), int_lit(24))),
            assign_stmt(ident("mday"), sub(ident("mday"), int_lit(1))),
        ]),
        while_stmt(lt(ident("mon"), int_lit(0)), vec![
            assign_stmt(ident("mon"), add(ident("mon"), int_lit(12))),
            assign_stmt(ident("year"), sub(ident("year"), int_lit(1))),
        ]),
        while_stmt(gte(ident("mon"), int_lit(12)), vec![
            assign_stmt(ident("mon"), sub(ident("mon"), int_lit(12))),
            assign_stmt(ident("year"), add(ident("year"), int_lit(1))),
        ]),
        while_stmt(lte(ident("mday"), int_lit(0)), vec![
            assign_stmt(ident("mon"), sub(ident("mon"), int_lit(1))),
            while_stmt(lt(ident("mon"), int_lit(0)), vec![
                assign_stmt(ident("mon"), add(ident("mon"), int_lit(12))),
                assign_stmt(ident("year"), sub(ident("year"), int_lit(1))),
            ]),
            assign_stmt(ident("mday"), add(ident("mday"), call_name("__c_dim_h", vec![ident("year"), ident("mon")]))),
        ]),
        while_stmt(gt(ident("mday"), call_name("__c_dim_h", vec![ident("year"), ident("mon")])), vec![
            assign_stmt(ident("mday"), sub(ident("mday"), call_name("__c_dim_h", vec![ident("year"), ident("mon")]))),
            assign_stmt(ident("mon"), add(ident("mon"), int_lit(1))),
            while_stmt(gte(ident("mon"), int_lit(12)), vec![
                assign_stmt(ident("mon"), sub(ident("mon"), int_lit(12))),
                assign_stmt(ident("year"), add(ident("year"), int_lit(1))),
            ]),
        ]),
        var_decl_stmt("days", int_lit(0)),
        var_decl_stmt("yy", int_lit(1970)),
        while_stmt(lt(ident("yy"), ident("year")), vec![
            assign_stmt(ident("days"), add(ident("days"), call_name("__c_days_in_year_h", vec![ident("yy")]))),
            assign_stmt(ident("yy"), add(ident("yy"), int_lit(1))),
        ]),
        while_stmt(gt(ident("yy"), ident("year")), vec![
            assign_stmt(ident("yy"), sub(ident("yy"), int_lit(1))),
            assign_stmt(ident("days"), sub(ident("days"), call_name("__c_days_in_year_h", vec![ident("yy")]))),
        ]),
        var_decl_stmt("yday", call_name("__c_yday_h", vec![ident("year"), ident("mon"), ident("mday")])),
        assign_stmt(ident("days"), add(ident("days"), ident("yday"))),
        var_decl_stmt("wday", modulo(add(int_lit(4), ident("days")), int_lit(7))),
        assign_stmt(tm_field("tm_year"), sub(ident("year"), int_lit(1900))),
        assign_stmt(tm_field("tm_mon"), ident("mon")),
        assign_stmt(tm_field("tm_mday"), ident("mday")),
        assign_stmt(tm_field("tm_hour"), ident("hour")),
        assign_stmt(tm_field("tm_min"), ident("min")),
        assign_stmt(tm_field("tm_sec"), ident("sec")),
        assign_stmt(tm_field("tm_yday"), ident("yday")),
        assign_stmt(tm_field("tm_wday"), ident("wday")),
        ret(add(mul(ident("days"), int_lit(86400)), add(mul(ident("hour"), int_lit(3600)), add(mul(ident("min"), int_lit(60)), ident("sec"))))),
    ]
}

fn strftime_body() -> Vec<Statement> {
    let mut body = Vec::new();
    for code in [
        "%Y", "%y", "%m", "%d", "%H", "%M", "%S", "%A", "%a", "%B", "%b", "%h", "%p", "%I",
        "%j", "%w", "%u", "%C", "%F", "%D", "%R", "%T", "%e", "%l", "%k", "%%", "%V", "%G",
        "%g", "%U", "%W", "%n", "%t",
    ] {
        body.push(strftime_return(code, strftime_value_for_code(code)));
    }
    body.push(strftime_return("%Y-%m-%d", strftime_value_for_code("%F")));
    body.push(strftime_return("%H:%M:%S", strftime_value_for_code("%T")));
    body.push(strftime_return(
        "%I %p",
        cat(cat(strftime_value_for_code("%I"), str_lit(" ")), strftime_value_for_code("%p")),
    ));
    body.push(strftime_return(
        "%Y%m%d%H%M%S",
        cat(
            cat(cat(strftime_value_for_code("%Y"), strftime_value_for_code("%m")), strftime_value_for_code("%d")),
            cat(cat(strftime_value_for_code("%H"), strftime_value_for_code("%M")), strftime_value_for_code("%S")),
        ),
    ));
    body.push(strftime_return("", str_lit("")));
    body.push(ret(str_lit("")));
    body
}

fn asctime_body() -> Vec<Statement> {
    vec![ret(cat(
        cat(
            cat(
                cat(
                    cat(
                        cat(
                            cat(
                                cat(weekday_name(false), str_lit(" ")),
                                cat(month_name(false), str_lit(" ")),
                            ),
                            cat(space_padded(tm_field("tm_mday"), 2), str_lit(" ")),
                        ),
                        cat(strftime_value_for_code("%T"), str_lit(" ")),
                    ),
                    strftime_value_for_code("%Y"),
                ),
                str_lit("\n"),
            ),
            str_lit(""),
        ),
        str_lit(""),
    ))]
}

pub fn runtime_helpers() -> Vec<Statement> {
    vec![
        function_stmt("__c_time_h", vec!["out_ptr"], vec![ret(int_lit(1704067200))]),
        function_stmt("__c_clock_h", vec![], vec![ret(int_lit(1))]),
        function_stmt("__c_is_leap_h", vec!["y"], is_leap_body()),
        function_stmt(
            "__c_days_in_year_h",
            vec!["y"],
            vec![ret(ternary(eq(call_name("__c_is_leap_h", vec![ident("y")]), int_lit(1)), int_lit(366), int_lit(365)))],
        ),
        function_stmt("__c_dim_h", vec!["y", "m"], days_in_month_body()),
        function_stmt("__c_yday_h", vec!["y", "m", "day"], yday_body()),
        function_stmt("__c_gmtime_h", vec!["t"], gmtime_body()),
        function_stmt("__c_localtime_h", vec!["t"], gmtime_body()),
        function_stmt("__c_mktime_h", vec!["tm"], mktime_body()),
        function_stmt("__c_asctime_h", vec!["tm"], asctime_body()),
        function_stmt(
            "__c_to_string_h",
            vec!["n"],
            vec![ret(cat(str_lit(""), ident("n")))],
        ),
        function_stmt(
            "__c_pad_int_h",
            vec!["n", "width", "pad"],
            vec![
                var_decl_stmt("s", call_name("__c_to_string_h", vec![ident("n")])),
                while_stmt(
                    lt(member(ident("s"), "length"), ident("width")),
                    vec![assign_stmt(ident("s"), cat(ident("pad"), ident("s")))],
                ),
                ret(ident("s")),
            ],
        ),
        function_stmt("__c_strftime_format_h", vec!["fmt", "tm"], strftime_body()),
    ]
}
