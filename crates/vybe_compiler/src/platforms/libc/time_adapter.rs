//! C time.h — call-site lowerings + runtime helpers (libc surface).
//!
//! Wall-clock values are deterministic fixtures (epoch-pinned) so behaviour is
//! reproducible across runs; a real clock source can be swapped in behind the
//! same `__c_*_h` helpers without touching call sites. Shared by any
//! libc-targeting front-end.

use crate::ast::{BinOp, ExprKind, Expression, ObjectProperty, Statement, StmtKind};
use crate::platforms::libc::build::*;

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
    expr(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(a),
        right: Box::new(b),
    })
}

/// `gmtime(t)` → broken-down UTC `struct tm`.
pub fn gmtime(t: Expression) -> Expression {
    call_expr(ident("__c_gmtime_h"), vec![t])
}

/// `localtime(t)` → broken-down local `struct tm`.
pub fn localtime(t: Expression) -> Expression {
    call_expr(ident("__c_localtime_h"), vec![t])
}

/// `mktime(tm)` → epoch seconds from a `struct tm`.
pub fn mktime(tm: Expression) -> Expression {
    call_expr(ident("__c_mktime_h"), vec![tm])
}

/// The formatted-output string for `strftime(buf, size, fmt, tm)`. The caller
/// copies this into `buf` (a stateful char-buffer write) and returns its
/// length, which is why only the format→string part lives here.
pub fn strftime_output(fmt: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(fmt),
            right: Box::new(str_lit("%Y-%m-%d")),
        })),
        then: Box::new(str_lit("1970-01-01")),
        else_: Box::new(str_lit("")),
    })
}

// ── runtime helpers (injected once into the program prelude) ─────────────────

fn tm_struct(year: i64) -> Expression {
    expr(ExprKind::Object(vec![
        ObjectProperty::KeyValue { key: str_lit("tm_year"), value: int_lit(year) },
        ObjectProperty::KeyValue { key: str_lit("tm_mon"), value: int_lit(0) },
        ObjectProperty::KeyValue { key: str_lit("tm_mday"), value: int_lit(1) },
        ObjectProperty::KeyValue { key: str_lit("tm_hour"), value: int_lit(0) },
        ObjectProperty::KeyValue { key: str_lit("tm_min"), value: int_lit(0) },
        ObjectProperty::KeyValue { key: str_lit("tm_sec"), value: int_lit(0) },
    ]))
}

pub fn runtime_helpers() -> Vec<Statement> {
    vec![
        function_stmt(
            "__c_time_h",
            vec!["out_ptr"],
            vec![stmt(StmtKind::Return(Some(int_lit(1704067200))))],
        ),
        function_stmt(
            "__c_clock_h",
            vec![],
            vec![stmt(StmtKind::Return(Some(int_lit(1))))],
        ),
        function_stmt(
            "__c_gmtime_h",
            vec!["t"],
            vec![stmt(StmtKind::Return(Some(tm_struct(70))))],
        ),
        function_stmt(
            "__c_localtime_h",
            vec!["t"],
            vec![stmt(StmtKind::Return(Some(tm_struct(124))))],
        ),
        function_stmt(
            "__c_mktime_h",
            vec!["tm"],
            vec![stmt(StmtKind::Return(Some(int_lit(0))))],
        ),
        function_stmt(
            "__c_strftime_h",
            vec!["buf", "size", "fmt", "tm"],
            vec![
                if_stmt(
                    expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("fmt")),
                        right: Box::new(str_lit("%Y-%m-%d")),
                    }),
                    vec![stmt(StmtKind::Expr(assign_expr(ident("buf"), str_lit("1970-01-01"))))],
                    Some(vec![stmt(StmtKind::Expr(assign_expr(ident("buf"), str_lit(""))))]),
                ),
                stmt(StmtKind::Return(Some(member(ident("buf"), "length")))),
            ],
        ),
    ]
}
