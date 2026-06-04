//! C stdio.h — I/O normalisation adapters.
//!
//! printf/fprintf → puts(sprintf(...))  (sprintf itself built by the user's
//! emitter/sprintf.rs — not reimplemented here).
//! puts → wasi:cli:log via the "print" profile emit.

use crate::ast::{Argument, ExprKind, Expression};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

/// `printf(fmt, args...)` → `puts(sprintf(fmt, args...))`.
/// The caller strips the stream argument first for `fprintf`.
pub fn printf_to_puts(fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let sprintf_call = call(ident("sprintf"), sprintf_args);
    call(ident("puts"), vec![sprintf_call])
}

/// `fprintf(stream, fmt, args...)` → `puts(sprintf(fmt, args...))`.
/// Stream is dropped; output goes to the WASI log.
pub fn fprintf_to_puts(fmt: Expression, rest: Vec<Expression>) -> Expression {
    printf_to_puts(fmt, rest)
}

/// `sprintf(buf, fmt, args...)` → `buf = sprintf(fmt, args...)`.
/// The buffer target is returned as the assign target; the RHS is the call.
pub fn sprintf_assign(buf: Expression, fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let rhs = call(ident("sprintf"), sprintf_args);
    e(ExprKind::Assign {
        target: Box::new(buf),
        value: Box::new(rhs),
    })
}
