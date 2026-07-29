//! Channel AST builders — shared normalization, not emission.
//!
//! A language's walker normalizes its channel syntax into these common AST
//! shapes; the compiler emits them (`vybe_compiler::primitives::channels`).
//! They construct `Expression`/`ExprKind` only — no chunks, no opcodes — so
//! they belong with the AST, where every frontend can reach them without
//! depending on a backend.
//!
//! This module exists because multiple languages were reimplementing the same
//! channel lowering. Centralizing it was the point; putting it in the AST is
//! where it belongs.

use crate::{Argument, ExprKind, Expression, ObjectProperty, UnaryOp};

fn channel_runtime_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

pub fn channel_new_expr(capacity: Option<Expression>) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("queue"),
            value: Expression::new(ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr: Box::new(Expression::new(ExprKind::Array(Vec::new()))),
            }),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("closed"),
            value: Expression::bool(false),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("capacity"),
            value: capacity.unwrap_or_else(|| Expression::int(0)),
        },
    ]))
}

pub fn channel_receive_expr(channel: Expression) -> Expression {
    channel_runtime_call("__vybe_channel_receive", vec![channel])
}

pub fn channel_send_expr(channel: Expression, value: Expression) -> Expression {
    channel_runtime_call("__vybe_channel_send", vec![channel, value])
}

pub fn channel_len_expr(channel: Expression) -> Expression {
    channel_runtime_call("__vybe_channel_len", vec![channel])
}

pub fn channel_cap_expr(channel: Expression) -> Expression {
    channel_runtime_call("__vybe_channel_cap", vec![channel])
}

pub fn channel_close_expr(channel: Expression) -> Expression {
    channel_runtime_call("__vybe_channel_close", vec![channel])
}
