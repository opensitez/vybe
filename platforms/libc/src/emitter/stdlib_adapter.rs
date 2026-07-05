//! C stdlib.h — memory allocation and conversion adapters.

use vybe_ast::{Argument, ArrayElement, BinOp, ExprKind, Expression, Literal};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

fn lit_float(f: f64) -> Expression {
    e(ExprKind::Lit(Literal::Float(f)))
}

fn array_elem(value: Expression) -> ArrayElement {
    ArrayElement {
        value,
        spread: false,
        key: None,
        by_ref: false,
    }
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
}

// ── Heap allocation ───────────────────────────────────────────────────────────

/// `malloc(n)` → empty GC-managed array (VM handles lifetime).
pub fn c_malloc() -> Expression {
    e(ExprKind::Array(Vec::new()))
}

/// `calloc(count, _size)` → flat zero-filled array of `count` elements.
/// The element size is ignored — the VM array is element-indexed, not byte-indexed.
pub fn c_calloc(count: usize) -> Expression {
    let zeros = (0..count).map(|_| array_elem(lit_int(0))).collect();
    e(ExprKind::Array(zeros))
}

/// `free(p)` → null (GC handles deallocation; noop at expression level).
pub fn c_free() -> Expression {
    e(ExprKind::Lit(Literal::Null))
}

// ── String → number conversions ───────────────────────────────────────────────

/// `atoi(s)` / `atol(s)` — parse leading decimal digits, return 0 for non-numeric.
/// `parseInt(s, 10) || 0` — stops at first non-digit, NaN → 0.
pub fn c_atoi(s: Expression) -> Expression {
    let parse = call(ident("parseInt"), vec![s, lit_int(10)]);
    e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(parse),
        right: Box::new(lit_int(0)),
    })
}

/// `atof(s)` — parse floating-point, return 0.0 for empty/non-numeric.
/// `parseFloat(s) || 0.0`
pub fn c_atof(s: Expression) -> Expression {
    let parse = call(ident("parseFloat"), vec![s]);
    e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(parse),
        right: Box::new(lit_float(0.0)),
    })
}
