//! C/Go pointer model — shared by any language with true pointer semantics.
//!
//! Two kinds of pointer, both tracked at runtime via a `__ref_kind` tag:
//!
//! **Scalar pointer** (existing cell mechanism in `emitter/references.rs`):
//!   `{__ref_kind: "cell", __value: T}`
//!   Used for `&scalar_var`. The compiler's `compile_address_of_expr` /
//!   `promote_local_binding_to_pointer_cell` handle this — nothing new needed.
//!   Walker emits `ExprKind::RefOf(PlaceExpr::Ident(name))` and
//!   `ExprKind::RefLoad(expr)` as Go already does.
//!
//! **Array pointer** (new, this module):
//!   `{__ref_kind: "carray", __base: Array, __idx: i32}`
//!   Used for `int *p = arr`, `char *p = text`, `p + n`, `p++`, `p - q`.
//!   Stores the base array and a current index so arithmetic and write-through
//!   both work correctly without copying.
//!
//! Walker usage:
//!   `&arr`       → `make_carray_ptr(ident("arr"), lit_int(0))`
//!   `&arr[n]`    → `make_carray_ptr(ident("arr"), n)`
//!   `*p`         → `carray_deref_read(p)` when p is a known carray pointer
//!   `*p = val`   → `carray_deref_write(p, val)`
//!   `p + n`      → `carray_advance(p, n)`
//!   `p - q`      → `carray_diff(p, q)`
//!   NULL         → `lit_int(0)` / null (unchanged)

use crate::ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Expression, Literal, ObjectProperty,
    Statement, StmtKind, VarDeclKind, VarDeclarator,
};

pub const REF_KIND_KEY: &str = "__ref_kind";
pub const CARRAY_BASE_KEY: &str = "__base";
pub const CARRAY_IDX_KEY: &str = "__idx";
pub const CARRAY_KIND: &str = "carray";

// ── AST helpers ───────────────────────────────────────────────────────────────

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
}

fn member(object: Expression, field: &str) -> Expression {
    e(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn obj_prop(key: &str, value: Expression) -> ObjectProperty {
    ObjectProperty::KeyValue {
        key: e(ExprKind::Lit(Literal::Str(key.to_string()))),
        value,
    }
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

// ── Public API ────────────────────────────────────────────────────────────────

/// Create a C array pointer over `base` starting at element `idx`.
/// `{__ref_kind: "carray", __base: base, __idx: idx}`
pub fn make_carray_ptr(base: Expression, idx: Expression) -> Expression {
    e(ExprKind::Object(vec![
        obj_prop(
            REF_KIND_KEY,
            e(ExprKind::Lit(Literal::Str(CARRAY_KIND.to_string()))),
        ),
        obj_prop(CARRAY_BASE_KEY, base),
        obj_prop(CARRAY_IDX_KEY, idx),
    ]))
}

/// Read the element the pointer currently points to: `ptr.__base[ptr.__idx]`.
pub fn carray_deref_read(ptr: Expression) -> Expression {
    e(ExprKind::Index {
        object: Box::new(member(ptr.clone(), CARRAY_BASE_KEY)),
        index: Box::new(member(ptr, CARRAY_IDX_KEY)),
        null_safe: false,
    })
}

/// Write through the pointer: `ptr.__base[ptr.__idx] = val`.
pub fn carray_deref_write(ptr: Expression, val: Expression) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(e(ExprKind::Index {
            object: Box::new(member(ptr.clone(), CARRAY_BASE_KEY)),
            index: Box::new(member(ptr, CARRAY_IDX_KEY)),
            null_safe: false,
        })),
        value: Box::new(val),
    })
}

/// Advance a pointer by `n` elements: new ptr with `__idx + n`.
/// The base array is shared — no copy, so writes through the advanced
/// pointer still affect the original array.
pub fn carray_advance(ptr: Expression, n: Expression) -> Expression {
    let new_idx = e(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(member(ptr.clone(), CARRAY_IDX_KEY)),
        right: Box::new(n),
    });
    e(ExprKind::Object(vec![
        obj_prop(
            REF_KIND_KEY,
            e(ExprKind::Lit(Literal::Str(CARRAY_KIND.to_string()))),
        ),
        obj_prop(CARRAY_BASE_KEY, member(ptr, CARRAY_BASE_KEY)),
        obj_prop(CARRAY_IDX_KEY, new_idx),
    ]))
}

/// Element distance between two pointers into the same array: `a.__idx - b.__idx`.
pub fn carray_diff(a: Expression, b: Expression) -> Expression {
    e(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(member(a, CARRAY_IDX_KEY)),
        right: Box::new(member(b, CARRAY_IDX_KEY)),
    })
}

/// In-place advance: emits `ptr.__idx += n` as an assignment expression.
/// Use for `p++`, `p += n` — mutates the pointer variable directly.
pub fn carray_advance_inplace(ptr_name: &str, n: Expression) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(member(ident(ptr_name), CARRAY_IDX_KEY)),
        value: Box::new(e(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(member(ident(ptr_name), CARRAY_IDX_KEY)),
            right: Box::new(n),
        })),
    })
}

/// In-place retreat: emits `ptr.__idx -= n` as an assignment expression.
/// Use for `p--`, `p -= n`.
pub fn carray_retreat_inplace(ptr_name: &str, n: Expression) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(member(ident(ptr_name), CARRAY_IDX_KEY)),
        value: Box::new(e(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(member(ident(ptr_name), CARRAY_IDX_KEY)),
            right: Box::new(n),
        })),
    })
}

/// Check if a pointer is null (0 or null): for `if (!p)` etc.
pub fn is_carray_ptr_kind(ptr: Expression) -> Expression {
    e(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(member(ptr, REF_KIND_KEY)),
        right: Box::new(e(ExprKind::Lit(Literal::Str(CARRAY_KIND.to_string())))),
    })
}

/// Convert a char carray pointer to a JS string for I/O (puts, printf %s).
/// Iterates from `ptr.__idx` until a 0 byte, building a string via
/// `String.fromCharCode`.  Emitted as a call to a runtime helper expression.
/// The walker should emit this wrapping puts()/log() arguments that are char*.
pub fn carray_chars_to_string(ptr: Expression) -> Expression {
    let base = member(ptr.clone(), CARRAY_BASE_KEY);
    let idx = member(ptr, CARRAY_IDX_KEY);
    let slice = call(member(base, "slice"), vec![idx]);
    let string = e(ExprKind::Call {
        callee: Box::new(e(ExprKind::Member {
            object: Box::new(ident("String")),
            field: "fromCharCode".to_string(),
            null_safe: false,
        })),
        args: vec![Argument {
            value: slice,
            name: None,
            by_ref: false,
            spread: true,
        }],
        optional: false,
    });
    e(ExprKind::Index {
        object: Box::new(call(
            member(string, "split"),
            vec![e(ExprKind::Lit(Literal::Str("\0".to_string())))],
        )),
        index: Box::new(lit_int(0)),
        null_safe: false,
    })
}

/// True if the raw initializer text indicates a scalar address-of (`&x`).
/// Used by walkers to decide between scalar cell (RefOf) and carray pointer.
pub fn init_is_addr_of(init_source_text: &str) -> bool {
    let t = init_source_text.trim();
    t.starts_with('&') && !t.starts_with("&&")
}
