//! Shared pointer model for languages with true pointer semantics.
//!
//! Two kinds of pointer, both tracked at runtime via a `__ref_kind` tag:
//!
//! **Scalar pointer** (existing cell mechanism in `primitives/references.rs`):
//!   `{__ref_kind: "cell", "__value": T}`
//!   Used for `&scalar_var`. The compiler's `compile_address_of_expr` /
//!   `promote_local_binding_to_pointer_cell` handle this.
//!
//! **Array pointer**:
//!   `{__ref_kind: "carray", "__base": Array, "__idx": i32}`
//!   Used for C-style decayed arrays and pointer arithmetic such as
//!   `int *p = arr`, `char *p = text`, `p + n`, `p++`, `p - q`.
//!   Stores the base array and a current index so arithmetic and write-through
//!   both work correctly without copying.

use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal, ObjectProperty};

pub const REF_KIND_KEY: &str = "__ref_kind";
pub const CARRAY_BASE_KEY: &str = "__base";
pub const CARRAY_IDX_KEY: &str = "__idx";
pub const CARRAY_KIND: &str = "carray";

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
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

fn is_zero(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Lit(Literal::Int(0)))
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    e(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

/// Create a C array pointer over `base` starting at element `idx`.
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

/// Read `ptr[n]`: `ptr.__base[ptr.__idx + n]`.
pub fn carray_indexed_read(ptr: Expression, n: Expression) -> Expression {
    e(ExprKind::Index {
        object: Box::new(member(ptr.clone(), CARRAY_BASE_KEY)),
        index: Box::new(binary(BinOp::Add, member(ptr, CARRAY_IDX_KEY), n)),
        null_safe: false,
    })
}

/// Advance a pointer by `n` elements. The base array is shared.
pub fn carray_advance(ptr: Expression, n: Expression) -> Expression {
    if is_zero(&n) {
        return ptr;
    }

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

/// Retreat a pointer by `n` elements.
pub fn carray_retreat(ptr: Expression, n: Expression) -> Expression {
    if is_zero(&n) {
        return ptr;
    }

    let new_idx = binary(BinOp::Sub, member(ptr.clone(), CARRAY_IDX_KEY), n);
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
    binary(
        BinOp::Sub,
        member(a, CARRAY_IDX_KEY),
        member(b, CARRAY_IDX_KEY),
    )
}

/// In-place advance for `p++` / `p += n`.
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

/// In-place retreat for `p--` / `p -= n`.
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

pub fn is_carray_ptr_kind(ptr: Expression) -> Expression {
    e(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(member(ptr, REF_KIND_KEY)),
        right: Box::new(e(ExprKind::Lit(Literal::Str(CARRAY_KIND.to_string())))),
    })
}

/// Convert a char carray pointer to a string through the libc runtime helper.
pub fn carray_chars_to_string(ptr: Expression) -> Expression {
    call(ident("__libc_char_to_str"), vec![ptr])
}

/// Convert a char array value to a string through the libc runtime helper.
pub fn code_array_to_string(arr: Expression) -> Expression {
    call(ident("__libc_char_to_str"), vec![arr])
}

/// True if the raw initializer text indicates a scalar address-of (`&x`).
pub fn init_is_addr_of(init_source_text: &str) -> bool {
    let t = init_source_text.trim();
    t.starts_with('&') && !t.starts_with("&&")
}
