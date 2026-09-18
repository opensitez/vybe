//! Shared memory-shaped AST constructors.
//!
//! These are not raw host allocations. They normalize language-level
//! allocation/free surfaces onto the runtime values the compiler already knows
//! how to dereference and eventually gives GC/accounting one common place to
//! grow from.

use vybe_ast::{Argument, ArrayElement, ExprKind, Expression, Literal, ObjectProperty};

use crate::primitives::pointers::{CELL_KIND, REF_KIND_KEY, REF_VALUE_KEY};

fn prop(key: &str, value: Expression) -> ObjectProperty {
    ObjectProperty::KeyValue {
        key: Expression::string(key),
        value,
    }
}

fn elem(value: Expression) -> ArrayElement {
    ArrayElement {
        value,
        spread: false,
        key: None,
        by_ref: false,
    }
}

pub fn heap_cell(value: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        prop(REF_KIND_KEY, Expression::string(CELL_KIND)),
        prop(REF_VALUE_KEY, value),
    ]))
}

pub fn heap_cell_null() -> Expression {
    heap_cell(Expression::new(ExprKind::Lit(Literal::Null)))
}

pub fn heap_array(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(values.into_iter().map(elem).collect()))
}

pub fn heap_zeroed_array(count: usize) -> Expression {
    heap_array((0..count).map(|_| Expression::int(0)).collect())
}

/// A zero-filled heap array whose length is only known at RUNTIME.
///
/// `heap_zeroed_array` takes a `usize`, so it can only serve a count the
/// compiler can already see. Every other count fell through to a zero-length
/// array, and since an unwritten slot reads back `undefined`, arithmetic on it
/// produced **NaN** where C guarantees a byte:
///
/// ```text
///            C                     native cc    before
///   malloc(16); a[0]                    0        NaN
///   calloc(w * h, 1); p[0]              0        NaN   (computed count)
/// ```
///
/// `calloc(w * h, 1)` is how chocolate-doom allocates its screens, so this sat
/// directly on the Doom path.
///
/// Lowered as the common primitive `Array(count, 0)`.
///
/// The builtin compiler recognizes this narrow call shape and emits
/// `newWithLength + fill` directly. Keeping the primitive as a call marker avoids
/// routing C-sized allocations through JS constructor/member dispatch
/// (`new Array(n).fill(0)`), which is both slower and far larger in bytecode.
pub fn heap_zeroed_array_sized(count: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("Array".to_string()))),
        args: vec![Argument::positional(count), Argument::positional(Expression::int(0))],
        optional: false,
    })
}

/// A zero-filled BYTE buffer with a runtime length — `new Uint8Array(n)`.
///
/// Typed element storage (`cmemoryplan.md` Stage 1): the declaration knows the
/// element type, so byte-shaped storage gets the byte-shaped backing. A
/// `Uint8Array` is zero by construction (no `.fill(0)` needed), truncates to
/// 8 bits on write the way C does, is the shape Python `bytes` already uses,
/// and is one byte per element instead of a tagged `Value`.
pub fn heap_zeroed_bytes_sized(count: Expression) -> Expression {
    Expression::new(ExprKind::New {
        class: Box::new(Expression::new(ExprKind::Ident("Uint8Array".to_string()))),
        args: vec![Argument::positional(count)],
    })
}

/// The element count of an allocation built by [`heap_zeroed_array_sized`].
///
/// A C allocation is sized in BYTES at the call site, because `malloc(n)` does
/// not know the type it will be assigned to. The DECLARATION does, so it
/// rescales the count — see [`rescale_heap_allocation`].
pub fn heap_allocation_count(expr: &Expression) -> Option<&Expression> {
    match &expr.kind {
        // `heap_zeroed_bytes_sized` — `new Uint8Array(n)`.
        ExprKind::New { class, args } if matches!(&class.kind, ExprKind::Ident(name) if name == "Uint8Array") => {
            args.first().map(|a| &a.value)
        }
        // `heap_zeroed_array_sized` — `Array(n, 0)`.
        ExprKind::Call { callee, args, .. } => {
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "Array") {
                return args.first().map(|a| &a.value);
            }
            None
        }
        _ => None,
    }
}

/// Reinterpret a byte-sized allocation as `bytes / element_size` elements.
///
/// Storage is ELEMENT-indexed — a pointer advances one element, not one byte —
/// so a byte count has to be divided by the element width before it becomes a
/// length. A literal folds; anything else divides at runtime.
pub fn rescale_heap_allocation(expr: Expression, element_size: i64) -> Expression {
    if element_size <= 1 {
        return expr;
    }
    let Some(bytes) = heap_allocation_count(&expr) else {
        return expr;
    };
    if let ExprKind::Lit(Literal::Int(n)) = &bytes.kind {
        return heap_zeroed_array_sized(Expression::int(n / element_size));
    }
    let divided = Expression::new(ExprKind::Binary {
        op: vybe_ast::BinOp::IDiv,
        left: Box::new(bytes.clone()),
        right: Box::new(Expression::int(element_size)),
    });
    heap_zeroed_array_sized(divided)
}

/// Rebuild a sized allocation with the BYTE backing, keeping its count.
///
/// The declaration calls this when it knows the element type is an 8-bit
/// unsigned byte — the count is already in elements (= bytes), only the
/// backing changes. Anything that is not a sized allocation passes through.
pub fn retype_heap_allocation_as_bytes(expr: Expression) -> Expression {
    let Some(count) = heap_allocation_count(&expr).cloned() else {
        return expr;
    };
    heap_zeroed_bytes_sized(count)
}

/// Packed byte read. Stack: [bytes, index] -> [integer byte].
///
/// Typed arrays are handled directly by `ARRAY_GET`; keeping this in the
/// memory primitive avoids routing C byte loads through ECMA host calls.
pub fn emit_bytes_get_item(chunks: &mut [vybe_runtime::Chunk], current: usize, line: u32) {
    chunks[current].emit_op(vybe_runtime::opcode::Op::ARRAY_GET, line);
}

/// Packed byte write. Stack: [bytes, index, value] -> [null].
///
/// This deliberately uses the VM opcode instead of `ecma:uint8array.set`.
/// `ObjectKind::TypedArray` is already handled directly by `ARRAY_SET`, so a
/// statically-known byte buffer does not cross the host boundary for every
/// element write. The trailing null preserves the common `SetItem` call
/// contract for assignment sites that drop the returned value.
pub fn emit_bytes_set_item(chunks: &mut [vybe_runtime::Chunk], current: usize, line: u32) {
    chunks[current].emit_op(vybe_runtime::opcode::Op::ARRAY_SET, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Does `expr` allocate array-backed storage through this module?
///
/// Callers classify a declaration by the SHAPE of its initializer — "is this
/// pointer array-backed?" — and they used to pattern-match `ExprKind::Array`
/// inline. That silently stopped matching the moment a sized allocation became
/// `new Array(n).fill(0)`: reads still worked, but an indexed WRITE fell
/// through to string surgery and threw. Asking the module that BUILDS the
/// shape keeps the answer in one place when the representation changes again.
pub fn is_heap_allocation(expr: &Expression) -> bool {
    match &expr.kind {
        // `heap_array` / `heap_zeroed_array` — a literal array.
        ExprKind::Array(_) => true,
        // `heap_zeroed_bytes_sized` — `new Uint8Array(n)`.
        ExprKind::New { class, .. } => {
            matches!(&class.kind, ExprKind::Ident(name) if name == "Uint8Array")
        }
        // `heap_zeroed_array_sized` — `Array(n, 0)`.
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) if name == "Array" => true,
            _ => false,
        },
        _ => false,
    }
}

pub fn free_value() -> Expression {
    Expression::new(ExprKind::Lit(Literal::Null))
}
