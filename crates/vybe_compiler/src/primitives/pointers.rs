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

use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal, ObjectProperty, UnaryOp};
use vybe_runtime::{Chunk, Op};

pub const REF_KIND_KEY: &str = "__ref_kind";
pub const REF_VALUE_KEY: &str = "__value";
pub const CELL_KIND: &str = "cell";
pub const CARRAY_BASE_KEY: &str = "__base";
pub const CARRAY_IDX_KEY: &str = "__idx";
pub const CARRAY_KIND: &str = "carray";
/// A word in SHARED linear memory — the storage the WASM atomics
/// (`i32.atomic.*`) act on. The third reference shape, and the first whose
/// base is an ADDRESS rather than an object graph: `(base, key)` unification
/// (referenceplan §10) rests on `ARRAY_GET`'s per-`ObjectKind` polymorphism,
/// and linear memory is not an `ObjectKind` — the same carve-out the plan
/// already makes for COBOL `REDEFINES`.
pub const SHARED_KIND: &str = "shared";
/// The word's byte address, allocated from the futex page.
pub const SHARED_ADDR_KEY: &str = "__addr";

/// Static linear-memory pointer backend.
///
/// Dynamic/GC-backed languages use the tagged `cell`/`carray` shapes above.
/// Static pointer languages such as C can keep a pointer as a raw wasm32 byte
/// address and route loads/stores through this backend. Pointer arithmetic is
/// still normalized by the frontend as address + scaled offset; the storage
/// operation itself is common and emits core wasm memory opcodes.
pub const LINEAR_KIND: &str = "linear";

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

fn carray_byte_index(ptr: Expression, offset: Expression, byte_width: i64) -> Expression {
    let byte_index = binary(BinOp::Add, member(ptr, CARRAY_IDX_KEY), offset);
    if byte_width <= 1 {
        byte_index
    } else {
        binary(BinOp::IDiv, byte_index, Expression::int(byte_width))
    }
}

/// Read from a byte-addressed pointer surface such as .NET Marshal.
///
/// The underlying carray remains element-addressed for cross-language pointer
/// interop; this helper converts byte offsets to element slots for typed reads.
pub fn carray_byte_offset_read(ptr: Expression, offset: Expression, byte_width: i64) -> Expression {
    e(ExprKind::Index {
        object: Box::new(member(ptr.clone(), CARRAY_BASE_KEY)),
        index: Box::new(carray_byte_index(ptr, offset, byte_width)),
        null_safe: false,
    })
}

/// Write through a byte-addressed pointer surface such as .NET Marshal.
pub fn carray_byte_offset_write(
    ptr: Expression,
    offset: Expression,
    byte_width: i64,
    val: Expression,
) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(e(ExprKind::Index {
            object: Box::new(member(ptr.clone(), CARRAY_BASE_KEY)),
            index: Box::new(carray_byte_index(ptr, offset, byte_width)),
            null_safe: false,
        })),
        value: Box::new(val),
    })
}

pub fn linear_addr_offset(ptr: Expression, offset: Expression) -> Expression {
    if is_zero(&offset) {
        return ptr;
    }
    binary(BinOp::Add, ptr, offset)
}

pub fn linear_scaled_offset(index: Expression, stride: i64) -> Expression {
    if stride <= 1 {
        return index;
    }
    binary(BinOp::Mul, index, Expression::int(stride))
}

pub fn linear_index_addr(ptr: Expression, index: Expression, stride: i64) -> Expression {
    linear_addr_offset(ptr, linear_scaled_offset(index, stride))
}

/// Offset a pointer whose runtime representation can be either a carray object
/// or a raw linear-memory address.
pub fn hybrid_offset(ptr: Expression, byte_offset: Expression) -> Expression {
    if is_zero(&byte_offset) {
        return ptr;
    }
    e(ExprKind::Ternary {
        cond: Box::new(is_carray_ptr_kind(ptr.clone())),
        then: Box::new(carray_advance(ptr.clone(), byte_offset.clone())),
        else_: Box::new(linear_addr_offset(ptr, byte_offset)),
    })
}

pub fn hybrid_retreat(ptr: Expression, byte_count: Expression) -> Expression {
    hybrid_offset(ptr, binary(BinOp::Sub, Expression::int(0), byte_count))
}

fn emit_value_preserving_store(chunk: &mut Chunk, op: Op, line: u32) {
    let value = chunk.alloc_scratch(1);
    let addr = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, addr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, addr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(op, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
}

pub fn emit_linear_i32_load(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_LOAD, line);
}

pub fn emit_linear_i32_load8_u(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_LOAD8_U, line);
}

pub fn emit_linear_i32_load8_s(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_LOAD8_S, line);
}

pub fn emit_linear_i32_load16_u(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_LOAD16_U, line);
}

pub fn emit_linear_i32_load16_s(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_LOAD16_S, line);
}

pub fn emit_linear_i64_load(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I64_LOAD, line);
}

pub fn emit_linear_i32_store(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_value_preserving_store(&mut chunks[current], Op::I32_STORE, line);
}

pub fn emit_linear_i32_store8(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_value_preserving_store(&mut chunks[current], Op::I32_STORE8, line);
}

pub fn emit_linear_i32_store16(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_value_preserving_store(&mut chunks[current], Op::I32_STORE16, line);
}

pub fn emit_linear_i64_store(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_value_preserving_store(&mut chunks[current], Op::I64_STORE, line);
}

pub fn emit_linear_memory_size(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_idx(Op::MEMORY_SIZE, 0u32, line);
}

pub fn emit_linear_memory_grow(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_idx(Op::MEMORY_GROW, 0u32, line);
}

pub fn emit_linear_memory_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_idx(Op::MEMORY_COPY, 0u32, line);
    chunks[current].emit_leb_u32(0u32, line);
    chunks[current].emit_i32_const(0, line);
}

/// Fill linear bytes in bulk. Stack: [address, byte, length] -> [address].
pub fn emit_linear_memory_fill(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let slots = chunk.alloc_scratch(3);
    for slot in (slots..slots + 3).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    for slot in slots..slots + 3 {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    }
    chunk.emit_op_idx(Op::MEMORY_FILL, 0u32, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slots, line);
}

/// Copy a byte range without rebinding either pointer. Stack: [dst, src, len]
/// -> [dst]. Array-backed operands contain bytes, not wider numeric elements.
/// Linear/linear and array/array copies stay bulk operations. Only crossing
/// between the disjoint linear and managed heaps needs a transfer loop.
pub fn emit_byte_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    use crate::primitives::{class_slots, collections, instructions::host, memory, ops};

    fn get(c: &mut Chunk, slot: u16, line: u32) {
        c.emit_op_u16(Op::LOCAL_GET, slot, line);
    }
    fn set(c: &mut Chunk, slot: u16, line: u32) {
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    fn field(c: &mut Chunk, obj: u16, name: &str, dest: u16, line: u32) {
        let key = class_slots::resolve(
            &class_slots::ClassSlot::internal(name),
            &class_slots::PlainNames,
        );
        class_slots::emit_class_get(
            c,
            class_slots::ObjSource::Local(obj),
            &key,
            class_slots::Dest::Local(dest),
            line,
        );
    }
    fn span(c: &mut Chunk, pointer: u16, base: u16, offset: u16, linear: u16, line: u32) {
        get(c, pointer, line);
        set(c, base, line);
        c.emit_i32_const(0, line);
        set(c, offset, line);
        get(c, base, line);
        host::emit(c, "wasm:js-number", "test", 1, line);
        set(c, linear, line);
        get(c, linear, line);
        c.emit_op(Op::I32_EQZ, line);
        let managed = c.emit_if(line);
        let kind = c.alloc_scratch(1);
        field(c, pointer, REF_KIND_KEY, kind, line);
        get(c, kind, line);
        c.emit_string_const(CARRAY_KIND, line);
        crate::primitives::strings::emit_equals_if_string(c, false, line);
        let wrapped = c.emit_if(line);
        field(c, pointer, CARRAY_BASE_KEY, base, line);
        field(c, pointer, CARRAY_IDX_KEY, offset, line);
        c.emit_end(line);
        c.patch_block(wrapped);
        c.emit_end(line);
        c.patch_block(managed);
    }

    let locals = chunks[current].alloc_scratch(10);
    let (dst, src, len) = (locals, locals + 1, locals + 2);
    let (db, di, dl) = (locals + 3, locals + 4, locals + 5);
    let (sb, si, sl) = (locals + 6, locals + 7, locals + 8);
    let tmp = locals + 9;
    let c = &mut chunks[current];
    set(c, len, line);
    set(c, src, line);
    set(c, dst, line);
    span(c, dst, db, di, dl, line);
    span(c, src, sb, si, sl, line);
    get(c, dl, line);
    get(c, sl, line);
    c.emit_op(Op::I32_AND, line);
    let both_linear = c.emit_if(line);
    get(c, db, line);
    get(c, sb, line);
    get(c, len, line);
    emit_linear_memory_copy(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    c.emit_else(line);
    get(c, dl, line);
    get(c, sl, line);
    c.emit_op(Op::I32_OR, line);
    let cross_heap = c.emit_if(line);
    get(c, sl, line);
    let source_linear = c.emit_if(line);
    // Select the transfer direction once; the byte loop only moves data.
    let byte = c.alloc_scratch(1);
    c.emit_i32_const(0, line);
    set(c, tmp, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, tmp, line);
    get(c, len, line);
    c.emit_op(Op::I32_GE_U, line);
    c.emit_br_if(1, line);
    get(c, sb, line);
    get(c, tmp, line);
    c.emit_op(Op::I32_ADD, line);
    emit_linear_i32_load8_u(chunks, current, line);
    let c = &mut chunks[current];
    set(c, byte, line);
    get(c, db, line);
    get(c, di, line);
    get(c, tmp, line);
    c.emit_op(Op::I32_ADD, line);
    get(c, byte, line);
    memory::emit_bytes_set_item(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    get(c, tmp, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, tmp, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    set(c, tmp, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, tmp, line);
    get(c, len, line);
    c.emit_op(Op::I32_GE_U, line);
    c.emit_br_if(1, line);
    get(c, sb, line);
    get(c, si, line);
    get(c, tmp, line);
    c.emit_op(Op::I32_ADD, line);
    memory::emit_bytes_get_item(chunks, current, line);
    let c = &mut chunks[current];
    set(c, byte, line);
    get(c, db, line);
    get(c, tmp, line);
    c.emit_op(Op::I32_ADD, line);
    get(c, byte, line);
    emit_linear_i32_store8(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    get(c, tmp, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, tmp, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    c.emit_end(line);
    c.patch_block(source_linear);
    c.emit_else(line);

    // TypedArray.set and array.copy snapshot before writing. A typed source
    // can remain a view; an ordinary array needs a range slice for set.
    let source_is_view = c.alloc_scratch(1);
    get(c, sb, line);
    host::emit(c, "ecma:arraybuffer", "isView", 1, line);
    crate::primitives::ops::emit_dyn_to_bool(c, line);
    c.emit_op_u16(Op::LOCAL_TEE, source_is_view, line);
    let typed_source = c.emit_if_value(line);
    get(c, sb, line);
    get(c, si, line);
    get(c, si, line);
    get(c, len, line);
    c.emit_op(Op::I32_ADD, line);
    host::emit(c, "ecma:uint8array", "subarray", 3, line);
    c.emit_else(line);
    get(c, sb, line);
    get(c, si, line);
    get(c, si, line);
    get(c, len, line);
    c.emit_op(Op::I32_ADD, line);
    collections::emit_slice(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_end(line);
    c.patch_block(typed_source);
    set(c, tmp, line);
    get(c, db, line);
    host::emit(c, "ecma:arraybuffer", "isView", 1, line);
    crate::primitives::ops::emit_dyn_to_bool(c, line);
    let typed_dest = c.emit_if(line);
    get(c, db, line);
    get(c, tmp, line);
    get(c, di, line);
    host::emit(c, "ecma:uint8array", "set", 3, line);
    c.emit_op(Op::DROP, line);
    c.emit_else(line);
    get(c, source_is_view, line);
    let convert_typed_source = c.emit_if(line);
    get(c, tmp, line);
    host::emit(c, "ecma:array", "from", 1, line);
    set(c, tmp, line);
    c.emit_end(line);
    c.patch_block(convert_typed_source);
    get(c, db, line);
    get(c, di, line);
    get(c, tmp, line);
    c.emit_i32_const(0, line);
    get(c, len, line);
    collections::emit_gc_array_copy(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_end(line);
    c.patch_block(typed_dest);
    c.emit_end(line);
    c.patch_block(cross_heap);
    c.emit_end(line);
    c.patch_block(both_linear);
    get(c, dst, line);
}

/// Bump allocate from wasm linear memory.
///
/// Stack: [byte_count] -> [address]
///
/// C malloc/calloc-style lowering should reach this primitive instead of
/// generating a C helper function. It keeps the hot path in core wasm:
/// globals for the heap cursor, integer arithmetic, memory.size/grow, and the
/// raw address result.
pub fn emit_linear_alloc(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let n = chunk.alloc_scratch(1);
    let addr = chunk.alloc_scratch(1);
    let new_ptr = chunk.alloc_scratch(1);
    let memory_bytes = chunk.alloc_scratch(1);
    let pages = chunk.alloc_scratch(1);

    // n = align_to(byte_count, 8)
    chunk.emit_i32_const(7, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_i32_const(!7, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);

    // addr = __c_heap_ptr
    crate::primitives::globals::emit_read(chunk, "__c_heap_ptr", line);
    chunk.emit_op_u16(Op::LOCAL_SET, addr, line);

    // new_ptr = addr + n
    chunk.emit_op_u16(Op::LOCAL_GET, addr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, new_ptr, line);

    // memory_bytes = memory.size(0) * 65536
    chunk.emit_op_idx(Op::MEMORY_SIZE, 0u32, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, memory_bytes, line);

    // Grow exactly enough pages if the allocation crosses the current memory.
    chunk.emit_op_u16(Op::LOCAL_GET, new_ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, memory_bytes, line);
    chunk.emit_op(Op::I32_GT_U, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, new_ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, memory_bytes, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(65535, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_DIV_U, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, pages, line);
    chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, new_ptr, line);
    crate::primitives::globals::emit_write(chunk, "__c_heap_ptr", line);
    chunk.emit_op_u16(Op::LOCAL_GET, addr, line);
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

/// Element distance for pointers that may be linear addresses or array-backed.
/// Array-backed indices already count elements; linear addresses count bytes.
pub fn hybrid_element_distance(
    a: Expression,
    b: Expression,
    stride: i64,
    a_is_array_base: bool,
    b_is_array_base: bool,
) -> Expression {
    let is_number = |value: Expression| {
        binary(
            BinOp::Eq,
            e(ExprKind::Unary {
                op: UnaryOp::Typeof,
                expr: Box::new(value),
            }),
            e(ExprKind::Lit(Literal::Str("number".to_string()))),
        )
    };
    let linear = binary(BinOp::Sub, a.clone(), b.clone());
    let linear = if stride == 1 {
        linear
    } else {
        binary(BinOp::Div, linear, e(ExprKind::Lit(Literal::Int(stride))))
    };
    e(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::And, is_number(a.clone()), is_number(b.clone()),)),
        then: Box::new(linear),
        else_: Box::new(carray_diff(
            if a_is_array_base { make_carray_ptr(a, e(ExprKind::Lit(Literal::Int(0)))) } else { a },
            if b_is_array_base { make_carray_ptr(b, e(ExprKind::Lit(Literal::Int(0)))) } else { b },
        )),
    })
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
    e(ExprKind::Ternary {
        cond: Box::new(e(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(e(ExprKind::Unary {
                op: UnaryOp::Typeof,
                expr: Box::new(ptr.clone()),
            })),
            right: Box::new(e(ExprKind::Lit(Literal::Str("object".to_string())))),
        })),
        then: Box::new(e(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(member(ptr, REF_KIND_KEY)),
            right: Box::new(e(ExprKind::Lit(Literal::Str(CARRAY_KIND.to_string())))),
        })),
        else_: Box::new(e(ExprKind::Lit(Literal::Bool(false)))),
    })
}

/// Treat `value` as an array pointer. Existing carray pointers pass through;
/// plain array-like values become a carray pointer at index zero.
pub fn ensure_carray_ptr(value: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(is_carray_ptr_kind(value.clone())),
        then: Box::new(value.clone()),
        else_: Box::new(make_carray_ptr(value, Expression::int(0))),
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
