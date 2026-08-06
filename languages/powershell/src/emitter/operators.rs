//! PowerShell operator lowerings that need a RUNTIME type test.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `@( … )` — the array subexpression operator. It guarantees an array and
/// flattens one level: `@(1..5)` is five elements, `@($arr)` is `$arr`'s
/// elements, and `@(7)` is one element.
///
/// Whether an operand flattens is a question about its RUNTIME value, so it
/// cannot be decided from the syntax: marking elements as spread drops a scalar
/// (nothing to iterate), and nesting them keeps a collection whole. `concat`
/// answers both — it flattens an array argument and appends a non-array one.
///
/// Stack: `[a0, …, aN-1]` → `[array]`.
pub fn emit_ensure_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let concat = chunk.add_import("ecma:array", "concat");

    // Arguments arrive with the LAST on top, so store them before rebuilding
    // the call in source order.
    let base = chunk.alloc_scratch(argc.max(1) as u16);
    for i in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
    }

    chunk.emit_array_new_fixed(0, 0, line);
    for i in 0..argc as u16 {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(concat, 2, line);
    }
}

/// PowerShell's `+`, which is three operations chosen by the LEFT operand:
///
/// | `$a` | `$a + $b` |
/// |------|-----------|
/// | array | `$b`'s elements appended — `@(1,2) + 3` is `@(1,2,3)` |
/// | string | concatenation — `'5' + 5` is `'55'` |
/// | number | arithmetic — `5 + '5'` is `10` |
///
/// No shared primitive answers this. `F64_ADD` and `dynamic_add` both coerce an
/// array operand to a number (`@(1,2) + 3` became `NaN`, and `$a += $x` — the
/// idiomatic way to grow an array in PowerShell — trapped in
/// `wasm:js-number.toF64`). `[builtin_slots.array] add` is not a way out: the
/// shared `compile_binop` consults it for EVERY `+`, so binding
/// `collections.concat` there made `10 + 5` evaluate to `5`.
///
/// The left-operand rule is why this cannot be `dynamic_add` either: that
/// concatenates whenever EITHER side is a string, so it answers `'55'` for
/// `5 + '5'` where PowerShell answers `10`.
///
/// Stack: `[a, b]` → `[result]`.
pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);

    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let concat = chunk.add_import("ecma:array", "concat");
    let is_string = chunk.add_import("wasm:js-string", "test");

    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    // `isArray` answers with a boxed boolean; the branch needs an i32.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    // Array on the left: append. `concat` flattens an array operand and
    // appends a scalar one, which is exactly what `+` means here.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(concat, 2, line);

    chunk.emit_else(line);

    // String on the left: concatenate, coercing the right. Note this tests only
    // the LEFT operand — `emit_dyn_add` would concatenate whenever EITHER side
    // is a string and so answer `'55'` for `5 + '5'`, where PowerShell answers
    // `10` because the left operand is a number.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    // `emit_dyn_add` concatenates whenever either operand is a string, and the
    // left one is — so here it is exactly right, coercions included.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);

    chunk.emit_else(line);

    // Number on the left: arithmetic. `F64_ADD` coerces BOTH operands through
    // `Value::as_f64`, which is what makes `5 + '5'` equal `10`.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_op(Op::F64_ADD, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}
