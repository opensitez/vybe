//! Fixed-width integer semantics over ecma BigInt — the parts ECMA does NOT
//! give you.
//!
//! `ecma:bigint` is arbitrary precision (ECMA-262 §21.2). A language whose
//! integer TYPE lowers to it — Kotlin `Long`, Java `long`, C# `long`/`ulong` —
//! carries the JLS-shaped quirks on top: results wrap to the declared width,
//! shift counts mask to `width - 1` (JLS §15.19), unsigned shifts read the
//! bits unsigned first, and narrowing keeps the low bits (JLS §5.1.3). Width
//! is a PARAMETER here; nothing in this module belongs to a language, and the
//! wrap operators are ECMA's own `BigInt.asIntN`/`asUintN` (§21.2.2).

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Which shift a [`emit_wrapped_shift`] call means.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ShiftKind {
    /// `<<` — result wraps back to the width (the sign bit is reachable).
    Shl,
    /// `>>` arithmetic — sign-propagating.
    Shr,
    /// `>>>` — the bits read unsigned first, then a plain right shift.
    Ushr,
}

/// A width-wrapped shift: `[value, count]` → `[bigint]`.
///
/// The count masks to `width - 1` in i32 (a Long-typed count coerces through
/// `Number` first — BigInt refuses implicit conversion, explicit `Number()`
/// is ECMA-legal), and the result wraps via `asIntN(width, …)` so
/// `1L shl 63` goes negative the way a JVM long does.
pub fn emit_wrapped_shift(chunk: &mut Chunk, width: u32, kind: ShiftKind, line: u32) {
    debug_assert!(
        width.is_power_of_two(),
        "shift-count mask needs a 2^n width"
    );
    let count = chunk.alloc_scratch(1);
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    let number = chunk.add_import("ecma:number", "Number");
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    chunk.emit_call(number, 1, line);
    chunk.emit_call(to_i32, 1, line);
    chunk.emit_i32_const(width as i32 - 1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);

    if kind == ShiftKind::Ushr {
        let as_uint_n = chunk.add_import("ecma:bigint", "asUintN");
        chunk.emit_f64_const(width as f64, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_call(as_uint_n, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    }

    let shift = chunk.add_import(
        "ecma:bigint",
        if kind == ShiftKind::Shl { "shl" } else { "shr" },
    );
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    chunk.emit_call(shift, 2, line);

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    emit_as_int_n_slot(chunk, width, value, line);
}

/// Narrow to `width` bits, signed: `[value]` → `[bigint]`.
/// `BigInt.asIntN(width, x)` — exact at any input width.
pub fn emit_as_int_n(chunk: &mut Chunk, width: u32, line: u32) {
    let t = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, t, line);
    emit_as_int_n_slot(chunk, width, t, line);
}

/// Narrow to `width` bits and hand back a NUMBER: `[value]` → `[f64]`.
///
/// For widths ≤ 32 the wrapped value fits f64 exactly, and a language whose
/// narrower type lives in the number model (Kotlin `Int`) needs the result
/// there — its `"int"` typing gates f64 paths downstream.
pub fn emit_as_int_n_number(chunk: &mut Chunk, width: u32, line: u32) {
    debug_assert!(width <= 32, "wider results do not fit the number model");
    emit_as_int_n(chunk, width, line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
}

/// Truncate toward zero and lift into BigInt: `[value]` → `[bigint]`.
///
/// The widening conversion (JLS §5.1.2-shaped): a Double receiver truncates,
/// a String parses through `Number` first, an existing BigInt passes through
/// the constructor unchanged.
pub fn emit_to_bigint_trunc(chunk: &mut Chunk, line: u32) {
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    crate::primitives::math::emit_trunc(chunk, line);
    let ctor = chunk.add_import("ecma:bigint", "BigInt");
    chunk.emit_call(ctor, 1, line);
}

/// A `±1` step that keeps the operand's numeric type — ECMA §13.4's
/// ToNumeric rule: a BigInt operand steps by 1n and STAYS BigInt, anything
/// else steps in f64. The shared `++` path (`emit_step_by_one`) gates its
/// equivalent on `Compiler::bigint_enabled`; language adapters that lower
/// `++` themselves (PHP's string-increment quirks) call this for their
/// numeric leg instead of a raw `F64_ADD` that traps on a BigInt value.
///
/// Stack: `[value]` → `[value ± 1]`.
pub fn emit_step(chunk: &mut Chunk, add: bool, line: u32) {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    let test_bi = chunk.add_import("wasm:js-bigint", "test");
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(test_bi, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_f64_const(1.0, line);
    let op = chunk.add_import("ecma:bigint", if add { "add" } else { "sub" });
    chunk.emit_call(op, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_f64_const(1.0, line);
    // `dyn_add`, not `F64_ADD`: the non-BigInt arm still meets null/bool
    // operands (PHP `null++` is 1), and the dynamic add coerces them the
    // same way `emit_step_by_one` always has. Subtraction keeps the raw op,
    // also matching that path.
    if add {
        crate::primitives::ops::emit_dyn_add(chunk, line);
    } else {
        chunk.emit_op(Op::F64_SUB, line);
    }
    chunk.emit_end(line);
}

fn emit_as_int_n_slot(chunk: &mut Chunk, width: u32, slot: u16, line: u32) {
    let as_int_n = chunk.add_import("ecma:bigint", "asIntN");
    chunk.emit_f64_const(width as f64, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(as_int_n, 2, line);
}
