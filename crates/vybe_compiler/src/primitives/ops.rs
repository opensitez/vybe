//! Dynamic-dispatch opcode emitters — spec-compliant replacements for DYN_*.
//!
//! Each function emits the WASM-standard type-dispatch sequence using only:
//!   - Standard WASM opcodes (Op::IF / Op::ELSE / Op::END for control flow)
//!   - `wasm:js-*` host imports (js-string-builtins + js-primitive-builtins proposals)
//!
//! Control flow uses actual WASM structured control flow:
//!   `Op::IF` (0x04) / `Op::ELSE` (0x05) / `Op::END` (0x0B)
//! No flat-offset BR_IF_FALSE / BR_IF_TRUE / BR_IF_NULL custom opcodes.

use std::sync::Arc;
use crate::primitives::class_slots;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ── helpers ────────────────────────────────────────────────────────────

fn call1(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 1, line);
}

fn call2(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 2, line);
}

fn alloc_locals(chunk: &mut Chunk, n: u16) -> u16 {
    chunk.alloc_scratch(n)
}

/// Convert the i32 (or Bool) on top of stack to a canonical JS Bool.
/// Stack: [i32 or Bool] → [Bool(true) if nonzero/true, Bool(false) otherwise]
fn i32_to_bool(chunk: &mut Chunk, line: u32) {
    // The constructor handles both a direct i32 and one carried through a
    // value-result join. A conditional here requires a raw i32 at the join
    // and expands every Boolean materialization into two constructor calls.
    let constructor = chunk.add_import("wasm:js-boolean", "fromI32");
    call1(chunk, constructor, line);
}

/// Convert an `i32` (0 or 1) on the stack to a JS `Value::Bool`.
///
/// Used wherever a WASM-level comparison opcode (`REF_IS_*`, `emit_dyn_eq`,
/// etc.) produces `i32` but the ECMA-262 runtime expects `boolean`.
/// Branch conditions (`if`/`br_if`/`emit_loop_cond`) accept both `i32` and
/// `Bool`, so this wrapper is only needed in *value* positions.
///
/// ⛔ This is the one that PRODUCES a boolean. [`emit_dyn_to_bool`] does not —
/// despite the name it CONSUMES any value and produces an `i32`. See its doc.
pub fn emit_i32_to_bool(chunk: &mut Chunk, line: u32) {
    i32_to_bool(chunk, line);
}

/// Compare a possibly absent runtime tag with a string literal, yielding i32.
/// Internal string tags do not need general numeric/object equality dispatch.
pub fn emit_string_slot_eq_literal(chunk: &mut Chunk, slot: u16, literal: &str, line: u32) {
    let test = chunk.add_import("wasm:js-string", "test");
    let equals = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(test, 1, line);
    chunk.emit_if_i32(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_string_const(literal, line);
    chunk.emit_call(equals, 2, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
}

fn save(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn load(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

/// Emit an ECMA BigInt COMPARISON on `[a, b]`, leaving the i32 the dynamic
/// ladders expect.
///
/// ⛔ `Op::I64_*` IS A REAL i64 INSTRUCTION, NOT A DYNAMIC BigInt OPERATION.
/// Every bigint arm in this file used to reach for one — `i64.eq` for
/// equality, `i64.add`, `0 - v` for negation, `i64.ne` against `i64.const 0`
/// for truthiness — with two BOXED BigInt values on the stack. Both halves of
/// the contract were broken in the same place:
///
/// * **VM ≠ wasm.** It only ran because the VM's `I64_EQ`/`I64_NE`/… do
///   `self.pop().as_i64()`, COERCING whatever `Value` they find. `i64.eq`
///   does not mean "coerce two values and compare"; it means "compare two
///   i64s". The interpreter was not executing wasm.
/// * **The operation never reached the host as ECMA.** BigInt comparison and
///   arithmetic are ECMA-262 §6.1.6.2 operations, and `ecma:bigint` has
///   provided the whole set — `eq ne lt le gt ge add sub mul neg …` — all
///   along. The arms simply were not asking for it.
///
/// It is also why nothing we emit can load: js-primitive-builtins specifies
/// `wasm:js-bigint` with `test` and NOTHING ELSE. `fromI64` / `wrapToI64` and
/// the arithmetic were considered and DELIBERATELY REMOVED (Overview.md,
/// "bigint operations"), so a module cannot extract an i64 from a JS BigInt at
/// all. There is no compliant lowering of `i64.eq` on two BigInts — the
/// operation has to be the host's, which is exactly where it belonged.
fn emit_bigint_cmp(chunk: &mut Chunk, name: &str, line: u32) {
    let op = chunk.add_import("ecma:bigint", name);
    call2(chunk, op, line); // [a, b] -> Bool
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    call1(chunk, cast, line); // Bool -> i32, the ladder's arm type
}

/// Emit an ECMA BigInt binary operation on `[a, b]`, leaving a BigInt.
/// See [`emit_bigint_cmp`] for why this is a host call and not `i64.*`.
fn emit_bigint_binop(chunk: &mut Chunk, name: &str, line: u32) {
    let op = chunk.add_import("ecma:bigint", name);
    call2(chunk, op, line); // [a, b] -> BigInt
}

/// Emit ECMA BigInt unary negation on `[a]`. The arms spelled this `0n - a`
/// with `i64.sub`; §6.1.6.2.1 BigInt::unaryMinus is one host operation.
fn emit_bigint_neg(chunk: &mut Chunk, line: u32) {
    let op = chunk.add_import("ecma:bigint", "neg");
    call1(chunk, op, line); // [a] -> BigInt
}


/// The `ecma:bigint` spelling of a relational operator — the host twin of
/// [`i64_cmp_op`], which is now only for real i64 operands.
fn bigint_cmp_name(op: &CmpOp) -> &'static str {
    match op {
        CmpOp::Lt => "lt",
        CmpOp::Gt => "gt",
        CmpOp::Le => "le",
        CmpOp::Ge => "ge",
    }
}

fn i32_const(chunk: &mut Chunk, v: i32, line: u32) {
    chunk.emit_i32_const(v, line);
}


fn f64_const(chunk: &mut Chunk, v: f64, line: u32) {
    chunk.emit_f64_const(v, line);
}

fn emit_object_field_to_slot(
    chunk: &mut Chunk,
    src_slot: u16,
    dst_slot: u16,
    field: &str,
    line: u32,
) {
    let field_key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal(field),
        &class_slots::PlainNames,
    );
    load(chunk, src_slot, line);
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &field_key,
        class_slots::Dest::Stack,
        line,
    );
    save(chunk, dst_slot, line);
}

fn emit_slot_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    load(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    load(chunk, slot, line);
    {
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_both_object_field_present(
    chunk: &mut Chunk,
    a_slot: u16,
    b_slot: u16,
    a_field_slot: u16,
    b_field_slot: u16,
    field: &str,
    line: u32,
) {
    emit_object_field_to_slot(chunk, a_slot, a_field_slot, field, line);
    emit_object_field_to_slot(chunk, b_slot, b_field_slot, field, line);
    emit_slot_is_null_or_undefined(chunk, a_field_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    emit_slot_is_null_or_undefined(chunk, b_field_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
}

/// ECMA-262 §7.1.4 ToNumber, as a value-ABI expression.
///
/// ⛔ EVERY ARM YIELDS A NUMBER, SO EVERY BLOCK IS `(result externref)` —
/// `emit_if_value`, not `emit_if`. A bare `emit_if` declares `(0,0)`, which
/// the VM tolerates (blocks share one operand stack) and WASM rejects with
/// `values remaining on stack at end of block`.
///
/// The arms produce RAW f64s (`f64.const`, `f64.convert_i32_s`,
/// `js-number.toF64`); the writer boxes each onto the value ABI at the block
/// boundary and the consuming f64 op unboxes both operands again — the same
/// round trip `f64.lt` already performs. That is why this needs no raw-f64
/// block kind alongside `block_i32_results`: the value ABI already carries it.
///
/// The result is therefore a boxed NUMBER VALUE, not a bare f64, and no caller
/// needs to change: an `f64` instruction has its operands unboxed by the
/// writer, and an import declared `(param f64)` — `wasm:js-string.fromF64`,
/// `wasm:js-number.fromF64` — has its argument unboxed at the call site by
/// `emit_unbox_raw_params`. Raw is not expressible across a block or a call
/// boundary here, and it does not need to be.
fn emit_js_to_number_f64(
    chunk: &mut Chunk,
    slot: u16,
    to_f64: u16,
    test_bool: u16,
    cast_bool: u16,
    line: u32,
) {
    load(chunk, slot, line);
    {
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_if_value(line);
    f64_const(chunk, f64::NAN, line);
    chunk.emit_else(line);

    load(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    f64_const(chunk, 0.0, line);
    chunk.emit_else(line);

    load(chunk, slot, line);
    call1(chunk, test_bool, line);
    chunk.emit_if_value(line);
    load(chunk, slot, line);
    call1(chunk, cast_bool, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    chunk.emit_else(line);
    load(chunk, slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}

// ── emit_dyn_to_bool ───────────────────────────────────────────────────

/// Truthy coercion — ECMA-262 §7.1.2 ToBoolean.
/// Stack: [v] → [i32: 0 or 1]
/// Calls the shared ToBoolean implementation and narrows its Boolean to i32.
///
/// ⛔ **THE NAME LIES ABOUT THE DIRECTION, AND IT IS NOT THE INVERSE OF
/// [`emit_i32_to_bool`].** Read the two together:
///
/// | helper | in | out | means |
/// |---|---|---|---|
/// | `emit_dyn_to_bool`  | any value | **i32** | ToBoolean — *narrow to a condition* |
/// | `emit_i32_to_bool`  | i32       | **`Value::Bool`** | *widen to a real boolean* |
///
/// So this one is right in a CONDITION position and wrong in a VALUE position:
/// using it to "make a boolean" leaves a **number**. That is exactly how
/// `c instanceof C` came to print `1` with `typeof` reporting `"number"` —
/// three rebuilds were spent blaming js-specific rendering for a helper that
/// was doing precisely what it says on the tin. If what you need is a value
/// ECMA-262 calls a `boolean`, you want [`emit_i32_to_bool`].
pub fn emit_dyn_to_bool(chunk: &mut Chunk, line: u32) {
    // Reuse the platform's ToBoolean implementation instead of expanding the
    // same type-dispatch tree at every condition. Keep the public contract an
    // i32 condition: the first call returns a language Boolean, the second
    // converts that known Boolean to 0/1.
    let convert = chunk.add_import("ecma:boolean", "toBoolean");
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    call1(chunk, convert, line);
    call1(chunk, cast, line);
}

/// Consume a condition whose producer already guarantees an i32 predicate.
/// This marker emits no guest instructions in production. Benchmark builds
/// can recreate the former generic conversion while compiling the same source.
#[inline(always)]
pub fn emit_proven_i32_condition(_code: &mut Chunk, _line: u32) {
    #[cfg(feature = "primitive-benchmarks")]
    if !REUSE_PROVEN_I32_CONDITIONS.with(|v| v.get()) {
        emit_dyn_to_bool(_code, _line);
    }
}

/// Compile-time control for recipes whose old predicate was a literal true.
#[inline(always)]
pub(crate) fn proven_i32_conditions_enabled() -> bool {
    #[cfg(feature = "primitive-benchmarks")]
    { REUSE_PROVEN_I32_CONDITIONS.with(|v| v.get()) }
    #[cfg(not(feature = "primitive-benchmarks"))]
    { true }
}

/// A compile-time true branch condition, with the same legacy benchmark arm.
pub fn emit_proven_true_condition(code: &mut Chunk, line: u32) {
    #[cfg(feature = "primitive-benchmarks")]
    if !REUSE_PROVEN_I32_CONDITIONS.with(|v| v.get()) {
        code.emit_bool_const(true, line);
        emit_dyn_to_bool(code, line);
        return;
    }
    code.emit_i32_const(1, line);
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static REUSE_PROVEN_I32_CONDITIONS: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}

#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_proven_i32_conditions<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { REUSE_PROVEN_I32_CONDITIONS.with(|v| v.set(self.0)); }
    }
    let _restore = Restore(REUSE_PROVEN_I32_CONDITIONS.with(|v| v.replace(enabled)));
    build()
}

/// Lua truthiness — only `nil` and `false` are falsy (§3.3.3).
/// Stack: [v] → [i32: 0 or 1]
pub fn emit_lua_to_bool(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 1);
    let v = slots;
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");

    save(chunk, v, line);

    load(chunk, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_i32(line);
    i32_const(chunk, 0, line);
    chunk.emit_else(line);

    load(chunk, v, line);
    call1(chunk, test_bool, line);
    chunk.emit_if_i32(line);
    load(chunk, v, line);
    call1(chunk, cast_bool, line);
    chunk.emit_else(line);
    i32_const(chunk, 1, line);

    chunk.emit_end(line); // bool
    chunk.emit_end(line); // null
}

// ── emit_dyn_not ──────────────────────────────────────────────────────

pub fn emit_dyn_not(chunk: &mut Chunk, line: u32) {
    emit_dyn_to_bool(chunk, line); // any → i32
    chunk.emit_op(Op::I32_EQZ, line); // negate: i32
    // Result is i32 — WASM-compliant
}

// ── emit_dyn_eq ───────────────────────────────────────────────────────

pub fn emit_dyn_eq(chunk: &mut Chunk, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        emit_comparison_helper_call(chunk, "__vybe_dyneq", line);
        return;
    }
    emit_dyn_eq_inline(chunk, line);
}

/// Both operands have been evaluated; one is a known null literal.
/// Keep only the other operand, preserving the evaluation order.
fn retain_other_null_operand(chunk: &mut Chunk, null_on_left: bool, line: u32) {
    if null_on_left {
        let value = alloc_locals(chunk, 1);
        save(chunk, value, line);
        chunk.emit_op(Op::DROP, line);
        load(chunk, value, line);
    } else {
        chunk.emit_op(Op::DROP, line);
    }
}

/// Specialize the nullish arms of emit_dyn_eq for a literal null operand.
pub fn emit_dyn_eq_null_operand(chunk: &mut Chunk, null_on_left: bool, line: u32) {
    retain_other_null_operand(chunk, null_on_left, line);
    let value = alloc_locals(chunk, 1);
    save(chunk, value, line);
    emit_slot_is_null_or_undefined(chunk, value, line);
}

fn emit_dyn_eq_inline(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 4);
    let b_slot = slots;
    let a_slot = slots + 1;
    let b_time_slot = slots + 2;
    let a_time_slot = slots + 3;

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);
    emit_dyn_eq_slots(chunk, a_slot, b_slot, a_time_slot, b_time_slot, line);
}

fn emit_dyn_eq_slots(
    chunk: &mut Chunk,
    a_slot: u16,
    b_slot: u16,
    a_time_slot: u16,
    b_time_slot: u16,
    line: u32,
) {

    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_eq = chunk.add_import("wasm:js-string", "equals");
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let abstract_eq = chunk.add_import("ecma:value", "abstractEq");

    // both number?
    load(chunk, a_slot, line);
    call1(chunk, test_num, line);
    load(chunk, b_slot, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line); // i32: 1 if both numbers
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);

    // both string?
    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line); // i32: 1 if both strings
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_eq, line);
    chunk.emit_else(line);

    // both boolean?
    load(chunk, a_slot, line);
    call1(chunk, test_bool, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bool, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    // cast_bool(v) → i32 (1=true, 0=false) for known Bool values
    load(chunk, a_slot, line);
    call1(chunk, cast_bool, line);
    load(chunk, b_slot, line);
    call1(chunk, cast_bool, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_else(line);

    // both bigint?
    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    emit_bigint_cmp(chunk, "eq", line);
    chunk.emit_else(line);
    // Matching primitive pairs are necessarily non-null. Keep nullish checks
    // on the fallback path so common primitive comparisons avoid those calls.
    // All preceding type tests are pure; object-field access order stays intact.
    emit_slot_is_null_or_undefined(chunk, a_slot, line);
    chunk.emit_if_i32(line);
    emit_slot_is_null_or_undefined(chunk, b_slot, line);
    chunk.emit_else(line);
    emit_slot_is_null_or_undefined(chunk, b_slot, line);
    chunk.emit_if_i32(line);
    i32_const(chunk, 0, line);
    chunk.emit_else(line);
    // DateTime-like comparable objects carry a numeric `__time` field.
    emit_both_object_field_present(
        chunk,
        a_slot,
        b_slot,
        a_time_slot,
        b_time_slot,
        "Ticks",
        line,
    );
    chunk.emit_if_i32(line);
    load(chunk, a_time_slot, line);
    call1(chunk, to_f64, line);
    load(chunk, b_time_slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);
    emit_both_object_field_present(
        chunk,
        a_slot,
        b_slot,
        a_time_slot,
        b_time_slot,
        "__value",
        line,
    );
    chunk.emit_if_i32(line);
    load(chunk, a_time_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_time_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_time_slot, line);
    load(chunk, b_time_slot, line);
    call2(chunk, str_eq, line);
    chunk.emit_else(line);
    load(chunk, a_time_slot, line);
    load(chunk, b_time_slot, line);
    chunk.emit_op(Op::REF_EQ, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // one `__value` wrapper and one raw primitive/value → wrapped equality
    emit_object_field_to_slot(chunk, a_slot, a_time_slot, "__value", line);
    emit_slot_is_null_or_undefined(chunk, a_time_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_i32(line);
    load(chunk, a_time_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, abstract_eq, line);
    call1(chunk, cast_bool, line);
    chunk.emit_else(line);
    emit_object_field_to_slot(chunk, b_slot, b_time_slot, "__value", line);
    emit_slot_is_null_or_undefined(chunk, b_time_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_time_slot, line);
    call2(chunk, abstract_eq, line);
    call1(chunk, cast_bool, line);
    chunk.emit_else(line);
    // object / cross-type → reference equality
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::REF_EQ, line);
    chunk.emit_end(line); // b.__value wrapper
    chunk.emit_end(line); // a.__value wrapper
    chunk.emit_end(line); // __value object
    chunk.emit_end(line); // comparable object
    chunk.emit_end(line); // b null
    chunk.emit_end(line); // a null
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // boolean
    chunk.emit_end(line); // string
    chunk.emit_end(line); // number
    // Result is i32 (0 or 1) — WASM-compliant for IF conditions
}

fn emit_slot_is_null_only(chunk: &mut Chunk, slot: u16, line: u32) {
    // Standard ref.is_null already excludes the non-null Undefined externref.
    load(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
}

/// Specialize strict ECMA equality against null, excluding undefined.
pub fn emit_js_strict_eq_null_operand(chunk: &mut Chunk, null_on_left: bool, line: u32) {
    retain_other_null_operand(chunk, null_on_left, line);
    let value = alloc_locals(chunk, 1);
    save(chunk, value, line);
    emit_slot_is_null_only(chunk, value, line);
}

pub fn emit_js_strict_eq(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num = chunk.add_import("wasm:js-number", "test");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_eq = chunk.add_import("wasm:js-string", "equals");
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    load(chunk, a_slot, line);
    {
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_if_i32(line);
    load(chunk, b_slot, line);
    {
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_else(line);

    emit_slot_is_null_only(chunk, a_slot, line);
    chunk.emit_if_i32(line);
    emit_slot_is_null_only(chunk, b_slot, line);
    chunk.emit_else(line);

    load(chunk, a_slot, line);
    call1(chunk, test_num, line);
    load(chunk, b_slot, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);

    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_eq, line);
    chunk.emit_else(line);

    load(chunk, a_slot, line);
    call1(chunk, test_bool, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bool, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    call1(chunk, cast_bool, line);
    load(chunk, b_slot, line);
    call1(chunk, cast_bool, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_else(line);

    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    emit_bigint_cmp(chunk, "eq", line);
    chunk.emit_else(line);

    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::REF_EQ, line);

    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // boolean
    chunk.emit_end(line); // string
    chunk.emit_end(line); // number
    chunk.emit_end(line); // null
    chunk.emit_end(line); // undefined
}

pub fn emit_dyn_ne(chunk: &mut Chunk, line: u32) {
    emit_dyn_eq(chunk, line); // i32
    chunk.emit_op(Op::I32_EQZ, line); // negate: 1 if not-equal
}

// ── comparison helpers ────────────────────────────────────────────────

enum CmpOp {
    Lt,
    Gt,
    Le,
    Ge,
}

fn f64_cmp_op(op: &CmpOp) -> Op {
    match op {
        CmpOp::Lt => Op::F64_LT,
        CmpOp::Gt => Op::F64_GT,
        CmpOp::Le => Op::F64_LE,
        CmpOp::Ge => Op::F64_GE,
    }
}

fn i32_cmp_op(op: &CmpOp) -> Op {
    match op {
        CmpOp::Lt => Op::I32_LT_S,
        CmpOp::Gt => Op::I32_GT_S,
        CmpOp::Le => Op::I32_LE_S,
        CmpOp::Ge => Op::I32_GE_S,
    }
}


fn emit_dyn_cmp(chunk: &mut Chunk, line: u32, op: CmpOp) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        let name = match op {
            CmpOp::Lt => "__vybe_dynlt",
            CmpOp::Gt => "__vybe_dyngt",
            CmpOp::Le => "__vybe_dynle",
            CmpOp::Ge => "__vybe_dynge",
        };
        emit_comparison_helper_call(chunk, name, line);
        return;
    }
    emit_dyn_cmp_inline(chunk, line, op);
}

fn emit_dyn_cmp_inline(chunk: &mut Chunk, line: u32, op: CmpOp) {
    let slots = alloc_locals(chunk, 4);
    let b_slot = slots;
    let a_slot = slots + 1;
    let b_time_slot = slots + 2;
    let a_time_slot = slots + 3;

    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // both number?
    load(chunk, a_slot, line);
    call1(chunk, test_num, line);
    load(chunk, b_slot, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_else(line);

    // both string?
    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_compare, line); // i32 (-1/0/1)
    i32_const(chunk, 0, line);
    chunk.emit_op(i32_cmp_op(&op), line);
    chunk.emit_else(line);

    // both bigint?
    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    emit_bigint_cmp(chunk, bigint_cmp_name(&op), line);
    chunk.emit_else(line);
    // DateTime-like comparable objects carry a numeric `__time` field.
    emit_both_object_field_present(
        chunk,
        a_slot,
        b_slot,
        a_time_slot,
        b_time_slot,
        "Ticks",
        line,
    );
    chunk.emit_if_i32(line);
    load(chunk, a_time_slot, line);
    call1(chunk, to_f64, line);
    load(chunk, b_time_slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_else(line);
    // fallback: coerce both to f64
    emit_js_to_number_f64(chunk, a_slot, to_f64, test_bool, cast_bool, line);
    emit_js_to_number_f64(chunk, b_slot, to_f64, test_bool, cast_bool, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_end(line); // comparable object
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // string
    chunk.emit_end(line); // number
    // ⛔ EVERY ARM OF THIS LADDER ENDS IN AN i32 COMPARISON, so every
    // block here is `(result i32)` — `emit_if_i32`, not `emit_if`. A bare
    // `emit_if` declares `(0,0)`: the VM tolerates it (one shared operand
    // stack), WASM does not (`values remaining on stack at end of block`),
    // and the writer, seeing no i32 marking, boxed each arm's raw i32 into
    // an externref the void block could not carry either.
    // Result is i32 (0 or 1) — WASM-compliant for IF/BR_IF conditions
}

pub fn emit_dyn_lt(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Lt);
}
pub fn emit_dyn_gt(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Gt);
}
pub fn emit_dyn_le(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Le);
}
pub fn emit_dyn_ge(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Ge);
}

/// JS Abstract Relational Comparison (ECMA-262 §7.2.13). Operands are
/// expected to already be primitives (the caller runs ToPrimitive first).
/// Both strings → lexicographic compare; both bigints → i64 compare;
/// otherwise ToNumber each operand and compare as f64.
///
/// The difference from [`emit_dyn_cmp`]: the mixed-type fallback uses
/// ECMA-262 §7.1.4 `ecma:value.toNumber` (which parses numeric strings,
/// yields NaN for non-numeric ones so `"foo" < 0` → NaN → false, and
/// throws a TypeError for symbols) instead of the strict
/// `wasm:js-number.toF64`, which traps ("not a number") on any
/// non-Number value.
///
/// MUST be emitted into the active code chunk — it resolves
/// `ecma:value.toNumber` via `chunk.add_import`, which targets *this*
/// chunk's import table. Never route it through an `_into` shared-runtime
/// chunk (that would add the import to the wrong table).
fn emit_js_relational_cmp(
    chunk: &mut Chunk,
    line: u32,
    op: CmpOp,
    a_known_num: bool,
    b_known_num: bool,
) {
    // ── Operands this emitter MATERIALISED ITSELF ──────────────────────
    //
    // `a_known_num` / `b_known_num` mean the caller pushed that operand as a
    // numeric constant (`f64.const 13`), so every question the ladder asks
    // about it has a compile-time answer: `wasm:js-number.test` is 1,
    // `wasm:js-number.toF64` is the identity, and — because a constant number
    // is neither — `both string?` and `both bigint?` are FALSE whenever
    // EITHER operand is one. Asking anyway cost 2 host calls per iteration on
    // the taken path and 4 more on the mixed path, for `i < 13`.
    //
    // `Literal::BigInt` is a separate AST variant and is never reported here,
    // so the bigint arm is only ever dropped when it is provably unreachable.
    if a_known_num && b_known_num {
        // Both constants: the number arm is the only reachable one, and the
        // operands are already on the stack in order.
        chunk.emit_op(f64_cmp_op(&op), line);
        return;
    }
    if a_known_num || b_known_num {
        let slots = alloc_locals(chunk, 2);
        let b_slot = slots;
        let a_slot = slots + 1;

        let test_num = chunk.add_import("wasm:js-number", "test");
        let to_number = chunk.add_import("ecma:value", "toNumber");

        save(chunk, b_slot, line);
        save(chunk, a_slot, line);

        // Only the UNKNOWN side is asked; the `i32.and` goes with the other.
        load(chunk, if a_known_num { b_slot } else { a_slot }, line);
        call1(chunk, test_num, line);
        chunk.emit_if_i32(line);
        // F64 comparison consumes Number operands directly; the test above
        // proves the unknown operand and the other is a numeric literal.
        load(chunk, a_slot, line);
        load(chunk, b_slot, line);
        chunk.emit_op(f64_cmp_op(&op), line);
        chunk.emit_else(line);
        // Mixed: ToNumber the unknown side (NaN-safe — `"foo" < 0` is false).
        load(chunk, a_slot, line);
        if !a_known_num {
            call1(chunk, to_number, line);
        }
        load(chunk, b_slot, line);
        if !b_known_num {
            call1(chunk, to_number, line);
        }
        chunk.emit_op(f64_cmp_op(&op), line);
        chunk.emit_end(line);
        return;
    }

    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num = chunk.add_import("wasm:js-number", "test");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let to_number = chunk.add_import("ecma:value", "toNumber");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // Both numbers: F64 comparison performs the required extraction itself.
    // Avoid explicit host calls that only re-prove/return the same Number.
    // both number?
    load(chunk, a_slot, line);
    call1(chunk, test_num, line);
    load(chunk, b_slot, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_else(line);

    // both string?
    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_compare, line); // i32 (-1/0/1)
    i32_const(chunk, 0, line);
    chunk.emit_op(i32_cmp_op(&op), line);
    chunk.emit_else(line);

    // both bigint?
    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    emit_bigint_cmp(chunk, bigint_cmp_name(&op), line);
    chunk.emit_else(line);
    // fallback: ToNumber both (NaN-safe), compare as f64.
    load(chunk, a_slot, line);
    call1(chunk, to_number, line);
    load(chunk, b_slot, line);
    call1(chunk, to_number, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // string
    chunk.emit_end(line); // number
    // Result is i32 (0 or 1) — WASM-compliant for IF/BR_IF conditions.
}

pub fn emit_js_lt(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Lt, a_known_num, b_known_num);
}
pub fn emit_js_gt(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Gt, a_known_num, b_known_num);
}
pub fn emit_js_le(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Le, a_known_num, b_known_num);
}
pub fn emit_js_ge(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Ge, a_known_num, b_known_num);
}

// ── emit_dyn_add ──────────────────────────────────────────────────────

/// Replace an OBJECT in `slot` with its primitive, in place — ECMA-262 §7.1.1
/// `ToPrimitive` with hint `"default"`, read from the DECLARED slots.
///
/// The three slots are already the chain: `ProtocolSlot::ToPrimitive`
/// (`Symbol.toPrimitive`), then `ValueOf`, then `ToString`. Nothing in the tree
/// read the first two, and `ToString` was read only by the render helpers — so
/// an object reaching an arithmetic operator was coerced straight to a float,
/// which is why `"a=" + someObject` trapped in `wasm:js-number.toF64`.
///
/// Asking the VALUE is the whole point: the slot travels with the object, so a
/// C# `ToString`, a PHP `__toString`, a Python `__str__`, a Ruby `to_s` and a
/// Dart `toString` all answer here, and an object crossing a language boundary
/// keeps answering. There is no language name and no profile flag in this
/// decision, and there cannot be one — the operand is the only thing that knows.
///
/// Two deliberate departures from §7.1.1, both toward the existing behaviour:
/// `valueOf` returning an OBJECT falls through to `toString` (the spec's
/// "must be primitive" retry), but an object declaring NONE of the three is
/// left untouched rather than raising `TypeError`, so every value that works
/// today keeps working. JS itself does not come through here at all — its
/// profile routes `+` to `ecma:value.add`, which is the host's exact §13.15.4.
fn emit_to_primitive_in_place(chunk: &mut Chunk, slot: u16, line: u32) {
    let method = alloc_locals(chunk, 1);

    load(chunk, slot, line);
    crate::primitives::instructions::recipes::is_object(chunk, line);
    chunk.emit_if(line);
    for protocol_slot in [
        vybe_ast::ProtocolSlot::ToPrimitive,
        vybe_ast::ProtocolSlot::ValueOf,
        vybe_ast::ProtocolSlot::ToString,
    ] {
        // Re-tested each round: an earlier link may have answered, and then
        // this one must not overwrite it. `is_object` is also the "did we
        // already get a primitive" test, which is what makes `valueOf`
        // returning an object fall through to `toString`.
        load(chunk, slot, line);
        crate::primitives::instructions::recipes::is_object(chunk, line);
        chunk.emit_if(line);

        // ⛔ WAS `add_constant(protocol_slot_key(..))` — see §2a; the owner
        // holds the spelling, callers hold the BINDING.
        let key = crate::primitives::class_slots::resolve_interned(
            chunk,
            &crate::primitives::class_slots::ClassSlot::Slot(protocol_slot),
            &crate::primitives::class_slots::PlainNames,
        );
        load(chunk, slot, line);
        crate::primitives::class_slots::emit_class_get(
            chunk,
            crate::primitives::class_slots::ObjSource::Stack,
            &key,
            crate::primitives::class_slots::Dest::Stack,
            line,
        );
        chunk.emit_op_u16(Op::LOCAL_SET, method, line);

        emit_slot_is_null_or_undefined(chunk, method, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        load(chunk, method, line);
        load(chunk, slot, line);
        crate::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        chunk.emit_end(line);

        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}

pub fn emit_dyn_add(chunk: &mut Chunk, line: u32) {
    emit_dyn_add_known(chunk, line, false, false);
}

/// `+` where the executing language states how a Boolean spells itself.
///
/// Separate entry point rather than a new parameter on `emit_dyn_add`: that
/// one has 140-odd callers, almost all of them internal emits adding numbers,
/// and none of them wants to answer a question about Booleans.
pub fn emit_dyn_add_with_bool_text(
    chunk: &mut Chunk,
    bool_text: Option<vybe_ast::BoolText>,
    line: u32,
) {
    emit_dyn_add_inner(chunk, line, false, false, bool_text);
}

/// Render the non-string operand of a legacy shared-helper string addition.
/// The caller proves the other operand is a string. Reuse the original
/// numeric/Boolean/null/undefined rendering and strict error behavior; this
/// is not the language-level String constructor or a new coercion policy.
pub fn emit_add_string_operand_from_local(chunk: &mut Chunk, slot: u16, line: u32) {
    emit_add_operand_as_string(chunk, slot, false, None, line);
}

/// Push operand `slot` coerced to a string, for the concat arm of `+`.
///
/// `known_num` says this emitter wrote the value itself as an f64 constant, so
/// "is it already a string?" has a compile-time answer — no — and the whole
/// test/cast branch collapses to the numeric side.
fn emit_add_operand_as_string(
    chunk: &mut Chunk,
    slot: u16,
    known_num: bool,
    bool_text: Option<vybe_ast::BoolText>,
    line: u32,
) {
    let str_from_f64 = chunk.add_import("wasm:js-string", "fromF64");
    if known_num {
        load(chunk, slot, line);
        call1(chunk, str_from_f64, line);
        return;
    }
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_cast = chunk.add_import("wasm:js-string", "cast");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    load(chunk, slot, line);
    call1(chunk, test_str, line); // i32: 1 if already string
    chunk.emit_if_value(line);
    load(chunk, slot, line);
    call1(chunk, str_cast, line);
    chunk.emit_else(line);
    // ⛔ A BOOLEAN IS NOT A NUMBER HERE. Without this arm the ToNumber ladder
    // below turns `true` into `1.0` and the concat reads `1` — which is what
    // every C# `bool` in a string printed. `Directives::bool_text` says how the
    // executing language spells it; `None` keeps the ECMA ladder exactly as it
    // was, so no language that did not ask for this changes a byte.
    if let Some((true_text, false_text)) = bool_text.map(vybe_ast::BoolText::texts) {
        load(chunk, slot, line);
        call1(chunk, test_bool, line);
        chunk.emit_if_value(line);
        load(chunk, slot, line);
        call1(chunk, cast_bool, line);
        chunk.emit_if_value(line);
        emit_string_const(chunk, true_text, line);
        chunk.emit_else(line);
        emit_string_const(chunk, false_text, line);
        chunk.emit_end(line);
        chunk.emit_else(line);
        emit_js_to_number_f64(chunk, slot, to_f64, test_bool, cast_bool, line);
        call1(chunk, str_from_f64, line);
        chunk.emit_end(line);
    } else {
        emit_js_to_number_f64(chunk, slot, to_f64, test_bool, cast_bool, line);
        call1(chunk, str_from_f64, line);
    }
    chunk.emit_end(line);
}

fn emit_string_const(chunk: &mut Chunk, text: &str, line: u32) {
    chunk.emit_string_const(text, line);
}

/// Push operand `slot` as a raw f64, for the numeric arm of `+`.
///
/// `known_num` collapses the entire ToNumber ladder — undefined? null?
/// boolean? unbox? — to a bare load, because the value in the slot is the
/// `f64.const` this emitter just wrote.
fn emit_add_operand_as_f64(chunk: &mut Chunk, slot: u16, known_num: bool, line: u32) {
    if known_num {
        load(chunk, slot, line);
        return;
    }
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    emit_js_to_number_f64(chunk, slot, to_f64, test_bool, cast_bool, line);
}

/// ECMA §13.15.3 `+`, with the operands the emitter already knows about.
///
/// `a_known_num` / `b_known_num` say that operand is a numeric constant THIS
/// emitter pushed (`emit_const(Value::F64(..))` / a number literal), not an
/// inferred or annotated type. Every runtime question the coercion path asks
/// about such an operand therefore has a compile-time answer:
///
/// - `emit_to_primitive_in_place` — a primitive is already primitive;
/// - `wasm:js-string test` — a number is not a string, so the `OR` collapses
///   to the other operand and, when BOTH are known, the concat arm is dead;
/// - `wasm:js-bigint test` — the bigint arm is `test(a) AND test(b)`, so ONE
///   known number makes the whole arm unreachable. `Literal::BigInt` is a
///   separate variant and is never marked known, so the arm is only ever
///   dropped when provably false.
/// - the ToNumber ladder — see `emit_add_operand_as_f64`.
///
/// `emit_dyn_add` is this with both flags `false`, so every existing caller
/// emits the same instruction sequence as before. △ Not byte-identical: the
/// operand-coercion imports moved into the two helpers below, so the import
/// TABLE is ordered differently. `add_import` dedups by name and every
/// `CallHost` still resolves to the right entry — but a bytecode-dump diff
/// will show shifted import indices, which is not a regression.
///
/// △ Only `emit_step_by_one` calls in today, always as `(false, true)`. The
/// `a_known_num` shapes and the both-known early return are therefore emitted
/// by no caller and covered by no test — whoever wires `compile_binop_operands`
/// through here is their first user and should verify them.
pub fn emit_dyn_add_known(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_dyn_add_inner(chunk, line, a_known_num, b_known_num, None);
}

fn emit_dyn_add_inner(
    chunk: &mut Chunk,
    line: u32,
    a_known_num: bool,
    b_known_num: bool,
    bool_text: Option<vybe_ast::BoolText>,
) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_concat = chunk.add_import("wasm:js-string", "concat");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let from_f64 = chunk.add_import("wasm:js-number", "fromF64");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // ToPrimitive BOTH operands first, then decide — the order §13.15.3 uses,
    // and it matters: the String test below is on the PRIMITIVES, not on the
    // originals, so an object whose `toString` yields a string concatenates
    // rather than being forced through a numeric coercion.
    if !a_known_num {
        emit_to_primitive_in_place(chunk, a_slot, line);
    }
    if !b_known_num {
        emit_to_primitive_in_place(chunk, b_slot, line);
    }

    // Both operands are known numbers: neither the concat arm nor the bigint
    // arm is reachable, so `+` is the f64 add it always was.
    if a_known_num && b_known_num {
        load(chunk, a_slot, line);
        load(chunk, b_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
        call1(chunk, from_f64, line);
        return;
    }

    // either is a string → string concatenation
    if !a_known_num {
        load(chunk, a_slot, line);
        call1(chunk, test_str, line);
    }
    if !b_known_num {
        load(chunk, b_slot, line);
        call1(chunk, test_str, line);
    }
    if !a_known_num && !b_known_num {
        chunk.emit_op(Op::I32_OR, line); // i32: 1 if either is string
    }
    chunk.emit_if_value(line);
    emit_add_operand_as_string(chunk, a_slot, a_known_num, bool_text, line);
    emit_add_operand_as_string(chunk, b_slot, b_known_num, bool_text, line);
    call2(chunk, str_concat, line);

    chunk.emit_else(line);
    // both bigint → i64.add. One known number makes the `AND` provably false,
    // so the arm is emitted only when neither operand is known.
    let bigint_arm = !a_known_num && !b_known_num;
    if bigint_arm {
        load(chunk, a_slot, line);
        call1(chunk, test_bigint, line);
        load(chunk, b_slot, line);
        call1(chunk, test_bigint, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if_value(line);
        load(chunk, a_slot, line);
        load(chunk, b_slot, line);
        emit_bigint_binop(chunk, "add", line);
        chunk.emit_else(line);
    }
    // number + number (or coerce) → f64.add
    emit_add_operand_as_f64(chunk, a_slot, a_known_num, line);
    emit_add_operand_as_f64(chunk, b_slot, b_known_num, line);
    chunk.emit_op(Op::F64_ADD, line);
    call1(chunk, from_f64, line);
    if bigint_arm {
        chunk.emit_end(line); // bigint
    }
    chunk.emit_end(line); // string
}

// ── emit_dyn_neg ──────────────────────────────────────────────────────

pub fn emit_dyn_neg(chunk: &mut Chunk, line: u32) {
    let v = alloc_locals(chunk, 1);

    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let from_f64 = chunk.add_import("wasm:js-number", "fromF64");
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");

    save(chunk, v, line);

    // bigint → i64 negation
    load(chunk, v, line);
    call1(chunk, test_bigint, line); // i32: 1 if bigint
    chunk.emit_if_value(line);
    load(chunk, v, line);
    emit_bigint_neg(chunk, line);
    chunk.emit_else(line);
    // number → f64 negation
    emit_js_to_number_f64(chunk, v, to_f64, test_bool, cast_bool, line);
    chunk.emit_op(Op::F64_NEG, line);
    call1(chunk, from_f64, line);
    chunk.emit_end(line);
}

// ── _into variants ────────────────────────────────────────────────────────
//
// These register imports on `imports` (the module-level chunk, chunks[0]) and
// emit bytecode into `code` (the function chunk being built).  This separation
// is REQUIRED for WASM compliance: a WASM module has ONE import section.
//
// Without this split, every emit_dyn_* call on a code chunk registers imports
// locally on that chunk (via chunk.add_import).  When emit_import_call_into
// also emits CALL_IMPORT using the module-level index, the two index spaces
// collide: CALL_IMPORT K in the code chunk resolves to the code chunk's local
// import K (e.g. wasm:js-boolean.test) instead of the module-level import K
// (e.g. ecma:array.new).  Use the _into variants in any stdlib builder that
// has a separate `imports: &mut Chunk` parameter.

pub fn emit_dyn_to_bool_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    // Helpers carry their own imports, just like ordinary chunks. Keep the
    // value-to-i32 contract identical instead of expanding a second truthiness
    // implementation (and allocating two scratch locals) in every helper.
    emit_dyn_to_bool(code, line);
}

#[cfg(test)]
mod truthiness_into_reuse_tests {
    use super::*;

    #[test]
    fn helper_truthiness_uses_the_same_emission_as_regular_chunks() {
        let mut direct = Chunk::new("direct_truthiness");
        let mut helper = Chunk::new("helper_truthiness");
        let mut imports = Chunk::new("imports");
        emit_dyn_to_bool(&mut direct, 17);
        emit_dyn_to_bool_into(&mut imports, &mut helper, 17);
        assert_eq!(helper.code, direct.code);
        assert_eq!(helper.local_count, direct.local_count);
        assert_eq!(helper.imports.len(), direct.imports.len());
    }
}

pub fn emit_dyn_not_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_to_bool_into(_imports, code, line);
    code.emit_op(Op::I32_EQZ, line); // i32
}

pub fn emit_dyn_eq_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let slots = alloc_locals(code, 4);
    let b_slot = slots;
    let a_slot = slots + 1;
    let b_time_slot = slots + 2;
    let a_time_slot = slots + 3;

    let test_num = code.add_import("wasm:js-number", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    let test_str = code.add_import("wasm:js-string", "test");
    let str_eq = code.add_import("wasm:js-string", "equals");
    let test_bool = code.add_import("wasm:js-boolean", "test");
    let cast_bool = code.add_import("wasm:js-boolean", "cast");
    let test_bigint = code.add_import("wasm:js-bigint", "test");

    save(code, b_slot, line);
    save(code, a_slot, line);

    emit_slot_is_null_or_undefined(code, a_slot, line);
    code.emit_if_i32(line);
    emit_slot_is_null_or_undefined(code, b_slot, line);
    code.emit_else(line);
    emit_slot_is_null_or_undefined(code, b_slot, line);
    code.emit_if_i32(line);
    i32_const(code, 0, line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_num, line);
    load(code, b_slot, line);
    call1(code, test_num, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_EQ, line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_str, line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    call2(code, str_eq, line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_bool, line);
    load(code, b_slot, line);
    call1(code, test_bool, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    call1(code, cast_bool, line); // Bool → i32
    load(code, b_slot, line);
    call1(code, cast_bool, line); // Bool → i32
    code.emit_op(Op::I32_EQ, line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_bigint, line);
    load(code, b_slot, line);
    call1(code, test_bigint, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    emit_bigint_cmp(code, "eq", line);
    code.emit_else(line);
    emit_both_object_field_present(
        code,
        a_slot,
        b_slot,
        a_time_slot,
        b_time_slot,
        "Ticks",
        line,
    );
    code.emit_if_i32(line);
    load(code, a_time_slot, line);
    call1(code, to_f64, line);
    load(code, b_time_slot, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_EQ, line);
    code.emit_else(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    code.emit_op(Op::REF_EQ, line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line); // symmetric null guard
    // Result is i32 — WASM-compliant
}

/// Dynamic equality against a string literal emitted by this primitive.
/// Preserve the general fallback's field reads and reference comparison for
/// non-string values; only impossible number/Boolean/BigInt pairs are removed.
pub fn emit_dyn_eq_string_literal(chunk: &mut Chunk, literal: &str, line: u32) {
    let slots = alloc_locals(chunk, 4);
    let b = slots;
    let a = slots + 1;
    let b_time = slots + 2;
    let a_time = slots + 3;
    chunk.emit_string_const(literal, line);
    save(chunk, b, line);
    save(chunk, a, line);
    emit_slot_is_null_or_undefined(chunk, a, line);
    chunk.emit_if_i32(line);
    i32_const(chunk, 0, line);
    chunk.emit_else(line);
    load(chunk, a, line);
    let test = chunk.add_import("wasm:js-string", "test");
    call1(chunk, test, line);
    chunk.emit_if_i32(line);
    load(chunk, a, line);
    load(chunk, b, line);
    let equals = chunk.add_import("wasm:js-string", "equals");
    call2(chunk, equals, line);
    chunk.emit_else(line);
    emit_both_object_field_present(chunk, a, b, a_time, b_time, "Ticks", line);
    chunk.emit_if_i32(line);
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    load(chunk, a_time, line);
    call1(chunk, to_f64, line);
    load(chunk, b_time, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);
    load(chunk, a, line);
    load(chunk, b, line);
    chunk.emit_op(Op::REF_EQ, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Dynamic equality against an emitted i32 literal, preserving the generic
/// fallback's field reads and reference identity for non-number operands.
pub fn emit_dyn_eq_i32_literal(chunk: &mut Chunk, literal: i32, line: u32) {
    let slots = alloc_locals(chunk, 4);
    let b = slots;
    let a = slots + 1;
    let b_time = slots + 2;
    let a_time = slots + 3;
    chunk.emit_i32_const(literal, line);
    save(chunk, b, line);
    save(chunk, a, line);
    emit_slot_is_null_or_undefined(chunk, a, line);
    chunk.emit_if_i32(line);
    i32_const(chunk, 0, line);
    chunk.emit_else(line);
    load(chunk, a, line);
    let test = chunk.add_import("wasm:js-number", "test");
    call1(chunk, test, line);
    chunk.emit_if_i32(line);
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    load(chunk, a, line);
    call1(chunk, to_f64, line);
    chunk.emit_f64_const(f64::from(literal), line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);
    emit_both_object_field_present(chunk, a, b, a_time, b_time, "Ticks", line);
    chunk.emit_if_i32(line);
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    load(chunk, a_time, line);
    call1(chunk, to_f64, line);
    load(chunk, b_time, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);
    load(chunk, a, line);
    load(chunk, b, line);
    chunk.emit_op(Op::REF_EQ, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}


pub fn emit_dyn_ne_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_eq_into(_imports, code, line); // i32
    code.emit_op(Op::I32_EQZ, line); // negate: 1 if not-equal, i32
}

fn emit_dyn_cmp_into(_imports: &mut Chunk, code: &mut Chunk, line: u32, op: CmpOp) {
    let slots = alloc_locals(code, 4);
    let (b_slot, a_slot, b_time_slot, a_time_slot) = (slots, slots + 1, slots + 2, slots + 3);
    save(code, b_slot, line);
    save(code, a_slot, line);
    emit_dyn_cmp_slots(code, line, op, a_slot, b_slot, a_time_slot, b_time_slot);
}

/// Only for a compiler-private RHS local that callbacks cannot assign. Source
/// locals are not eligible: coercion may mutate captured source bindings.
fn emit_dyn_cmp_private_rhs(code: &mut Chunk, rhs: u16, line: u32, op: CmpOp) {
    assert!(rhs < code.local_count);
    let slots = alloc_locals(code, 3);
    let (a, a_time, b_time) = (slots, slots + 1, slots + 2);
    save(code, a, line);
    emit_dyn_cmp_slots(code, line, op, a, rhs, a_time, b_time);
}
pub(crate) fn emit_dyn_lt_private_rhs_into(code: &mut Chunk, rhs: u16, line: u32) {
    emit_dyn_cmp_private_rhs(code, rhs, line, CmpOp::Lt);
}
pub(crate) fn emit_dyn_gt_private_rhs_into(code: &mut Chunk, rhs: u16, line: u32) {
    emit_dyn_cmp_private_rhs(code, rhs, line, CmpOp::Gt);
}
/// Both operands are compiler-private retained values, not mutable source
/// bindings. Comparator/coercion callbacks cannot assign these local slots.
pub(crate) fn emit_dyn_gt_private_slots_into(code: &mut Chunk, lhs: u16, rhs: u16, line: u32) {
    assert!(lhs < code.local_count && rhs < code.local_count);
    #[cfg(feature = "primitive-benchmarks")]
    if !comparison_slots_reuse_enabled() {
        load(code, lhs, line);
        load(code, rhs, line);
        let mut imports = Chunk::new("imports");
        emit_dyn_gt_into(&mut imports, code, line);
        return;
    }
    let temps = alloc_locals(code, 2);
    emit_dyn_cmp_slots(code, line, CmpOp::Gt, lhs, rhs, temps, temps + 1);
}
fn emit_dyn_cmp_slots(code: &mut Chunk, line: u32, op: CmpOp, a_slot: u16, b_slot: u16, a_time_slot: u16, b_time_slot: u16) {
    let test_num = code.add_import("wasm:js-number", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    load(code, a_slot, line);
    call1(code, test_num, line);
    load(code, b_slot, line);
    call1(code, test_num, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_else(line);
    emit_dyn_cmp_fallback_slots(code, line, op, a_slot, b_slot, a_time_slot, b_time_slot);
    code.emit_end(line);
}
fn emit_dyn_cmp_fallback_slots(code: &mut Chunk, line: u32, op: CmpOp, a_slot: u16, b_slot: u16, a_time_slot: u16, b_time_slot: u16) {
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    let test_str = code.add_import("wasm:js-string", "test");
    let str_compare = code.add_import("wasm:js-string", "compare");
    let test_bool = code.add_import("wasm:js-boolean", "test");
    let cast_bool = code.add_import("wasm:js-boolean", "cast");
    let test_bigint = code.add_import("wasm:js-bigint", "test");


    load(code, a_slot, line);
    call1(code, test_str, line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    call2(code, str_compare, line);
    i32_const(code, 0, line);
    code.emit_op(i32_cmp_op(&op), line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_bigint, line);
    load(code, b_slot, line);
    call1(code, test_bigint, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    emit_bigint_cmp(code, bigint_cmp_name(&op), line);
    code.emit_else(line);
    emit_both_object_field_present(
        code,
        a_slot,
        b_slot,
        a_time_slot,
        b_time_slot,
        "Ticks",
        line,
    );
    code.emit_if_i32(line);
    load(code, a_time_slot, line);
    call1(code, to_f64, line);
    load(code, b_time_slot, line);
    call1(code, to_f64, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_else(line);
    emit_js_to_number_f64(code, a_slot, to_f64, test_bool, cast_bool, line);
    emit_js_to_number_f64(code, b_slot, to_f64, test_bool, cast_bool, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
}
/// Legacy dynamic less-than with a physically proven Number left operand.
/// Preserve fallback Ticks property reads: prototype getters can observe them.
pub fn emit_dyn_lt_number_lhs_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let slots = alloc_locals(code, 4);
    let (b, a, b_time, a_time) = (slots, slots + 1, slots + 2, slots + 3);
    save(code, b, line);
    save(code, a, line);
    let test_num = code.add_import("wasm:js-number", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    let test_bool = code.add_import("wasm:js-boolean", "test");
    let cast_bool = code.add_import("wasm:js-boolean", "cast");
    load(code, b, line);
    call1(code, test_num, line);
    code.emit_if_i32(line);
    load(code, a, line);
    call1(code, to_f64, line);
    load(code, b, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_LT, line);
    code.emit_else(line);
    // Number x non-Number cannot enter string-pair or BigInt-pair branches.
    emit_both_object_field_present(code, a, b, a_time, b_time, "Ticks", line);
    code.emit_if_i32(line);
    load(code, a_time, line);
    call1(code, to_f64, line);
    load(code, b_time, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_LT, line);
    code.emit_else(line);
    load(code, a, line);
    call1(code, to_f64, line);
    emit_js_to_number_f64(code, b, to_f64, test_bool, cast_bool, line);
    code.emit_op(Op::F64_LT, line);
    code.emit_end(line);
    code.emit_end(line);
}

pub fn emit_dyn_lt_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(_imports, code, line, CmpOp::Lt);
}
/// Legacy `value > Number(0)`, with a physically emitted zero RHS.
/// Primitive prototype Ticks getters remain observable in fallback order.
pub fn emit_dyn_gt_zero_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    #[cfg(feature = "primitive-benchmarks")]
    if !ZERO_COMPARISON_REUSE.with(|v| v.get()) {
        code.emit_i32_const(0, line);
        emit_dyn_gt_into(imports, code, line);
        return;
    }
    let slots = alloc_locals(code, 3);
    let (a, a_time, b_time) = (slots, slots + 1, slots + 2);
    save(code, a, line);
    let test_num = code.add_import("wasm:js-number", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    let test_bool = code.add_import("wasm:js-boolean", "test");
    let cast_bool = code.add_import("wasm:js-boolean", "cast");
    load(code, a, line);
    call1(code, test_num, line);
    code.emit_if_i32(line);
    load(code, a, line);
    call1(code, to_f64, line);
    code.emit_f64_const(0.0, line);
    code.emit_op(Op::F64_GT, line);
    code.emit_else(line);
    // A Number RHS excludes string/string and BigInt/BigInt branches.
    // The RHS work slot becomes its Ticks payload after its last operand read.
    code.emit_i32_const(0, line);
    save(code, b_time, line);
    emit_both_object_field_present(code, a, b_time, a_time, b_time, "Ticks", line);
    code.emit_if_i32(line);
    load(code, a_time, line);
    call1(code, to_f64, line);
    load(code, b_time, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_GT, line);
    code.emit_else(line);
    emit_js_to_number_f64(code, a, to_f64, test_bool, cast_bool, line);
    code.emit_f64_const(0.0, line);
    code.emit_op(Op::F64_GT, line);
    code.emit_end(line);
    code.emit_end(line);
}
#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static ZERO_COMPARISON_REUSE: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_zero_comparison<R>(reuse: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { ZERO_COMPARISON_REUSE.with(|v| v.set(self.0)); }
    }
    let _restore = Restore(ZERO_COMPARISON_REUSE.with(|v| v.replace(reuse)));
    build()
}

pub fn emit_dyn_gt_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(_imports, code, line, CmpOp::Gt);
}
pub fn emit_dyn_le_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(_imports, code, line, CmpOp::Le);
}
pub fn emit_dyn_ge_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(_imports, code, line, CmpOp::Ge);
}

pub fn emit_dyn_add_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_add_into_impl(code, line, false);
}

/// Same legacy addition contract, with a physically proven Number RHS.
/// Call only for numeric constants/private counters, never advisory hints.
pub fn emit_dyn_add_number_rhs_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_add_into_impl(code, line, true);
}

fn emit_dyn_add_into_impl(code: &mut Chunk, line: u32, number_rhs: bool) {
    let slots = alloc_locals(code, if number_rhs { 3 } else { 4 });
    let (b, a, a_string, b_string) = (slots, slots + 1, slots + 2, slots + 3);
    save(code, b, line);
    save(code, a, line);
    let number_test = code.add_import("wasm:js-number", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    load(code, a, line);
    call1(code, number_test, line);
    if !number_rhs {
        load(code, b, line);
        call1(code, number_test, line);
        code.emit_op(Op::I32_AND, line);
    }
    code.emit_if_value(line);
    load(code, a, line);
    load(code, b, line);
    code.emit_op(Op::F64_ADD, line);
    code.emit_else(line);
    if !number_rhs {
        let bigint_test = code.add_import("wasm:js-bigint", "test");
        load(code, a, line);
        call1(code, bigint_test, line);
        load(code, b, line);
        call1(code, bigint_test, line);
        code.emit_op(Op::I32_AND, line);
        code.emit_if_value(line);
        load(code, a, line);
        load(code, b, line);
        emit_bigint_binop(code, "add", line);
        code.emit_else(line);
    }
    // Preserve strings; normalize every other operand once, in left-to-right
    // order. Both concatenation and arithmetic consume that same number.
    let string_test = code.add_import("wasm:js-string", "test");
    let bool_test = code.add_import("wasm:js-boolean", "test");
    let bool_cast = code.add_import("wasm:js-boolean", "cast");
    for (value, is_string) in [(a, a_string), (b, b_string)] {
        if number_rhs && value == b { continue; }
        load(code, value, line);
        call1(code, string_test, line);
        code.emit_op_u16(Op::LOCAL_TEE, is_string, line);
        code.emit_op(Op::I32_EQZ, line);
        code.emit_if(line);
        emit_js_to_number_f64(code, value, to_f64, bool_test, bool_cast, line);
        save(code, value, line);
        code.emit_end(line);
    }
    load(code, a_string, line);
    if !number_rhs {
        load(code, b_string, line);
        code.emit_op(Op::I32_OR, line);
    }
    code.emit_if_value(line);
    let string_cast = code.add_import("wasm:js-string", "cast");
    let string_from_f64 = code.add_import("wasm:js-string", "fromF64");
    for (value, is_string) in [(a, a_string), (b, b_string)] {
        if number_rhs && value == b {
            load(code, value, line);
            call1(code, string_from_f64, line);
            continue;
        }
        load(code, is_string, line);
        code.emit_if_value(line);
        load(code, value, line);
        call1(code, string_cast, line);
        code.emit_else(line);
        load(code, value, line);
        call1(code, string_from_f64, line);
        code.emit_end(line);
    }
    let concat = code.add_import("wasm:js-string", "concat");
    call2(code, concat, line);
    code.emit_else(line);
    load(code, a, line);
    load(code, b, line);
    code.emit_op(Op::F64_ADD, line);
    code.emit_end(line);
    if !number_rhs { code.emit_end(line); }
    code.emit_end(line);
}

/// Accumulate into a private total initialized to Number(0). Legacy addition
/// preserves the invariant Number-or-String; once String it remains String.
/// `string_total` is an exact private i32 tag, initially0, maintained here.
/// RHS is already evaluated on stack. No advisory source type hints are used.
pub(crate) fn emit_sum_accumulate(code: &mut Chunk, total: u16, string_total: u16, line: u32) {
    assert!(total < code.local_count && string_total < code.local_count);
    let slots = alloc_locals(code, 2);
    let (b, b_string) = (slots, slots + 1);
    save(code, b, line);
    let number_test = code.add_import("wasm:js-number", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    load(code, string_total, line);
    code.emit_op(Op::I32_EQZ, line);
    load(code, b, line);
    call1(code, number_test, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_value(line);
    load(code, total, line);
    load(code, b, line);
    code.emit_op(Op::F64_ADD, line);
    code.emit_else(line);
    // A Number-or-String total cannot enter the BigInt-pair branch and never
    // needs ToNumber normalization. Preserve the original RHS normalization.
    let string_test = code.add_import("wasm:js-string", "test");
    let bool_test = code.add_import("wasm:js-boolean", "test");
    let bool_cast = code.add_import("wasm:js-boolean", "cast");
    load(code, b, line);
    call1(code, string_test, line);
    code.emit_op_u16(Op::LOCAL_TEE, b_string, line);
    code.emit_op(Op::I32_EQZ, line);
    code.emit_if(line);
    emit_js_to_number_f64(code, b, to_f64, bool_test, bool_cast, line);
    save(code, b, line);
    code.emit_end(line);
    load(code, string_total, line);
    load(code, b_string, line);
    code.emit_op(Op::I32_OR, line);
    code.emit_if_value(line);
    let string_cast = code.add_import("wasm:js-string", "cast");
    let string_from_f64 = code.add_import("wasm:js-string", "fromF64");
    for (value, is_string) in [(total, string_total), (b, b_string)] {
        load(code, is_string, line);
        code.emit_if_value(line);
        load(code, value, line);
        call1(code, string_cast, line);
        code.emit_else(line);
        load(code, value, line);
        call1(code, string_from_f64, line);
        code.emit_end(line);
    }
    let concat = code.add_import("wasm:js-string", "concat");
    call2(code, concat, line);
    // Update only after successful concatenation; leave its result on stack.
    code.emit_i32_const(1, line);
    save(code, string_total, line);
    code.emit_else(line);
    load(code, total, line);
    load(code, b, line);
    code.emit_op(Op::F64_ADD, line);
    code.emit_end(line);
    code.emit_end(line);
}

pub fn emit_dyn_neg_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let v = alloc_locals(code, 1);

    let test_bigint = code.add_import("wasm:js-bigint", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    let from_f64 = code.add_import("wasm:js-number", "fromF64");
    let test_bool = code.add_import("wasm:js-boolean", "test");
    let cast_bool = code.add_import("wasm:js-boolean", "cast");

    save(code, v, line);

    load(code, v, line);
    call1(code, test_bigint, line);
    code.emit_if_value(line);
    load(code, v, line);
    emit_bigint_neg(code, line);
    code.emit_else(line);
    emit_js_to_number_f64(code, v, to_f64, test_bool, cast_bool, line);
    code.emit_op(Op::F64_NEG, line);
    call1(code, from_f64, line);
    code.emit_end(line);
}

// ── Linkable chunk builders ──────────────────────────────────────────────────

/// The type-dispatch body is identical at every use. Link it once per module
/// through the existing helper mechanism, retaining the public i32 result.
pub(crate) fn emit_comparison_helper_call(chunk: &mut Chunk, global: &str, line: u32) {
    let base = alloc_locals(chunk, 2);
    save(chunk, base + 1, line);
    save(chunk, base, line);
    emit_comparison_helper_call_from_slots_impl(chunk, global, base, base + 1, line);
}

/// Operands already reside in caller locals. Reading locals cannot invoke user
/// code; keep the global lookup and invocation at the original call point.
pub(crate) fn emit_comparison_helper_call_from_slots(
    chunk: &mut Chunk, global: &str, lhs: u16, rhs: u16, line: u32,
) {
    #[cfg(feature = "primitive-benchmarks")]
    if !comparison_slots_reuse_enabled() {
        load(chunk, lhs, line);
        load(chunk, rhs, line);
        emit_comparison_helper_call(chunk, global, line);
        return;
    }
    emit_comparison_helper_call_from_slots_impl(chunk, global, lhs, rhs, line);
}

fn emit_comparison_helper_call_from_slots_impl(
    chunk: &mut Chunk, global: &str, lhs: u16, rhs: u16, line: u32,
) {
    assert!(lhs < chunk.local_count && rhs < chunk.local_count);
    crate::primitives::globals::emit_read(chunk, global, line);
    load(chunk, lhs, line);
    load(chunk, rhs, line);
    crate::primitives::callable::emit_direct_invoke_chunk(chunk, 2, line);
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");
    call1(chunk, to_i32, line);
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static COMPARISON_SLOTS_REUSE: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub(crate) fn comparison_slots_reuse_enabled() -> bool {
    COMPARISON_SLOTS_REUSE.with(|v| v.get())
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_comparison_slots<R>(reuse: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { COMPARISON_SLOTS_REUSE.with(|v| v.set(self.0)); }
    }
    let _restore = Restore(COMPARISON_SLOTS_REUSE.with(|v| v.replace(reuse)));
    build()
}

pub fn build_dyn_comparison(name: &str) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = 2;
    chunk.local_count = 2;
    if name == "__stdlib_dyneq" {
        // Parameters already occupy locals 0/1. Only wrapper payloads need
        // scratch space; avoid loading and spilling both operands on entry.
        let temps = alloc_locals(&mut chunk, 2);
        emit_dyn_eq_slots(&mut chunk, 0, 1, temps, temps + 1, 0);
        chunk.emit_op(Op::RETURN, 0);
        return chunk;
    }
    load(&mut chunk, 0, 0);
    load(&mut chunk, 1, 0);
    match name {
        "__stdlib_dynlt" => emit_dyn_cmp_inline(&mut chunk, 0, CmpOp::Lt),
        "__stdlib_dyngt" => emit_dyn_cmp_inline(&mut chunk, 0, CmpOp::Gt),
        "__stdlib_dynle" => emit_dyn_cmp_inline(&mut chunk, 0, CmpOp::Le),
        "__stdlib_dynge" => emit_dyn_cmp_inline(&mut chunk, 0, CmpOp::Ge),
        _ => unreachable!("unknown comparison helper"),
    }
    chunk.emit_op(Op::RETURN, 0);
    chunk
}
//
// Linkable chunk builders — the standalone-chunk packaging of what the
// `emit_*` forms splice inline. A language prefix in a name records which
// frontend first needed a linkable chunk, not a language-specific meaning.

// ── dynMul(a, b) → string repeat or numeric multiply ─────────
pub fn build_dyn_mul() -> Chunk {
    let mut c = Chunk::new("__stdlib_dynmul");
    c.arity = 2;
    c.local_count = 2;
    let str_tag = c.add_constant(Value::String(Arc::from("string")));
    // if typeof(a) == "string": return str_repeat(a, b)
    let a_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::primitives::reflection::emit_typeof_in_chunk(&mut c, 0);
    crate::primitives::expressions::emit_const_index(&mut c, str_tag, 0);
    crate::primitives::strings::emit_str_equals(&mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0); // skip if a is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    {
        let idx = c.add_import("ecma:string", "repeat");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(a_block_p);
    // if typeof(b) == "string": return str_repeat(b, a)
    let b_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    crate::primitives::reflection::emit_typeof_in_chunk(&mut c, 0);
    crate::primitives::expressions::emit_const_index(&mut c, str_tag, 0);
    crate::primitives::strings::emit_str_equals(&mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0); // skip if b is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    {
        let idx = c.add_import("ecma:string", "repeat");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(b_block_p);
    // numeric
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
