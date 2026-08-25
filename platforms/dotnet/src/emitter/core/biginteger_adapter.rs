//! `System.Numerics.BigInteger` — a CLASS whose payload is a native BigInt.
//!
//! ⛔ **It is a class, not a primitive spelling.** `BigInteger` is a .NET type
//! with members and operators, so it is registered in the tree and its value is
//! an OBJECT — `{ __type: "BigInteger", __bi: <native BigInt> }` — exactly the
//! shape `TimeSpan` uses. That object is what carries the protocol slots, which
//! is what makes `a * b` work: `emit_rich_binop` reads
//! `protocol_slot_key(Mul)` off the LEFT OPERAND, and a bare `Value::BigInt`
//! has no property bag and no prototype to read it from.
//!
//! The payload stays a NATIVE ECMA BigInt (`Value::BigInt`, arbitrary
//! precision), so the arithmetic underneath is exact — `ecma:bigint` registers
//! `add sub mul div rem and or xor` through its `binop!`/`divlike!` macros,
//! which is why a literal grep of that file finds only `neg`/`pow`/`shl`.
//!
//! Nothing shared changed for this: the slot vocabulary, `emit_bind_method`
//! and the `emit_rich_binop` reader already existed.

use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use std::sync::Arc;

const TYPE_KEY: &str = "__type";
const PAYLOAD: &str = "__bi";

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_set(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, idx, line);
}

/// `[obj] → [bigint]` — the payload, accepting a bare BigInt or a Number too so
/// a literal (`Dim b As BigInteger = -999`) works wherever a value is expected.
fn unwrap_payload(chunk: &mut Chunk, line: u32) {
    let v = chunk.alloc_scratch(1);
    set(chunk, v, line);
    get(chunk, v, line);
    call(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("object", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, v, line);
    let k = chunk.add_constant(Value::String(Arc::from(PAYLOAD)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    chunk.emit_else(line);
    get(chunk, v, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    chunk.emit_end(line);
}

/// `[value] → [bigint]` — the boundary coercion every entry point makes.
fn to_bigint(chunk: &mut Chunk, line: u32) {
    unwrap_payload(chunk, line);
}

/// `[bigint] → [BigInteger]` — wrap a payload as the class value, slots bound.
fn wrap(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let payload = chunk.alloc_scratch(2);
    let obj = payload + 1;
    set(chunk, payload, line);
    chunk.emit_struct_new(0, 0, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    chunk.emit_string_const("BigInteger", line);
    field_set(chunk, TYPE_KEY, line);
    get(chunk, obj, line);
    get(chunk, payload, line);
    field_set(chunk, PAYLOAD, line);

    // ⛔ The predicates are stored as FIELDS, not bound as methods. .NET spells
    // them as PROPERTIES (`v.IsZero`, no parentheses), and a bound method read
    // as a property yields the funcref rather than the answer. A BigInteger
    // value is immutable — every operator mints a new object — so computing
    // them once at construction is correct, not a cache that can go stale.
    emit_predicate_fields(chunks, current, obj, payload, line);

    bind_operator_slots(chunks, current, obj, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
}

/// `IsZero` / `IsOne` / `IsEven` / `Sign`, computed from the payload.
fn emit_predicate_fields(
    chunks: &mut Vec<Chunk>,
    current: usize,
    obj: u16,
    payload: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let zero = chunk.alloc_scratch(1);
    core_wasm::i32_const(chunk, line, 0);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    set(chunk, zero, line);

    for (key, rhs) in [("IsZero", 0i32), ("IsOne", 1)] {
        get(chunk, obj, line);
        get(chunk, payload, line);
        core_wasm::i32_const(chunk, line, rhs);
        call(chunk, "ecma:bigint", "BigInt", 1, line);
        ops::emit_dyn_eq(chunk, line);
        ops::emit_i32_to_bool(chunk, line);
        field_set(chunk, key, line);
    }

    // IsEven — the LOW BIT. `asUintN(1, v)` reads two's-complement bits, so it
    // is right for negatives too, and there is no modulo to reach for.
    get(chunk, obj, line);
    core_wasm::i32_const(chunk, line, 1);
    get(chunk, payload, line);
    call(chunk, "ecma:bigint", "asUintN", 2, line);
    core_wasm::i32_const(chunk, line, 0);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
    field_set(chunk, "IsEven", line);

    // Sign — an ordinary Number, which is what .NET returns.
    get(chunk, obj, line);
    get(chunk, payload, line);
    get(chunk, zero, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    core_wasm::i32_const(chunk, line, -1);
    chunk.emit_else(line);
    get(chunk, payload, line);
    get(chunk, zero, line);
    ops::emit_dyn_gt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_else(line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_end(line);
    chunk.emit_end(line);
    field_set(chunk, "Sign", line);
}

/// The five operator chunks, created together and memoised by name.
///
/// ⛔ TWO PHASES, and it has to be. Each chunk's RESULT must itself carry the
/// slots, or `(a + b) * c` reads a slotless object and silently falls back to
/// the Number path — losing every digit past 2^53, which is the one thing a
/// big integer exists to prevent. A chunk cannot bind indices that do not
/// exist yet, so all five are created first and only then does each one get
/// the binds and its `RETURN`.
const BIGINT_OPS: [(&str, &str, vybe_ast::ProtocolSlot); 6] = [
    ("__bigint_add", "add", vybe_ast::ProtocolSlot::Add),
    ("__bigint_sub", "sub", vybe_ast::ProtocolSlot::Sub),
    ("__bigint_mul", "mul", vybe_ast::ProtocolSlot::Mul),
    ("__bigint_div", "div", vybe_ast::ProtocolSlot::Div),
    ("__bigint_rem", "rem", vybe_ast::ProtocolSlot::Mod),
    // ⛔ `\` reads `IDiv`, `/` reads `Div` — two slots, one operation here:
    // `ecma:bigint.div` TRUNCATES (`divlike!("div", |q, _r| q)`), which is what
    // .NET's `BigInteger` division and VB's `\` both mean. Binding only `Div`
    // left `a \ b` answering NaN while `a / b` was exact.
    ("__bigint_idiv", "div", vybe_ast::ProtocolSlot::IDiv),
];

fn ensure_operator_chunks(chunks: &mut Vec<Chunk>, line: u32) -> [usize; 6] {
    if let Some(first) = chunks.iter().position(|c| c.name == BIGINT_OPS[0].0) {
        let mut found = [first; 6];
        for (i, (name, _, _)) in BIGINT_OPS.iter().enumerate() {
            found[i] = chunks
                .iter()
                .position(|c| c.name == *name)
                .expect("operator chunks are created together");
        }
        return found;
    }

    // Phase 1 — arithmetic and the result object, no binds, no return.
    let mut idxs = [0usize; 6];
    let mut obj_slots = [0u16; 6];
    for (i, (name, op, _)) in BIGINT_OPS.iter().enumerate() {
        let mut method = create_function_chunk(name, 2);
        method.local_count = 2;
        get(&mut method, 0, line);
        unwrap_payload(&mut method, line);
        get(&mut method, 1, line);
        unwrap_payload(&mut method, line);
        call(&mut method, "ecma:bigint", op, 2, line);
        let payload = method.alloc_scratch(2);
        let obj = payload + 1;
        set(&mut method, payload, line);
        method.emit_struct_new(0, 0, line);
        set(&mut method, obj, line);
        get(&mut method, obj, line);
        method.emit_string_const("BigInteger", line);
        field_set(&mut method, TYPE_KEY, line);
        get(&mut method, obj, line);
        get(&mut method, payload, line);
        field_set(&mut method, PAYLOAD, line);
        chunks.push(method);
        idxs[i] = chunks.len() - 1;
        obj_slots[i] = obj;
    }

    // Phase 2 — now every index exists, so each result can carry the slots.
    for i in 0..BIGINT_OPS.len() {
        let target = idxs[i];
        let obj = obj_slots[i];
        for (j, (_, _, slot)) in BIGINT_OPS.iter().enumerate() {
            let method_idx = idxs[j];
            vybe_compiler::primitives::object::emit_bind_method(
                &mut chunks[target],
                obj,
                &vybe_ast::protocol_slot_key(*slot),
                method_idx,
                line,
            );
        }
        // ToString too — an operator RESULT is exactly the value with no
        // declared type, so it is the one that most needs the slot.
        let ts = to_string_chunk(chunks, line);
        vybe_compiler::primitives::object::emit_bind_method(
            &mut chunks[target],
            obj,
            &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
            ts,
            line,
        );
        vybe_compiler::primitives::object::emit_bind_method(
            &mut chunks[target],
            obj,
            "ToString",
            ts,
            line,
        );
        get(&mut chunks[target], obj, line);
        chunks[target].emit_op(Op::RETURN, line);
    }
    idxs
}

/// The `ToString` slot — one unary chunk, memoised.
///
/// ⛔ Needed because the RESULT of an operator has no declared type: `(a + b)`
/// is a `BigInteger` object the tree cannot name, so `.ToString()` fell through
/// to the generic object one and printed `[object BigInteger]`. Publishing the
/// slot on the value answers it wherever the value goes.
fn to_string_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__bigint_to_string";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut method = create_function_chunk(NAME, 1);
    method.local_count = 1;
    get(&mut method, 0, line);
    unwrap_payload(&mut method, line);
    call(&mut method, "ecma:bigint", "toString", 1, line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

/// Publish `Add`/`Sub`/`Mul`/`Div`/`Mod` on the value in `obj_slot`.
fn bind_operator_slots(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let idxs = ensure_operator_chunks(chunks, line);
    for (i, (_, _, slot)) in BIGINT_OPS.iter().enumerate() {
        vybe_compiler::primitives::object::emit_bind_method(
            &mut chunks[current],
            obj_slot,
            &vybe_ast::protocol_slot_key(*slot),
            idxs[i],
            line,
        );
    }
    let ts = to_string_chunk(chunks, line);
    vybe_compiler::primitives::object::emit_bind_method(
        &mut chunks[current],
        obj_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        ts,
        line,
    );
    // The .NET spelling too: `x.ToString()` is a member call, not the shared
    // stringification slot, and both have to answer.
    vybe_compiler::primitives::object::emit_bind_method(
        &mut chunks[current],
        obj_slot,
        "ToString",
        ts,
        line,
    );
}


/// Pop `argc` operands into consecutive slots, each unwrapped to its payload.
/// Returns the base slot; operand `i` is at `base + i`.
fn stash_bigints(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        unwrap_payload(chunk, line);
        set(chunk, base + offset, line);
    }
    base
}

// ── one-operand statics ─────────────────────────────────────────────────────

/// `BigInteger.Abs(v)` — `v < 0 ? -v : v`, in big-integer arithmetic.
pub fn emit_abs(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    get(chunk, v, line);
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, v, line);
    call(chunk, "ecma:bigint", "neg", 1, line);
    chunk.emit_else(line);
    get(chunk, v, line);
    chunk.emit_end(line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

pub fn emit_negate(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    get(chunk, v, line);
    call(chunk, "ecma:bigint", "neg", 1, line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

/// `BigInteger.Parse(s)` — `BigInt` already parses a decimal string, and it
/// THROWS on a malformed one, which is what `Parse` owes its caller.
pub fn emit_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    to_bigint(&mut chunks[current], line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

/// `v.ToString()` — the decimal spelling. `ecma:bigint.toString` is exact for
/// every magnitude, where a `Number` round-trip would lose digits past 2^53.
pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    get(chunk, v, line);
    call(chunk, "ecma:bigint", "toString", 1, line);
}

/// `v.Sign` — −1, 0 or 1 as an ordinary Number, which is what .NET returns.
pub fn emit_sign(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    let zero = chunk.alloc_scratch(1);
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    set(chunk, zero, line);

    get(chunk, v, line);
    get(chunk, zero, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    core_wasm::i32_const(chunk, line, -1);
    chunk.emit_else(line);
    get(chunk, v, line);
    get(chunk, zero, line);
    ops::emit_dyn_gt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_else(line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_compare_to_const(chunks: &mut [Chunk], current: usize, rhs: i32, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    get(chunk, v, line);
    core_wasm::i32_const(chunk, line, rhs);
    to_bigint(chunk, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_is_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_compare_to_const(chunks, current, 0, line);
}

pub fn emit_is_one(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_compare_to_const(chunks, current, 1, line);
}

/// `v.IsEven` — the LOW BIT, read with `asUintN(1, v)`.
///
/// ⛔ Not `v % 2 == 0`: `ecma:bigint` exposes no binary arithmetic (only `neg`,
/// `not`, `pow`, `shl`, `shr` and the `asIntN`/`asUintN` width ops), so there
/// is no modulo to call. The low bit answers the same question exactly, for
/// negative values too, because `asUintN` reads the two's-complement bits.
pub fn emit_is_even(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    core_wasm::i32_const(chunk, line, 1);
    get(chunk, v, line);
    call(chunk, "ecma:bigint", "asUintN", 2, line);
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

// ── two-operand statics ─────────────────────────────────────────────────────

pub fn emit_pow(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = stash_bigints(chunks, current, 2, line);
    let chunk = &mut chunks[current];
    get(chunk, base, line);
    get(chunk, base + 1, line);
    call(chunk, "ecma:bigint", "pow", 2, line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

/// `[a, b] → [a op b]` on two BigInts.
///
/// ⛔ `ecma:bigint` registers these through the `binop!`/`divlike!` MACROS, so
/// a literal grep for `"mul"` in that file finds nothing and the surface looks
/// like it only has `neg`/`pow`/`shl`. It has `add sub mul div rem and or xor`.
fn bigop(chunk: &mut Chunk, name: &str, a: u16, b: u16, line: u32) {
    get(chunk, a, line);
    get(chunk, b, line);
    call(chunk, "ecma:bigint", name, 2, line);
}

/// `BigInteger.ModPow(value, exponent, modulus)`.
///
/// ⛔ SQUARE-AND-MULTIPLY, reducing at every step — not `pow(v, e) % m`. The
/// direct form materialises the whole power first, and for the exponents these
/// tests use that is a number with millions of digits: not slow, hung.
pub fn emit_mod_pow(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = stash_bigints(chunks, current, 3, line);
    let value = base;
    let exp = base + 1;
    let modulus = base + 2;
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(3);
    let result = scratch;
    let two = scratch + 1;
    let zero = scratch + 2;

    core_wasm::i32_const(chunk, line, 1);
    to_bigint(chunk, line);
    set(chunk, result, line);
    core_wasm::i32_const(chunk, line, 2);
    to_bigint(chunk, line);
    set(chunk, two, line);
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    set(chunk, zero, line);

    bigop(chunk, "rem", value, modulus, line);
    set(chunk, value, line);

    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, exp, line);
    get(chunk, zero, line);
    ops::emit_dyn_gt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);

    // odd exponent → fold `value` into the result
    core_wasm::i32_const(chunk, line, 1);
    get(chunk, exp, line);
    call(chunk, "ecma:bigint", "asUintN", 2, line);
    core_wasm::i32_const(chunk, line, 1);
    to_bigint(chunk, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    bigop(chunk, "mul", result, value, line);
    set(chunk, result, line);
    bigop(chunk, "rem", result, modulus, line);
    set(chunk, result, line);
    chunk.emit_end(line);

    bigop(chunk, "mul", value, value, line);
    set(chunk, value, line);
    bigop(chunk, "rem", value, modulus, line);
    set(chunk, value, line);
    bigop(chunk, "div", exp, two, line);
    set(chunk, exp, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    get(chunk, result, line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

/// `BigInteger.GreatestCommonDivisor(a, b)` — Euclid on absolute values, which
/// is what .NET returns: a GCD is never negative.
pub fn emit_gcd(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = stash_bigints(chunks, current, 2, line);
    let a = base;
    let b = base + 1;
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(2);
    let tmp = scratch;
    let zero = scratch + 1;
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    set(chunk, zero, line);

    for slot in [a, b] {
        get(chunk, slot, line);
        get(chunk, zero, line);
        ops::emit_dyn_lt(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        get(chunk, slot, line);
        call(chunk, "ecma:bigint", "neg", 1, line);
        set(chunk, slot, line);
        chunk.emit_end(line);
    }

    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, b, line);
    get(chunk, zero, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, b, line);
    set(chunk, tmp, line);
    bigop(chunk, "rem", a, b, line);
    set(chunk, b, line);
    get(chunk, tmp, line);
    set(chunk, a, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    get(chunk, a, line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

fn emit_pick(chunks: &mut Vec<Chunk>, current: usize, want_smaller: bool, line: u32) {
    let base = stash_bigints(chunks, current, 2, line);
    let chunk = &mut chunks[current];
    get(chunk, base, line);
    get(chunk, base + 1, line);
    if want_smaller {
        ops::emit_dyn_lt(chunk, line);
    } else {
        ops::emit_dyn_gt(chunk, line);
    }
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, base, line);
    chunk.emit_else(line);
    get(chunk, base + 1, line);
    chunk.emit_end(line);
}

pub fn emit_min(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_pick(chunks, current, true, line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

pub fn emit_max(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_pick(chunks, current, false, line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

/// `a.CompareTo(b)` / `BigInteger.Compare(a, b)` → −1, 0 or 1.
pub fn emit_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_bigints(chunks, current, 2, line);
    let chunk = &mut chunks[current];
    get(chunk, base, line);
    get(chunk, base + 1, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    core_wasm::i32_const(chunk, line, -1);
    chunk.emit_else(line);
    get(chunk, base, line);
    get(chunk, base + 1, line);
    ops::emit_dyn_gt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_else(line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_bigints(chunks, current, 2, line);
    let chunk = &mut chunks[current];
    get(chunk, base, line);
    get(chunk, base + 1, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

/// The declared constants. Each is a fresh BigInt rather than a Number, so
/// `BigInteger.Zero + x` stays big-integer arithmetic.
fn emit_const(chunks: &mut Vec<Chunk>, current: usize, value: i32, line: u32) {
    let chunk = &mut chunks[current];
    core_wasm::i32_const(chunk, line, value);
    to_bigint(chunk, line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

pub fn emit_zero(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_const(chunks, current, 0, line);
}

pub fn emit_one(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_const(chunks, current, 1, line);
}

pub fn emit_minus_one(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_const(chunks, current, -1, line);
}
