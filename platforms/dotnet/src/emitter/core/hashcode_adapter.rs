//! `System.HashCode` — .NET's xxHash32-based combiner, bytecode-only.
//!
//! # One algorithm, so the contract holds by construction
//!
//! .NET guarantees that `HashCode.Combine(a, b)` equals `hc.Add(a); hc.Add(b);
//! hc.ToHashCode()` — verified against the SDK for every arity 1..8 before this
//! was written. The obvious implementation gives `Combine` a closed-form
//! unrolled fold and `Add`/`ToHashCode` an incremental state machine, and then
//! that equality is a COINCIDENCE that has to be re-proved after every edit.
//!
//! Here there is exactly one fold. `Add` appends a per-value hash to an array
//! on the instance; `Combine` builds the same array from its arguments. Both
//! then run [`emit_fold`]. The equality the corpus asserts is structural.
//!
//! # The seed is fixed, and that is a deliberate divergence
//!
//! .NET randomizes `s_seed` per PROCESS, which is why its own documentation
//! refuses to guarantee a value across runs — and why the corpus tests assert
//! only relational properties (stable within a run, order-sensitive, `Add`
//! agrees with `Combine`). We seed with a constant, so our values are stable
//! ACROSS runs too. That is strictly more determinism than .NET promises.
//!
//! ⛔ **No test may ever pin a literal hash value.** A test that does would be
//! asserting something real .NET explicitly does not guarantee, and it would
//! freeze the seed as an observable.
//!
//! # Per-value hashing honours a user's override
//!
//! `HashCode.Combine(x)` calls `x.GetHashCode()`, so a class that overrides it
//! decides its own contribution — measured against the SDK: two instances with
//! equal `GetHashCode` but different `ToString` combine EQUAL, and the converse
//! combines DIFFERENT. Hashing by stringifying would get both backwards, and
//! would still pass all four corpus tests, because every assertion in them
//! compares two values that travel the same path. See [`emit_value_hash`].

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::{collections, errors, object, ops};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use super::object_fields::field_slot;

/// The queued per-value hash codes. `ToHashCode` folds exactly this.
const VALUES_KEY: &str = "__hc_values";
const TYPE_KEY: &str = "__type";

// xxHash32's primes, as the i32 bit patterns of .NET's `uint` constants.
const PRIME1: i32 = -1640531535; // 2654435761
const PRIME2: i32 = -2048144777; // 2246822519
const PRIME3: i32 = -1028477379; // 3266489917
const PRIME4: i32 = 668265263;
const PRIME5: i32 = 374761393;

/// .NET's process-random `s_seed`, fixed here — see the module header.
const SEED: i32 = 0;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// `slot = rotl(slot + <top of stack> * mul, rot) * post` — xxHash32's `Round`
/// and `QueueRound` differ only in their four constants.
fn mix_into(chunk: &mut Chunk, slot: u16, mul: i32, rot: i32, post: i32, line: u32) {
    chunk.emit_i32_const(mul, line);
    chunk.emit_op(Op::I32_MUL, line);
    get(chunk, slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_i32_const(rot, line);
    chunk.emit_op(Op::I32_ROTL, line);
    chunk.emit_i32_const(post, line);
    chunk.emit_op(Op::I32_MUL, line);
    set(chunk, slot, line);
}

/// `hash ^= hash >>> shift; hash *= mul` — one step of `MixFinal`.
fn mix_final_step(chunk: &mut Chunk, hash: u16, shift: i32, mul: Option<i32>, line: u32) {
    get(chunk, hash, line);
    get(chunk, hash, line);
    chunk.emit_i32_const(shift, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_op(Op::I32_XOR, line);
    if let Some(mul) = mul {
        chunk.emit_i32_const(mul, line);
        chunk.emit_op(Op::I32_MUL, line);
    }
    set(chunk, hash, line);
}

/// `values[index + offset]` onto the stack.
fn element(
    chunks: &mut [Chunk],
    current: usize,
    values: u16,
    index: u16,
    offset: i32,
    line: u32,
) {
    get(&mut chunks[current], values, line);
    get(&mut chunks[current], index, line);
    if offset != 0 {
        chunks[current].emit_i32_const(offset, line);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    collections::emit_get(chunks, current, line);
}

/// `value?.GetHashCode() ?? 0`. Stack `[value] -> [i32]`.
///
/// Reads the bound `Hash` ROLE, exactly as `runtime_adapter`'s `ToString`
/// dispatch reads the `ToString` role — `ClassSlot::Slot` is the one place a
/// binding becomes a storage name, and a language binds a slot rather than
/// naming one. `languages/csharp/src/protocol.rs` already maps `GetHashCode`
/// onto `ProtocolSlot::Hash`, so an override is bound and this is its reader.
///
/// ⛔ THREE NAME-KEYED LOOKUPS WERE TRIED HERE FIRST AND ALL THREE FOUND
/// NOTHING: `collections::emit_get` (`ARRAY_GET`, own properties only),
/// `ecma:object.get` (§7.3.2 `GetV`, walks the prototype chain), and
/// `invoke::emit_invoke_method` called unconditionally. A method is not
/// reachable under its SPELLING here; it is reachable under its ROLE. Worth
/// keeping because the failure was silent in the discriminating direction —
/// the fallback hashes two instances of one class IDENTICALLY (`String(obj)`
/// is `[object Plain]` for both), so "equal hash, different ToString" passed
/// for entirely the wrong reason while "different hash, same ToString" failed.
///
/// A value with no `Hash` role — every number and string — falls through to
/// the shared `object::emit_hash_code`, which is also what
/// `comparer.GetHashCode(x)` answers, so the two agree on a value.
pub fn emit_value_hash(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let (value, func, result, handled, ty, is_num) =
        (base, base + 1, base + 2, base + 3, base + 4, base + 5);
    let hash_role = class_slots::resolve(
        &class_slots::ClassSlot::Slot(vybe_ast::ProtocolSlot::Hash),
        &class_slots::PlainNames,
    );
    set(&mut chunks[current], value, line);

    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], handled, line);

    let typeof_fn = chunks[current].add_import("ecma:value", "typeof");
    get(&mut chunks[current], value, line);
    chunks[current].emit_call(typeof_fn, 1, line);
    set(&mut chunks[current], ty, line);

    // null / undefined — .NET's `value?.GetHashCode() ?? 0`.
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    set(&mut chunks[current], handled, line);
    chunks[current].emit_end(line);

    // `bool.GetHashCode()` is 1 / 0 in .NET, measured on the SDK.
    unhandled(&mut chunks[current], handled, line);
    type_is(&mut chunks[current], ty, "boolean", line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_i32_const(1, line);
    set(&mut chunks[current], handled, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    // `int.GetHashCode()` IS the int in .NET — measured, including the sign
    // (`(-5).GetHashCode()` is `-5`). Restricted to values that are integral
    // AND fit i32: a `long` hashes by folding its halves and a `double` by its
    // bits, neither of which this branch computes, so both take the structural
    // path below rather than a truncation that would collide silently.
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], is_num, line);
    for name in ["number", "i32", "i64"] {
        get(&mut chunks[current], is_num, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        type_is(&mut chunks[current], ty, name, line);
        chunks[current].emit_if(line);
        chunks[current].emit_i32_const(1, line);
        set(&mut chunks[current], is_num, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
    unhandled(&mut chunks[current], handled, line);
    get(&mut chunks[current], is_num, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(i32::MIN as f64, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(i32::MAX as f64, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_i32_const(1, line);
    set(&mut chunks[current], handled, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    // An object's own `GetHashCode` override, read as the bound ROLE.
    unhandled(&mut chunks[current], handled, line);
    type_is(&mut chunks[current], ty, "object", line);
    chunks[current].emit_if(line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(value),
        &hash_role,
        Dest::Local(func),
        line,
    );
    get(&mut chunks[current], func, line);
    host::emit(&mut chunks[current], "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], func, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_i32_const(1, line);
    set(&mut chunks[current], handled, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    // Strings, non-integral numbers, objects with no override.
    unhandled(&mut chunks[current], handled, line);
    get(&mut chunks[current], value, line);
    object::emit_hash_code(&mut chunks[current], line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], result, line);
}

/// `if handled == 0 {` — every branch above is a fallthrough guard, and an
/// `else if` chain of them would nest six deep.
fn unhandled(chunk: &mut Chunk, handled: u16, line: u32) {
    get(chunk, handled, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
}

/// `typeof(value) == name`, as an i32 condition.
fn type_is(chunk: &mut Chunk, ty: u16, name: &str, line: u32) {
    get(chunk, ty, line);
    chunk.emit_string_const(name, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
}

/// .NET's `ToHashCode` over an array of already-hashed values.
/// Stack `[array] -> [i32]`.
///
/// Straight from the framework: four-at-a-time `Round`s into `v1..v4`, then
/// `MixState` (or `MixEmptyState` when fewer than four values ever arrived),
/// then `+= length * 4`, then a `QueueRound` per leftover, then `MixFinal`.
fn emit_fold(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let (values, n, i, v1, v2, v3, v4, hash) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
    );

    set(&mut chunks[current], values, line);
    get(&mut chunks[current], values, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], n, line);

    // Initialize(out v1, out v2, out v3, out v4).
    for (slot, init) in [
        (v1, SEED.wrapping_add(PRIME1).wrapping_add(PRIME2)),
        (v2, SEED.wrapping_add(PRIME2)),
        (v3, SEED),
        (v4, SEED.wrapping_sub(PRIME1)),
    ] {
        chunks[current].emit_i32_const(init, line);
        set(&mut chunks[current], slot, line);
    }
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], i, line);

    // while i + 4 <= n: consume one four-value block.
    let outer = chunks[current].emit_block(line);
    let (block_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_br_if(1, line);
    for (offset, slot) in [(0, v1), (1, v2), (2, v3), (3, v4)] {
        element(chunks, current, values, i, offset, line);
        mix_into(&mut chunks[current], slot, PRIME2, 13, PRIME1, line);
    }
    get(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(block_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    // MixEmptyState() when no block ever ran, MixState(v1..v4) otherwise.
    get(&mut chunks[current], n, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(SEED.wrapping_add(PRIME5), line);
    chunks[current].emit_else(line);
    for (idx, (slot, rot)) in [(v1, 1), (v2, 7), (v3, 12), (v4, 18)].iter().enumerate() {
        get(&mut chunks[current], *slot, line);
        chunks[current].emit_i32_const(*rot, line);
        chunks[current].emit_op(Op::I32_ROTL, line);
        if idx > 0 {
            chunks[current].emit_op(Op::I32_ADD, line);
        }
    }
    chunks[current].emit_end(line);
    set(&mut chunks[current], hash, line);

    // hash += length * 4 — the BYTE count, four per value.
    get(&mut chunks[current], hash, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], hash, line);

    // One QueueRound per value the blocks left behind.
    let tail = chunks[current].emit_block(line);
    let (tail_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    element(chunks, current, values, i, 0, line);
    mix_into(&mut chunks[current], hash, PRIME3, 17, PRIME4, line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(tail_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(tail);

    // MixFinal.
    mix_final_step(&mut chunks[current], hash, 15, Some(PRIME2), line);
    mix_final_step(&mut chunks[current], hash, 13, Some(PRIME3), line);
    mix_final_step(&mut chunks[current], hash, 16, None, line);
    get(&mut chunks[current], hash, line);
}

/// `new HashCode()` — an empty queue.
pub fn emit_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    set(&mut chunks[current], obj, line);

    chunks[current].emit_string_const("HashCode", line);
    let tag = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tag, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Local(obj),
        &field_slot(TYPE_KEY),
        ValueSource::Local(tag),
        line,
    );

    collections::emit_array_new(chunks, current, 0, line);
    let empty = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], empty, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Local(obj),
        &field_slot(VALUES_KEY),
        ValueSource::Local(empty),
        line,
    );

    get(&mut chunks[current], obj, line);
}

/// `hc.Add(value)` — queue the value's own hash code. Returns null (void).
pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (recv, value) = (base, base + 1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], recv, line);

    get(&mut chunks[current], value, line);
    emit_value_hash(chunks, current, line);
    let hashed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], hashed, line);

    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(recv),
        &field_slot(VALUES_KEY),
        Dest::Stack,
        line,
    );
    get(&mut chunks[current], hashed, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `hc.Add(value, comparer)` — the comparer decides the value's hash.
///
/// Routed through `comparer_adapter::emit_get_hash_code`, the same helper
/// `comparer.GetHashCode(x)` uses, so an ignore-case comparer hashes `"A"` and
/// `"a"` alike here too — .NET's rule that equal values hash equally.
pub fn emit_add_with_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let (recv, value, comparer) = (base, base + 1, base + 2);
    set(&mut chunks[current], comparer, line);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], recv, line);

    get(&mut chunks[current], comparer, line);
    get(&mut chunks[current], value, line);
    super::comparer_adapter::emit_get_hash_code(chunks, current, line);
    let hashed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], hashed, line);

    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(recv),
        &field_slot(VALUES_KEY),
        Dest::Stack,
        line,
    );
    get(&mut chunks[current], hashed, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `hc.ToHashCode()`.
pub fn emit_to_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], recv, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(recv),
        &field_slot(VALUES_KEY),
        Dest::Stack,
        line,
    );
    emit_fold(chunks, current, line);
}

/// `HashCode.Combine(v1, …, vN)` — .NET declares arities 1 through 8.
///
/// Identical to queueing all N with `Add` and calling `ToHashCode`, because it
/// IS that: the arguments are hashed, packed into the same array shape, and
/// handed to the same fold.
pub fn emit_combine(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        collections::emit_array_new(chunks, current, 0, line);
        emit_fold(chunks, current, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc).rev() {
        set(&mut chunks[current], base + offset as u16, line);
    }
    for offset in 0..argc {
        let slot = base + offset as u16;
        get(&mut chunks[current], slot, line);
        emit_value_hash(chunks, current, line);
        set(&mut chunks[current], slot, line);
    }
    for offset in 0..argc {
        get(&mut chunks[current], base + offset as u16, line);
    }
    collections::emit_array_new(chunks, current, argc as u16, line);
    emit_fold(chunks, current, line);
}

/// `hc.GetHashCode()` and `hc.Equals(obj)` — both throw in .NET, and the
/// framework's own message says why: the mutable accumulator is not a value to
/// compare or to hash. Verified against the SDK; both answer
/// `NotSupportedException`.
pub fn emit_unsupported(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..=argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    // ⛔ The message alone is NOT the constructor sequence: an exception is
    // allocated, duplicated, then finalized with its message. Passing only the
    // message produced a throwable whose `GetType().Name` answered `Object` —
    // caught, but not catchable AS `NotSupportedException`.
    class_slots::emit_class_alloc(&mut chunks[current], line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_string_const(
        "HashCode is a mutable struct and should not be compared with other HashCodes.",
        line,
    );
    errors::emit_exception_new_finalize(&mut chunks[current], "NotSupportedException", line);
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
