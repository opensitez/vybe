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

use super::object_fields;
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};

use super::object_fields::field_slot;

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
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
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
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(PAYLOAD),
        Dest::Stack,
        line,
    );
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
    let obj = emit_value_body(&mut chunks[current], line);
    bind_value_slots(chunks, current, obj, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
}

/// `[bigint] → []`, leaving the finished value in the returned slot: the type
/// stamp, the payload, and the property fields. Binding the slots needs chunk
/// indices that may not exist yet and is therefore a separate phase.
///
/// The one place a BigInteger value's SHAPE is decided — every mint goes
/// through it, so a constructed value, a parsed one and an operator result are
/// indistinguishable to everything downstream.
fn emit_value_body(chunk: &mut Chunk, line: u32) -> u16 {
    let payload = chunk.alloc_scratch(2);
    let obj = payload + 1;
    set(chunk, payload, line);
    class_slots::emit_class_alloc(chunk, line);
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
    emit_predicate_fields(chunk, obj, payload, line);
    obj
}

/// `[] → [bool]` — `v > 0 && (v & (v - 1)) == 0` on the payload in `payload`,
/// against the zero bigint already in `zero`.
///
/// ⛔ ONE implementation, emitted by BOTH the value's `IsPowerOfTwo` field and
/// the tree's static. Writing the predicate twice is the shape that cost this
/// platform `Double.IsNaN`: two spellings of one question drift, and the one
/// that drifts is whichever the tests happen not to reach.
///
/// ⛔ The `> 0` guard is load-bearing for ZERO: `0 & -1` is 0, so the bit test
/// alone calls zero a power of two where .NET answers False. Negatives the bit
/// test already rejects; the guard states the rule outright rather than leaning
/// on that.
fn emit_power_of_two_test(chunk: &mut Chunk, payload: u16, zero: u16, line: u32) {
    get(chunk, payload, line);
    get(chunk, zero, line);
    ops::emit_dyn_gt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, payload, line);
    get(chunk, payload, line);
    core_wasm::i32_const(chunk, line, 1);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    call(chunk, "ecma:bigint", "sub", 2, line);
    call(chunk, "ecma:bigint", "and", 2, line);
    get(chunk, zero, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    core_wasm::i32_const(chunk, line, 0);
    ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);
}

/// `IsZero` / `IsOne` / `IsEven` / `IsPowerOfTwo` / `Sign` from the payload.
fn emit_predicate_fields(chunk: &mut Chunk, obj: u16, payload: u16, line: u32) {
    let scratch = chunk.alloc_scratch(2);
    let zero = scratch;
    let answer = scratch + 1;
    core_wasm::i32_const(chunk, line, 0);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    set(chunk, zero, line);

    // ⛔ Each answer lands in a slot first so it can be stored under BOTH
    // spellings. VB folds the member name, so a PascalCase-only field answers
    // `undefined` there while C# reads the same object correctly — see
    // `object_fields`.
    for (key, rhs) in [("IsZero", 0i32), ("IsOne", 1)] {
        get(chunk, payload, line);
        core_wasm::i32_const(chunk, line, rhs);
        call(chunk, "ecma:bigint", "BigInt", 1, line);
        ops::emit_dyn_eq(chunk, line);
        ops::emit_i32_to_bool(chunk, line);
        set(chunk, answer, line);
        object_fields::set_both_spellings(chunk, obj, answer, key, line);
    }

    // IsEven — the LOW BIT. `asUintN(1, v)` reads two's-complement bits, so it
    // is right for negatives too, and there is no modulo to reach for.
    core_wasm::i32_const(chunk, line, 1);
    get(chunk, payload, line);
    call(chunk, "ecma:bigint", "asUintN", 2, line);
    get(chunk, zero, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
    set(chunk, answer, line);
    object_fields::set_both_spellings(chunk, obj, answer, "IsEven", line);

    emit_power_of_two_test(chunk, payload, zero, line);
    set(chunk, answer, line);
    object_fields::set_both_spellings(chunk, obj, answer, "IsPowerOfTwo", line);

    // Sign — an ordinary Number, which is what .NET returns.
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
    set(chunk, answer, line);
    object_fields::set_both_spellings(chunk, obj, answer, "Sign", line);
}

/// The binary operator chunks, memoised by name. Built with the unary one by
/// [`ensure_value_chunks`], which owns the two-phase order.
const BIGINT_OPS: [(&str, &str, vybe_ast::ProtocolSlot); 11] = [
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
    // ⛔ The BITWISE slots were never bound, so `a And b` on two BigIntegers
    // fell through to the i32 path, which coerces a BigInt to 0 — every one of
    // `And`/`Or`/`Xor`/`<<`/`>>` answered 0. `ecma:bigint` has had `and`, `or`,
    // `xor`, `shl` and `shr` all along.
    ("__bigint_and", "and", vybe_ast::ProtocolSlot::And),
    ("__bigint_or", "or", vybe_ast::ProtocolSlot::Or),
    ("__bigint_xor", "xor", vybe_ast::ProtocolSlot::Xor),
    ("__bigint_shl", "shl", vybe_ast::ProtocolSlot::LShift),
    ("__bigint_shr", "shr", vybe_ast::ProtocolSlot::RShift),
];

/// Unary negation, published on the VALUE. `emit_rich_unary` already reads
/// `ProtocolSlot::Neg` for every language — the slot was simply never bound, so
/// `-x` fell through to the numeric path and answered 0.
const BIGINT_NEG: &str = "__bigint_neg";

/// `a.CompareTo(b)` → −1 / 0 / 1, published on the VALUE.
///
/// Binding this ONE method answers `=`, `<>`, `<`, `<=`, `>` and `>=` together:
/// when an operator's own slot is absent `emit_rich_compare_locals` reads a
/// `CompareTo` off the left operand and compares its result against zero. That
/// is also why it must be a NUMBER and not a wrapped value.
const BIGINT_COMPARE_TO: &str = "__bigint_compare_to";

/// Every chunk that answers a slot on a BigInteger value: the six operators
/// then unary negation. All seven RETURN a BigInteger, which is why they are
/// built together — see [`ensure_value_chunks`].
/// One chunk per binary operator, plus unary negation last.
const VALUE_CHUNKS: usize = BIGINT_OPS.len() + 1;

fn ensure_value_chunks(chunks: &mut Vec<Chunk>, line: u32) -> [usize; VALUE_CHUNKS] {
    let mut names: [&str; VALUE_CHUNKS] = [BIGINT_NEG; VALUE_CHUNKS];
    for (i, (name, _, _)) in BIGINT_OPS.iter().enumerate() {
        names[i] = name;
    }
    names[VALUE_CHUNKS - 1] = BIGINT_NEG;
    if chunks.iter().any(|c| c.name == BIGINT_NEG) {
        let mut found = [0usize; VALUE_CHUNKS];
        for (i, name) in names.iter().enumerate() {
            found[i] = chunks
                .iter()
                .position(|c| c.name == *name)
                .expect("value chunks are created together");
        }
        return found;
    }

    // Phase 1 — arithmetic and the result value, no binds, no return.
    let mut idxs = [0usize; VALUE_CHUNKS];
    let mut obj_slots = [0u16; VALUE_CHUNKS];
    for (i, (name, op, _)) in BIGINT_OPS.iter().enumerate() {
        let mut method = create_function_chunk(name, 2);
        method.local_count = 2;
        get(&mut method, 0, line);
        unwrap_payload(&mut method, line);
        get(&mut method, 1, line);
        unwrap_payload(&mut method, line);
        call(&mut method, "ecma:bigint", op, 2, line);
        obj_slots[i] = emit_value_body(&mut method, line);
        chunks.push(method);
        idxs[i] = chunks.len() - 1;
    }
    let mut neg = create_function_chunk(BIGINT_NEG, 1);
    neg.local_count = 1;
    get(&mut neg, 0, line);
    unwrap_payload(&mut neg, line);
    call(&mut neg, "ecma:bigint", "neg", 1, line);
    obj_slots[VALUE_CHUNKS - 1] = emit_value_body(&mut neg, line);
    chunks.push(neg);
    idxs[VALUE_CHUNKS - 1] = chunks.len() - 1;

    // Phase 2 — now every index exists, so each result can carry the slots.
    //
    // ⛔ IT HAS TO BE TWO PHASES. Each chunk's RESULT must itself carry the
    // slots, or `(a + b) * c` reads a slotless object and silently falls back
    // to the Number path — losing every digit past 2^53, which is the one thing
    // a big integer exists to prevent. A chunk cannot bind indices that do not
    // exist yet.
    let ts = to_string_chunk(chunks, line);
    let cmp = compare_to_chunk(chunks, line);
    for i in 0..idxs.len() {
        let target = idxs[i];
        bind_slots_on(chunks, target, obj_slots[i], &idxs, ts, cmp, line);
        get(&mut chunks[target], obj_slots[i], line);
        chunks[target].emit_op(Op::RETURN, line);
    }
    idxs
}

/// Publish every slot a BigInteger value answers on the object in `obj_slot`.
fn bind_slots_on(
    chunks: &mut Vec<Chunk>,
    target: usize,
    obj_slot: u16,
    idxs: &[usize; VALUE_CHUNKS],
    to_string: usize,
    compare_to: usize,
    line: u32,
) {
    let bind = |chunks: &mut Vec<Chunk>, key: &str, method: usize| {
        vybe_compiler::primitives::object::emit_bind_method(
            &mut chunks[target],
            obj_slot,
            key,
            method,
            line,
        );
    };
    for (i, (_, _, slot)) in BIGINT_OPS.iter().enumerate() {
        bind(chunks, &vybe_ast::protocol_slot_key(*slot), idxs[i]);
    }
    bind(
        chunks,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Neg),
        idxs[VALUE_CHUNKS - 1],
    );
    // ToString twice — the shared stringification slot AND the .NET spelling,
    // which is a member call, not the slot. Both have to answer, and an
    // operator RESULT is the value that most needs them: `(a + b)` is a
    // BigInteger the tree cannot name, so `.ToString()` otherwise fell through
    // to the generic object one and printed `[object BigInteger]`.
    bind(
        chunks,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        to_string,
    );
    bind(chunks, "ToString", to_string);
    bind(
        chunks,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Compare),
        compare_to,
    );
    bind(chunks, "CompareTo", compare_to);
}

/// The `ToString` slot — one unary chunk, memoised.
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

/// The `CompareTo` chunk — `[a, b] → [-1 | 0 | 1]` as an ordinary Number.
fn compare_to_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == BIGINT_COMPARE_TO) {
        return idx;
    }
    let mut method = create_function_chunk(BIGINT_COMPARE_TO, 2);
    method.local_count = 2;

    // ⛔ NULL-SAFE, and it has to be. Binding `CompareTo` is what answers `=`,
    // `<>`, `<`, `<=`, `>` and `>=` on a BigInteger, so `x <> Nothing` arrives
    // HERE with `Nothing` on the right. `typeof null` is `"object"` in ECMA,
    // so `unwrap_payload` read `__bi` off it and the comparison TRAPPED
    // (`toF64 — not a number`) instead of answering — which is how
    // `BigInteger.TryParse`, whose desugar is a `<> Nothing` test, died. A
    // present value sorts AFTER an absent one: `x <> Nothing` is True.
    get(&mut method, 1, line);
    method.emit_op(Op::REF_IS_NULL, line);
    method.emit_if(line);
    core_wasm::i32_const(&mut method, line, 1);
    method.emit_op(Op::RETURN, line);
    method.emit_end(line);

    let a = method.alloc_scratch(2);
    let b = a + 1;
    get(&mut method, 0, line);
    unwrap_payload(&mut method, line);
    set(&mut method, a, line);
    get(&mut method, 1, line);
    unwrap_payload(&mut method, line);
    set(&mut method, b, line);

    get(&mut method, a, line);
    get(&mut method, b, line);
    ops::emit_dyn_lt(&mut method, line);
    ops::emit_dyn_to_bool(&mut method, line);
    method.emit_if_value(line);
    core_wasm::i32_const(&mut method, line, -1);
    method.emit_else(line);
    get(&mut method, a, line);
    get(&mut method, b, line);
    ops::emit_dyn_gt(&mut method, line);
    ops::emit_dyn_to_bool(&mut method, line);
    method.emit_if_value(line);
    core_wasm::i32_const(&mut method, line, 1);
    method.emit_else(line);
    core_wasm::i32_const(&mut method, line, 0);
    method.emit_end(line);
    method.emit_end(line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

/// Publish every value slot on the object in `obj_slot` of the CURRENT chunk.
fn bind_value_slots(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let idxs = ensure_value_chunks(chunks, line);
    let ts = to_string_chunk(chunks, line);
    let cmp = compare_to_chunk(chunks, line);
    bind_slots_on(chunks, current, obj_slot, &idxs, ts, cmp, line);
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

/// `BigInteger.Parse(s)` / `Parse(s, NumberStyles)` — `BigInt` already parses a
/// decimal string, and it THROWS on a malformed one, which is what `Parse` owes
/// its caller.
///
/// Also the CONSTRUCTOR's backing, so `New BigInteger(54)` and `Parse("54")`
/// mint the same value.
pub fn emit_parse(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        emit_hex_aware_parse(chunks, current, line);
    } else {
        emit_scalar_or_bytes(&mut chunks[current], line);
    }
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

/// `[s, styles] → [bigint]`, reading the styles for `AllowHexSpecifier`.
///
/// ⛔ The test is `styles >= 512`, not a bit AND. `AllowHexSpecifier` IS 512
/// and it is the HIGHEST flag in `NumberStyles`, so every combination that
/// includes it is ≥ 512 and every combination that omits it is ≤ 511 — the
/// comparison is exact over the whole enum, and the styles arrive as an f64
/// where a bitwise AND would need an integer round-trip first.
fn emit_hex_aware_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(2);
    let styles = scratch;
    let text = scratch + 1;
    set(chunk, styles, line);
    set(chunk, text, line);

    get(chunk, styles, line);
    core_wasm::f64_const(chunk, line, 512.0);
    ops::emit_dyn_ge(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // `BigInt` reads a `0x` prefix as hexadecimal, which is how the arbitrary
    // width survives — `parseInt(s, 16)` would round past 2^53.
    chunk.emit_string_const("0x", line);
    get(chunk, text, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    chunk.emit_else(line);
    get(chunk, text, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    chunk.emit_end(line);
}

/// `BigInteger.TryParse(s)` — the hidden one-arg core `try_parse_desugar`
/// builds. NULL on a malformed string; the desugar turns that into `False` and
/// the out-param write.
pub fn emit_try_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let text = chunk.alloc_scratch(1);
    set(chunk, text, line);

    // A decimal spelling, sign optional — `BigInt` THROWS on anything else and
    // `TryParse` may not throw. `ecma:regexp.test` takes the PATTERN first.
    chunk.emit_string_const("^\\s*[+-]?[0-9]+\\s*$", line);
    get(chunk, text, line);
    call(chunk, "ecma:string", "String", 1, line);
    call(chunk, "ecma:regexp", "test", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, text, line);
    to_bigint(chunk, line);
    wrap(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
}

/// `v.ToString()` / `v.ToString(format)` — the decimal spelling, or a radix
/// when the format asks for one. `ecma:bigint.toString` is exact for every
/// magnitude, where a `Number` round-trip would lose digits past 2^53.
pub fn emit_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        let v = stash_bigints(chunks, current, 1, line);
        let chunk = &mut chunks[current];
        get(chunk, v, line);
        call(chunk, "ecma:bigint", "toString", 1, line);
        return;
    }

    let chunk = &mut chunks[current];
    let format = chunk.alloc_scratch(1);
    set(chunk, format, line);
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];

    // `X` / `x` — hexadecimal, in the case of the format letter itself.
    get(chunk, format, line);
    call(chunk, "ecma:string", "toUpperCase", 1, line);
    chunk.emit_string_const("X", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, v, line);
    core_wasm::i32_const(chunk, line, 16);
    call(chunk, "ecma:bigint", "toString", 2, line);
    let digits = chunk.alloc_scratch(1);
    set(chunk, digits, line);
    get(chunk, format, line);
    chunk.emit_string_const("X", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, digits, line);
    call(chunk, "ecma:string", "toUpperCase", 1, line);
    chunk.emit_else(line);
    get(chunk, digits, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    get(chunk, v, line);
    call(chunk, "ecma:bigint", "toString", 1, line);
    chunk.emit_end(line);
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

/// `v.IsPowerOfTwo` — the tree's answer for a receiver it can name. The value
/// carries the same answer as a FIELD; both paths exist because .NET spells it
/// as a property and only one of the two reaches an untyped local.
pub fn emit_is_power_of_two(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    let zero = chunk.alloc_scratch(1);
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    set(chunk, zero, line);
    emit_power_of_two_test(chunk, v, zero, line);
}

/// `[] → [f64]` — log base ten of the bigint in `v`, from its DECIMAL SPELLING.
///
/// ⛔ Not `Math.log10(Number(v))`. A `Value::BigInt` reaching `as_f64()`
/// answers NaN, and even a working conversion is `Infinity` past 1.8e308 while
/// .NET's `BigInteger.Log10` stays finite at any magnitude. The exact decimal
/// string carries both halves of the answer: its LENGTH is the exponent and its
/// leading digits are the mantissa, so the precision is the Double's rather
/// than the conversion's.
fn emit_log10_core(chunk: &mut Chunk, v: u16, line: u32) {
    let scratch = chunk.alloc_scratch(2);
    let text = scratch;
    let head = scratch + 1;

    get(chunk, v, line);
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // .NET answers NaN for a negative, and the logarithm of a negative has no
    // real value to answer with.
    core_wasm::f64_const(chunk, line, f64::NAN);
    chunk.emit_else(line);

    get(chunk, v, line);
    call(chunk, "ecma:bigint", "toString", 1, line);
    set(chunk, text, line);
    // Seventeen digits is more than a Double's 15-to-17 significant ones, so
    // the head carries every bit the result can hold.
    get(chunk, text, line);
    core_wasm::i32_const(chunk, line, 0);
    core_wasm::i32_const(chunk, line, 17);
    call(chunk, "ecma:string", "substring", 3, line);
    set(chunk, head, line);

    get(chunk, head, line);
    call(chunk, "ecma:number", "Number", 1, line);
    call(chunk, "ecma:math", "log10", 1, line);
    get(chunk, text, line);
    call(chunk, "ecma:string", "length", 1, line);
    call(chunk, "ecma:number", "Number", 1, line);
    chunk.emit_op(Op::F64_ADD, line);
    get(chunk, head, line);
    call(chunk, "ecma:string", "length", 1, line);
    call(chunk, "ecma:number", "Number", 1, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_end(line);
}

/// `BigInteger.Log10(value)` → Double.
pub fn emit_log10(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    emit_log10_core(&mut chunks[current], v, line);
}

/// `BigInteger.Log(value)` → the NATURAL logarithm, `BigInteger.Log(value,
/// baseValue)` → an arbitrary base. Both are derived from the base-ten answer
/// so the arbitrary-magnitude handling lives in exactly one place.
pub fn emit_log(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        let v = stash_bigints(chunks, current, 1, line);
        let chunk = &mut chunks[current];
        emit_log10_core(chunk, v, line);
        core_wasm::f64_const(chunk, line, std::f64::consts::LN_10);
        chunk.emit_op(Op::F64_MUL, line);
        return;
    }

    // ⛔ The base is an ordinary Double, so it must come off the stack BEFORE
    // `stash_bigints`, which unwraps every operand it pops as a payload.
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(1);
    set(chunk, base, line);
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    emit_log10_core(chunk, v, line);
    get(chunk, base, line);
    call(chunk, "ecma:number", "Number", 1, line);
    call(chunk, "ecma:math", "log10", 1, line);
    chunk.emit_op(Op::F64_DIV, line);
}

/// `BigInteger.Clamp(value, min, max)`.
pub fn emit_clamp(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = stash_bigints(chunks, current, 3, line);
    let chunk = &mut chunks[current];
    get(chunk, base, line);
    get(chunk, base + 1, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, base + 1, line);
    chunk.emit_else(line);
    get(chunk, base, line);
    get(chunk, base + 2, line);
    ops::emit_dyn_gt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, base + 2, line);
    chunk.emit_else(line);
    get(chunk, base, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    // The result is a BigInteger VALUE — wrapped, with its
    // operator slots bound, so `BigInteger.Abs(x) * y` works.
    wrap(chunks, current, line);
}

/// `v.ToByteArray()` — LITTLE-ENDIAN two's complement, which is .NET's own
/// layout and exactly what `New BigInteger(bytes)` reads back. The two are
/// written together for that reason: a round-trip is the only thing that
/// pins the convention, and either one alone can be self-consistently wrong.
pub fn emit_to_byte_array(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = stash_bigints(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(5);
    let arr = scratch;
    let rest = scratch + 1;
    let byte = scratch + 2;
    let zero = scratch + 3;
    let minus_one = scratch + 4;

    core_wasm::i32_const(chunk, line, 0);
    call(chunk, "vybe:js-array", "newWithLength", 1, line);
    set(chunk, arr, line);
    get(chunk, v, line);
    set(chunk, rest, line);
    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    set(chunk, zero, line);
    core_wasm::i32_const(chunk, line, -1);
    to_bigint(chunk, line);
    set(chunk, minus_one, line);

    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);

    // The low eight bits, as an ordinary Number. `asUintN` reads the
    // TWO'S-COMPLEMENT bits, so a negative value yields 0..255 like any other.
    core_wasm::i32_const(chunk, line, 8);
    get(chunk, rest, line);
    call(chunk, "ecma:bigint", "asUintN", 2, line);
    call(chunk, "ecma:bigint", "toString", 1, line);
    call(chunk, "ecma:number", "Number", 1, line);
    set(chunk, byte, line);
    get(chunk, arr, line);
    get(chunk, byte, line);
    call(chunk, "ecma:array", "push", 2, line);
    chunk.emit_op(Op::DROP, line);

    get(chunk, rest, line);
    core_wasm::i32_const(chunk, line, 8);
    to_bigint(chunk, line);
    call(chunk, "ecma:bigint", "shr", 2, line);
    set(chunk, rest, line);

    // Done once the SIGN is representable in the bytes already written:
    // nothing left and the top bit clear (positive), or −1 left and the top
    // bit set (negative). Stopping at "nothing left" alone drops the sign and
    // 200 reads back as −56.
    get(chunk, rest, line);
    get(chunk, zero, line);
    ops::emit_dyn_eq(chunk, line);
    get(chunk, byte, line);
    core_wasm::f64_const(chunk, line, 128.0);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    get(chunk, rest, line);
    get(chunk, minus_one, line);
    ops::emit_dyn_eq(chunk, line);
    get(chunk, byte, line);
    core_wasm::f64_const(chunk, line, 128.0);
    ops::emit_dyn_ge(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_br_if(1, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    get(chunk, arr, line);
}

/// `[bytes] → [bigint]` — the read side of [`emit_to_byte_array`].
fn emit_bytes_to_bigint(chunk: &mut Chunk, bytes: u16, line: u32) {
    let scratch = chunk.alloc_scratch(6);
    let acc = scratch;
    let index = scratch + 1;
    let top = scratch + 2;
    let eight = scratch + 3;
    let len = scratch + 4;
    let byte = scratch + 5;

    core_wasm::i32_const(chunk, line, 0);
    to_bigint(chunk, line);
    set(chunk, acc, line);
    core_wasm::i32_const(chunk, line, 8);
    to_bigint(chunk, line);
    set(chunk, eight, line);
    get(chunk, bytes, line);
    call(chunk, "ecma:array", "length", 1, line);
    call(chunk, "ecma:number", "Number", 1, line);
    set(chunk, len, line);
    get(chunk, len, line);
    set(chunk, index, line);

    // ⛔ The SIGN byte is the LAST one, read here rather than in the loop: the
    // loop walks downwards and its final read is `bytes[0]`, the LEAST
    // significant. Testing that one made every value with an odd low byte
    // negative.
    get(chunk, bytes, line);
    get(chunk, len, line);
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op(Op::F64_SUB, line);
    call(chunk, "ecma:array", "get", 2, line);
    call(chunk, "ecma:number", "Number", 1, line);
    set(chunk, top, line);

    // Most significant byte first — the array is little-endian, so walk it
    // backwards and shift each byte in from the bottom.
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, index, line);
    core_wasm::f64_const(chunk, line, 0.0);
    ops::emit_dyn_le(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, index, line);
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, index, line);

    get(chunk, bytes, line);
    get(chunk, index, line);
    call(chunk, "ecma:array", "get", 2, line);
    call(chunk, "ecma:number", "Number", 1, line);
    set(chunk, byte, line);

    get(chunk, acc, line);
    get(chunk, eight, line);
    call(chunk, "ecma:bigint", "shl", 2, line);
    get(chunk, byte, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    call(chunk, "ecma:bigint", "or", 2, line);
    set(chunk, acc, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);

    // ⛔ The layout is TWO'S COMPLEMENT: a top byte of 128 or more means the
    // value is negative, and it is `acc - 2^(8n)`. Reading the bytes as an
    // unsigned magnitude makes `New BigInteger(x.ToByteArray())` disagree with
    // `x` for every negative `x`.
    get(chunk, top, line);
    core_wasm::f64_const(chunk, line, 128.0);
    ops::emit_dyn_ge(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, acc, line);
    core_wasm::i32_const(chunk, line, 1);
    to_bigint(chunk, line);
    get(chunk, len, line);
    core_wasm::f64_const(chunk, line, 8.0);
    chunk.emit_op(Op::F64_MUL, line);
    call(chunk, "ecma:bigint", "BigInt", 1, line);
    call(chunk, "ecma:bigint", "shl", 2, line);
    call(chunk, "ecma:bigint", "sub", 2, line);
    set(chunk, acc, line);
    chunk.emit_end(line);
    get(chunk, acc, line);
}

/// `[value] → [bigint]` for the CONSTRUCTOR and one-argument `Parse`.
///
/// ⛔ An array has to be tested for FIRST. `New BigInteger(bytes)` is a real
/// .NET constructor, and `typeof [] === "object"` — so `unwrap_payload` read a
/// `__bi` field off the array and answered `undefined`, silently.
fn emit_scalar_or_bytes(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    set(chunk, value, line);
    get(chunk, value, line);
    call(chunk, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    emit_bytes_to_bigint(chunk, value, line);
    chunk.emit_else(line);
    get(chunk, value, line);
    to_bigint(chunk, line);
    chunk.emit_end(line);
}

/// `BigInteger.DivRem(a, b)` — the two-argument TUPLE overload, which is also
/// what the walkers' out-param desugar calls. Both halves are wrapped VALUES,
/// so `BigInteger.DivRem(a, b).Quotient * c` stays big-integer arithmetic.
pub fn emit_div_rem(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = stash_bigints(chunks, current, 2, line);
    bigop(&mut chunks[current], "div", base, base + 1, line);
    wrap(chunks, current, line);
    let chunk = &mut chunks[current];
    let slots = chunk.alloc_scratch(3);
    let quotient = slots;
    let remainder = slots + 1;
    let pair = slots + 2;
    set(chunk, quotient, line);

    bigop(&mut chunks[current], "rem", base, base + 1, line);
    wrap(chunks, current, line);
    let chunk = &mut chunks[current];
    set(chunk, remainder, line);

    class_slots::emit_class_alloc(chunk, line);
    set(chunk, pair, line);
    // `Quotient`/`Remainder` are the .NET 7 tuple's names and `Item1`/`Item2`
    // the positional ones. A `ValueTuple` answers to both, so both are stored.
    for (key, slot) in [
        ("Quotient", quotient),
        ("Item1", quotient),
        ("Remainder", remainder),
        ("Item2", remainder),
    ] {
        get(chunk, pair, line);
        get(chunk, slot, line);
        field_set(chunk, key, line);
    }
    get(chunk, pair, line);
}
