//! Integer bit operations — one home for a family every language has.
//!
//! `i32.popcnt` is ONE wasm instruction, and before this module it was reached
//! three different ways: `common:fortran.popcnt` through a Fortran-local
//! adapter (i32 lane), `common:go.bits_ones_count` through a Go-local adapter
//! (i64 lane), and `opcode:i32_popcnt` from wast. Two emitters, opposite lanes,
//! neither aware of the other — the shape `directives.md` §9.2 describes, where
//! N spellings drift and each copy looks locally correct.
//!
//! The operations are [`UnaryOp::PopCount`] / [`LeadingZeros`] /
//! [`TrailingZeros`] and [`BinOp::RotL`] / [`RotR`], each carrying its
//! [`BitLane`], because the lane is a property of the operand's declared type
//! and wasm spells the two lanes as different instructions.
//!
//! Number-backed operands are converted into the declared integer lane.
//! Arithmetic returns a number; comparisons return an i32 predicate.
//!
//! [`LeadingZeros`]: UnaryOp::LeadingZeros
//! [`TrailingZeros`]: UnaryOp::TrailingZeros
//! [`RotR`]: BinOp::RotR

use vybe_ast::{BitLane, IntOp, NumericRepr, ShiftOverflow};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static FOLD_LITERAL_WORDS: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_literal_word_folding<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore { fn drop(&mut self) { FOLD_LITERAL_WORDS.with(|v| v.set(self.0)); } }
    let _restore = Restore(FOLD_LITERAL_WORDS.with(|v| v.replace(enabled)));
    build()
}

/// Fold only the literal domain already emitted as exact W32 words. Wider
/// integer literals can round during their Number conversion, or trap, so
/// their original conversion must remain. Trapping operations also stay live.
pub(crate) fn fold_literal_word(op: IntOp, a: i64, b: i64) -> Option<i32> {
    #[cfg(feature = "primitive-benchmarks")]
    if !FOLD_LITERAL_WORDS.with(|v| v.get()) { return None; }
    let range = i64::from(i32::MIN)..=i64::from(u32::MAX);
    if !range.contains(&a) || !range.contains(&b) { return None; }
    let (a,b) = (a as i32,b as i32);
    Some(match op {
        IntOp::Add => a.wrapping_add(b),
        IntOp::Sub => a.wrapping_sub(b),
        IntOp::Mul => a.wrapping_mul(b),
        IntOp::DivS => a.checked_div(b)?,
        IntOp::DivU => (a as u32).checked_div(b as u32)? as i32,
        IntOp::RemS => { if b==0 {return None;} a.wrapping_rem(b) },
        IntOp::RemU => (a as u32).checked_rem(b as u32)? as i32,
        IntOp::Shl => a.wrapping_shl(b as u32),
        IntOp::ShrS => a.wrapping_shr(b as u32),
        IntOp::ShrU => (a as u32).wrapping_shr(b as u32) as i32,
        IntOp::And => a & b, IntOp::Or => a | b, IntOp::Xor => a ^ b,
        IntOp::Eq => i32::from(a==b), IntOp::Ne => i32::from(a!=b),
        IntOp::LtS => i32::from(a<b), IntOp::LtU => i32::from((a as u32)<b as u32),
        IntOp::LeS => i32::from(a<=b), IntOp::LeU => i32::from((a as u32)<=b as u32),
        IntOp::GtS => i32::from(a>b), IntOp::GtU => i32::from((a as u32)>b as u32),
        IntOp::GeS => i32::from(a>=b), IntOp::GeU => i32::from((a as u32)>=b as u32),
    })
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static FOLD_WORD_EXPRESSIONS: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_word_expression_folding<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore { fn drop(&mut self) { FOLD_WORD_EXPRESSIONS.with(|v| v.set(self.0)); } }
    let _restore = Restore(FOLD_WORD_EXPRESSIONS.with(|v| v.replace(enabled)));
    build()
}

// Fold bottom-up from what the emitter actually produced. Recognizing a
// constant visits at most nine bytes; nonconstant ASTs are never rescanned.
fn word_expression_folding_enabled() -> bool {
    #[cfg(feature = "primitive-benchmarks")]
    { return FOLD_WORD_EXPRESSIONS.with(|v| v.get()) && FOLD_LITERAL_WORDS.with(|v| v.get()); }
    #[cfg(not(feature = "primitive-benchmarks"))]
    { true }
}
fn emitted_word(code: &[u8]) -> Option<i32> {
    if !(5..=9).contains(&code.len()) || Op::from_code_at(code, 0) != Op::I32_CONST {
        return None;
    }
    let mut at = 4;
    let value = vybe_runtime::opcode::read_leb_i32(code, &mut at);
    (at == code.len()).then_some(value)
}
fn replace_constant_tail(chunk: &mut Chunk, start: usize, value: i32, line: u32) {
    // Only i32.const instructions are removed. They carry no relocation,
    // control-flow, call-site or result-type metadata. Preserve the byte-indexed
    // source map using its sparse-aware split, without materializing it.
    debug_assert_eq!(chunk.code.len(), chunk.lines.len());
    chunk.code.truncate(start);
    drop(chunk.lines.split_off_table(start));
    chunk.emit_i32_const(value, line);
}
pub(crate) fn fold_emitted_word_binary(
    chunk: &mut Chunk, start: usize, middle: usize, op: &vybe_ast::BinOp, line: u32,
) -> bool {
    if !word_expression_folding_enabled() { return false; }
    let Some(a) = emitted_word(&chunk.code[start..middle]) else { return false; };
    let Some(b) = emitted_word(&chunk.code[middle..]) else { return false; };
    let value = match op {
        vybe_ast::BinOp::Integer(operation, BitLane::W32) => fold_literal_word(*operation, a.into(), b.into()),
        vybe_ast::BinOp::RotL(BitLane::W32) => Some(a.rotate_left(b as u32)),
        vybe_ast::BinOp::RotR(BitLane::W32) => Some(a.rotate_right(b as u32)),
        _ => None,
    };
    let Some(value) = value else { return false; };
    replace_constant_tail(chunk, start, value, line);
    true
}
pub(crate) fn fold_emitted_word_count(
    chunk: &mut Chunk, start: usize, op: &vybe_ast::UnaryOp, line: u32,
) -> bool {
    if !word_expression_folding_enabled() { return false; }
    let Some(a) = emitted_word(&chunk.code[start..]) else { return false; };
    let value = match op {
        vybe_ast::UnaryOp::PopCount(BitLane::W32) => a.count_ones(),
        vybe_ast::UnaryOp::LeadingZeros(BitLane::W32) => a.leading_zeros(),
        vybe_ast::UnaryOp::TrailingZeros(BitLane::W32) => a.trailing_zeros(),
        _ => return false,
    };
    replace_constant_tail(chunk, start, value as i32, line);
    true
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static WORD_RIGHT_CONSTANT: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_word_right_constant<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore { fn drop(&mut self) { WORD_RIGHT_CONSTANT.with(|v| v.set(self.0)); } }
    let _restore = Restore(WORD_RIGHT_CONSTANT.with(|v| v.replace(enabled)));
    build()
}

/// A pure constant on top of a live operand can be emitted after that operand's
/// normalization. This preserves evaluation/trap order without a scratch local.
/// `left_is_word` means an actual word producer, not just a numeric type hint.
pub(crate) fn emit_word_right_constant(
    chunk: &mut Chunk, right_start: usize, op: &vybe_ast::BinOp,
    left_is_word: bool, line: u32,
) -> bool {
    #[cfg(feature = "primitive-benchmarks")]
    if !WORD_RIGHT_CONSTANT.with(|v| v.get()) { return false; }
    use vybe_ast::BinOp;
    if !matches!(op, BinOp::Integer(_, BitLane::W32)
        | BinOp::RotL(BitLane::W32) | BinOp::RotR(BitLane::W32)) { return false; }
    let Some(right) = emitted_word(&chunk.code[right_start..]) else { return false; };
    let identity = match op {
        BinOp::Integer(IntOp::Add | IntOp::Sub | IntOp::Or | IntOp::Xor, _) => right == 0,
        BinOp::Integer(IntOp::Mul | IntOp::DivS | IntOp::DivU, _) => right == 1,
        BinOp::Integer(IntOp::And, _) => right == -1,
        BinOp::Integer(IntOp::Shl | IntOp::ShrS | IntOp::ShrU, _)
            | BinOp::RotL(_) | BinOp::RotR(_) => (right as u32 & 31) == 0,
        _ => false,
    };
    if left_is_word && !identity { return false; }
    debug_assert_eq!(chunk.code.len(), chunk.lines.len());
    chunk.code.truncate(right_start);
    drop(chunk.lines.split_off_table(right_start));
    if !left_is_word { narrow(chunk, BitLane::W32, line); }
    if !identity {
        chunk.emit_i32_const(right, line);
        chunk.emit_op(match op {
            BinOp::Integer(operation, _) => integer_opcode(BitLane::W32, *operation),
            BinOp::RotL(_) => Op::I32_ROTL,
            BinOp::RotR(_) => Op::I32_ROTR,
            _ => unreachable!(),
        }, line);
    }
    true
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static STATIC_WORD_CONSTANTS: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_static_word_constants<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore { fn drop(&mut self) { STATIC_WORD_CONSTANTS.with(|v| v.set(self.0)); } }
    let _restore = Restore(STATIC_WORD_CONSTANTS.with(|v| v.replace(enabled)));
    build()
}
pub(crate) fn static_word_constants_enabled() -> bool {
    #[cfg(feature = "primitive-benchmarks")]
    { STATIC_WORD_CONSTANTS.with(|v| v.get()) }
    #[cfg(not(feature = "primitive-benchmarks"))]
    { true }
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static LINEAR_NUMERIC_INFERENCE: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_linear_numeric_inference<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore { fn drop(&mut self) { LINEAR_NUMERIC_INFERENCE.with(|v| v.set(self.0)); } }
    let _restore = Restore(LINEAR_NUMERIC_INFERENCE.with(|v| v.replace(enabled)));
    build()
}
#[cfg(feature = "primitive-benchmarks")]
pub(crate) fn linear_numeric_inference_enabled() -> bool {
    LINEAR_NUMERIC_INFERENCE.with(|v| v.get())
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static WORD_SCALAR_REFERENCE_PROOF: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_word_scalar_reference_proof<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore { fn drop(&mut self) { WORD_SCALAR_REFERENCE_PROOF.with(|v| v.set(self.0)); } }
    let _restore = Restore(WORD_SCALAR_REFERENCE_PROOF.with(|v| v.replace(enabled)));
    build()
}
pub(crate) fn word_scalar_reference_proof_enabled() -> bool {
    #[cfg(feature = "primitive-benchmarks")]
    { WORD_SCALAR_REFERENCE_PROOF.with(|v| v.get()) }
    #[cfg(not(feature = "primitive-benchmarks"))]
    { true }
}

/// Raw locals cannot reuse a slot that previously held a value: Wasm local
/// types cover the entire function, including earlier statements.
pub(crate) fn raw_i64_scratch(chunk: &mut Chunk) -> u16 {
    if chunk.i64_scratch_slot.is_none() {
        chunk.local_count = chunk.local_count.max(chunk.scratch_high_water);
    }
    chunk.i64_scratch()
}

/// Bring the value on the stack into `lane`'s integer type.
///
/// Convert the numeric value to its integer lane. A direct i32 coercion can
/// saturate values above INT32_MAX; wrapping through i64 preserves all 32 bits.
fn narrow(chunk: &mut Chunk, lane: BitLane, line: u32) {
    match lane {
        BitLane::W32 => {
            chunk.emit_op(Op::I64_TRUNC_F64_S, line);
            chunk.emit_op(Op::I32_WRAP_I64, line);
        }
        BitLane::W64 => chunk.emit_op(Op::I64_TRUNC_F64_S, line),
    }
}

/// Return `lane`'s integer to the number representation the runtime carries.
/// The 32-bit lane is already there.
fn widen(chunk: &mut Chunk, lane: BitLane, line: u32) {
    if lane == BitLane::W64 {
        chunk.emit_op(Op::F64_CONVERT_I64_S, line);
    }
}

fn narrow_comparison_word(chunk: &mut Chunk, lane: BitLane, line: u32) {
    if lane == BitLane::W32 {
        narrow(chunk, lane, line);
        return;
    }
    let word = retain_wide_comparison_word(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
}

// Keep the input below the converted operand live on the stack. No source
// expression runs during either conversion, and conversion order is unchanged.
fn retain_wide_comparison_word(chunk: &mut Chunk, line: u32) -> u16 {
    // Number-backed unsigned words can occupy the upper half of u64. Both
    // conversions produce the same i64 bit lane; signed ordering is selected
    // by the comparison opcode, not by lossy conversion through i32.
    let value = chunk.alloc_scratch(1);
    let word = raw_i64_scratch(chunk);
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
    chunk.emit_f64_const(9_223_372_036_854_775_808.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);
    chunk.emit_end(line);
    word
}

fn emit_wide_word_comparison(chunk: &mut Chunk, op: IntOp, line: u32) {
    emit_wide_comparison_with_signed_word_operands(chunk, op, (false, false), line);
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static WIDE_SIGNED_WORDS: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_wide_signed_words<R>(enabled: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore { fn drop(&mut self) { WIDE_SIGNED_WORDS.with(|v| v.set(self.0)); } }
    let _restore = Restore(WIDE_SIGNED_WORDS.with(|v| v.replace(enabled)));
    build()
}
pub(crate) fn wide_signed_words_enabled() -> bool {
    #[cfg(feature = "primitive-benchmarks")]
    { WIDE_SIGNED_WORDS.with(|v| v.get()) }
    #[cfg(not(feature = "primitive-benchmarks"))]
    { true }
}

/// `signed_words` proves each operand's original numeric value is a physical
/// signed i32, not merely a word-shaped replacement for an unsigned Number.
/// The generic public normalization flags keep their existing lane contract.
pub fn emit_wide_comparison_with_signed_word_operands(
    chunk: &mut Chunk, op: IntOp, signed_words: (bool, bool), line: u32,
) {
    assert!(op.is_comparison());
    let signed_words = if signed_words != (false, false) && !wide_signed_words_enabled() {
        (false, false)
    } else { signed_words };
    if signed_words == (true, true) {
        // Sign extension preserves signed ordering and unsigned ordering
        // within the signed-i32 domain, as well as equality and inequality.
        chunk.emit_op(integer_opcode(BitLane::W32, op), line);
        return;
    }
    // Keep right-before-left conversion/trap order for unproven values.
    let right = if signed_words.1 {
        let slot = raw_i64_scratch(chunk);
        chunk.emit_op(Op::I64_EXTEND_I32_S, line);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    } else {
        retain_wide_comparison_word(chunk, line)
    };
    if signed_words.0 {
        chunk.emit_op(Op::I64_EXTEND_I32_S, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right, line);
        chunk.emit_op(integer_opcode(BitLane::W64, op), line);
        return;
    }
    let left = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, left, line);
    chunk.emit_f64_const(9_223_372_036_854_775_808.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if_i32(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    chunk.emit_op(integer_opcode(BitLane::W64, op), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    chunk.emit_op(integer_opcode(BitLane::W64, op), line);
    chunk.emit_end(line);
}

/// Sign extension commutes with bitwise AND/OR/XOR. Preserve the W64
/// arithmetic result's Number carrier even though its integer work is i32.
pub fn emit_wide_bitwise_with_signed_word_operands(
    chunk: &mut Chunk, op: IntOp, signed_words: (bool, bool), line: u32,
) {
    assert!(matches!(op, IntOp::And | IntOp::Or | IntOp::Xor));
    if signed_words == (true, true) && wide_signed_words_enabled() {
        chunk.emit_op(integer_opcode(BitLane::W32, op), line);
        chunk.emit_op(Op::F64_FROM_I32, line);
    } else {
        emit_integer_binary(chunk, BitLane::W64, op, line);
    }
}

/// Two signed i32 values have an exact sum/difference within 33 bits. f64
/// produces the same Number as signed i64 arithmetic followed by conversion,
/// without wrapping at i32 width or overflowing i64 width.
pub fn emit_wide_sum_with_signed_word_operands(
    chunk: &mut Chunk, op: IntOp, signed_words: (bool, bool), line: u32,
) {
    assert!(matches!(op, IntOp::Add | IntOp::Sub));
    if signed_words == (true, true) && wide_signed_words_enabled() {
        chunk.emit_op(if op == IntOp::Add { Op::F64_ADD } else { Op::F64_SUB }, line);
    } else {
        emit_integer_binary(chunk, BitLane::W64, op, line);
    }
}

pub fn emit_normalize_integer(chunk: &mut Chunk, lane: BitLane, line: u32) {
    narrow(chunk, lane, line);
    widen(chunk, lane, line);
}

/// One unary bit-count instruction. Skip narrowing only for a proven W32
/// operand; the W64 number representation always retains its conversions.
pub fn emit_count_with_normalized_operand(
    chunk: &mut Chunk, lane: BitLane, w32: Op, w64: Op, normalized: bool, line: u32,
) {
    if lane != BitLane::W32 || !normalized { narrow(chunk, lane, line); }
    chunk.emit_op(if lane == BitLane::W32 { w32 } else { w64 }, line);
    widen(chunk, lane, line);
}

pub fn emit_pop_count(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count_with_normalized_operand(chunk, lane, Op::I32_POPCNT, Op::I64_POPCNT, false, line);
}

pub fn emit_leading_zeros(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count_with_normalized_operand(chunk, lane, Op::I32_CLZ, Op::I64_CLZ, false, line);
}

pub fn emit_trailing_zeros(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count_with_normalized_operand(chunk, lane, Op::I32_CTZ, Op::I64_CTZ, false, line);
}

/// `rotl` / `rotr` — stack on entry is `[value, count]`.
///
/// wasm's rotate already takes the count modulo the width, which is what every
/// language spelling this wants: a rotation by the full width IS the identity,
/// so [`ShiftOverflow`] does not apply here the way it does to a shift.
pub fn emit_rotate(chunk: &mut Chunk, lane: BitLane, left: bool, line: u32) {
    emit_rotate_with_normalized_operands(chunk, lane, left, (false, false), line);
}

/// Skip conversion only when the emitter proved an operand is already a W32 word.
/// A normalized left word can remain on the stack while the count is narrowed.
pub fn emit_rotate_with_normalized_operands(
    chunk: &mut Chunk, lane: BitLane, left: bool, normalized: (bool, bool), line: u32,
) {
    if lane == BitLane::W64 {
        // Only the count must leave the operand stack. Reserve a persistent raw
        // i64 local; ordinary scratch is externref and can be reclaimed later.
        let count = raw_i64_scratch(chunk);
        narrow(chunk, lane, line);
        chunk.emit_op_u16(Op::LOCAL_SET, count, line);
        narrow(chunk, lane, line);
        chunk.emit_op_u16(Op::LOCAL_GET, count, line);
        chunk.emit_op(if left { Op::I64_ROTL } else { Op::I64_ROTR }, line);
        widen(chunk, lane, line);
        return;
    }
    if normalized.0 {
        if !normalized.1 { narrow(chunk, lane, line); }
        chunk.emit_op(if left { Op::I32_ROTL } else { Op::I32_ROTR }, line);
        return;
    }
    let count = chunk.alloc_scratch(1);
    // Preserve the count while narrowing the left operand underneath it.
    if lane != BitLane::W32 || !normalized.1 { narrow(chunk, lane, line); }
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);
    if lane != BitLane::W32 || !normalized.0 { narrow(chunk, lane, line); }
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    let op = match (lane, left) {
        (BitLane::W32, true) => Op::I32_ROTL,
        (BitLane::W32, false) => Op::I32_ROTR,
        (BitLane::W64, true) => Op::I64_ROTL,
        (BitLane::W64, false) => Op::I64_ROTR,
    };
    chunk.emit_op(op, line);
    widen(chunk, lane, line);
}

/// Compare numeric word operands after preserving their complete 32-bit lanes.
pub fn emit_word_comparison(chunk: &mut Chunk, opcode: Op, line: u32) {
    emit_word_comparison_with_normalized_operands(chunk, opcode, (false, false), line);
}

pub fn emit_word_comparison_with_normalized_operands(
    chunk: &mut Chunk, opcode: Op, normalized: (bool, bool), line: u32,
) {
    if normalized.0 {
        if !normalized.1 { narrow(chunk, BitLane::W32, line); }
        chunk.emit_op(opcode, line);
        return;
    }
    let right = chunk.alloc_scratch(1);
    if !normalized.1 { narrow(chunk, BitLane::W32, line); }
    chunk.emit_op_u16(Op::LOCAL_SET, right, line);
    if !normalized.0 { narrow(chunk, BitLane::W32, line); }
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    chunk.emit_op(opcode, line);
}

/// Stack `[left, right]` -> one numeric result, using the declared integer
/// lane instead of the profile's default arithmetic slot.
pub fn emit_integer_binary(chunk: &mut Chunk, lane: BitLane, op: IntOp, line: u32) {
    emit_integer_binary_with_normalized_operands(chunk, lane, op, (false, false), line);
}

/// Operand proofs come from explicit integer-lane expressions, not type names.
pub fn emit_integer_binary_with_normalized_operands(
    chunk: &mut Chunk, lane: BitLane, op: IntOp, normalized: (bool, bool), line: u32,
) {
    if lane == BitLane::W32 && normalized.0 {
        // The left word needs no conversion. Normalize only the top operand;
        // both words then remain in the instruction's original stack order.
        if !normalized.1 {
            narrow(chunk, lane, line);
        }
        chunk.emit_op(integer_opcode(lane, op), line);
        return;
    }
    if lane == BitLane::W64 && op.is_comparison() {
        emit_wide_word_comparison(chunk, op, line);
        return;
    }
    {
        let right = if lane == BitLane::W64 {
            raw_i64_scratch(chunk)
        } else {
            chunk.alloc_scratch(1)
        };
        let narrow_operand = |chunk: &mut Chunk| {
            if op.is_comparison() {
                narrow_comparison_word(chunk, lane, line);
            } else {
                narrow(chunk, lane, line);
            }
        };
        if !normalized.1 { narrow_operand(chunk); }
        chunk.emit_op_u16(Op::LOCAL_SET, right, line);
        if !normalized.0 { narrow_operand(chunk); }
        chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    }
    chunk.emit_op(integer_opcode(lane, op), line);
    if !op.is_comparison() {
        widen(chunk, lane, line);
    }
}

pub(crate) fn integer_opcode(lane: BitLane, op: IntOp) -> Op {
    match (lane, op) {
        (BitLane::W32, IntOp::Add) => Op::I32_ADD,
        (BitLane::W32, IntOp::Sub) => Op::I32_SUB,
        (BitLane::W32, IntOp::Mul) => Op::I32_MUL,
        (BitLane::W32, IntOp::DivS) => Op::I32_DIV_S,
        (BitLane::W32, IntOp::DivU) => Op::I32_DIV_U,
        (BitLane::W32, IntOp::RemS) => Op::I32_REM_S,
        (BitLane::W32, IntOp::RemU) => Op::I32_REM_U,
        (BitLane::W32, IntOp::Shl) => Op::I32_SHL,
        (BitLane::W32, IntOp::ShrS) => Op::I32_SHR_S,
        (BitLane::W32, IntOp::ShrU) => Op::I32_SHR_U,
        (BitLane::W32, IntOp::And) => Op::I32_AND,
        (BitLane::W32, IntOp::Or) => Op::I32_OR,
        (BitLane::W32, IntOp::Xor) => Op::I32_XOR,
        (BitLane::W32, IntOp::Eq) => Op::I32_EQ,
        (BitLane::W32, IntOp::Ne) => Op::I32_NE,
        (BitLane::W32, IntOp::LtS) => Op::I32_LT_S,
        (BitLane::W32, IntOp::LtU) => Op::I32_LT_U,
        (BitLane::W32, IntOp::LeS) => Op::I32_LE_S,
        (BitLane::W32, IntOp::LeU) => Op::I32_LE_U,
        (BitLane::W32, IntOp::GtS) => Op::I32_GT_S,
        (BitLane::W32, IntOp::GtU) => Op::I32_GT_U,
        (BitLane::W32, IntOp::GeS) => Op::I32_GE_S,
        (BitLane::W32, IntOp::GeU) => Op::I32_GE_U,
        (BitLane::W64, IntOp::Add) => Op::I64_ADD,
        (BitLane::W64, IntOp::Sub) => Op::I64_SUB,
        (BitLane::W64, IntOp::Mul) => Op::I64_MUL,
        (BitLane::W64, IntOp::DivS) => Op::I64_DIV_S,
        (BitLane::W64, IntOp::DivU) => Op::I64_DIV_U,
        (BitLane::W64, IntOp::RemS) => Op::I64_REM_S,
        (BitLane::W64, IntOp::RemU) => Op::I64_REM_U,
        (BitLane::W64, IntOp::Shl) => Op::I64_SHL,
        (BitLane::W64, IntOp::ShrS) => Op::I64_SHR_S,
        (BitLane::W64, IntOp::ShrU) => Op::I64_SHR_U,
        (BitLane::W64, IntOp::And) => Op::I64_AND,
        (BitLane::W64, IntOp::Or) => Op::I64_OR,
        (BitLane::W64, IntOp::Xor) => Op::I64_XOR,
        (BitLane::W64, IntOp::Eq) => Op::I64_EQ,
        (BitLane::W64, IntOp::Ne) => Op::I64_NE,
        (BitLane::W64, IntOp::LtS) => Op::I64_LT_S,
        (BitLane::W64, IntOp::LtU) => Op::I64_LT_U,
        (BitLane::W64, IntOp::LeS) => Op::I64_LE_S,
        (BitLane::W64, IntOp::LeU) => Op::I64_LE_U,
        (BitLane::W64, IntOp::GtS) => Op::I64_GT_S,
        (BitLane::W64, IntOp::GtU) => Op::I64_GT_U,
        (BitLane::W64, IntOp::GeS) => Op::I64_GE_S,
        (BitLane::W64, IntOp::GeU) => Op::I64_GE_U,
    }
}

/// Read a value's STORAGE as another type of the same width — a bit cast.
///
/// The fifth member of the family, and the one every language spells
/// differently: Fortran `TRANSFER`, Go `math.Float32bits`, Java
/// `Float.floatToIntBits`, C#/VB `BitConverter.SingleToInt32Bits`, C `union`,
/// C++ `std::bit_cast`, Rust `transmute`, JS `DataView`. Before this there were
/// five implementations — two correct, one that returned its argument
/// unchanged (VB `BitConverter` answered `1` for `SingleToInt32Bits(1.0f)`
/// where .NET answers `1065353216`), and two that did not resolve at all.
///
/// The runtime carries numbers as f64, so the 32-bit lanes narrow on the way in
/// and widen on the way out; the 64-bit lanes are already the right width.
pub fn emit_reinterpret(chunk: &mut Chunk, to: NumericRepr, line: u32) {
    match to {
        // f32 → i32
        NumericRepr::I32 => {
            chunk.emit_op(Op::F32_DEMOTE_F64, line);
            chunk.emit_op(Op::I32_REINTERPRET_F32, line);
        }
        // i32 → f32
        NumericRepr::F32 => {
            chunk.emit_op(Op::F32_REINTERPRET_I32, line);
            chunk.emit_op(Op::F64_PROMOTE_F32, line);
        }
        // f64 → i64
        NumericRepr::I64 => chunk.emit_op(Op::I64_REINTERPRET_F64, line),
        // i64 → f64
        NumericRepr::F64 => chunk.emit_op(Op::F64_REINTERPRET_I64, line),
    }
}

/// Does this region's [`ShiftOverflow`] policy need a guard around a shift?
///
/// wasm masks the count, so `Mask` is already what the bare instruction does
/// and costs nothing. `Zero` is the one that needs code.
pub fn shift_needs_guard(policy: ShiftOverflow) -> bool {
    policy == ShiftOverflow::Zero
}

// ── Elementwise bit operations over a boolean ARRAY ────────────────────────
//
// A bit set is a boolean sequence, and combining two of them elementwise is the
// same operation the scalar helpers above perform on one word: .NET's
// `BitArray`, Java's `BitSet`, Python's `bitarray`, C++'s `vector<bool>`.
//
// The operators mutate the receiver and return it, which is what those APIs
// document and what lets calls chain.

use crate::primitives::collections::emit_import_call as bits_import_call;

/// The combining instruction for one elementwise pass, or `None` to negate.
fn emit_elementwise(chunks: &mut [Chunk], current: usize, op: Option<Op>, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let me = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    if op.is_some() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, other, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, me, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
    bits_import_call(chunks, current, "ecma:array", "length", 1, line);
    // `| 0` is the narrowing every i32 site in this file uses: the `I32_*`
    // instructions coerce their operands dynamically.
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    let done = chunks[current].emit_block(line);
    let (again, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    bits_import_call(chunks, current, "ecma:array", "get", 2, line);
    bits_import_call(chunks, current, "wasm:js-boolean", "cast", 1, line);
    match op {
        Some(instr) => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, other, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
            bits_import_call(chunks, current, "ecma:array", "get", 2, line);
            bits_import_call(chunks, current, "wasm:js-boolean", "cast", 1, line);
            chunks[current].emit_op(instr, line);
        }
        None => chunks[current].emit_op(Op::I32_EQZ, line),
    }
    // A real boolean, not the i32: under `materialize_bool_results` an i32 1
    // does not compare equal to the language's `true`.
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);

    bits_import_call(chunks, current, "ecma:array", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(again);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
}

/// Elementwise AND, in place. Stack: `[a, b]` → `[a]`.
pub fn emit_array_and(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, Some(Op::I32_AND), line);
}

/// Elementwise OR, in place. Stack: `[a, b]` → `[a]`.
pub fn emit_array_or(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, Some(Op::I32_OR), line);
}

/// Elementwise XOR, in place. Stack: `[a, b]` → `[a]`.
pub fn emit_array_xor(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, Some(Op::I32_XOR), line);
}

/// Elementwise NOT, in place. Stack: `[a]` → `[a]`.
pub fn emit_array_not(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, None, line);
}

/// Position of the highest set bit — `31 - clz` in the 32-bit lane.
///
/// An integer operation, not a floating-point logarithm: `log2(64)` is exactly
/// 6 where `ln(64)/ln(2)` can land a fraction below it.
///
/// Stack: `[x]` → `[n]`.
pub fn emit_log2(chunks: &mut [Chunk], current: usize, line: u32) {
    let tmp = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_CLZ, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tmp, line);
    chunk.emit_i32_const(31, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tmp, line);
    chunk.emit_op(Op::I32_SUB, line);
}

/// Whether exactly one bit is set. Zero is NOT a power of two, which the
/// `x & (x - 1)` test alone would report as one.
///
/// Stack: `[x]` → `[bool]`.
pub fn emit_is_pow2(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// The smallest power of two at or above `x`.
///
/// 0 and 1 are returned unchanged: 0 has no power of two at or above it that
/// fits the lane, and both would otherwise shift by the lane's full 32, which
/// wasm takes modulo the width and turns into a shift of zero.
///
/// Stack: `[x]` → `[n]`.
pub fn emit_round_up_pow2(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_LE_S, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_CLZ, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_SHL, line);
    chunk.emit_end(line);
}

/// Reinterpret the 32-bit lane's signed result as unsigned.
///
/// The lane's instructions are signed, so a value with the high bit set comes
/// back negative. Languages whose bit types are unsigned — .NET's `uint`,
/// Java's `Integer.toUnsignedLong`, JS `>>> 0` — need the other reading, which
/// is what `f64.convert_i32_u` performs.
///
/// Stack: `[i32]` → `[n]`.
pub fn emit_as_unsigned32(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
}

/// Wrap an integral f64 through the low 32 bits and expose the unsigned value.
/// Stack: `[f64]` -> `[f64]`. Callers must keep inputs within the i64 range.
pub fn emit_wrap_unsigned32(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
}

/// 32-bit rotate with an UNSIGNED result. Stack: `[value, count]` → `[n]`.
pub fn emit_rotl32_unsigned(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_rotate(&mut chunks[current], BitLane::W32, true, line);
    emit_as_unsigned32(&mut chunks[current], line);
}

/// 32-bit rotate right with an UNSIGNED result. Stack: `[value, count]` → `[n]`.
pub fn emit_rotr32_unsigned(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_rotate(&mut chunks[current], BitLane::W32, false, line);
    emit_as_unsigned32(&mut chunks[current], line);
}
