// Frozen pre-optimization classifier emitter for differential checks.
use vybe_runtime::{Chunk, Op};
fn emit_length(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "length");
    c.emit_call(idx, 1, line);
}
/// One code-point class, keyed to a generated range table.
#[derive(Clone, Copy)]
enum CharClass {
    Digit,
    Alpha,
    Alnum,
    Space,
    Upper,
    Lower,
}

/// Append one range boundary as two sub-surrogate UTF-16 units.
fn push_boundary(out: &mut String, b: u32) {
    out.push(char::from_u32(b >> 12).expect("boundary hi unit"));
    out.push(char::from_u32(b & 0xFFF).expect("boundary lo unit"));
}

/// Flatten a predicate over the whole code-point space into sorted
/// `[start, end)` boundary pairs. A code point is IN the class iff the
/// count of boundaries `<= cp` is odd.
fn build_class_table(pred: fn(u32) -> bool) -> String {
    let mut out = String::new();
    let mut start: Option<u32> = None;
    for cp in 0..=0x10FFFF_u32 {
        match (pred(cp), start) {
            (true, None) => start = Some(cp),
            (false, Some(s)) => {
                push_boundary(&mut out, s);
                push_boundary(&mut out, cp);
                start = None;
            }
            _ => {}
        }
    }
    if let Some(s) = start {
        push_boundary(&mut out, s);
        push_boundary(&mut out, 0x110000);
    }
    out
}

fn class_table(class: CharClass) -> &'static str {
    use std::sync::OnceLock;
    static DIGIT: OnceLock<String> = OnceLock::new();
    static ALPHA: OnceLock<String> = OnceLock::new();
    static ALNUM: OnceLock<String> = OnceLock::new();
    static SPACE: OnceLock<String> = OnceLock::new();
    static UPPER: OnceLock<String> = OnceLock::new();
    static LOWER: OnceLock<String> = OnceLock::new();
    fn digit(cp: u32) -> bool {
        (0x30..=0x39).contains(&cp)
    }
    fn alpha(cp: u32) -> bool {
        char::from_u32(cp).is_some_and(|c| c.is_alphabetic())
    }
    fn alnum(cp: u32) -> bool {
        alpha(cp) || digit(cp)
    }
    fn space(cp: u32) -> bool {
        char::from_u32(cp).is_some_and(|c| c.is_whitespace())
    }
    fn upper(cp: u32) -> bool {
        char::from_u32(cp).is_some_and(|c| c.is_uppercase())
    }
    fn lower(cp: u32) -> bool {
        char::from_u32(cp).is_some_and(|c| c.is_lowercase())
    }
    match class {
        CharClass::Digit => DIGIT.get_or_init(|| build_class_table(digit)),
        CharClass::Alpha => ALPHA.get_or_init(|| build_class_table(alpha)),
        CharClass::Alnum => ALNUM.get_or_init(|| build_class_table(alnum)),
        CharClass::Space => SPACE.get_or_init(|| build_class_table(space)),
        CharClass::Upper => UPPER.get_or_init(|| build_class_table(upper)),
        CharClass::Lower => LOWER.get_or_init(|| build_class_table(lower)),
    }
}

/// Membership test. Stack: `[cp: i32]` → `[i32 0/1]` — binary search for
/// the first boundary `> cp`; inside a range iff that index is odd.
fn emit_class_test(chunks: &mut [Chunk], current: usize, class: CharClass, line: u32) {
    let table = class_table(class);
    let n_bounds = (table.chars().count() / 2) as i32;
    let base = chunks[current].alloc_scratch(4);
    let (cp, lo, hi, mid) = (base, base + 1, base + 2, base + 3);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_SET, cp, line);
        c.emit_i32_const(0, line);
        c.emit_op_u16(Op::LOCAL_SET, lo, line);
        c.emit_i32_const(n_bounds, line);
        c.emit_op_u16(Op::LOCAL_SET, hi, line);
    }
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, lo, line);
        c.emit_op_u16(Op::LOCAL_GET, hi, line);
        c.emit_op(Op::I32_LT_S, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    {
        let c = &mut chunks[current];
        let cca = c.add_import("wasm:js-string", "charCodeAt");
        // mid = (lo + hi) >> 1
        c.emit_op_u16(Op::LOCAL_GET, lo, line);
        c.emit_op_u16(Op::LOCAL_GET, hi, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_SHR_U, line);
        c.emit_op_u16(Op::LOCAL_SET, mid, line);
        // boundary = charCodeAt(T, 2*mid) << 12 | charCodeAt(T, 2*mid + 1)
        c.emit_string_const(table, line);
        c.emit_op_u16(Op::LOCAL_GET, mid, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_SHL, line);
        c.emit_call(cca, 2, line);
        c.emit_i32_const(12, line);
        c.emit_op(Op::I32_SHL, line);
        c.emit_string_const(table, line);
        c.emit_op_u16(Op::LOCAL_GET, mid, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_SHL, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_call(cca, 2, line);
        c.emit_op(Op::I32_OR, line);
        // boundary <= cp ? lo = mid + 1 : hi = mid
        c.emit_op_u16(Op::LOCAL_GET, cp, line);
        c.emit_op(Op::I32_LE_S, line);
        c.emit_if(line);
        c.emit_op_u16(Op::LOCAL_GET, mid, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_SET, lo, line);
        c.emit_else(line);
        c.emit_op_u16(Op::LOCAL_GET, mid, line);
        c.emit_op_u16(Op::LOCAL_SET, hi, line);
        c.emit_end(line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, lo, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_AND, line);
    }
}

/// Non-empty AND every code point in `class`. Stack: `[string]` → `[bool]`.
/// The `i += cp > 0xFFFF ? 2 : 1` step is the `emit_byte_length` walk.
fn emit_all_in_class(chunks: &mut [Chunk], current: usize, class: CharClass, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let (s, i, n, ok, cp) = (base, base + 1, base + 2, base + 3, base + 4);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_SET, s, line);
        c.emit_i32_const(0, line);
        c.emit_op_u16(Op::LOCAL_SET, i, line);
        c.emit_op_u16(Op::LOCAL_GET, s, line);
        emit_length(c, line);
        c.emit_op_u16(Op::LOCAL_SET, n, line);
        c.emit_op_u16(Op::LOCAL_GET, n, line);
        c.emit_i32_const(0, line);
        c.emit_op(Op::I32_GT_S, line);
        c.emit_op_u16(Op::LOCAL_SET, ok, line);
    }
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, i, line);
        c.emit_op_u16(Op::LOCAL_GET, n, line);
        c.emit_op(Op::I32_LT_S, line);
        c.emit_op_u16(Op::LOCAL_GET, ok, line);
        c.emit_op(Op::I32_AND, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    {
        let c = &mut chunks[current];
        let cpa = c.add_import("wasm:js-string", "codePointAt");
        c.emit_op_u16(Op::LOCAL_GET, s, line);
        c.emit_op_u16(Op::LOCAL_GET, i, line);
        c.emit_call(cpa, 2, line);
        c.emit_op_u16(Op::LOCAL_SET, cp, line);
        c.emit_op_u16(Op::LOCAL_GET, cp, line);
    }
    emit_class_test(chunks, current, class, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_SET, ok, line);
        c.emit_op_u16(Op::LOCAL_GET, i, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_GET, cp, line);
        c.emit_i32_const(65535, line);
        c.emit_op(Op::I32_GT_S, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_SET, i, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, ok, line);
        vybe_compiler::primitives::ops::emit_i32_to_bool(c, line);
    }
}

/// At least one cased code point and none of the opposite case — the
/// shared shape of `is_upper` (`want_upper`) and `is_lower`. Empty and
/// caseless strings answer false in both directions.
fn emit_cased_class(chunks: &mut [Chunk], current: usize, want_upper: bool, line: u32) {
    let (same, opposite) = if want_upper {
        (CharClass::Upper, CharClass::Lower)
    } else {
        (CharClass::Lower, CharClass::Upper)
    };
    let base = chunks[current].alloc_scratch(6);
    let (s, i, n, bad, has, cp) = (base, base + 1, base + 2, base + 3, base + 4, base + 5);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_SET, s, line);
        c.emit_i32_const(0, line);
        c.emit_op_u16(Op::LOCAL_SET, i, line);
        c.emit_op_u16(Op::LOCAL_GET, s, line);
        emit_length(c, line);
        c.emit_op_u16(Op::LOCAL_SET, n, line);
        c.emit_i32_const(0, line);
        c.emit_op_u16(Op::LOCAL_SET, bad, line);
        c.emit_i32_const(0, line);
        c.emit_op_u16(Op::LOCAL_SET, has, line);
    }
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, i, line);
        c.emit_op_u16(Op::LOCAL_GET, n, line);
        c.emit_op(Op::I32_LT_S, line);
        c.emit_op_u16(Op::LOCAL_GET, bad, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_op(Op::I32_AND, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    {
        let c = &mut chunks[current];
        let cpa = c.add_import("wasm:js-string", "codePointAt");
        c.emit_op_u16(Op::LOCAL_GET, s, line);
        c.emit_op_u16(Op::LOCAL_GET, i, line);
        c.emit_call(cpa, 2, line);
        c.emit_op_u16(Op::LOCAL_SET, cp, line);
        c.emit_op_u16(Op::LOCAL_GET, cp, line);
    }
    emit_class_test(chunks, current, opposite, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_SET, bad, line);
        c.emit_op_u16(Op::LOCAL_GET, has, line);
        c.emit_op_u16(Op::LOCAL_GET, cp, line);
    }
    emit_class_test(chunks, current, same, line);
    {
        let c = &mut chunks[current];
        c.emit_op(Op::I32_OR, line);
        c.emit_op_u16(Op::LOCAL_SET, has, line);
        c.emit_op_u16(Op::LOCAL_GET, i, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_GET, cp, line);
        c.emit_i32_const(65535, line);
        c.emit_op(Op::I32_GT_S, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_SET, i, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, bad, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_op_u16(Op::LOCAL_GET, has, line);
        c.emit_op(Op::I32_AND, line);
        vybe_compiler::primitives::ops::emit_i32_to_bool(c, line);
    }
}

/// `common:str_is_digit` — ASCII `0-9` (see the section note). `[string]` → `[bool]`.
pub fn emit_is_digit(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_all_in_class(chunks, current, CharClass::Digit, line);
}

/// `common:str_is_alpha` — Unicode Alphabetic. `[string]` → `[bool]`.
pub fn emit_is_alpha(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_all_in_class(chunks, current, CharClass::Alpha, line);
}

/// `common:str_is_alnum` — Alphabetic or ASCII digit. `[string]` → `[bool]`.
pub fn emit_is_alnum(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_all_in_class(chunks, current, CharClass::Alnum, line);
}

/// `common:str_is_space` — Unicode White_Space. `[string]` → `[bool]`.
pub fn emit_is_space(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_all_in_class(chunks, current, CharClass::Space, line);
}

/// `common:str_is_upper` — some cased, none lowercase. `[string]` → `[bool]`.
pub fn emit_is_upper(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_cased_class(chunks, current, true, line);
}

/// `common:str_is_lower` — some cased, none uppercase. `[string]` → `[bool]`.
pub fn emit_is_lower(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_cased_class(chunks, current, false, line);
}
