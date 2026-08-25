//! Python `colorsys` — RGB ↔ YIQ / HLS / HSV conversions.
//!
//! Pure arithmetic, so it is pure bytecode: no host function exists for any
//! of it and none is needed. The formulae are CPython's `Lib/colorsys.py`
//! verbatim, including the two places that look like simplifiable algebra but
//! are not — `2.0-maxc-minc` (CPython gh-106498, NOT `2.0-sumc`) and the
//! truncating `int(h*6.0)` in `hsv_to_rgb`.
//!
//! Everything runs on raw `f64`: arguments are coerced once through
//! `wasm:js-number.toF64`, and `Value::F64` is what the arithmetic opcodes
//! produce, so the results drop straight into a tuple.

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::tuples;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const ONE_THIRD: f64 = 1.0 / 3.0;
const ONE_SIXTH: f64 = 1.0 / 6.0;
const TWO_THIRD: f64 = 2.0 / 3.0;

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn konst(chunk: &mut Chunk, value: f64, line: u32) {
    core_wasm::f64_const(chunk, line, value);
}

/// Pop `argc` arguments into three consecutive scratch slots, each coerced to
/// a raw f64. Missing arguments (an under-applied call) read as 0.0.
fn three_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> (u16, u16, u16) {
    let base = chunks[current].alloc_scratch(3);
    for offset in (0..3u16).rev() {
        if (offset as u8) < argc {
            lset(&mut chunks[current], base + offset, line);
        } else {
            konst(&mut chunks[current], 0.0, line);
            lset(&mut chunks[current], base + offset, line);
        }
    }
    for _ in 3..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    for offset in 0..3u16 {
        lget(&mut chunks[current], base + offset, line);
        chunks[current].emit_call(to_f64, 1, line);
        lset(&mut chunks[current], base + offset, line);
    }
    (base, base + 1, base + 2)
}

/// Python's `x % 1.0` — the fractional part with the sign of the DIVISOR, so
/// `-0.25 % 1.0` is `0.75`. `x - floor(x)` is exactly that. Stack: `[x]` → `[x']`.
fn emit_mod1(chunk: &mut Chunk, line: u32) {
    chunk.emit_dup(line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op(Op::F64_SUB, line);
}

/// `max(a, b, c)` / `min(a, b, c)` over three slots. Stack: `[]` → `[f64]`.
fn emit_max3(chunk: &mut Chunk, a: u16, b: u16, c: u16, line: u32) {
    lget(chunk, a, line);
    lget(chunk, b, line);
    chunk.emit_op(Op::F64_MAX, line);
    lget(chunk, c, line);
    chunk.emit_op(Op::F64_MAX, line);
}

fn emit_min3(chunk: &mut Chunk, a: u16, b: u16, c: u16, line: u32) {
    lget(chunk, a, line);
    lget(chunk, b, line);
    chunk.emit_op(Op::F64_MIN, line);
    lget(chunk, c, line);
    chunk.emit_op(Op::F64_MIN, line);
}

/// `(x - y) * k`. Stack: `[]` → `[f64]`.
fn emit_diff_scaled(chunk: &mut Chunk, x: u16, y: u16, k: f64, line: u32) {
    lget(chunk, x, line);
    lget(chunk, y, line);
    chunk.emit_op(Op::F64_SUB, line);
    konst(chunk, k, line);
    chunk.emit_op(Op::F64_MUL, line);
}

/// The hue both `rgb_to_hls` and `rgb_to_hsv` compute, given the channels plus
/// their max and range (which the callers already have). Stack: `[]` → `[h]`.
fn emit_hue(chunk: &mut Chunk, r: u16, g: u16, b: u16, maxc: u16, rangec: u16, line: u32) {
    let rc = chunk.alloc_scratch(3);
    let gc = rc + 1;
    let bc = rc + 2;
    for (slot, channel) in [(rc, r), (gc, g), (bc, b)] {
        lget(chunk, maxc, line);
        lget(chunk, channel, line);
        chunk.emit_op(Op::F64_SUB, line);
        lget(chunk, rangec, line);
        chunk.emit_op(Op::F64_DIV, line);
        lset(chunk, slot, line);
    }

    lget(chunk, r, line);
    lget(chunk, maxc, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    // h = bc - gc
    lget(chunk, bc, line);
    lget(chunk, gc, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_else(line);
    lget(chunk, g, line);
    lget(chunk, maxc, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    // h = 2.0 + rc - bc
    konst(chunk, 2.0, line);
    lget(chunk, rc, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, bc, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_else(line);
    // h = 4.0 + gc - rc
    konst(chunk, 4.0, line);
    lget(chunk, gc, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, rc, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // h = (h / 6.0) % 1.0
    konst(chunk, 6.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    emit_mod1(chunk, line);
}

/// `colorsys.rgb_to_yiq(r, g, b)` → `(y, i, q)`.
pub fn emit_rgb_to_yiq(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (r, g, b) = three_args(chunks, current, argc, line);
    let chunk = &mut chunks[current];
    let y = chunk.alloc_scratch(1);

    // y = 0.30*r + 0.59*g + 0.11*b
    lget(chunk, r, line);
    konst(chunk, 0.30, line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, g, line);
    konst(chunk, 0.59, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, b, line);
    konst(chunk, 0.11, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, y, line);

    lget(chunk, y, line);
    // i = 0.74*(r-y) - 0.27*(b-y)
    emit_diff_scaled(chunk, r, y, 0.74, line);
    emit_diff_scaled(chunk, b, y, 0.27, line);
    chunk.emit_op(Op::F64_SUB, line);
    // q = 0.48*(r-y) + 0.41*(b-y)
    emit_diff_scaled(chunk, r, y, 0.48, line);
    emit_diff_scaled(chunk, b, y, 0.41, line);
    chunk.emit_op(Op::F64_ADD, line);

    tuples::emit_tuple(chunks, current, 3, line);
}

/// `colorsys.yiq_to_rgb(y, i, q)` → `(r, g, b)`, each clamped to [0, 1].
pub fn emit_yiq_to_rgb(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (y, i, q) = three_args(chunks, current, argc, line);
    let chunk = &mut chunks[current];

    // Coefficients are CPython's precomputed inverse of the FCC NTSC matrix.
    for (ki, kq) in [
        (0.9468822170900693_f64, 0.6235565819861433_f64),
        (-0.27478764629897834, -0.6356910791873801),
        (-1.1085450346420322, 1.7090069284064666),
    ] {
        lget(chunk, y, line);
        lget(chunk, i, line);
        konst(chunk, ki, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        lget(chunk, q, line);
        konst(chunk, kq, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        // clamp
        konst(chunk, 0.0, line);
        chunk.emit_op(Op::F64_MAX, line);
        konst(chunk, 1.0, line);
        chunk.emit_op(Op::F64_MIN, line);
    }

    tuples::emit_tuple(chunks, current, 3, line);
}

/// `colorsys.rgb_to_hls(r, g, b)` → `(h, l, s)`.
pub fn emit_rgb_to_hls(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (r, g, b) = three_args(chunks, current, argc, line);
    let chunk = &mut chunks[current];
    let maxc = chunk.alloc_scratch(4);
    let minc = maxc + 1;
    let rangec = maxc + 2;
    let l = maxc + 3;

    emit_max3(chunk, r, g, b, line);
    lset(chunk, maxc, line);
    emit_min3(chunk, r, g, b, line);
    lset(chunk, minc, line);
    lget(chunk, maxc, line);
    lget(chunk, minc, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, rangec, line);
    // l = (maxc + minc) / 2.0
    lget(chunk, maxc, line);
    lget(chunk, minc, line);
    chunk.emit_op(Op::F64_ADD, line);
    konst(chunk, 2.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    lset(chunk, l, line);

    lget(chunk, minc, line);
    lget(chunk, maxc, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    // grey: (0.0, l, 0.0)
    konst(chunk, 0.0, line);
    lget(chunk, l, line);
    konst(chunk, 0.0, line);
    tuples::emit_tuple(chunks, current, 3, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    emit_hue(chunk, r, g, b, maxc, rangec, line);
    lget(chunk, l, line);
    // s = l <= 0.5 ? rangec/(maxc+minc) : rangec/(2.0-maxc-minc)
    lget(chunk, l, line);
    konst(chunk, 0.5, line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_if_value(line);
    lget(chunk, rangec, line);
    lget(chunk, maxc, line);
    lget(chunk, minc, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_else(line);
    lget(chunk, rangec, line);
    konst(chunk, 2.0, line);
    lget(chunk, maxc, line);
    chunk.emit_op(Op::F64_SUB, line);
    lget(chunk, minc, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_end(line);
    tuples::emit_tuple(chunks, current, 3, line);
    chunks[current].emit_end(line);
}

/// CPython's `colorsys._v(m1, m2, hue)`. `hue` is on the stack; `m1`/`m2` are
/// slots. Stack: `[hue]` → `[component]`.
fn emit_v(chunk: &mut Chunk, m1: u16, m2: u16, line: u32) {
    let hue = chunk.alloc_scratch(1);
    emit_mod1(chunk, line);
    lset(chunk, hue, line);

    lget(chunk, hue, line);
    konst(chunk, ONE_SIXTH, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    // m1 + (m2-m1) * hue * 6.0
    lget(chunk, m1, line);
    lget(chunk, m2, line);
    lget(chunk, m1, line);
    chunk.emit_op(Op::F64_SUB, line);
    lget(chunk, hue, line);
    chunk.emit_op(Op::F64_MUL, line);
    konst(chunk, 6.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);

    lget(chunk, hue, line);
    konst(chunk, 0.5, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    lget(chunk, m2, line);
    chunk.emit_else(line);

    lget(chunk, hue, line);
    konst(chunk, TWO_THIRD, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    // m1 + (m2-m1) * (TWO_THIRD - hue) * 6.0
    lget(chunk, m1, line);
    lget(chunk, m2, line);
    lget(chunk, m1, line);
    chunk.emit_op(Op::F64_SUB, line);
    konst(chunk, TWO_THIRD, line);
    lget(chunk, hue, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_MUL, line);
    konst(chunk, 6.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    lget(chunk, m1, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `colorsys.hls_to_rgb(h, l, s)` → `(r, g, b)`.
pub fn emit_hls_to_rgb(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (h, l, s) = three_args(chunks, current, argc, line);
    let chunk = &mut chunks[current];
    let m1 = chunk.alloc_scratch(2);
    let m2 = m1 + 1;

    lget(chunk, s, line);
    konst(chunk, 0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    lget(chunk, l, line);
    lget(chunk, l, line);
    lget(chunk, l, line);
    tuples::emit_tuple(chunks, current, 3, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    // m2 = l <= 0.5 ? l*(1.0+s) : l+s-(l*s)
    lget(chunk, l, line);
    konst(chunk, 0.5, line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_if_value(line);
    lget(chunk, l, line);
    konst(chunk, 1.0, line);
    lget(chunk, s, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_else(line);
    lget(chunk, l, line);
    lget(chunk, s, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, l, line);
    lget(chunk, s, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_end(line);
    lset(chunk, m2, line);
    // m1 = 2.0*l - m2
    konst(chunk, 2.0, line);
    lget(chunk, l, line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, m2, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, m1, line);

    for offset in [ONE_THIRD, 0.0, -ONE_THIRD] {
        lget(chunk, h, line);
        if offset != 0.0 {
            konst(chunk, offset, line);
            chunk.emit_op(Op::F64_ADD, line);
        }
        emit_v(chunk, m1, m2, line);
    }
    tuples::emit_tuple(chunks, current, 3, line);
    chunks[current].emit_end(line);
}

/// `colorsys.rgb_to_hsv(r, g, b)` → `(h, s, v)`.
pub fn emit_rgb_to_hsv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (r, g, b) = three_args(chunks, current, argc, line);
    let chunk = &mut chunks[current];
    let maxc = chunk.alloc_scratch(3);
    let minc = maxc + 1;
    let rangec = maxc + 2;

    emit_max3(chunk, r, g, b, line);
    lset(chunk, maxc, line);
    emit_min3(chunk, r, g, b, line);
    lset(chunk, minc, line);
    lget(chunk, maxc, line);
    lget(chunk, minc, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, rangec, line);

    lget(chunk, minc, line);
    lget(chunk, maxc, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    // grey: (0.0, 0.0, v) where v is maxc
    konst(chunk, 0.0, line);
    konst(chunk, 0.0, line);
    lget(chunk, maxc, line);
    tuples::emit_tuple(chunks, current, 3, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    emit_hue(chunk, r, g, b, maxc, rangec, line);
    // s = rangec / maxc
    lget(chunk, rangec, line);
    lget(chunk, maxc, line);
    chunk.emit_op(Op::F64_DIV, line);
    // v = maxc
    lget(chunk, maxc, line);
    tuples::emit_tuple(chunks, current, 3, line);
    chunks[current].emit_end(line);
}

/// `colorsys.hsv_to_rgb(h, s, v)` → `(r, g, b)`.
pub fn emit_hsv_to_rgb(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (h, s, v) = three_args(chunks, current, argc, line);
    let chunk = &mut chunks[current];
    let i = chunk.alloc_scratch(5);
    let f = i + 1;
    let p = i + 2;
    let q = i + 3;
    let t = i + 4;

    lget(chunk, s, line);
    konst(chunk, 0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    lget(chunk, v, line);
    lget(chunk, v, line);
    lget(chunk, v, line);
    tuples::emit_tuple(chunks, current, 3, line);
    let mut chunk = &mut chunks[current];
    chunk.emit_else(line);

    // i = int(h*6.0) — CPython truncates, so f64.trunc, not floor.
    lget(chunk, h, line);
    konst(chunk, 6.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    lset(chunk, i, line);
    // f = h*6.0 - i
    lget(chunk, h, line);
    konst(chunk, 6.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, i, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, f, line);
    // p = v*(1.0-s)
    lget(chunk, v, line);
    konst(chunk, 1.0, line);
    lget(chunk, s, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_MUL, line);
    lset(chunk, p, line);
    // q = v*(1.0 - s*f)
    lget(chunk, v, line);
    konst(chunk, 1.0, line);
    lget(chunk, s, line);
    lget(chunk, f, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_MUL, line);
    lset(chunk, q, line);
    // t = v*(1.0 - s*(1.0-f))
    lget(chunk, v, line);
    konst(chunk, 1.0, line);
    lget(chunk, s, line);
    konst(chunk, 1.0, line);
    lget(chunk, f, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_MUL, line);
    lset(chunk, t, line);
    // i = i % 6 (Python modulo — i is non-negative for h in [0,1], and
    // floor-mod keeps a negative h behaving like CPython's)
    lget(chunk, i, line);
    lget(chunk, i, line);
    konst(chunk, 6.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    konst(chunk, 6.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i, line);

    // The sextant table, as a chain of equality tests.
    let sextants: [(f64, [u16; 3]); 6] = [
        (0.0, [v, t, p]),
        (1.0, [q, v, p]),
        (2.0, [p, v, t]),
        (3.0, [p, q, v]),
        (4.0, [t, p, v]),
        (5.0, [v, p, q]),
    ];
    for (index, (which, _)) in sextants.iter().enumerate() {
        if index == sextants.len() - 1 {
            break;
        }
        lget(chunk, i, line);
        konst(chunk, *which, line);
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_if_value(line);
        for slot in sextants[index].1 {
            lget(chunk, slot, line);
        }
        tuples::emit_tuple(chunks, current, 3, line);
        chunk = &mut chunks[current];
        chunk.emit_else(line);
    }
    // i == 5 — the only remaining case.
    for slot in sextants[5].1 {
        lget(chunk, slot, line);
    }
    tuples::emit_tuple(chunks, current, 3, line);
    let chunk = &mut chunks[current];
    for _ in 0..sextants.len() - 1 {
        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}
