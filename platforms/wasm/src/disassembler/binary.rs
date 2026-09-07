//! WASM binary → WAT text.
//!
//! A `(module binary "…")` fixture states a module in its binary form. It is a
//! module like any other: it instantiates, its exports are reachable, and a
//! later `(invoke "f")` calls into it. Rendering it back to text is what lets
//! the SAME front end handle it — the text walker builds the class, the export
//! tables, the globals and the segments, and nothing about the module's
//! spelling reaches any of that.
//!
//! Two rules keep the round trip exact:
//!
//! - **Names are never emitted, only indices.** Every index space is written
//!   positionally, so no `$id` is invented and none can collide with another
//!   module's.
//! - **A float is rendered from its BITS, as a hex literal.** A decimal
//!   rendering would re-introduce the double rounding this front end already
//!   had to remove: the digits that decide a tie do not survive a round trip
//!   through decimal. `v128.const` goes further and is always written as
//!   `i8x16` byte lanes, which involve no float parsing at all.
//!
//! Anything this renderer cannot express is an ERROR, never a silent omission
//! — a module that decoded but rendered wrong would instantiate as something
//! the fixture did not state.

use crate::reader::WasmError;
use std::fmt::Write;

// ── Byte reader ───────────────────────────────────────────────────────────────

struct Reader<'a> {
    d: &'a [u8],
    p: usize,
}

impl<'a> Reader<'a> {
    fn new(d: &'a [u8]) -> Self {
        Reader { d, p: 0 }
    }

    fn done(&self) -> bool {
        self.p >= self.d.len()
    }

    fn byte(&mut self) -> Result<u8, WasmError> {
        let b = *self
            .d
            .get(self.p)
            .ok_or_else(|| WasmError::from("unexpected end of section"))?;
        self.p += 1;
        Ok(b)
    }

    fn bytes(&mut self, n: usize) -> Result<&'a [u8], WasmError> {
        let end = self
            .p
            .checked_add(n)
            .filter(|e| *e <= self.d.len())
            .ok_or_else(|| WasmError::from("unexpected end of section"))?;
        let s = &self.d[self.p..end];
        self.p = end;
        Ok(s)
    }

    fn u32(&mut self) -> Result<u32, WasmError> {
        let mut n: u64 = 0;
        let mut shift = 0;
        loop {
            let b = self.byte()?;
            n |= ((b & 0x7F) as u64) << shift;
            shift += 7;
            if b & 0x80 == 0 {
                break;
            }
            if shift > 35 {
                return Err("integer representation too long".into());
            }
        }
        Ok(n as u32)
    }

    fn u64(&mut self) -> Result<u64, WasmError> {
        let mut n: u128 = 0;
        let mut shift = 0;
        loop {
            let b = self.byte()?;
            n |= ((b & 0x7F) as u128) << shift;
            shift += 7;
            if b & 0x80 == 0 {
                break;
            }
            if shift > 70 {
                return Err("integer representation too long".into());
            }
        }
        Ok(n as u64)
    }

    fn i64(&mut self) -> Result<i64, WasmError> {
        let mut n: i64 = 0;
        let mut shift = 0u32;
        loop {
            let b = self.byte()?;
            n |= ((b & 0x7F) as i64).wrapping_shl(shift);
            shift += 7;
            if b & 0x80 == 0 {
                if shift < 64 && b & 0x40 != 0 {
                    n |= -1i64 << shift;
                }
                break;
            }
            if shift > 70 {
                return Err("integer representation too long".into());
            }
        }
        Ok(n)
    }

    fn i32(&mut self) -> Result<i32, WasmError> {
        Ok(self.i64()? as i32)
    }

    fn f32_bits(&mut self) -> Result<u32, WasmError> {
        let b = self.bytes(4)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    fn f64_bits(&mut self) -> Result<u64, WasmError> {
        let b = self.bytes(8)?;
        Ok(u64::from_le_bytes([
            b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7],
        ]))
    }

    /// A `name` is a length-prefixed byte string. It is emitted back verbatim,
    /// byte for byte — an export name is not required to be printable, and
    /// re-encoding it as text would change which name the module exports.
    fn name(&mut self) -> Result<Vec<u8>, WasmError> {
        let n = self.u32()? as usize;
        Ok(self.bytes(n)?.to_vec())
    }
}

// ── Literals ──────────────────────────────────────────────────────────────────

/// A WAT string literal for arbitrary bytes: every byte as `\XX`.
fn wat_string(bytes: &[u8]) -> String {
    let mut s = String::with_capacity(bytes.len() * 4 + 2);
    s.push('"');
    for b in bytes {
        let _ = write!(s, "\\{b:02x}");
    }
    s.push('"');
    s
}

/// An f32 as an EXACT hex literal, built from the bits.
fn f32_literal(bits: u32) -> String {
    let sign = if bits >> 31 == 1 { "-" } else { "" };
    let exp = ((bits >> 23) & 0xFF) as i32;
    let mant = bits & 0x7F_FFFF;
    // 23 mantissa bits do not fill whole hex digits; one left shift makes 24.
    let hex = format!("{:06x}", mant << 1);
    match exp {
        0xFF if mant == 0 => format!("{sign}inf"),
        0xFF => format!("{sign}nan:0x{mant:x}"),
        0 if mant == 0 => format!("{sign}0x0p+0"),
        0 => format!("{sign}0x0.{hex}p-126"),
        _ => format!("{sign}0x1.{hex}p{:+}", exp - 127),
    }
}

/// An f64 as an EXACT hex literal, built from the bits.
fn f64_literal(bits: u64) -> String {
    let sign = if bits >> 63 == 1 { "-" } else { "" };
    let exp = ((bits >> 52) & 0x7FF) as i32;
    let mant = bits & 0xF_FFFF_FFFF_FFFF;
    let hex = format!("{mant:013x}");
    match exp {
        0x7FF if mant == 0 => format!("{sign}inf"),
        0x7FF => format!("{sign}nan:0x{mant:x}"),
        0 if mant == 0 => format!("{sign}0x0p+0"),
        0 => format!("{sign}0x0.{hex}p-1022"),
        _ => format!("{sign}0x1.{hex}p{:+}", exp - 1023),
    }
}

// ── Types ─────────────────────────────────────────────────────────────────────

/// An ABSTRACT heap type's single byte. A concrete one is a type index encoded
/// as a non-negative s33 and is read by `heap_type`.
fn abs_heap_type(b: u8) -> Option<&'static str> {
    Some(match b {
        0x69 => "exn",
        0x6A => "array",
        0x6B => "struct",
        0x6C => "i31",
        0x6D => "eq",
        0x6E => "any",
        0x6F => "extern",
        0x70 => "func",
        0x71 => "none",
        0x72 => "noextern",
        0x73 => "nofunc",
        0x74 => "noexn",
        _ => return None,
    })
}

/// A heap type: one abstract byte, or a type index as a non-negative s33.
fn heap_type(r: &mut Reader) -> Result<String, WasmError> {
    let b = *r
        .d
        .get(r.p)
        .ok_or_else(|| WasmError::from("unexpected end of heap type"))?;
    if let Some(name) = abs_heap_type(b) {
        r.p += 1;
        return Ok(name.to_string());
    }
    let idx = r.i64()?;
    if idx < 0 {
        return Err(format!("cannot render heap type 0x{b:02x}").into());
    }
    Ok(idx.to_string())
}

/// A reference type: `0x63 ht` is nullable, `0x64 ht` is not, and a bare
/// abstract heap byte is the nullable short form.
fn ref_type(r: &mut Reader) -> Result<String, WasmError> {
    let b = r.byte()?;
    match b {
        0x63 => Ok(format!("(ref null {})", heap_type(r)?)),
        0x64 => Ok(format!("(ref {})", heap_type(r)?)),
        0x70 => Ok("funcref".to_string()),
        0x6F => Ok("externref".to_string()),
        _ => match abs_heap_type(b) {
            Some(ht) => Ok(format!("(ref null {ht})")),
            None => Err(format!("cannot render reference type 0x{b:02x}").into()),
        },
    }
}

/// A value type: a number type, `v128`, or a reference type.
fn val_type(r: &mut Reader) -> Result<String, WasmError> {
    let b = *r
        .d
        .get(r.p)
        .ok_or_else(|| WasmError::from("unexpected end of value type"))?;
    match b {
        0x7F => {
            r.p += 1;
            Ok("i32".to_string())
        }
        0x7E => {
            r.p += 1;
            Ok("i64".to_string())
        }
        0x7D => {
            r.p += 1;
            Ok("f32".to_string())
        }
        0x7C => {
            r.p += 1;
            Ok("f64".to_string())
        }
        0x7B => {
            r.p += 1;
            Ok("v128".to_string())
        }
        _ => ref_type(r),
    }
}

/// `limits` — the flags byte carries `has max` (bit 0), `shared` (bit 1) and
/// `64-bit addresses` (bit 2), and the text form spells the last two out.
fn limits(r: &mut Reader) -> Result<String, WasmError> {
    let flags = r.byte()?;
    // The core binary format spells exactly four: i32/i64 addresses × with or
    // without a maximum. `shared` belongs to the threads proposal and has no
    // spelling in the text this front end parses, so it is reported as
    // unrenderable rather than written back as an unshared memory.
    if !matches!(flags, 0x00 | 0x01 | 0x04 | 0x05) {
        return Err(format!("cannot render limits flags 0x{flags:02x}").into());
    }
    let min = r.u64()?;
    let max = if flags & 0x01 != 0 {
        Some(r.u64()?)
    } else {
        None
    };
    let mut s = String::new();
    if flags & 0x04 != 0 {
        s.push_str("i64 ");
    }
    let _ = write!(s, "{min}");
    if let Some(m) = max {
        let _ = write!(s, " {m}");
    }
    Ok(s)
}

fn global_type(r: &mut Reader) -> Result<String, WasmError> {
    let t = val_type(r)?;
    let mutable = match r.byte()? {
        0x00 => false,
        0x01 => true,
        b => return Err(format!("cannot render mutability 0x{b:02x}").into()),
    };
    Ok(if mutable {
        format!("(mut {t})")
    } else {
        t.to_string()
    })
}

fn table_type(r: &mut Reader) -> Result<String, WasmError> {
    let rt = ref_type(r)?;
    let l = limits(r)?;
    Ok(format!("{l} {rt}"))
}

// ── Instructions ──────────────────────────────────────────────────────────────

/// The opcodes with no immediates, by their spec mnemonics.
///
/// Generated from the spec's own instruction index appendix, with the four
/// `min`/`max` names taken from the text format (the appendix spells the
/// operator `fmin`/`fmax`, which is not what a module says).
fn nullary_name(op: u8) -> Option<&'static str> {
    Some(match op {
        0x00 => "unreachable",
        0x01 => "nop",
        0x0F => "return",
        0x1A => "drop",
        0x1B => "select",
        0x45 => "i32.eqz",
        0x46 => "i32.eq",
        0x47 => "i32.ne",
        0x48 => "i32.lt_s",
        0x49 => "i32.lt_u",
        0x4A => "i32.gt_s",
        0x4B => "i32.gt_u",
        0x4C => "i32.le_s",
        0x4D => "i32.le_u",
        0x4E => "i32.ge_s",
        0x4F => "i32.ge_u",
        0x50 => "i64.eqz",
        0x51 => "i64.eq",
        0x52 => "i64.ne",
        0x53 => "i64.lt_s",
        0x54 => "i64.lt_u",
        0x55 => "i64.gt_s",
        0x56 => "i64.gt_u",
        0x57 => "i64.le_s",
        0x58 => "i64.le_u",
        0x59 => "i64.ge_s",
        0x5A => "i64.ge_u",
        0x5B => "f32.eq",
        0x5C => "f32.ne",
        0x5D => "f32.lt",
        0x5E => "f32.gt",
        0x5F => "f32.le",
        0x60 => "f32.ge",
        0x61 => "f64.eq",
        0x62 => "f64.ne",
        0x63 => "f64.lt",
        0x64 => "f64.gt",
        0x65 => "f64.le",
        0x66 => "f64.ge",
        0x67 => "i32.clz",
        0x68 => "i32.ctz",
        0x69 => "i32.popcnt",
        0x6A => "i32.add",
        0x6B => "i32.sub",
        0x6C => "i32.mul",
        0x6D => "i32.div_s",
        0x6E => "i32.div_u",
        0x6F => "i32.rem_s",
        0x70 => "i32.rem_u",
        0x71 => "i32.and",
        0x72 => "i32.or",
        0x73 => "i32.xor",
        0x74 => "i32.shl",
        0x75 => "i32.shr_s",
        0x76 => "i32.shr_u",
        0x77 => "i32.rotl",
        0x78 => "i32.rotr",
        0x79 => "i64.clz",
        0x7A => "i64.ctz",
        0x7B => "i64.popcnt",
        0x7C => "i64.add",
        0x7D => "i64.sub",
        0x7E => "i64.mul",
        0x7F => "i64.div_s",
        0x80 => "i64.div_u",
        0x81 => "i64.rem_s",
        0x82 => "i64.rem_u",
        0x83 => "i64.and",
        0x84 => "i64.or",
        0x85 => "i64.xor",
        0x86 => "i64.shl",
        0x87 => "i64.shr_s",
        0x88 => "i64.shr_u",
        0x89 => "i64.rotl",
        0x8A => "i64.rotr",
        0x8B => "f32.abs",
        0x8C => "f32.neg",
        0x8D => "f32.ceil",
        0x8E => "f32.floor",
        0x8F => "f32.trunc",
        0x90 => "f32.nearest",
        0x91 => "f32.sqrt",
        0x92 => "f32.add",
        0x93 => "f32.sub",
        0x94 => "f32.mul",
        0x95 => "f32.div",
        0x96 => "f32.min",
        0x97 => "f32.max",
        0x98 => "f32.copysign",
        0x99 => "f64.abs",
        0x9A => "f64.neg",
        0x9B => "f64.ceil",
        0x9C => "f64.floor",
        0x9D => "f64.trunc",
        0x9E => "f64.nearest",
        0x9F => "f64.sqrt",
        0xA0 => "f64.add",
        0xA1 => "f64.sub",
        0xA2 => "f64.mul",
        0xA3 => "f64.div",
        0xA4 => "f64.min",
        0xA5 => "f64.max",
        0xA6 => "f64.copysign",
        0xA7 => "i32.wrap_i64",
        0xA8 => "i32.trunc_f32_s",
        0xA9 => "i32.trunc_f32_u",
        0xAA => "i32.trunc_f64_s",
        0xAB => "i32.trunc_f64_u",
        0xAC => "i64.extend_i32_s",
        0xAD => "i64.extend_i32_u",
        0xAE => "i64.trunc_f32_s",
        0xAF => "i64.trunc_f32_u",
        0xB0 => "i64.trunc_f64_s",
        0xB1 => "i64.trunc_f64_u",
        0xB2 => "f32.convert_i32_s",
        0xB3 => "f32.convert_i32_u",
        0xB4 => "f32.convert_i64_s",
        0xB5 => "f32.convert_i64_u",
        0xB6 => "f32.demote_f64",
        0xB7 => "f64.convert_i32_s",
        0xB8 => "f64.convert_i32_u",
        0xB9 => "f64.convert_i64_s",
        0xBA => "f64.convert_i64_u",
        0xBB => "f64.promote_f32",
        0xBC => "i32.reinterpret_f32",
        0xBD => "i64.reinterpret_f64",
        0xBE => "f32.reinterpret_i32",
        0xBF => "f64.reinterpret_i64",
        0xC0 => "i32.extend8_s",
        0xC1 => "i32.extend16_s",
        0xC2 => "i64.extend8_s",
        0xC3 => "i64.extend16_s",
        0xC4 => "i64.extend32_s",
        0xD1 => "ref.is_null",
        _ => return None,
    })
}

/// A memory access opcode's mnemonic and its NATURAL alignment exponent — the
/// text form omits `align=` only when it equals the natural one, and writing a
/// smaller-than-natural alignment back as natural would change the module.
fn memop(op: u8) -> Option<(&'static str, u32)> {
    Some(match op {
        0x28 => ("i32.load", 2),
        0x29 => ("i64.load", 3),
        0x2A => ("f32.load", 2),
        0x2B => ("f64.load", 3),
        0x2C => ("i32.load8_s", 0),
        0x2D => ("i32.load8_u", 0),
        0x2E => ("i32.load16_s", 1),
        0x2F => ("i32.load16_u", 1),
        0x30 => ("i64.load8_s", 0),
        0x31 => ("i64.load8_u", 0),
        0x32 => ("i64.load16_s", 1),
        0x33 => ("i64.load16_u", 1),
        0x34 => ("i64.load32_s", 2),
        0x35 => ("i64.load32_u", 2),
        0x36 => ("i32.store", 2),
        0x37 => ("i64.store", 3),
        0x38 => ("f32.store", 2),
        0x39 => ("f64.store", 3),
        0x3A => ("i32.store8", 0),
        0x3B => ("i32.store16", 1),
        0x3C => ("i64.store8", 0),
        0x3D => ("i64.store16", 1),
        0x3E => ("i64.store32", 2),
        _ => return None,
    })
}

/// The `0xFC` prefix space: saturating truncation, then the bulk operations.
fn fc_name(sub: u32) -> Option<(&'static str, FcImm)> {
    use FcImm::*;
    Some(match sub {
        0 => ("i32.trunc_sat_f32_s", Nullary),
        1 => ("i32.trunc_sat_f32_u", Nullary),
        2 => ("i32.trunc_sat_f64_s", Nullary),
        3 => ("i32.trunc_sat_f64_u", Nullary),
        4 => ("i64.trunc_sat_f32_s", Nullary),
        5 => ("i64.trunc_sat_f32_u", Nullary),
        6 => ("i64.trunc_sat_f64_s", Nullary),
        7 => ("i64.trunc_sat_f64_u", Nullary),
        8 => ("memory.init", SegThenSpace),
        9 => ("data.drop", Segment),
        10 => ("memory.copy", TwoSpaces),
        11 => ("memory.fill", OneSpace),
        12 => ("table.init", SegThenSpace),
        13 => ("elem.drop", Segment),
        14 => ("table.copy", TwoSpaces),
        15 => ("table.grow", OneSpace),
        16 => ("table.size", OneSpace),
        17 => ("table.fill", OneSpace),
        _ => return None,
    })
}

/// The immediate shape of a `0xFC` instruction.
///
/// ⛔ THE TWO ORDERS DISAGREE. `memory.init` and `table.init` encode the
/// SEGMENT index first and the memory/table index second, and the text form
/// writes them the other way round (`table.init <table> <elem>`). Reading them
/// out in encoding order and printing them in that order swaps the two.
enum FcImm {
    /// No immediates.
    Nullary,
    /// A data/elem segment index, always written.
    Segment,
    /// One memory/table index; omitted from the text when it is 0.
    OneSpace,
    /// Two memory/table indices; omitted from the text when both are 0.
    TwoSpaces,
    /// A segment index then a memory/table index — written in the opposite
    /// order, and the space index omitted when it is 0.
    SegThenSpace,
}

/// An index-space immediate, written only when it is not the default. Index 0
/// is what a module without the multi-memory / reference-types spellings means,
/// and the text form for it is the bare mnemonic.
fn idx_suffix(idx: u32) -> String {
    if idx == 0 {
        String::new()
    } else {
        format!(" {idx}")
    }
}

/// A block type: nothing, a single result, or a type index.
fn block_type(r: &mut Reader) -> Result<String, WasmError> {
    let b = *r
        .d
        .get(r.p)
        .ok_or_else(|| WasmError::from("unexpected end of block type"))?;
    if b == 0x40 {
        r.p += 1;
        return Ok(String::new());
    }
    // A single value type is a one-result block; anything else is a signed
    // 33-bit type index. The two are told apart by the byte, not by trying.
    if matches!(b, 0x7B..=0x7F) || b == 0x63 || b == 0x64 || abs_heap_type(b).is_some() {
        return Ok(format!(" (result {})", val_type(r)?));
    }
    let idx = r.i64()?;
    if idx < 0 {
        return Err(format!("cannot render block type 0x{b:02x}").into());
    }
    Ok(format!(" (type {idx})"))
}

/// Render one instruction. Returns `false` at the `end` that closes the body.
fn instr(r: &mut Reader, out: &mut String, depth: &mut usize) -> Result<bool, WasmError> {
    let op = r.byte()?;
    let pad = "  ".repeat(*depth + 2);
    if let Some(name) = nullary_name(op) {
        let _ = writeln!(out, "{pad}{name}");
        return Ok(true);
    }
    if let Some((name, natural)) = memop(op) {
        let align_flags = r.u32()?;
        // Bit 6 of the align field announces an explicit memory index, which
        // sits BETWEEN the align and the offset.
        let has_memidx = align_flags & 0x40 != 0;
        let align_exp = align_flags & 0x3F;
        let memidx = if has_memidx { r.u32()? } else { 0 };
        let offset = r.u64()?;
        let _ = write!(out, "{pad}{name}");
        if memidx != 0 {
            let _ = write!(out, " {memidx}");
        }
        if offset != 0 {
            let _ = write!(out, " offset={offset}");
        }
        if align_exp != natural {
            let _ = write!(out, " align={}", 1u64 << align_exp);
        }
        out.push('\n');
        return Ok(true);
    }
    match op {
        0x02 | 0x03 | 0x04 => {
            let name = match op {
                0x02 => "block",
                0x03 => "loop",
                _ => "if",
            };
            let bt = block_type(r)?;
            let _ = writeln!(out, "{pad}{name}{bt}");
            *depth += 1;
        }
        0x05 => {
            let outer = "  ".repeat(depth.saturating_sub(1) + 2);
            let _ = writeln!(out, "{outer}else");
        }
        0x0B => {
            if *depth == 0 {
                return Ok(false);
            }
            *depth -= 1;
            let outer = "  ".repeat(*depth + 2);
            let _ = writeln!(out, "{outer}end");
        }
        0x0C | 0x0D => {
            let l = r.u32()?;
            let name = if op == 0x0C { "br" } else { "br_if" };
            let _ = writeln!(out, "{pad}{name} {l}");
        }
        0x0E => {
            let n = r.u32()?;
            let mut labels = Vec::with_capacity(n as usize + 1);
            for _ in 0..n {
                labels.push(r.u32()?.to_string());
            }
            labels.push(r.u32()?.to_string());
            let _ = writeln!(out, "{pad}br_table {}", labels.join(" "));
        }
        0x10 => {
            let x = r.u32()?;
            let _ = writeln!(out, "{pad}call {x}");
        }
        0x11 => {
            let ty = r.u32()?;
            let tbl = r.u32()?;
            let _ = writeln!(out, "{pad}call_indirect{} (type {ty})", idx_suffix(tbl));
        }
        0x1C => {
            let n = r.u32()?;
            let mut ts = Vec::new();
            for _ in 0..n {
                ts.push(val_type(r)?);
            }
            let _ = writeln!(out, "{pad}select (result {})", ts.join(" "));
        }
        0x20..=0x22 => {
            let x = r.u32()?;
            let name = match op {
                0x20 => "local.get",
                0x21 => "local.set",
                _ => "local.tee",
            };
            let _ = writeln!(out, "{pad}{name} {x}");
        }
        0x23 | 0x24 => {
            let x = r.u32()?;
            let name = if op == 0x23 { "global.get" } else { "global.set" };
            let _ = writeln!(out, "{pad}{name} {x}");
        }
        0x25 | 0x26 => {
            let x = r.u32()?;
            let name = if op == 0x25 { "table.get" } else { "table.set" };
            let _ = writeln!(out, "{pad}{name} {x}");
        }
        0x3F | 0x40 => {
            let m = r.u32()?;
            let name = if op == 0x3F { "memory.size" } else { "memory.grow" };
            let _ = writeln!(out, "{pad}{name}{}", idx_suffix(m));
        }
        0x41 => {
            let v = r.i32()?;
            let _ = writeln!(out, "{pad}i32.const {v}");
        }
        0x42 => {
            let v = r.i64()?;
            let _ = writeln!(out, "{pad}i64.const {v}");
        }
        0x43 => {
            let v = r.f32_bits()?;
            let _ = writeln!(out, "{pad}f32.const {}", f32_literal(v));
        }
        0x44 => {
            let v = r.f64_bits()?;
            let _ = writeln!(out, "{pad}f64.const {}", f64_literal(v));
        }
        0xD0 => {
            let ht = heap_type(r)?;
            let _ = writeln!(out, "{pad}ref.null {ht}");
        }
        0xD2 => {
            let x = r.u32()?;
            let _ = writeln!(out, "{pad}ref.func {x}");
        }
        0xFC => {
            let sub = r.u32()?;
            let (name, imm) = fc_name(sub)
                .ok_or_else(|| WasmError::from(format!("cannot render instruction 0xfc {sub}")))?;
            match imm {
                FcImm::Nullary => {
                    let _ = writeln!(out, "{pad}{name}");
                }
                FcImm::Segment => {
                    let x = r.u32()?;
                    let _ = writeln!(out, "{pad}{name} {x}");
                }
                FcImm::OneSpace => {
                    let x = r.u32()?;
                    let _ = writeln!(out, "{pad}{name}{}", idx_suffix(x));
                }
                FcImm::TwoSpaces => {
                    let a = r.u32()?;
                    let b = r.u32()?;
                    if a == 0 && b == 0 {
                        let _ = writeln!(out, "{pad}{name}");
                    } else {
                        let _ = writeln!(out, "{pad}{name} {a} {b}");
                    }
                }
                FcImm::SegThenSpace => {
                    let seg = r.u32()?;
                    let space = r.u32()?;
                    let _ = writeln!(out, "{pad}{name}{} {seg}", idx_suffix(space));
                }
            }
        }
        0xFD => {
            let sub = r.u32()?;
            if sub != 12 {
                return Err(format!("cannot render instruction 0xfd {sub}").into());
            }
            // ⛔ ALWAYS `i8x16`. Byte lanes are the vector's bits verbatim; any
            // float shape would put the constant through a float parser and
            // the fixture that pins single-rounding would answer differently.
            let b = r.bytes(16)?;
            let lanes: Vec<String> = b.iter().map(|x| x.to_string()).collect();
            let _ = writeln!(out, "{pad}v128.const i8x16 {}", lanes.join(" "));
        }
        _ => return Err(format!("cannot render instruction 0x{op:02x}").into()),
    }
    Ok(true)
}

/// A constant expression, rendered folded so it can sit where the text form
/// expects one `(…)` — a segment offset, a global initialiser, an elem item.
fn const_expr(r: &mut Reader) -> Result<String, WasmError> {
    let mut parts: Vec<String> = Vec::new();
    loop {
        let mut line = String::new();
        let mut depth = 0usize;
        if !instr(r, &mut line, &mut depth)? {
            break;
        }
        parts.push(line.trim().to_string());
    }
    if parts.len() == 1 {
        Ok(format!("({})", parts[0]))
    } else {
        // More than one instruction: the text form nests them as a sequence
        // inside an explicit `(offset …)`/`(item …)` wrapper, which every
        // caller here already supplies.
        Ok(parts.join(" "))
    }
}

// ── Sections ──────────────────────────────────────────────────────────────────

/// Render a WASM module's bytes as WAT text.
pub fn wat_from_binary(data: &[u8]) -> Result<String, WasmError> {
    if data.len() < 8 || &data[0..4] != b"\0asm" {
        return Err("magic header not detected".into());
    }
    if &data[4..8] != &[0x01, 0x00, 0x00, 0x00] {
        return Err("unknown binary version".into());
    }

    // Fields are buffered by KIND and emitted in the text format's own order,
    // not the binary's. The two disagree where it matters: a `(data …)` or
    // `(elem …)` segment naming a function, and an `(export … (func N))`, are
    // written by the binary before the code section that defines the function.
    let mut s_types = String::new();
    let mut s_imports = String::new();
    let mut s_tables = String::new();
    let mut s_memories = String::new();
    let mut s_globals = String::new();
    let mut s_funcs = String::new();
    let mut s_exports = String::new();
    let mut s_start = String::new();
    let mut s_elems = String::new();
    let mut s_datas = String::new();
    // The func section names each defined function's type; the code section
    // supplies its body. They are separate sections and have to be joined.
    let mut func_types: Vec<u32> = Vec::new();
    let mut code_bodies: Vec<Vec<u8>> = Vec::new();
    let mut type_sigs: Vec<(Vec<String>, Vec<String>)> = Vec::new();

    let mut p = 8usize;
    while p < data.len() {
        let mut hdr = Reader::new(&data[p..]);
        let id = hdr.byte()?;
        let size = hdr.u32()? as usize;
        let body_start = p + hdr.p;
        let body_end = body_start
            .checked_add(size)
            .filter(|e| *e <= data.len())
            .ok_or_else(|| WasmError::from("unexpected end of section or function"))?;
        let body = &data[body_start..body_end];
        p = body_end;
        let mut r = Reader::new(body);
        match id {
            // A custom section carries no module structure.
            0 => {}
            1 => {
                let n = r.u32()?;
                for _ in 0..n {
                    if r.byte()? != 0x60 {
                        return Err("cannot render a non-function type".into());
                    }
                    let np = r.u32()?;
                    let mut params = Vec::new();
                    for _ in 0..np {
                        params.push(val_type(&mut r)?);
                    }
                    let nr = r.u32()?;
                    let mut results = Vec::new();
                    for _ in 0..nr {
                        results.push(val_type(&mut r)?);
                    }
                    let _ = writeln!(s_types, "  (type (func{}))", sig_text(&params, &results));
                    type_sigs.push((params, results));
                }
            }
            2 => {
                let n = r.u32()?;
                for _ in 0..n {
                    let module = r.name()?;
                    let field = r.name()?;
                    let kind = r.byte()?;
                    let desc = match kind {
                        0x00 => format!("(func (type {}))", r.u32()?),
                        0x01 => format!("(table {})", table_type(&mut r)?),
                        0x02 => format!("(memory {})", limits(&mut r)?),
                        0x03 => format!("(global {})", global_type(&mut r)?),
                        0x04 => {
                            let attr = r.byte()?;
                            if attr != 0x00 {
                                return Err("cannot render tag attribute".into());
                            }
                            format!("(tag (type {}))", r.u32()?)
                        }
                        _ => return Err(format!("cannot render import kind {kind}").into()),
                    };
                    let _ = writeln!(
                        s_imports,
                        "  (import {} {} {desc})",
                        wat_string(&module),
                        wat_string(&field)
                    );
                }
            }
            3 => {
                let n = r.u32()?;
                for _ in 0..n {
                    func_types.push(r.u32()?);
                }
            }
            4 => {
                let n = r.u32()?;
                for _ in 0..n {
                    // A table may carry an INITIALISER, and the encoding for
                    // that puts a `0x40 0x00` prefix ahead of the table type
                    // (`Btable`). Without it the prefix was read as the table's
                    // reference type and rejected.
                    let has_init = r.d.get(r.p) == Some(&0x40) && r.d.get(r.p + 1) == Some(&0x00);
                    if has_init {
                        r.p += 2;
                        let tt = table_type(&mut r)?;
                        let init = const_expr(&mut r)?;
                        let _ = writeln!(s_tables, "  (table {tt} {init})");
                    } else {
                        let _ = writeln!(s_tables, "  (table {})", table_type(&mut r)?);
                    }
                }
            }
            5 => {
                let n = r.u32()?;
                for _ in 0..n {
                    let _ = writeln!(s_memories, "  (memory {})", limits(&mut r)?);
                }
            }
            6 => {
                let n = r.u32()?;
                for _ in 0..n {
                    let ty = global_type(&mut r)?;
                    let init = const_expr(&mut r)?;
                    let _ = writeln!(s_globals, "  (global {ty} {init})");
                }
            }
            7 => {
                let n = r.u32()?;
                for _ in 0..n {
                    let name = r.name()?;
                    let kind = r.byte()?;
                    let idx = r.u32()?;
                    let what = match kind {
                        0x00 => "func",
                        0x01 => "table",
                        0x02 => "memory",
                        0x03 => "global",
                        0x04 => "tag",
                        _ => return Err(format!("cannot render export kind {kind}").into()),
                    };
                    let _ = writeln!(
                        s_exports,
                        "  (export {} ({what} {idx}))",
                        wat_string(&name)
                    );
                }
            }
            8 => {
                let _ = writeln!(s_start, "  (start {})", r.u32()?);
            }
            9 => render_elem_section(&mut r, &mut s_elems)?,
            10 => {
                let n = r.u32()?;
                for _ in 0..n {
                    let size = r.u32()? as usize;
                    code_bodies.push(r.bytes(size)?.to_vec());
                }
            }
            11 => render_data_section(&mut r, &mut s_datas)?,
            // The data count section is implied by the segments themselves.
            12 => {}
            13 => {
                let n = r.u32()?;
                for _ in 0..n {
                    let attr = r.byte()?;
                    if attr != 0x00 {
                        return Err("cannot render tag attribute".into());
                    }
                    let _ = writeln!(s_imports, "  (tag (type {}))", r.u32()?);
                }
            }
            _ => return Err(format!("cannot render section {id}").into()),
        }
    }

    if func_types.len() != code_bodies.len() {
        return Err("function and code section have inconsistent lengths".into());
    }
    for (ty, body) in func_types.iter().zip(code_bodies.iter()) {
        let sig = type_sigs
            .get(*ty as usize)
            .ok_or_else(|| WasmError::from("unknown type"))?;
        let _ = write!(s_funcs, "  (func (type {ty}){}", sig_text(&sig.0, &sig.1));
        render_func_body(body, &mut s_funcs)?;
        s_funcs.push_str("  )\n");
    }

    let mut out = String::from("(module\n");
    for part in [
        &s_types,
        &s_imports,
        &s_tables,
        &s_memories,
        &s_globals,
        &s_funcs,
        &s_exports,
        &s_start,
        &s_elems,
        &s_datas,
    ] {
        out.push_str(part);
    }
    out.push_str(")\n");
    Ok(out)
}

fn sig_text(params: &[String], results: &[String]) -> String {
    let mut s = String::new();
    for p in params {
        let _ = write!(s, " (param {p})");
    }
    for r in results {
        let _ = write!(s, " (result {r})");
    }
    s
}

fn render_func_body(body: &[u8], out: &mut String) -> Result<(), WasmError> {
    let mut r = Reader::new(body);
    let groups = r.u32()?;
    let mut locals: Vec<String> = Vec::new();
    for _ in 0..groups {
        let count = r.u32()?;
        let t = val_type(&mut r)?;
        for _ in 0..count {
            locals.push(t.clone());
        }
    }
    if !locals.is_empty() {
        let _ = write!(out, " (local {})", locals.join(" "));
    }
    out.push('\n');
    let mut depth = 0usize;
    while !r.done() {
        if !instr(&mut r, out, &mut depth)? {
            break;
        }
    }
    Ok(())
}

/// The eight element-segment encodings of the binary format. The low two bits
/// select the mode (active / passive / active-with-table / declarative); bit 2
/// selects whether the items are function indices or expressions.
fn render_elem_section(r: &mut Reader, out: &mut String) -> Result<(), WasmError> {
    let n = r.u32()?;
    for _ in 0..n {
        let mode = r.u32()?;
        let mut s = String::from("  (elem");
        // Placement.
        match mode {
            0 | 4 => {
                let off = const_expr(r)?;
                let _ = write!(s, " (offset {off})");
            }
            2 | 6 => {
                let table = r.u32()?;
                let off = const_expr(r)?;
                let _ = write!(s, " (table {table}) (offset {off})");
            }
            1 | 5 => {}
            3 | 7 => s.push_str(" declare"),
            _ => return Err(format!("cannot render element segment mode {mode}").into()),
        }
        // Element type, then the items.
        if mode & 0x04 == 0 {
            // Function indices, tagged by an elemkind byte on every mode but 0.
            if mode != 0 {
                let kind = r.byte()?;
                if kind != 0x00 {
                    return Err(format!("cannot render element kind {kind}").into());
                }
            }
            s.push_str(" func");
            let count = r.u32()?;
            for _ in 0..count {
                let _ = write!(s, " {}", r.u32()?);
            }
        } else {
            // Expressions, tagged by a reftype on every mode but 4.
            let rt = if mode == 4 {
                "funcref".to_string()
            } else {
                ref_type(r)?
            };
            let _ = write!(s, " {rt}");
            let count = r.u32()?;
            for _ in 0..count {
                let _ = write!(s, " (item {})", const_expr(r)?);
            }
        }
        s.push(')');
        let _ = writeln!(out, "{s}");
    }
    Ok(())
}

/// The three data-segment encodings: active on memory 0, passive, and active
/// on an explicit memory.
fn render_data_section(r: &mut Reader, out: &mut String) -> Result<(), WasmError> {
    let n = r.u32()?;
    for _ in 0..n {
        let mode = r.u32()?;
        let mut s = String::from("  (data");
        match mode {
            0 => {
                let off = const_expr(r)?;
                let _ = write!(s, " (offset {off})");
            }
            1 => {}
            2 => {
                let mem = r.u32()?;
                let off = const_expr(r)?;
                let _ = write!(s, " (memory {mem}) (offset {off})");
            }
            _ => return Err(format!("cannot render data segment mode {mode}").into()),
        }
        let len = r.u32()? as usize;
        let bytes = r.bytes(len)?;
        let _ = write!(s, " {}", wat_string(bytes));
        s.push(')');
        let _ = writeln!(out, "{s}");
    }
    Ok(())
}
