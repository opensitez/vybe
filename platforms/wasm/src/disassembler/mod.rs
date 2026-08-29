//! WAT (WebAssembly Text) emission — the human-readable, sexpr-based
//! form of a WASM module. Useful for inspecting what our compiler
//! actually produces before engines swallow the binary.
//!
//! This is a **direct renderer from `Chunk`s**, not a disassembler on
//! the binary. Some low-level details the binary emitter adds
//! (per-import type indices, exact boxing calls, custom sections)
//! are deliberately abstracted into readable names — the goal is to
//! understand shape and control flow, not to re-produce byte-exact
//! output. For byte-exact inspection run `wasm-tools print` on the
//! binary from `write_wasm`.
//!
//! Output sketch:
//! ```wat
//! (module
//!   (import "wasm:js-number" "fromI32" (func $fromI32 ...))
//!   ...
//!   (func $chunk_0_main (param $p0 externref) (result externref)
//!     local.get $p0
//!     ...
//!     return
//!   )
//! )
//! ```

use std::fmt::Write;
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::{Op, OperandFormat, read_leb_u32, read_leb_u64};

/// Render a collection of chunks as a single WAT module. The output
/// is human-readable; indentation reflects nesting depth of WASM
/// structured control flow (`block` / `loop` / `if` / `try_table`).
pub fn write_wat(chunks: &[Chunk]) -> String {
    let mut out = String::new();
    out.push_str("(module\n");

    // ── Imports ────────────────────────────────────────────────────
    // Each chunk's host imports go here; we dedupe across chunks.
    let mut seen: Vec<(&str, &str)> = Vec::new();
    for chunk in chunks {
        for imp in &chunk.imports {
            let key = (imp.module.as_str(), imp.name.as_str());
            if !seen.iter().any(|k| k == &key) {
                seen.push(key);
                let _ = writeln!(
                    out,
                    "  (import \"{}\" \"{}\" (func ${}__{}))",
                    imp.module,
                    imp.name,
                    sanitize_ident(&imp.module),
                    sanitize_ident(&imp.name),
                );
            }
        }
    }

    // ── Functions ──────────────────────────────────────────────────
    for (ci, chunk) in chunks.iter().enumerate() {
        let id = format!("$c{}_{}", ci, sanitize_ident(&chunk.name));
        let _ = write!(out, "\n  (func {id}");
        for p in 0..chunk.arity {
            let _ = write!(out, " (param $p{p} externref)");
        }
        let result_arity = chunk.result_arity.max(1);
        let _ = write!(out, " (result");
        for _ in 0..result_arity {
            out.push_str(" externref");
        }
        out.push(')');
        if chunk.is_async {
            out.push_str(" (; async ;)");
        }
        out.push('\n');

        let extras = (chunk.local_count as u32).saturating_sub(chunk.arity as u32);
        if extras > 0 {
            let _ = write!(out, "    (local");
            for _ in 0..extras {
                out.push_str(" externref");
            }
            out.push_str(")\n");
        }

        // Body
        render_body(&mut out, chunk, 4);
        out.push_str("  )\n");
    }

    out.push_str(")\n");
    out
}

/// Render a single chunk as WAT — convenience for disassembly-style
/// diagnostics on one function at a time.
pub fn write_wat_chunk(chunk: &Chunk) -> String {
    let mut out = String::new();
    let _ = write!(out, "(func ${}", sanitize_ident(&chunk.name));
    for p in 0..chunk.arity {
        let _ = write!(out, " (param $p{p} externref)");
    }
    let result_arity = chunk.result_arity.max(1);
    out.push_str(" (result");
    for _ in 0..result_arity {
        out.push_str(" externref");
    }
    out.push_str(")\n");
    render_body(&mut out, chunk, 2);
    out.push_str(")\n");
    out
}

fn render_body(out: &mut String, chunk: &Chunk, base_indent: usize) {
    let mut ip = 0;
    let mut depth = base_indent;
    while ip < chunk.code.len() {
        if ip + 3 >= chunk.code.len() {
            break;
        }
        let group = ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16;
        let sub = ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16;
        let Some(op) = Op::decode(group, sub) else {
            let _ = writeln!(
                out,
                "{:indent$};; unknown 0x{:04X} 0x{:04X}",
                "",
                group,
                sub,
                indent = depth,
            );
            ip += 4;
            continue;
        };
        // Closing-bracket ops de-indent first.
        if op == Op::END {
            depth = depth.saturating_sub(2);
        }
        let _ = write!(out, "{:indent$}", "", indent = depth);
        render_instruction(out, chunk, op, ip);
        out.push('\n');
        // Opening ops indent for their body.
        if op == Op::BLOCK || op == Op::LOOP || op == Op::TRY_TABLE {
            depth += 2;
        }
        ip += crate::writer::code::opcode_size(op, &chunk.code, ip);
    }
}

fn render_instruction(out: &mut String, chunk: &Chunk, op: Op, ip: usize) {
    let name = op.wasm_name();
    out.push_str(name);
    match op.operand_format() {
        OperandFormat::None => {}
        OperandFormat::U8 => {
            let v = chunk.code[ip + 4];
            let _ = write!(out, " {v}");
        }
        OperandFormat::U8_U8 => {
            let a = chunk.code[ip + 4];
            let b = chunk.code[ip + 5];
            let _ = write!(out, " {a} {b}");
        }
        // u16 tableidx + u16 tag-name index + u8 argc.
        OperandFormat::U16_U16_U8 => {
            let a = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
            let b = ((chunk.code[ip + 6] as u16) << 8) | chunk.code[ip + 7] as u16;
            let c = chunk.code[ip + 8];
            let _ = write!(out, " {a} {b} {c}");
        }
        OperandFormat::U8_U8_U8 => {
            let a = chunk.code[ip + 4];
            let b = chunk.code[ip + 5];
            let c = chunk.code[ip + 6];
            let _ = write!(out, " {a} {b} {c}");
        }
        OperandFormat::U16 => {
            let v = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
            let _ = write!(out, " {v}");
        }
        OperandFormat::I16 => {
            let raw = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
            let v = raw as i16;
            let _ = write!(out, " {v}");
        }
        OperandFormat::U16_U8 => {
            let hi = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
            let lo = chunk.code[ip + 6];
            let _ = write!(out, " {hi} {lo}");
        }
        OperandFormat::U16_U16 => {
            let a = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
            let b = ((chunk.code[ip + 6] as u16) << 8) | chunk.code[ip + 7] as u16;
            let _ = write!(out, " {a} {b}");
        }
        OperandFormat::U16_I16 => {
            let a = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
            let raw = ((chunk.code[ip + 6] as u16) << 8) | chunk.code[ip + 7] as u16;
            let _ = write!(out, " {a} {}", raw as i16);
        }
        OperandFormat::U32Leb => {
            let mut operand_ip = ip + 4;
            let value = read_leb_u32(&chunk.code, &mut operand_ip);
            let _ = write!(out, " {value}");
        }
        OperandFormat::U32Leb_U32Leb => {
            let mut operand_ip = ip + 4;
            let a = read_leb_u32(&chunk.code, &mut operand_ip);
            let b = read_leb_u32(&chunk.code, &mut operand_ip);
            let _ = write!(out, " {a} {b}");
        }
        OperandFormat::MemArg => {
            let mut operand_ip = ip + 4;
            let align = read_leb_u32(&chunk.code, &mut operand_ip);
            let offset = read_leb_u32(&chunk.code, &mut operand_ip);
            let _ = write!(out, " align={align} offset={offset}");
            if align & 0x40 != 0 {
                let memidx = read_leb_u32(&chunk.code, &mut operand_ip);
                let _ = write!(out, " memory={memidx}");
            }
        }
        OperandFormat::MemArg64 => {
            let mut operand_ip = ip + 4;
            let align = read_leb_u32(&chunk.code, &mut operand_ip);
            let offset = read_leb_u64(&chunk.code, &mut operand_ip);
            let _ = write!(out, " align={align} offset={offset}");
            if align & 0x40 != 0 {
                let memidx = read_leb_u32(&chunk.code, &mut operand_ip);
                let _ = write!(out, " memory={memidx}");
            }
        }
        OperandFormat::SimdMemArg => {
            // Optional marker-tagged memarg — absent contributes zero bytes.
            let mut operand_ip = ip + 4;
            let align = read_leb_u32(&chunk.code, &mut operand_ip);
            if align & 0x80 != 0 {
                let offset = read_leb_u64(&chunk.code, &mut operand_ip);
                let _ = write!(out, " align={} offset={offset}", align & !0x1C0);
                if align & 0x40 != 0 {
                    let memidx = read_leb_u32(&chunk.code, &mut operand_ip);
                    let _ = write!(out, " memory={memidx}");
                }
            }
        }
        OperandFormat::MemLane => {
            // Optional marker-tagged memarg, then the mandatory lane byte.
            let mut operand_ip = ip + 4;
            let align = read_leb_u32(&chunk.code, &mut operand_ip);
            if align & 0x80 != 0 {
                let offset = read_leb_u64(&chunk.code, &mut operand_ip);
                let _ = write!(out, " align={} offset={offset}", align & !0x1C0);
                if align & 0x40 != 0 {
                    let memidx = read_leb_u32(&chunk.code, &mut operand_ip);
                    let _ = write!(out, " memory={memidx}");
                }
            } else {
                operand_ip = ip + 4;
            }
            let lane = chunk.code.get(operand_ip).copied().unwrap_or(0);
            let _ = write!(out, " {lane}");
        }
        OperandFormat::BrTable => {
            let mut operand_ip = ip + 4;
            let count = read_leb_u32(&chunk.code, &mut operand_ip);
            for _ in 0..count {
                let label = read_leb_u32(&chunk.code, &mut operand_ip);
                let _ = write!(out, " {label}");
            }
            let default = read_leb_u32(&chunk.code, &mut operand_ip);
            let _ = write!(out, " {default}");
        }
        OperandFormat::V128Const | OperandFormat::Shuffle => {
            let _ = write!(out, " <{} bytes>", 16);
        }
        OperandFormat::Closure | OperandFormat::TryTable => {
            let _ = write!(out, " <variable-length>");
        }
        OperandFormat::SlI32 => {
            let mut pos = ip + 4;
            let val = vybe_runtime::opcode::read_leb_i32(&chunk.code, &mut pos);
            let _ = write!(out, " {val}");
        }
        OperandFormat::SlI64 => {
            let mut pos = ip + 4;
            let val = vybe_runtime::opcode::read_leb_i64(&chunk.code, &mut pos);
            let _ = write!(out, " {val}");
        }
        OperandFormat::RawF32 => {
            let start = ip + 4;
            let bytes: [u8; 4] = chunk.code[start..start + 4].try_into().unwrap_or([0; 4]);
            let _ = write!(out, " {}", f32::from_le_bytes(bytes));
        }
        OperandFormat::RawF64 => {
            let start = ip + 4;
            let bytes: [u8; 8] = chunk.code[start..start + 8].try_into().unwrap_or([0; 8]);
            let _ = write!(out, " {}", f64::from_le_bytes(bytes));
        }
    }
}

fn sanitize_ident(s: &str) -> String {
    s.chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect()
}
