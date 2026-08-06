//! JSPI custom-section emission — verifies that chunks flagged
//! `is_async = true` show up in the `vybe.jspi` custom section of the
//! emitted `.wasm`, and chunks without the flag don't.

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op};
use vybe_platform_wasm::write_wasm;

/// Minimal LEB128-decoding walk of the resulting wasm. Finds the
/// `vybe.jspi` custom section, returns the decoded
/// (promising_indices, suspending_indices) pair. Returns `None` if no
/// such section exists.
fn extract_jspi_section(wasm: &[u8]) -> Option<(Vec<u32>, Vec<u32>)> {
    if wasm.len() < 8 || &wasm[..4] != b"\0asm" {
        return None;
    }
    let mut i = 8;
    while i < wasm.len() {
        let sid = wasm[i];
        i += 1;
        let (sz, read) = read_leb(wasm, i);
        i += read;
        let end = i + sz;
        if sid == 0 {
            // Custom section — name length + name + payload
            let (nl, r) = read_leb(wasm, i);
            let mut p = i + r;
            let name = std::str::from_utf8(&wasm[p..p + nl]).unwrap_or("");
            p += nl;
            if name == "vybe.jspi" {
                let (pc, r) = read_leb(wasm, p);
                p += r;
                let mut promising = Vec::with_capacity(pc);
                for _ in 0..pc {
                    let (idx, r) = read_leb(wasm, p);
                    p += r;
                    promising.push(idx as u32);
                }
                let (sc, r) = read_leb(wasm, p);
                p += r;
                let mut suspending = Vec::with_capacity(sc);
                for _ in 0..sc {
                    let (idx, r) = read_leb(wasm, p);
                    p += r;
                    suspending.push(idx as u32);
                }
                return Some((promising, suspending));
            }
        }
        i = end;
    }
    None
}

fn read_leb(buf: &[u8], mut i: usize) -> (usize, usize) {
    let mut v = 0usize;
    let mut shift = 0;
    let start = i;
    loop {
        let b = buf[i];
        i += 1;
        v |= ((b & 0x7f) as usize) << shift;
        if b & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    (v, i - start)
}

#[test]
fn jspi_section_absent_when_no_async_chunks() {
    let mut script = Chunk::new("<script>");
    script.emit_i32_const(42, 0);
    script.emit_op(Op::RETURN, 0);
    let wasm = write_wasm(&[script]);
    assert!(
        extract_jspi_section(&wasm).is_none(),
        "module with no async chunks must not emit vybe.jspi section",
    );
}

#[test]
fn jspi_section_lists_async_chunk() {
    // Script chunk (not async) + one async user function. The async
    // function's wasm func_idx = import_count + rt_imports_len + 1
    // (script is chunk 0). For a module with zero user-declared
    // imports, the only prefix that raises the chunk-level index is
    // `rt_imports` (js-primitive-builtins globals).
    let mut script = Chunk::new("<script>");
    script.emit_op(Op::RETURN, 0);
    let mut async_fn = Chunk::new("doSomethingAsync");
    async_fn.arity = 0;
    async_fn.is_async = true;
    async_fn.emit_i32_const(7, 0);
    async_fn.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&[script, async_fn]);
    let (promising, suspending) = extract_jspi_section(&wasm)
        .expect("vybe.jspi section must be present when a chunk is async");
    assert_eq!(
        promising.len(),
        1,
        "exactly one promising export expected, got {:?}",
        promising
    );
    assert!(
        suspending.is_empty(),
        "no suspending imports expected in this first cut"
    );
}

#[test]
fn jspi_section_lists_multiple_async_chunks() {
    let mut script = Chunk::new("<script>");
    script.emit_op(Op::RETURN, 0);
    let mut a = Chunk::new("a");
    a.is_async = true;
    a.emit_op(Op::RETURN, 0);
    let mut b = Chunk::new("b");
    b.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("c");
    c.is_async = true;
    c.emit_op(Op::RETURN, 0);
    let wasm = write_wasm(&[script, a, b, c]);
    let (promising, _) = extract_jspi_section(&wasm).expect("vybe.jspi section must be present");
    assert_eq!(
        promising.len(),
        2,
        "exactly two async chunks expected, got {:?}",
        promising
    );
    // Indices must be strictly increasing (chunk order preserved).
    assert!(
        promising[0] < promising[1],
        "promising indices must come out in chunk order"
    );
}
