use vybe_runtime::{Chunk, Op, VM};

// Compile the owned emitter directly so this test also works with shared debug
// dependencies while another worker is rebuilding the libc crate.
#[path = "../../../platforms/libc/src/emitter/heap_memory.rs"]
mod heap;

fn emit(chunks: &mut [Chunk], realloc: bool, baseline: bool) {
    let start = chunks[0].code.len();
    let first = chunks[0].local_count;
    if realloc { heap::emit_realloc(chunks, 0, 0); }
    else { heap::emit_alloc(chunks, 0, 0); }
    if !baseline { return; }
    // Reconstruct exactly the former set/get transport and externref branch
    // contracts, leaving the allocator primitive's pre-existing tee untouched.
    let slots = if realloc {
        vec![first + 2, first + 3, first + 4, first + 5]
    } else { vec![first, first + 1] };
    let c = &mut chunks[0];
    let end = c.code.len();
    let mut positions = Vec::new();
    let mut ip = start;
    while ip < end {
        let op = Op::decode(c.read_u16(ip), c.read_u16(ip + 2)).unwrap();
        if op == Op::LOCAL_TEE && slots.contains(&c.read_u16(ip + 4)) {
            positions.push((ip, c.read_u16(ip + 4)));
        }
        ip += 4 + op.operand_format().size_in(&c.code, ip + 4);
    }
    c.block_i32_results.retain(|at| *at < start || *at >= end);
    for (at, slot) in positions.into_iter().rev() {
        let mut replacement = Chunk::new("<baseline-transport>");
        replacement.emit_op_u16(Op::LOCAL_SET, slot, 0);
        replacement.emit_op_u16(Op::LOCAL_GET, slot, 0);
        c.code.splice(at..at + 6, replacement.code);
    }
    c.lines.resize(c.code.len(), 0);
}

fn runtime() -> VM {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

fn setup() -> Vec<Chunk> {
    let mut c = Chunk::new("<script>");
    c.memory_min_pages.push(1);
    c.emit_i32_const(1, 0);
    c.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_i32_const(1024, 0);
    vybe_compiler::primitives::globals::emit_write(&mut c, "__c_heap_ptr", 0);
    vec![c]
}

fn finish(chunks: &mut Vec<Chunk>) -> Vec<u8> {
    chunks[0].emit_op(Op::RETURN, 0);
    vybe_compiler::primitives::bundle::finalize_with_runtime_helpers(chunks);
    vybe_compiler::primitives::globals::finalize_global_table(chunks);
    let wasm = vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm).unwrap();
    wasm
}

fn kernel(size: f64, realloc: bool, baseline: bool, repeats: i32) -> Vec<Chunk> {
    let mut chunks = setup();
    let remaining = chunks[0].alloc_scratch(1);
    let result = chunks[0].alloc_scratch(1);
    let c = &mut chunks[0];
    c.emit_i32_const(repeats, 0);
    c.emit_op_u16(Op::LOCAL_SET, remaining, 0);
    c.emit_loop_s(0);
    if realloc {
        c.emit_f64_const(8.0, 0);
        emit(&mut chunks, false, baseline);
        let old = chunks[0].alloc_scratch(1);
        let c = &mut chunks[0];
        c.emit_op_u16(Op::LOCAL_TEE, old, 0);
        c.emit_i32_const(129, 0);
        c.emit_i32_const(8, 0);
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
        c.emit_op_u16(Op::LOCAL_GET, old, 0);
    }
    chunks[0].emit_f64_const(size, 0);
    emit(&mut chunks, realloc, baseline);
    let c = &mut chunks[0];
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, remaining, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_TEE, remaining, 0);
    c.emit_br_if(0, 0);
    c.emit_end(0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    finish(&mut chunks);
    chunks
}

#[test]
fn heap_transport_preserves_headers_prefixes_and_size_guards() {
    for baseline in [true, false] {
        for realloc in [false, true] {
            for size in [-0.0, 0.0, 1.0, 7.0, 7.5, 8.0, 17.0, 65537.0,
                -1.0, f64::NAN, f64::INFINITY, (u32::MAX - 14) as f64, u32::MAX as f64] {
                let mut vm = runtime();
                let result = vm.run(kernel(size, realloc, baseline, 1)).unwrap().as_i32();
                if !size.is_finite() || size < 0.0 || size > (u32::MAX - 15) as f64
                    || realloc && size == 0.0 {
                    assert_eq!(result, 0);
                } else {
                    assert!(result >= 1032);
                    vm.memory.with_buffer(|bytes| {
                        let at = result as usize;
                        assert_eq!(&bytes[at - 8..at - 4], &(size as u32).to_le_bytes());
                        if realloc {
                            assert!(bytes[at..at + (size as usize).min(8)].iter()
                                .all(|byte| *byte == 129));
                        }
                    });
                }
                if realloc {
                    vm.memory.with_buffer(|bytes| assert_eq!(&bytes[1032..1040], &[129; 8]));
                }
            }
        }
    }
}

#[test]
fn heap_transport_null_and_invalid_old_pointer_trap_parity() {
    for baseline in [true, false] {
        for (old, size, traps) in [(0, 17.0, false), (0, 0.0, false),
            (16777220, 17.0, true), (16777220, 0.0, false), (16777220, -1.0, false)] {
            let mut chunks = setup();
            chunks[0].emit_f64_const(f64::from(old), 0);
            chunks[0].emit_f64_const(size, 0);
            emit(&mut chunks, true, baseline);
            finish(&mut chunks);
            let mut vm = runtime();
            let result = vm.run(chunks);
            if traps {
                assert!(result.unwrap_err().to_string().contains("out of bounds"));
                vm.memory.with_buffer(|bytes| assert_eq!(&bytes[1024..1028], &17u32.to_le_bytes()));
            } else {
                assert_eq!(result.unwrap().as_i32(), if size > 0.0 { 1032 } else { 0 });
            }
        }
    }
}

#[test]
#[cfg(debug_assertions)]
fn heap_transport_debug_measurements() {
    use std::sync::atomic::Ordering;
    use vybe_runtime::{DebugCommand, DebugRequest};
    for realloc in [false, true] {
        for size in [0.0, 8.0, 4096.0] {
            let mut before = None;
            for baseline in [true, false] {
                let chunks = kernel(size, realloc, baseline, 1);
                let wasm = vybe_platform_wasm::write_wasm_with_options(&chunks,
                    vybe_platform_wasm::WasmWriteOptions {
                        include_vybe_metadata: false, include_name_section: true,
                    });
                let mut calls = 0;
                let mut operators = 0;
                for payload in wasmparser::Parser::new(0).parse_all(&wasm) {
                    if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
                        let mut reader = body.get_operators_reader().unwrap();
                        while !reader.eof() {
                            if matches!(reader.read().unwrap(), wasmparser::Operator::Call { .. }) { calls += 1; }
                            operators += 1;
                        }
                    }
                }
                if baseline {
                    before = Some((operators, calls));
                } else {
                    let (old_ops, old_calls) = before.unwrap();
                    assert_eq!(old_ops - operators, if realloc { 8 } else { 3 });
                    assert_eq!(old_calls - calls, if realloc { 2 } else { 1 });
                }
                if let Some(dir) = std::env::var_os("LIBC_HEAP_WASM_DIR") {
                    std::fs::write(std::path::Path::new(&dir).join(
                        format!("heap-realloc-{realloc}-size-{size}-baseline-{baseline}.wasm")), &wasm).unwrap();
                }
                let mut vm = runtime();
                let (tx, rx) = std::sync::mpsc::channel();
                let (events, _rx) = std::sync::mpsc::channel();
                vm.attach_debugger(rx, events, true);
                for command in [DebugCommand::ExceptionBreak { on_throw: true, on_uncaught: true }, DebugCommand::Continue] {
                    let (reply, _rx) = std::sync::mpsc::channel();
                    tx.send(DebugRequest { command, reply }).unwrap();
                }
                let report = vm.debug_report().unwrap();
                vm.run(chunks).unwrap();
                let instructions = report.lock().unwrap().live.instructions.load(Ordering::Relaxed);
                let mut times = Vec::new();
                for _ in 0..7 {
                    let chunks = kernel(size, realloc, baseline, 32);
                    let mut vm = runtime();
                    let start = std::time::Instant::now();
                    vm.run(chunks).unwrap();
                    times.push(start.elapsed());
                }
                times.sort();
                eprintln!("heap transport: realloc={realloc} size={size} baseline={baseline} instructions={instructions} wasm_ops={operators} wasm_calls={calls} median32={:?}", times[3]);
            }
        }
    }
}

#[test]
fn heap_transport_c_fixture_equivalence() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/libc_heap_transport_equivalence.c");
    let mut vm = runtime();
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    let wasm = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm).unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
