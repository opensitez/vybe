use std::path::Path;
use vybe_runtime::{Chunk, Op, VM};

fn runtime() -> VM {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

pub(crate) fn string_copy_kernel(duplicate: bool, length: i32, capacity: i32, repeats: i32) -> Vec<Chunk> {
    use vybe_compiler::primitives::{bundle, globals};
    let mut chunks = vec![Chunk::new("<script>")];
    let c = &mut chunks[0];
    c.emit_i32_const(2, 0);
    c.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_i32_const(1024, 0);
    c.emit_i32_const(129, 0);
    c.emit_i32_const(length, 0);
    c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    c.emit_i32_const(1024 + length, 0);
    c.emit_i32_const(0, 0);
    c.emit_op(Op::I32_STORE8, 0);
    c.emit_i32_const(16384, 0);
    globals::emit_write(c, "__c_heap_ptr", 0);
    let remaining = c.alloc_scratch(1);
    c.emit_i32_const(repeats, 0);
    c.emit_op_u16(Op::LOCAL_SET, remaining, 0);
    let (repeat, _) = c.emit_loop_s(0);
    // Exercise C's Number-backed value ABI, not only native word constants.
    if !duplicate { c.emit_f64_const(16384.0, 0); }
    c.emit_f64_const(1024.0, 0);
    if duplicate {
        vybe_platform_libc::emitter::string_memory::emit_strdup(&mut chunks, 0, 0);
    } else {
        chunks[0].emit_f64_const(f64::from(capacity), 0);
        vybe_platform_libc::emitter::string_memory::emit_strncpy(&mut chunks, 0, 0);
    }
    let c = &mut chunks[0];
    let result = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, remaining, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_TEE, remaining, 0);
    c.emit_br_if(0, 0);
    c.emit_end(0);
    c.patch_loop(repeat);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::finalize_global_table(&mut chunks);
    chunks
}

fn finish_string_kernel(chunks: &mut Vec<Chunk>) {
    vybe_compiler::primitives::bundle::finalize_with_runtime_helpers(chunks);
    vybe_compiler::primitives::globals::finalize_global_table(chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
}

#[test]
fn string_linear_copy_preserves_aliases_and_trap_phase_writes() {
    for destination in [15, 16, 17, 40] {
        let mut expected: Vec<u8> = (0..128).map(|i| (i * 37 + 129) as u8).collect();
        expected[16..24].fill(129);
        expected[24] = 0;
        let mut c = Chunk::new("<script>");
        c.emit_i32_const(1, 0);
        c.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        c.emit_op(Op::DROP, 0);
        for (index, byte) in expected.iter().enumerate() {
            c.emit_i32_const(index as i32, 0);
            c.emit_i32_const(i32::from(*byte), 0);
            c.emit_op(Op::I32_STORE8, 0);
        }
        expected.copy_within(16..24, destination as usize);
        expected[destination as usize + 8..destination as usize + 12].fill(0);
        for value in [destination, 16, 12] { c.emit_i32_const(value, 0); }
        let mut chunks = vec![c];
        vybe_platform_libc::emitter::string_memory::emit_strncpy(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        finish_string_kernel(&mut chunks);
        let mut vm = runtime();
        assert_eq!(vm.run(chunks).unwrap().as_i32(), destination);
        vm.memory.with_buffer(|bytes| assert_eq!(&bytes[..128], &expected));
    }
    // Scan failure, atomic copy failure, then padding failure after a copied byte.
    for (source, destination, length, expected_byte) in [
        (65535, 16, 1, 7), (16, 65535, 2, 7), (16, 65535, 1, 129),
    ] {
        let mut c = Chunk::new("<script>");
        c.emit_i32_const(1, 0);
        c.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(destination, 0);
        c.emit_i32_const(7, 0);
        c.emit_op(Op::I32_STORE8, 0);
        for index in 0..length {
            c.emit_i32_const(source + index, 0);
            c.emit_i32_const(129, 0);
            c.emit_op(Op::I32_STORE8, 0);
        }
        if source + length < 65536 {
            c.emit_i32_const(source + length, 0);
            c.emit_i32_const(0, 0);
            c.emit_op(Op::I32_STORE8, 0);
        }
        for value in [destination, source, 2] { c.emit_i32_const(value, 0); }
        let mut chunks = vec![c];
        vybe_platform_libc::emitter::string_memory::emit_strncpy(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        finish_string_kernel(&mut chunks);
        let mut vm = runtime();
        let error = vm.run(chunks).unwrap_err();
        assert!(error.to_string().contains("out of bounds"), "{error}");
        vm.memory.with_buffer(|bytes| assert_eq!(bytes[destination as usize], expected_byte));
    }
}

#[test]
fn string_copy_managed_source_keeps_utf8_and_owned_duplicate() {
    for duplicate in [false, true] {
        let mut chunks = vec![Chunk::new("<script>")];
        let c = &mut chunks[0];
        c.emit_i32_const(1, 0);
        c.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(1024, 0);
        vybe_compiler::primitives::globals::emit_write(c, "__c_heap_ptr", 0);
        if !duplicate { c.emit_i32_const(1024, 0); }
        c.emit_string_const("\u{e9}\u{ff}", 0);
        if duplicate {
            vybe_platform_libc::emitter::string_memory::emit_strdup(&mut chunks, 0, 0);
        } else {
            chunks[0].emit_i32_const(7, 0);
            vybe_platform_libc::emitter::string_memory::emit_strncpy(&mut chunks, 0, 0);
        }
        chunks[0].emit_op(Op::RETURN, 0);
        finish_string_kernel(&mut chunks);
        let mut vm = runtime();
        let address = vm.run(chunks).unwrap().as_i32() as usize;
        vm.memory.with_buffer(|bytes| {
            assert_eq!(&bytes[address..address + 5], &[0xc3, 0xa9, 0xc3, 0xbf, 0]);
            if duplicate {
                assert_eq!(&bytes[address - 8..address - 4], &5u32.to_le_bytes());
            } else {
                assert_eq!(&bytes[address + 5..address + 7], &[0, 0]);
            }
        });
    }
}

#[test]
fn string_linear_copy_debug_instruction_workloads() {
    use std::sync::atomic::Ordering;
    use vybe_runtime::{DebugCommand, DebugRequest};
    for duplicate in [false, true] {
        for length in [0, 8, 4096] {
            let mut vm = runtime();
            let (commands, rx) = std::sync::mpsc::channel();
            let (events, _rx) = std::sync::mpsc::channel();
            vm.attach_debugger(rx, events, true);
            for command in [DebugCommand::ExceptionBreak { on_throw: true, on_uncaught: true },
                DebugCommand::Continue] {
                let (reply, _response) = std::sync::mpsc::channel();
                commands.send(DebugRequest { command, reply }).unwrap();
            }
            let report = vm.debug_report().unwrap();
            let chunks = string_copy_kernel(duplicate, length, length + 3, 1);
            let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                });
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes).unwrap();
            let address = vm.run(chunks).unwrap().as_i32() as usize;
            let count = report.lock().unwrap().live.instructions.load(Ordering::Relaxed);
            eprintln!("libc linear string copy: strdup={duplicate}, length={length}, instructions={count}");
            vm.memory.with_buffer(|bytes| {
                assert!(bytes[address..address + length as usize].iter().all(|byte| *byte == 129));
                let padding = if duplicate { 1 } else { 3 };
                assert!(bytes[address + length as usize..address + length as usize + padding]
                    .iter().all(|byte| *byte == 0));
            });
            let mut timings = Vec::new();
            for _ in 0..5 {
                let chunks = string_copy_kernel(duplicate, length, length + 3, 16);
                let mut vm = runtime();
                let start = std::time::Instant::now();
                vm.run(chunks).unwrap();
                timings.push(start.elapsed());
            }
            timings.sort();
            eprintln!("libc linear string copy: strdup={duplicate}, length={length}, \
                median 16 calls/5 debug runs={:?}", timings[2]);
        }
    }
}

#[test]
fn existing_copy_fill_realloc_and_sdl_tick_paths_remain_valid() {
    for (name, required) in [
        ("dynamic_memcpy_bulk", Some("memory.copy")),
        ("byte_pointer_memset_bulk", Some("memory.fill")),
        ("realloc_null_static_guard", Some("memory.copy")),
        ("sdl_tick_comparison", None),
    ] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{name}.c"));
        let mut vm = runtime();
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path).unwrap();
        let dump = compiled.chunks.iter().map(vybe_runtime::debug::disassemble)
            .collect::<Vec<_>>().join("\n");
        if let Some(required) = required {
            assert!(dump.contains(required), "{name}: missing {required}");
        } else {
            assert!(!dump.contains("dyneq"), "{name}: dynamic tick comparison");
        }
        let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled).unwrap();
        assert_eq!(vm.pending_exit_code, 0, "{name}");
    }
}

#[test]
#[ignore = "computed strncpy destination is rewritten before the libc emitter; frontend owner gate"]
fn computed_string_copy_destinations_preserve_pointer_results() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/libc_linear_string_copy_alias.c");
    let mut vm = runtime();
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&path).unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn sdl_bulk_blit_preserves_overlap_bytes_and_conversion_fallbacks() {
    for name in ["libc_bulk_sdl_byte_equivalence", "sdl_blit_pitch_masks",
        "sdl_indexed_to_rgba_full_blit"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{name}.c"));
        let mut vm = runtime();
        let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_path(&path).unwrap();
        let dumps = compiled.chunks.iter().map(vybe_runtime::debug::disassemble)
            .collect::<Vec<_>>().join("\n");
        assert!(dumps.contains("memory.copy"), "{name}: no standard bulk copy");
        assert!(compiled.chunks.iter().any(|chunk| chunk.imports.iter()
            .any(|import| import.module == "libc:sdl" && import.name == "bulkBlitLength")),
            "{name}: bulk-blit proof was not emitted");
        let wasm = vybe_platform_wasm::write_wasm_with_options(&compiled.chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&wasm).unwrap();
        vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .run_compiled(compiled).unwrap();
        assert_eq!(vm.pending_exit_code, 0, "{name}");
    }
}

#[test]
fn column_bulk_preserves_forward_aliasing_and_unsigned_bytes() {
    for (destination, source, stride, count) in [
        (65, 64, 1, 8), (64, 65, 1, 8), (64, 64, 1, 8),
        (80, 64, 1, 8), (72, 64, 1, 8), (40, 64, -3, 8),
        (16, 64, 0, 8), (-1, -1, 1, 0),
    ] {
        let mut expected: Vec<u8> = (0..128).map(|i| (i * 37 + 129) as u8).collect();
        for i in 0..count {
            expected[(destination + i * stride) as usize] = expected[(source + i) as usize];
        }
        let mut chunk = Chunk::new("<script>");
        chunk.emit_i32_const(1, 0);
        chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        chunk.emit_op(Op::DROP, 0);
        for i in 0..128 {
            chunk.emit_i32_const(i, 0);
            chunk.emit_i32_const((i * 37 + 129) & 255, 0);
            chunk.emit_op(Op::I32_STORE8, 0);
        }
        for value in [destination, source, stride, count] { chunk.emit_i32_const(value, 0); }
        let mut chunks = vec![chunk];
        vybe_platform_libc::emitter::byte_memory::emit_copy_column(&mut chunks, 0, 0);
        chunks[0].emit_i32_const(destination, 0);
        chunks[0].emit_op(Op::I32_NE, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        let mut vm = runtime();
        assert_eq!(vm.run(chunks).unwrap().as_i32(), 0);
        vm.memory.with_buffer(|bytes| assert_eq!(&bytes[..128], &expected));
    }
}

#[test]
fn column_copy_retains_partial_writes_before_bounds_trap() {
    for (destination, source, touched) in [(16, 65534, 16), (65534, 16, 65534)] {
        let mut chunk = Chunk::new("<script>");
        chunk.emit_i32_const(1, 0);
        chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        chunk.emit_op(Op::DROP, 0);
        for i in 0..2 {
            chunk.emit_i32_const(source + i, 0);
            chunk.emit_i32_const(129 + i, 0);
            chunk.emit_op(Op::I32_STORE8, 0);
        }
        for value in [destination, source, 1, 3] {
            chunk.emit_i32_const(value, 0);
        }
        let mut chunks = vec![chunk];
        vybe_platform_libc::emitter::byte_memory::emit_copy_column(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        let mut vm = runtime();
        let error = vm.run(chunks).unwrap_err();
        assert!(error.to_string().contains("out of bounds"), "{error}");
        vm.memory.with_buffer(|bytes| {
            assert_eq!(&bytes[touched as usize..touched as usize + 2], &[129, 130]);
        });
    }
}
