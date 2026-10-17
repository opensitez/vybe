use std::path::Path;
use std::sync::atomic::Ordering;
use vybe_runtime::{Chunk, DebugCommand, DebugRequest, Op, VM};

fn runtime() -> VM {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

fn run_counted(chunks: Vec<Chunk>) -> u64 {
    run_counted_with_precision(chunks, false)
}

fn run_counted_with_precision(chunks: Vec<Chunk>, exact: bool) -> u64 {
    let mut vm = runtime();
    let (commands, rx) = std::sync::mpsc::channel();
    let (events, _rx) = std::sync::mpsc::channel();
    vm.attach_debugger(rx, events, exact);
    if exact {
        // An active stop condition publishes every instruction, rather than
        // the free-running mode's 64-instruction snapshots. Kernels do not throw.
        let (reply, _response) = std::sync::mpsc::channel();
        commands
            .send(DebugRequest {
                command: DebugCommand::ExceptionBreak {
                    on_throw: true,
                    on_uncaught: true,
                },
                reply,
            })
            .unwrap();
    }
    // Global initializers can request an entry pause even in free-running mode.
    let (reply, _response) = std::sync::mpsc::channel();
    commands
        .send(DebugRequest {
            command: DebugCommand::Continue,
            reply,
        })
        .unwrap();
    let report = vm.debug_report().unwrap();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    let count = report
        .lock()
        .unwrap()
        .live
        .instructions
        .load(Ordering::Relaxed);
    count
}

#[test]
fn column_directive_matches_portable_scalar_reference_with_fewer_opcodes() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_byte_column.c");
    vybe_language_c::register();
    let mut bundle = vybe_compiler::projects::load(&path).unwrap();
    let scalar = run_counted(bundle.compile().unwrap());
    bundle.sources[0].code = bundle.sources[0].code.replace(
        "#define COLUMN scalar_column",
        "#define COLUMN __c_ptr_copy_column",
    );
    let chunks = bundle.compile().unwrap();
    let main = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
    let dump = vybe_runtime::debug::disassemble(main);
    assert!(dump.contains("i32.load8_u"), "{dump}");
    let emitted = run_counted(chunks);
    eprintln!("column fixture VM instructions: scalar={scalar}, emitted={emitted}");
    assert!(emitted < scalar, "scalar={scalar}, emitted={emitted}");
}

#[test]
fn column_emission_validates_and_runs_without_metadata() {
    for (destination, source, stride, count) in [
        (16, 64, 3, 8),
        (40, 64, -3, 8),
        (16, 64, 0, 8),
        (65, 64, 1, 8),
        (64, 65, 1, 8),
        (64, 64, 1, 8),
        (80, 64, 1, 8),
        (72, 64, 1, 8),
        (-1, -1, 1, 0),
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
        for value in [destination, source, stride, count] {
            chunk.emit_i32_const(value, 0);
        }
        let mut chunks = vec![chunk];
        vybe_platform_libc::emitter::byte_memory::emit_copy_column(&mut chunks, 0, 0);
        let chunk = &mut chunks[0];
        chunk.emit_i32_const(destination, 0);
        chunk.emit_op(Op::I32_NE, 0);
        for (i, byte) in expected.into_iter().enumerate() {
            chunk.emit_i32_const(i as i32, 0);
            chunk.emit_op(Op::I32_LOAD8_U, 0);
            chunk.emit_i32_const(byte as i32, 0);
            chunk.emit_op(Op::I32_NE, 0);
            chunk.emit_op(Op::I32_OR, 0);
        }
        chunk.emit_op(Op::RETURN, 0);
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(
            !dump.contains("CallHost") && !dump.contains("call_ref"),
            "{dump}"
        );
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            &chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        assert_eq!(runtime().run(chunks).unwrap().as_i32(), 0, "{dump}");
        let reloaded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
        let mut vm = runtime();
        vm.run(reloaded).unwrap();
        let mut function = vybe_runtime::value::Object::new();
        function.kind = vybe_runtime::value::ObjectKind::Function(vybe_runtime::value::Function {
            name: None,
            arity: 0,
            chunk_index: 1,
            upvalues: Vec::new(),
        });
        let function = vybe_runtime::value::Value::Object(std::sync::Arc::new(
            std::sync::Mutex::new(function),
        ));
        assert_eq!(
            vm.invoke(&function, &[]).unwrap().as_i32(),
            0,
            "reloaded {dump}"
        );
    }
}

#[test]
fn column_loop_has_seventeen_standard_opcodes_and_retains_memory_traps() {
    for (destination, source) in [(16, -1), (-1, 16)] {
        let mut chunk = Chunk::new("<script>");
        chunk.emit_i32_const(1, 0);
        chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        chunk.emit_op(Op::DROP, 0);
        for value in [destination, source, 1, 1] {
            chunk.emit_i32_const(value, 0);
        }
        let mut chunks = vec![chunk];
        vybe_platform_libc::emitter::byte_memory::emit_copy_column(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        let code = &chunks[0].code;
        let mut offset = 0;
        let mut body = Vec::new();
        let mut in_loop = false;
        while offset < code.len() {
            let op = Op::from_code_at(code, offset);
            if in_loop {
                body.push(op);
                if op == Op::BR_IF {
                    break;
                }
            }
            in_loop |= op == Op::LOOP;
            offset += 4 + op.operand_format().size_in(code, offset + 4);
        }
        assert_eq!(body.len(), 17, "{body:?}");
        let error = runtime().run(chunks).unwrap_err();
        assert!(error.to_string().contains("out of bounds"), "{error}");
    }
}

#[test]
fn contiguous_column_cost_is_constant_and_strided_cost_is_seventeen_per_byte() {
    let column = |count, stride| {
        let mut chunk = Chunk::new("<script>");
        chunk.emit_i32_const(2, 0);
        chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        chunk.emit_op(Op::DROP, 0);
        for value in [1024, 16, stride, count] {
            chunk.emit_i32_const(value, 0);
        }
        let mut chunks = vec![chunk];
        vybe_platform_libc::emitter::byte_memory::emit_copy_column(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        chunks
    };
    let bulk_one = run_counted_with_precision(column(1, 1), true);
    let bulk_many = run_counted_with_precision(column(256, 1), true);
    let strided_one = run_counted_with_precision(column(1, 320), true);
    let strided_many = run_counted_with_precision(column(256, 320), true);
    eprintln!(
        "column kernel VM instructions: bulk(1)={bulk_one}, bulk(256)={bulk_many}, \
        strided(1)={strided_one}, strided(256)={strided_many}"
    );
    assert_eq!(bulk_many, bulk_one);
    assert_eq!(strided_many - strided_one, 255 * 17);
    assert!(bulk_many < strided_many);
}
