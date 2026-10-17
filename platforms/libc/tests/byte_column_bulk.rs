use std::sync::atomic::Ordering;
use vybe_runtime::{Chunk, DebugCommand, DebugRequest, Op, VM};

#[path = "../src/emitter/byte_memory.rs"]
mod byte_memory;

type Emitter = fn(&mut [Chunk], usize, u32);

fn kernel(
    emit: Emitter,
    destination: i32,
    source: i32,
    stride: i32,
    count: i32,
    repeats: i32,
) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<script>")];
    let c = &mut chunks[0];
    c.emit_i32_const(1, 0);
    c.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_i32_const(0, 0);
    c.emit_i32_const(129, 0);
    c.emit_i32_const(65536, 0);
    c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    for i in 0..128 {
        c.emit_i32_const(i, 0);
        c.emit_i32_const((i * 37 + 129) & 255, 0);
        c.emit_op(Op::I32_STORE8, 0);
    }
    let remaining = c.alloc_scratch(1);
    c.emit_i32_const(repeats, 0);
    c.emit_op_u16(Op::LOCAL_SET, remaining, 0);
    let (repeat, _) = c.emit_loop_s(0);
    for value in [destination, source, stride, count] {
        c.emit_f64_const(f64::from(value), 0);
    }
    emit(&mut chunks, 0, 0);
    let c = &mut chunks[0];
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, remaining, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_TEE, remaining, 0);
    c.emit_br_if(0, 0);
    c.emit_end(0);
    c.patch_loop(repeat);
    c.emit_i32_const(0, 0);
    c.emit_op(Op::RETURN, 0);
    chunks
}

fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
}

pub fn measure(emit: Emitter) {
    for (label, destination, stride, count) in [
        ("strided", 2048, 320, 64),
        ("forward-dependent", 1025, 1, 64),
        ("bulk", 2048, 1, 64),
        ("empty", -1, 1, 0),
    ] {
        let source = if count == 0 { -1 } else { 1024 };
        let chunks = kernel(emit, destination, source, stride, count, 1);
        validate(&chunks);
        let mut vm = VM::new();
        let (commands, rx) = std::sync::mpsc::channel();
        let (events, _rx) = std::sync::mpsc::channel();
        vm.attach_debugger(rx, events, true);
        for command in [
            DebugCommand::ExceptionBreak {
                on_throw: true,
                on_uncaught: true,
            },
            DebugCommand::Continue,
        ] {
            let (reply, _response) = std::sync::mpsc::channel();
            commands.send(DebugRequest { command, reply }).unwrap();
        }
        let report = vm.debug_report().unwrap();
        assert_eq!(vm.run(chunks).unwrap().as_i32(), 0);
        let instructions = report
            .lock()
            .unwrap()
            .live
            .instructions
            .load(Ordering::Relaxed);
        let mut times = Vec::new();
        for _ in 0..5 {
            let chunks = kernel(emit, destination, source, stride, count, 128);
            let mut vm = VM::new();
            let start = std::time::Instant::now();
            assert_eq!(vm.run(chunks).unwrap().as_i32(), 0);
            times.push(start.elapsed());
        }
        times.sort();
        eprintln!(
            "column {label}: instructions={instructions}, median 128 calls={:?}",
            times[2]
        );
    }
}

#[test]
fn column_guard_debug_workloads() {
    measure(byte_memory::emit_copy_column);
}

#[test]
fn column_scalar_bounds_and_aliases_keep_ordered_writes() {
    for (destination, source, stride, count) in [
        (65, 64, 1, 8),
        (64, 65, 1, 8),
        (64, 64, 1, 8),
        (80, 64, 1, 8),
        (72, 64, 1, 8),
        (40, 64, -3, 8),
        (16, 64, 0, 8),
        (-1, -1, 1, 0),
        (65534, 16, 1, 3),
        (16, 65534, 1, 3),
        (65533, 16, 2, 3),
        (16, 65534, 2, 3),
        (65, 64, 1, 65536),
        (0, 64, -1, 3),
        (-1, 16, 1, 3),
        (16, -1, 1, 3),
    ] {
        let mut expected = vec![129u8; 65536];
        for (i, byte) in expected[..128].iter_mut().enumerate() {
            *byte = (i * 37 + 129) as u8;
        }
        let mut trapped = false;
        for i in 0..count {
            let from = (source as u32).wrapping_add(i as u32) as usize;
            let to =
                (destination as u32).wrapping_add((i as u32).wrapping_mul(stride as u32)) as usize;
            if from >= expected.len() || to >= expected.len() {
                trapped = true;
                break;
            }
            expected[to] = expected[from];
        }
        let chunks = kernel(
            byte_memory::emit_copy_column,
            destination,
            source,
            stride,
            count,
            1,
        );
        validate(&chunks);
        let mut vm = VM::new();
        let result = vm.run(chunks);
        assert_eq!(
            result.is_err(),
            trapped,
            "{destination}/{source}/{stride}/{count}"
        );
        if let Err(error) = result {
            assert!(error.to_string().contains("out of bounds"), "{error}");
        }
        vm.memory.with_buffer(|bytes| assert_eq!(bytes, &expected));
    }
}
