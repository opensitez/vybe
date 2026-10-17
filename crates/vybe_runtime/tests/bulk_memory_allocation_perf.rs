//! Allocation traffic for a VM loop, not isolated kernel or Doom timing.
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use vybe_runtime::{VM, Value};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;

struct CountingAllocator;
static ENABLED: AtomicBool = AtomicBool::new(false);
static CALLS: AtomicUsize = AtomicUsize::new(0);
static BYTES: AtomicUsize = AtomicUsize::new(0);

fn record(pointer: *mut u8, size: usize) {
    if !pointer.is_null() && ENABLED.load(Ordering::Relaxed) {
        CALLS.fetch_add(1, Ordering::Relaxed);
        BYTES.fetch_add(size, Ordering::Relaxed);
    }
}

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        record(pointer, layout.size());
        pointer
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc_zeroed(layout) };
        record(pointer, layout.size());
        pointer
    }
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        let pointer = unsafe { System.realloc(pointer, layout, size) };
        record(pointer, size);
        pointer
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn constant(chunk: &mut Chunk, mut value: i32) {
    chunk.emit_op(Op::I32_CONST, 0);
    loop {
        let byte = (value as u8) & 0x7f;
        value >>= 7;
        let done = (value == 0 && byte & 0x40 == 0)
            || (value == -1 && byte & 0x40 != 0);
        chunk.code.push(byte | if done { 0 } else { 0x80 });
        if done { break; }
    }
}

fn workload(rounds: i32) -> Chunk {
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.result_arity = 1;
    chunk.memory_min_pages = vec![1];
    chunk.data_segments = vec![vec![0x33; 4096]];
    constant(&mut chunk, rounds);
    chunk.emit_op(Op::LOCAL_SET, 0);
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 0);
    for (op, destination, source) in [
        (Op::MEMORY_FILL, 0, 0x5a),
        (Op::MEMORY_COPY, 4096, 0),
        (Op::MEMORY_INIT, 8192, 0),
    ] {
        constant(&mut chunk, destination);
        constant(&mut chunk, source);
        constant(&mut chunk, 4096);
        chunk.emit_op(op, 0);
        chunk.code.push(0);
        if op != Op::MEMORY_FILL { chunk.code.push(0); }
    }
    chunk.emit_op(Op::LOCAL_GET, 0);
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
    constant(&mut chunk, 1);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op(Op::LOCAL_TEE, 0);
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
    chunk.emit_op(Op::BR_IF, 0);
    chunk.code.push(0);
    chunk.emit_op(Op::END, 0);
    for (index, address) in [0, 4096, 8192].into_iter().enumerate() {
        constant(&mut chunk, address);
        chunk.emit_op(Op::I32_LOAD8_U, 0);
        chunk.code.extend_from_slice(&[0x80, 0x01, 0]);
        if index != 0 { chunk.emit_op(Op::I32_XOR, 0); }
    }
    chunk.emit_op(Op::RETURN, 0);
    chunk.lines.resize(chunk.code.len(), 0);
    chunk
}

#[test]
#[ignore = "allocation measurement; run alone with one test thread"]
fn bulk_loop_does_not_allocate_per_operation() {
    let mut baseline = None;
    for rounds in [128, 1024, 8192] {
        let chunk = workload(rounds);
        let mut vm = VM::new();
        CALLS.store(0, Ordering::Relaxed);
        BYTES.store(0, Ordering::Relaxed);
        ENABLED.store(true, Ordering::Relaxed);
        let result = vm.run(vec![chunk]);
        ENABLED.store(false, Ordering::Relaxed);
        let calls = CALLS.load(Ordering::Relaxed);
        let bytes = BYTES.load(Ordering::Relaxed);
        assert_eq!(result.unwrap(), Value::I32(0x33));
        if let Some(first) = baseline {
            assert!(calls <= first + 2, "allocation calls grew with loop count");
        } else {
            baseline = Some(calls);
        }
        eprintln!("bulk rounds={rounds} operations={} allocations={calls} requested_bytes={bytes}", rounds * 3);
    }
}

#[test]
#[ignore = "execution-counter diagnostic; separate from allocation measurement"]
fn bulk_loop_execution_counters() {
    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    assert_eq!(vm.run(vec![workload(128)]).unwrap(), Value::I32(0x33));
    let counters = vm.take_runtime_perf().expect("enabled runtime counters");
    let report = counters.format_report();
    let count = |key: &str| -> u64 {
        report.split_whitespace()
            .find_map(|token| token.strip_prefix(key)?.strip_prefix('='))
            .unwrap_or_else(|| panic!("missing counter {key}"))
            .parse().expect("numeric runtime counter")
    };
    assert_eq!(count("memory_copy"), 128);
    assert_eq!(count("memory_fill"), 128);
    assert_eq!(count("memory_copy_bytes"), 128 * 4096);
    assert_eq!(count("memory_fill_bytes"), 128 * 4096);
    assert_eq!(count("br_if"), 128);
    assert_eq!(count("br_if_taken"), 127);
    assert_eq!(count("br_if_not_taken"), 1);
    assert_eq!(count("i32_load8_u"), 3);
    assert_eq!(count("host_call"), 0);
    eprintln!("{report}\n{}", counters.bridge_summary());
}
