//! Requested allocation traffic, not RSS or a timing benchmark.
//! Run alone with --test-threads=1: the allocator counters are process-wide.
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Value, VM};

struct CountingAllocator;
static CALLS: AtomicUsize = AtomicUsize::new(0);
static REQUESTED: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            CALLS.fetch_add(1, Ordering::Relaxed);
            REQUESTED.fetch_add(layout.size(), Ordering::Relaxed);
        }
        ptr
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc_zeroed(layout) };
        if !ptr.is_null() {
            CALLS.fetch_add(1, Ordering::Relaxed);
            REQUESTED.fetch_add(layout.size(), Ordering::Relaxed);
        }
        ptr
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        let result = unsafe { System.realloc(ptr, layout, size) };
        if !result.is_null() {
            CALLS.fetch_add(1, Ordering::Relaxed);
            REQUESTED.fetch_add(size, Ordering::Relaxed);
        }
        result
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op(op, 0);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 0);
}

fn fixture(rounds: i32, simd: bool) -> Chunk {
    let mut chunk = Chunk::new("memory-assignment-allocation");
    chunk.local_count = 3;
    chunk.scratch_high_water = 3;
    chunk.result_arity = 1;
    chunk.emit_i32_const(rounds, 0);
    local(&mut chunk, Op::LOCAL_SET, 0);
    chunk.emit_i32_const(42, 0);
    local(&mut chunk, Op::LOCAL_SET, 1);
    chunk.emit_op(Op::LOOP, 0);
    chunk.code.extend_from_slice(&[0, 0]);
    chunk.lines.resize(chunk.code.len(), 0);
    for _ in 0..32 {
        if simd {
            chunk.emit_i32_const(64, 0);
            chunk.emit_i32_const(3, 0);
            chunk.emit_op(Op::V128_LOAD, 0);
            chunk.emit_op(Op::V128_STORE, 0);
        } else {
            local(&mut chunk, Op::LOCAL_GET, 1);
            local(&mut chunk, Op::LOCAL_SET, 2);
        }
    }
    local(&mut chunk, Op::LOCAL_GET, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    local(&mut chunk, Op::LOCAL_TEE, 0);
    chunk.emit_op(Op::BR_IF, 0);
    chunk.code.push(0);
    chunk.lines.resize(chunk.code.len(), 0);
    chunk.emit_op(Op::END, 0);
    if simd {
        chunk.emit_i32_const(42, 0);
    } else {
        local(&mut chunk, Op::LOCAL_GET, 2);
    }
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

fn measure(rounds: i32, simd: bool) -> (usize, usize) {
    let mut vm = VM::new();
    vm.memory.resize(128, 0);
    vm.memory.write_bytes(3, &[0x5a; 16]);
    let chunks = vec![fixture(rounds, simd)];
    let calls = CALLS.load(Ordering::Relaxed);
    let bytes = REQUESTED.load(Ordering::Relaxed);
    let result = vm.run(chunks);
    let traffic = (
        CALLS.load(Ordering::Relaxed) - calls,
        REQUESTED.load(Ordering::Relaxed) - bytes,
    );
    assert!(matches!(result.unwrap(), Value::I32(42)));
    if simd {
        let mut output = [0; 16];
        assert_eq!(vm.memory.read_bytes(64, &mut output), 16);
        assert_eq!(output, [0x5a; 16]);
    }
    traffic
}

#[test]
#[ignore = "process-wide allocation measurement; run alone with one test thread"]
fn allocation_traffic_by_iteration_count() {
    for simd in [false, true] {
        measure(1, simd);
        for rounds in [128, 1024, 8192] {
            let (calls, bytes) = measure(rounds, simd);
            eprintln!("simd={simd} operations={} allocations={calls} requested_bytes={bytes}", rounds * 32);
        }
    }
}

fn measure_multicast(rounds: i32, handler_count: usize) -> (usize, usize) {
    use std::sync::Arc;
    use vybe_runtime::value::{Object, ObjectKind};
    let mut vm = VM::new();
    let observed = Arc::new(AtomicUsize::new(0));
    let callback_count = observed.clone();
    vm.register_host_fn("test:multicast", "allocation_probe", Box::new(move |_, args| {
        assert_eq!(args, &[Value::I32(7)]);
        callback_count.fetch_add(1, Ordering::Relaxed);
        Value::I32(42)
    }));
    let index = vm.host_registry[&("test:multicast".to_string(), "allocation_probe".to_string())];
    let mut function = Object::new();
    function.kind = ObjectKind::HostFunction(index);
    let handler = Value::Object(Arc::new(std::sync::Mutex::new(function)));
    let callee = Object::new_array(vec![handler; handler_count]);
    vm.set_global("__multicast_allocation_probe", Value::Object(Arc::new(std::sync::Mutex::new(callee))));

    let mut chunk = Chunk::new("multicast-allocation");
    chunk.local_count = 1;
    chunk.scratch_high_water = 1;
    chunk.result_arity = 1;
    let callee_key = chunk.add_constant(Value::String(Arc::from("__multicast_allocation_probe")));
    chunk.emit_i32_const(rounds, 0);
    local(&mut chunk, Op::LOCAL_SET, 0);
    chunk.emit_op(Op::LOOP, 0);
    chunk.code.extend_from_slice(&[0, 0]);
    chunk.lines.resize(chunk.code.len(), 0);
    for _ in 0..32 {
        chunk.emit_op_u32(Op::GLOBAL_GET, callee_key, 0);
        chunk.emit_i32_const(0, 0); // Receiver slot, consumed by the free host function.
        chunk.emit_i32_const(7, 0);
        chunk.emit_op_u8_u8(Op::CALL_REF, 2, 1, 0);
        chunk.emit_op(Op::DROP, 0);
    }
    local(&mut chunk, Op::LOCAL_GET, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    local(&mut chunk, Op::LOCAL_TEE, 0);
    chunk.emit_op(Op::BR_IF, 0);
    chunk.code.push(0);
    chunk.lines.resize(chunk.code.len(), 0);
    chunk.emit_op(Op::END, 0);
    chunk.emit_i32_const(42, 0);
    chunk.emit_op(Op::RETURN, 0);
    let chunks = vec![chunk];
    let calls = CALLS.load(Ordering::Relaxed);
    let bytes = REQUESTED.load(Ordering::Relaxed);
    let result = vm.run(chunks);
    let traffic = (CALLS.load(Ordering::Relaxed) - calls, REQUESTED.load(Ordering::Relaxed) - bytes);
    assert!(matches!(result.unwrap(), Value::I32(42)));
    assert_eq!(observed.load(Ordering::Relaxed), rounds as usize * 32 * handler_count);
    traffic
}

#[test]
#[ignore = "complete multicast allocation measurement; run alone with one test thread"]
fn multicast_allocation_traffic_by_iteration_count() {
    for handlers in [0, 1, 2, 4] {
        measure_multicast(1, handlers);
        for rounds in [128, 1024, 8192] {
            let (calls, bytes) = measure_multicast(rounds, handlers);
            eprintln!("multicast handlers={handlers} invocations={} allocations={calls} requested_bytes={bytes}", rounds * 32);
        }
    }
}
