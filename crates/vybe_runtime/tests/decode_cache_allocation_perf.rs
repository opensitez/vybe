//! Requested Rust allocation bytes, not RSS or uninstrumented timing.
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use vybe_runtime::{Chunk, Value, VM};
use vybe_runtime::opcode::Op;

struct CountingAllocator;
static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
static RECORDING: AtomicBool = AtomicBool::new(false);
static EVENT_COUNT: AtomicUsize = AtomicUsize::new(0);
static EVENT_SIZES: [AtomicUsize; 128] = [const { AtomicUsize::new(0) }; 128];
static TRACE_SIZE: AtomicUsize = AtomicUsize::new(0);

fn record_allocation(size: usize) {
    if RECORDING.load(Ordering::Relaxed) {
        let index = EVENT_COUNT.fetch_add(1, Ordering::Relaxed);
        if let Some(slot) = EVENT_SIZES.get(index) {
            slot.store(size, Ordering::Relaxed);
        }
        if size == TRACE_SIZE.load(Ordering::Relaxed) {
            // Diagnostic only: backtrace construction allocates. Disable event
            // recording to avoid recursion; this mode is not a byte benchmark.
            RECORDING.store(false, Ordering::Relaxed);
            eprintln!("allocation-owner size={size}\n{}", std::backtrace::Backtrace::force_capture());
            RECORDING.store(true, Ordering::Relaxed);
        }
    }
}

fn change(old: usize, new: usize) {
    let live = if new >= old {
        LIVE.fetch_add(new - old, Ordering::Relaxed) + new - old
    } else {
        LIVE.fetch_sub(old - new, Ordering::Relaxed) - (old - new)
    };
    PEAK.fetch_max(live, Ordering::Relaxed);
}

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() { change(0, layout.size()); record_allocation(layout.size()); }
        ptr
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc_zeroed(layout) };
        if !ptr.is_null() { change(0, layout.size()); record_allocation(layout.size()); }
        ptr
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        change(layout.size(), 0);
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let next = unsafe { System.realloc(ptr, layout, new_size) };
        if !next.is_null() { change(layout.size(), new_size); record_allocation(new_size); }
        next
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn workload(code_bytes: usize) -> Chunk {
    let mut chunk = Chunk::new("decode-allocation-baseline");
    chunk.result_arity = 1;
    chunk.code.reserve(code_bytes + 16);
    chunk.lines.reserve(code_bytes + 16);
    for _ in 0..code_bytes / 4 { chunk.emit_op(Op::NOP, 0); }
    chunk.emit_i32_const(42, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

#[test]
fn allocation_fixture_preserves_result() {
    assert_eq!(VM::new().run(vec![workload(64)]).unwrap(), Value::I32(42));
}

fn measure_call_free_preparation() -> (usize, usize) {
    assert_eq!(VM::new().run(vec![workload(64)]).unwrap(), Value::I32(42));
    let mut vm = VM::new();
    let chunk = workload(1 << 20);
    let code_bytes = chunk.code.len();
    let chunks = vec![chunk];
    let before = LIVE.load(Ordering::Relaxed);
    assert_eq!(vm.run(chunks).unwrap(), Value::I32(42));
    let retained = LIVE.load(Ordering::Relaxed).saturating_sub(before);
    (code_bytes, retained)
}

#[test]
#[ignore = "allocation budget; run alone with one test thread"]
fn call_free_preparation_avoids_native_call_table_storage() {
    let (code_bytes, retained) = measure_call_free_preparation();
    // The native-call table currently costs four bytes per code byte.
    // Removing that optional table leaves approximately nine; allow one
    // additional byte per code byte for fixed preparation overhead.
    assert!(retained <= code_bytes * 10,
        "call-free preparation retained {retained} bytes for {code_bytes} code bytes; optional call-table budget exceeded");
}

#[test]
#[ignore = "allocation budget; run alone with one test thread"]
fn call_free_preparation_avoids_result_consumer_table_storage() {
    let (code_bytes, retained) = measure_call_free_preparation();
    // Dropping the unused result-consumer row saves another two bytes per
    // code byte, leaving about seven plus fixed preparation overhead.
    assert!(retained <= code_bytes * 8,
        "call-free preparation retained {retained} bytes for {code_bytes} code bytes; optional result-table budget exceeded");
}

#[test]
#[ignore = "allocation baseline; run alone with one test thread, not a timing benchmark"]
fn decode_metadata_requested_allocation_baseline() {
    let trace_size = std::env::var("VYBE_TRACE_ALLOCATION_SIZE")
        .ok().and_then(|value| value.parse::<usize>().ok()).unwrap_or(0);
    TRACE_SIZE.store(trace_size, Ordering::Relaxed);
    if trace_size != 0 {
        eprintln!("diagnostic allocation traces enabled; byte totals are not comparable to the baseline");
    }
    assert_eq!(VM::new().run(vec![workload(64)]).unwrap(), Value::I32(42));
    for requested in [64 << 10, 256 << 10, 1 << 20] {
        let mut vm = VM::new();
        let chunk = workload(requested);
        let code_bytes = chunk.code.len();
        let input_code_capacity = chunk.code.capacity();
        let input_line_bytes = std::mem::size_of_val(chunk.lines.as_slice());
        let input_line_capacity_bytes = chunk.lines.capacity()
            * std::mem::size_of_val(&chunk.lines[0]);
        let chunks = vec![chunk];
        let before = LIVE.load(Ordering::Relaxed);
        PEAK.store(before, Ordering::Relaxed);
        EVENT_COUNT.store(0, Ordering::Relaxed);
        RECORDING.store(true, Ordering::Relaxed);
        assert_eq!(vm.run(chunks).unwrap(), Value::I32(42));
        RECORDING.store(false, Ordering::Relaxed);
        let retained = LIVE.load(Ordering::Relaxed).saturating_sub(before);
        let peak = PEAK.load(Ordering::Relaxed).saturating_sub(before);
        eprintln!("decode-allocation code_bytes={code_bytes} retained_extra={retained} peak_extra={peak} retained_per_code_byte={:.3} peak_per_code_byte={:.3}",
            retained as f64 / code_bytes as f64, peak as f64 / code_bytes as f64);
        let events = EVENT_COUNT.load(Ordering::Relaxed);
        let mut sizes = std::collections::BTreeMap::<usize, usize>::new();
        for slot in EVENT_SIZES.iter().take(events.min(EVENT_SIZES.len())) {
            *sizes.entry(slot.load(Ordering::Relaxed)).or_default() += 1;
        }
        eprintln!("run-allocation input_code_capacity={input_code_capacity} input_line_bytes={input_line_bytes} input_line_capacity_bytes={input_line_capacity_bytes} events={events} captured={} size_histogram={sizes:?}", events.min(EVENT_SIZES.len()));
        drop(vm);
    }
}
