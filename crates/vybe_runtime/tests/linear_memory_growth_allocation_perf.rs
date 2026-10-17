//! Allocation evidence for linear-memory growth, not VM/Doom timing.
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use vybe_runtime::shared_memory::SharedMemory;

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

#[test]
#[ignore = "allocation measurement; run this integration test alone with one test thread"]
fn repeated_growth_is_amortized_and_zero_growth_does_not_allocate() {
    const PAGE: usize = 65536;
    for shared in [false, true] {
        for rounds in [1, 16, 128] {
            let memory = SharedMemory::new(PAGE);
            let alias = shared.then(|| memory.clone());
            memory.store_u8(0, 0x91).unwrap();
            CALLS.store(0, Ordering::Relaxed);
            BYTES.store(0, Ordering::Relaxed);
            ENABLED.store(true, Ordering::Relaxed);
            for old_pages in 1..=rounds {
                assert_eq!(memory.grow(1), old_pages);
            }
            ENABLED.store(false, Ordering::Relaxed);
            let calls = CALLS.load(Ordering::Relaxed);
            let bytes = BYTES.load(Ordering::Relaxed);
            assert!(calls <= 1 + rounds.ilog2() as usize,
                "one-page growth must retain amortized capacity: {calls} allocations");
            assert_eq!(memory.load_u8(0).unwrap(), 0x91);
            memory.with_buffer(|buffer| {
                assert_eq!(buffer.len(), (rounds + 1) * PAGE);
                assert!(buffer[PAGE..].iter().all(|&byte| byte == 0));
            });
            if let Some(alias) = alias {
                assert_eq!(alias.len(), (rounds + 1) * PAGE);
            }
            eprintln!("growth shared={shared} rounds={rounds} allocations={calls} requested_bytes={bytes}");
            CALLS.store(0, Ordering::Relaxed);
            ENABLED.store(true, Ordering::Relaxed);
            for _ in 0..65536 {
                assert_eq!(memory.grow(0), rounds + 1);
            }
            ENABLED.store(false, Ordering::Relaxed);
            assert_eq!(CALLS.load(Ordering::Relaxed), 0);
        }
    }
}
