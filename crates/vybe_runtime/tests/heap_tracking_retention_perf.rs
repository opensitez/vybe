//! Retained allocation evidence for reset-capable tracked heaps, not Doom timing.
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicIsize, Ordering};
use vybe_runtime::{heap, value::Object};

struct LiveAllocator;
static LIVE_BYTES: AtomicIsize = AtomicIsize::new(0);

unsafe impl GlobalAlloc for LiveAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            LIVE_BYTES.fetch_add(layout.size() as isize, Ordering::Relaxed);
        }
        pointer
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if !pointer.is_null() {
            LIVE_BYTES.fetch_add(layout.size() as isize, Ordering::Relaxed);
        }
        pointer
    }
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        let pointer = unsafe { System.realloc(pointer, layout, size) };
        if !pointer.is_null() {
            LIVE_BYTES.fetch_add(size as isize - layout.size() as isize, Ordering::Relaxed);
        }
        pointer
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        LIVE_BYTES.fetch_sub(layout.size() as isize, Ordering::Relaxed);
        unsafe { System.dealloc(pointer, layout) };
    }
}

#[global_allocator]
static ALLOCATOR: LiveAllocator = LiveAllocator;

#[test]
#[ignore = "retained-memory measurement; run alone with one test thread"]
fn dropped_tracked_object_backing_before_and_after_reset() {
    for objects in [4096, 32768] {
        let (before_reset, after_reset) = std::thread::spawn(move || {
            heap::enable_tracking();
            let snapshot = heap::snapshot();
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            for _ in 0..objects {
                drop(heap::alloc(Object::new()));
            }
            assert_eq!(heap::live_count(), 0);
            let before = LIVE_BYTES.load(Ordering::Relaxed) - baseline;
            heap::restore(&snapshot);
            assert_eq!(heap::live_count(), 0);
            let after = LIVE_BYTES.load(Ordering::Relaxed) - baseline;
            assert!(after <= before);
            (before, after)
        }).join().unwrap();
        eprintln!("tracked dropped_objects={objects} retained_before_reset={before_reset} retained_after_reset={after_reset}");
    }
}
