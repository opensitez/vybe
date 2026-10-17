//! Run alone with --test-threads=1: allocation counting is process-wide.
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_runtime::typedef::{TypeDef, TypeRegistry};

struct CountingAllocator;
static CALLS: AtomicUsize = AtomicUsize::new(0);
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() { CALLS.fetch_add(1, Ordering::Relaxed); }
        ptr
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc_zeroed(layout) };
        if !ptr.is_null() { CALLS.fetch_add(1, Ordering::Relaxed); }
        ptr
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        let result = unsafe { System.realloc(ptr, layout, size) };
        if !result.is_null() { CALLS.fetch_add(1, Ordering::Relaxed); }
        result
    }
}
#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

#[test]
fn lookup_preserves_ascii_unicode_and_duplicate_binding() {
    let mut registry = TypeRegistry::new();
    for name in ["System.Buffer", "Stra\u{df}e", "\u{039f}\u{03a3}", "\u{0130}"] {
        let id = registry.register(TypeDef::new(name));
        assert_eq!(registry.get_id(name), Some(id));
        assert_eq!(registry.get_id(&name.to_lowercase()), Some(id));
    }
    let replacement = registry.register(TypeDef::new("system.buffer"));
    assert_eq!(registry.get_id("System.Buffer"), Some(replacement));
    assert_eq!(registry.get_id("system.buffer"), Some(replacement));
    assert_eq!(registry.get_id("missing"), None);
    assert_eq!(registry.get_id(""), None);
}

#[test]
#[ignore = "process-wide allocation counter; run alone with one test thread"]
fn lowercase_ascii_type_lookup_avoids_allocations() {
    use std::hint::black_box;
    let mut registry = TypeRegistry::new();
    let id = registry.register(TypeDef::new("System.Buffer"));
    let before = CALLS.load(Ordering::Relaxed);
    for _ in 0..100_000 {
        assert_eq!(registry.get_id(black_box("system.buffer")), Some(id));
        assert_eq!(registry.get_id(black_box("missing.type")), None);
    }
    let optimized = CALLS.load(Ordering::Relaxed) - before;
    let before = CALLS.load(Ordering::Relaxed);
    for _ in 0..100_000 {
        black_box(black_box("system.buffer").to_lowercase());
        black_box(black_box("missing.type").to_lowercase());
    }
    let previous_normalization = CALLS.load(Ordering::Relaxed) - before;
    eprintln!("200000 lookups: current allocations={optimized}; previous normalization allocations={previous_normalization}");
    assert_eq!(optimized, 0);
    assert_eq!(previous_normalization, 200_000);
}
