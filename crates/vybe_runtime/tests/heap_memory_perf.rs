use std::cell::RefCell;
use std::hint::black_box;
use std::sync::{Arc, Mutex, Weak};
use std::time::Instant;

use vybe_runtime::{heap, value::{Object, Value}};

const PAIRS: usize = 21;

fn paired_measurement(label: &str, mut run: impl FnMut(bool) -> u128) {
    run(false);
    run(true);
    let mut previous = Vec::with_capacity(PAIRS);
    let mut current = Vec::with_capacity(PAIRS);
    let mut ratios = Vec::with_capacity(PAIRS);
    for round in 0..PAIRS {
        let (before, after) = if round % 2 == 0 {
            (run(false), run(true))
        } else {
            let after = run(true);
            (run(false), after)
        };
        previous.push(before);
        current.push(after);
        ratios.push(before as f64 / after as f64);
    }
    previous.sort_unstable();
    current.sort_unstable();
    ratios.sort_by(f64::total_cmp);
    println!(
        "{label}: previous_p50_ns={}, current_p50_ns={}, previous_p95_ns={}, \
         current_p95_ns={}, paired_ratio_p10={:.4}, paired_ratio_p50={:.4}, \
         paired_ratio_p90={:.4}",
        previous[10], current[10], previous[19], current[19],
        ratios[2], ratios[10], ratios[18],
    );
}

// Exact registry access shape used before the disabled-tracking fast path.
// Keep the reference local to this benchmark, not in production allocation.
#[derive(Default)]
struct PreviousHeap {
    tracking: bool,
    objects: Vec<Weak<Mutex<Object>>>,
}

thread_local! {
    static PREVIOUS: RefCell<PreviousHeap> = RefCell::new(PreviousHeap::default());
}

#[inline]
fn previous_alloc(object: Object) -> Arc<Mutex<Object>> {
    let arc = Arc::new(Mutex::new(object));
    PREVIOUS.with(|heap| {
        let mut heap = heap.borrow_mut();
        if heap.tracking {
            heap.objects.push(Arc::downgrade(&arc));
        }
    });
    arc
}

#[test]
#[ignore = "alternating loaded-machine allocation measurement; no timing assertions"]
fn allocation_registry_loaded_benchmark() {
    // Each benchmark owns a fresh TLS heap, independent of test-worker reuse.
    std::thread::spawn(|| {
        assert!(!heap::is_tracking());
        for tracked in [false, true] {
            if tracked {
                heap::enable_tracking();
                PREVIOUS.with(|heap| heap.borrow_mut().tracking = true);
            }
            let count = if tracked { 100_000 } else { 500_000 };
            paired_measurement(&format!("allocation tracking={tracked} count={count}"), |current| {
                let start = Instant::now();
                if current {
                    for _ in 0..count {
                        drop(black_box(heap::alloc(black_box(Object::new()))));
                    }
                } else {
                    for _ in 0..count {
                        drop(black_box(previous_alloc(black_box(Object::new()))));
                    }
                }
                let elapsed = start.elapsed().as_nanos();
                // Registry cleanup is outside the allocation timing. Both
                // registries retain their vector capacity between samples.
                if current {
                    heap::collect_since(0);
                    assert_eq!(heap::live_count(), 0);
                } else {
                    PREVIOUS.with(|heap| heap.borrow_mut().objects.clear());
                }
                elapsed
            });
        }
    })
    .join()
    .unwrap();
}

fn previous_restore(object: &Weak<Mutex<Object>>, saved: &Object) {
    // The current benchmark has exactly one baseline object. Both paths run
    // identical generation collection before upgrading/locking their object.
    heap::collect_since(1);
    if let Some(object) = object.upgrade() {
        if let Ok(mut object) = object.lock() {
            object.properties = saved.properties.clone();
            object.fields = saved.fields.clone();
            object.type_id = saved.type_id;
            object.kind = saved.kind.clone();
        }
    }
}

#[test]
#[ignore = "alternating loaded-machine restore measurement; no timing assertions"]
fn snapshot_restore_loaded_benchmark() {
    std::thread::spawn(|| {
        assert!(!heap::is_tracking());
        let mut saved = Object::new_typed_with_fields(42, 512);
        for (index, field) in saved.fields.iter_mut().enumerate() {
            *field = Value::I32(index as i32);
        }
        for index in 0..32 {
            saved.properties.insert(format!("field{index}"), Value::I32(index));
        }
        // The reference object is intentionally outside the tracked heap;
        // only the current object's saved state belongs to the snapshot.
        let previous = Arc::new(Mutex::new(saved.clone()));
        let weak = Arc::downgrade(&previous);
        heap::enable_tracking();
        let current = heap::alloc(saved.clone());
        assert_eq!(heap::live_count(), 1);
        let snapshot = heap::snapshot();
        paired_measurement("restore fields=512 properties=32 count=20000", |current| {
            let start = Instant::now();
            for _ in 0..20_000 {
                if current {
                    heap::restore(black_box(&snapshot));
                } else {
                    previous_restore(black_box(&weak), black_box(&saved));
                }
            }
            start.elapsed().as_nanos()
        });
        let previous = previous.lock().unwrap();
        let current = current.lock().unwrap();
        assert_eq!(previous.fields, current.fields);
        assert_eq!(previous.properties, current.properties);
        assert_eq!(previous.type_id, current.type_id);
        assert_eq!(
            previous.properties.keys().collect::<Vec<_>>(),
            current.properties.keys().collect::<Vec<_>>(),
        );
    })
    .join()
    .unwrap();
}
