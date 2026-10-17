//! Allocation/reset timing controls; not a guest-program or Doom benchmark.
use std::time::Instant;
use vybe_runtime::{heap, value::Object};

#[test]
#[ignore = "allocation timing control; run alone with one test thread"]
fn allocation_and_reset_timing() {
    for tracked in [false, true] {
        for objects in [8192, 65536] {
            for sample in 0..5 {
                let (allocation_ms, total_ms) = std::thread::spawn(move || {
                    if tracked { heap::enable_tracking(); }
                    let snapshot = heap::snapshot();
                    let start = Instant::now();
                    for _ in 0..objects { drop(heap::alloc(Object::new())); }
                    let allocation_ms = start.elapsed().as_secs_f64() * 1000.0;
                    heap::restore(&snapshot);
                    let total_ms = start.elapsed().as_secs_f64() * 1000.0;
                    assert_eq!(heap::live_count(), 0);
                    (allocation_ms, total_ms)
                }).join().unwrap();
                eprintln!("heap tracked={tracked} objects={objects} sample={sample} allocation_ms={allocation_ms:.3} total_ms={total_ms:.3}");
            }
        }
    }
}
