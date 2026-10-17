//! Closure upvalue machinery.
//!
//! `capture_upvalue` promotes a stack slot to an `Upvalue` (shared with
//! any closure that captures the same slot). `close_upvalues` copies the
//! stack value into the Upvalue's `closed` slot when the enclosing frame
//! returns, so closures can outlive their defining frame.

use crate::value::{Upvalue, UpvalueLocation, Value};
use crate::vm::VM;
use std::sync::{Arc, Mutex};

impl VM {
    pub(crate) fn capture_upvalue(&mut self, stack_idx: usize) -> Arc<Mutex<Upvalue>> {
        for uv in &self.open_upvalues {
            if let UpvalueLocation::Open(idx) = uv.lock().unwrap().location {
                if idx == stack_idx { return uv.clone(); }
            }
        }
        let uv = Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Open(stack_idx),
        }));
        self.open_upvalues.push(uv.clone());
        uv
    }

    #[cfg(test)]
    #[inline]
    fn capture_upvalue_candidate(&mut self, stack_idx: usize) -> Arc<Mutex<Upvalue>> {
        // Preserve the cheapest, common first-entry hit before any hints.
        if let Some(first) = self.open_upvalues.first() {
            if let UpvalueLocation::Open(idx) = first.lock().unwrap().location {
                if idx == stack_idx { return first.clone(); }
            }
        }
        self.capture_upvalue_tail(stack_idx)
    }

    // Keep scan/allocation state out of the common first-entry call frame.
    #[cfg(test)]
    #[inline(never)]
    fn capture_upvalue_tail(&mut self, stack_idx: usize) -> Arc<Mutex<Upvalue>> {
        let mut entries = self.open_upvalues.iter();
        let _ = entries.next();
        // A shared/weak-observed tail suggests ordinary lock scanning is
        // cheaper. Counts are only a performance hint, never an exclusivity
        // proof: every lock bypass below still requires Arc::get_mut.
        let try_exclusive = self.open_upvalues.last().is_some_and(|uv|
            Arc::strong_count(uv) == 1 && Arc::weak_count(uv) == 0);
        if try_exclusive {
            // Release the immutable iterator only on the exclusive path.
            // The shared path below continues the original iterator directly.
            drop(entries);
            let (_, remaining) = self.open_upvalues.split_first_mut().unwrap();
            for uv in remaining.iter_mut() {
                if Arc::strong_count(uv) == 1 {
                    if let Some(upvalue) = Arc::get_mut(uv) {
                        if matches!(upvalue.get_mut().unwrap().location,
                            UpvalueLocation::Open(idx) if idx == stack_idx) {
                            return uv.clone();
                        }
                        continue;
                    }
                }
                if let UpvalueLocation::Open(idx) = uv.lock().unwrap().location {
                    if idx == stack_idx { return uv.clone(); }
                }
            }
        } else {
            for uv in entries {
                if let UpvalueLocation::Open(idx) = uv.lock().unwrap().location {
                    if idx == stack_idx { return uv.clone(); }
                }
            }
        }
        let uv = Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Open(stack_idx),
        }));
        self.open_upvalues.push(uv.clone());
        uv
    }

    pub(crate) fn close_upvalues(&mut self, from: usize) {
        if self.open_upvalues.is_empty() {
            return;
        }
        let stack = &self.stack;
        self.open_upvalues.retain_mut(|uv| {
            // Most live captures are shared with a closure: keep them on the
            // existing lock-and-close path without probing Arc's weak lock.
            // The count is only a negative hint; get_mut proves exclusivity,
            // including exclusion of weak handles and concurrent upgrades.
            if Arc::strong_count(uv) == 1 {
                if let Some(upvalue) = Arc::get_mut(uv) {
                    if let UpvalueLocation::Open(idx) = upvalue.get_mut().unwrap().location {
                        // No observer can see this capture close. Retire it
                        // directly instead of cloning a value just to drop it.
                        return idx < from;
                    }
                    return true;
                }
            }
            let mut upvalue = uv.lock().unwrap();
            if let UpvalueLocation::Open(idx) = upvalue.location {
                if idx >= from {
                    // Lazy-locals convention: an unwritten captured slot may
                    // lie beyond the materialized stack.
                    upvalue.location = UpvalueLocation::Closed(
                        stack.get(idx).cloned().unwrap_or(Value::Null)
                    );
                    return false;
                }
            }
            true
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn original_capture(vm: &mut VM, stack_idx: usize) -> Arc<Mutex<Upvalue>> {
        for uv in &vm.open_upvalues {
            if let UpvalueLocation::Open(idx) = uv.lock().unwrap().location {
                if idx == stack_idx { return uv.clone(); }
            }
        }
        let uv = Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Open(stack_idx),
        }));
        vm.open_upvalues.push(uv.clone());
        uv
    }

    #[test]
    fn exclusive_capture_candidate_preserves_shared_and_weak_identity() {
        for weak_observer in [false, true] {
            let mut vm = VM::new();
            drop(vm.capture_upvalue(1));
            let capture = vm.capture_upvalue(9);
            let weak = weak_observer.then(|| Arc::downgrade(&capture));
            drop(capture);
            let candidate = vm.capture_upvalue_candidate(9);
            assert_eq!(vm.open_upvalues.len(), 2);
            assert!(Arc::ptr_eq(&candidate, &vm.open_upvalues[1]));
            if let Some(weak) = weak {
                assert!(Arc::ptr_eq(&candidate, &weak.upgrade().unwrap()));
            }
            assert!(Arc::ptr_eq(&candidate, &vm.capture_upvalue_candidate(9)));
        }
    }

    #[test]
    #[ignore = "rejected capture candidate timing; production retains lock-only scan"]
    fn exclusive_capture_scan_benchmark() {
        use std::{hint::black_box, time::Instant};
        for count in [8usize, 64, 1024] {
            for shared in [false, true] {
                let mut vm = VM::new();
                let observers: Vec<_> = (0..count).map(|slot| vm.capture_upvalue(slot)).collect();
                if !shared { drop(observers); }
                for slot in [0, count - 1] {
                    let rounds = if count == 1024 { 512 } else { 8192 };
                    let measure = |vm: &mut VM, candidate: bool| {
                        let t = Instant::now();
                        for _ in 0..rounds {
                            let uv = if candidate { vm.capture_upvalue_candidate(black_box(slot)) }
                                else { original_capture(vm, black_box(slot)) };
                            black_box(uv);
                        }
                        t.elapsed().as_secs_f64()
                    };
                    let mut ratios = Vec::new();
                    for pair in 0..21 {
                        let (old, new) = if pair % 2 == 0 {
                            (measure(&mut vm, false), measure(&mut vm, true))
                        } else {
                            let new = measure(&mut vm, true);
                            (measure(&mut vm, false), new)
                        };
                        assert_eq!(vm.open_upvalues.len(), count);
                        ratios.push(new / old);
                    }
                    ratios.sort_by(f64::total_cmp);
                    println!("capture entries={count} shared={shared} slot={slot} candidate_over_previous={:.4}", ratios[10]);
                }
            }
        }
    }

    #[test]
    fn capture_mixed_owners_keeps_registry_order_and_first_duplicate_identity() {
        let mut vm = VM::new();
        drop(vm.capture_upvalue(0));
        let shared = vm.capture_upvalue(4);
        let observed = vm.capture_upvalue(8);
        let weak = Arc::downgrade(&observed);
        drop(observed);
        vm.open_upvalues.push(Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Closed(Value::I32(99)),
        })));
        vm.open_upvalues.push(Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Open(4),
        })));
        drop(vm.capture_upvalue(12));
        let identities: Vec<_> = vm.open_upvalues.iter().map(Arc::as_ptr).collect();
        for (slot, index) in [(0, 0), (4, 1), (8, 2), (12, 5)] {
            let capture = vm.capture_upvalue(slot);
            assert_eq!(Arc::as_ptr(&capture), identities[index]);
            assert_eq!(vm.open_upvalues.len(), identities.len());
        }
        assert!(Arc::ptr_eq(&shared, &vm.capture_upvalue(4)));
        assert!(Arc::ptr_eq(&weak.upgrade().unwrap(), &vm.capture_upvalue(8)));
        for (entry, identity) in vm.open_upvalues.iter().zip(identities) {
            assert_eq!(Arc::as_ptr(entry), identity);
        }
    }

    #[test]
    fn exclusive_capture_preserves_poisoned_lock_failure() {
        use std::panic::{catch_unwind, AssertUnwindSafe};
        let poisoned = Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Open(5),
        }));
        let _ = catch_unwind(AssertUnwindSafe(|| {
            let _guard = poisoned.lock().unwrap();
            panic!("test poison");
        }));
        let mut vm = VM::new();
        drop(vm.capture_upvalue(0));
        vm.open_upvalues.push(poisoned);
        vm.open_upvalues.push(Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Open(12),
        })));
        assert!(catch_unwind(AssertUnwindSafe(|| vm.capture_upvalue(99))).is_err());
        assert_eq!(vm.open_upvalues.len(), 3);
    }

    #[test]
    fn close_keeps_outer_and_closed_entries_while_retiring_unobserved_captures() {
        let mut vm = VM::new();
        let text: Arc<str> = "captured reference".into();
        vm.stack = vec![Value::I32(17), Value::String(text.clone())];
        drop(vm.capture_upvalue(0));
        drop(vm.capture_upvalue(1));
        let shared = vm.capture_upvalue(2);
        let closed = Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Closed(Value::I32(99)),
        }));
        vm.open_upvalues.push(closed);
        assert_eq!(Arc::strong_count(&vm.open_upvalues[0]), 1);
        assert_eq!(Arc::strong_count(&vm.open_upvalues[1]), 1);
        vm.close_upvalues(1);
        assert_eq!(vm.open_upvalues.len(), 2);
        assert!(matches!(vm.open_upvalues[0].lock().unwrap().location,
            UpvalueLocation::Open(0)));
        assert!(matches!(vm.open_upvalues[1].lock().unwrap().location,
            UpvalueLocation::Closed(Value::I32(99))));
        assert!(matches!(shared.lock().unwrap().location,
            UpvalueLocation::Closed(Value::Null)));
        assert_eq!(Arc::strong_count(&text), 2);
    }

    #[test]
    fn weak_capture_observer_sees_closed_reference_identity() {
        let mut vm = VM::new();
        let text: Arc<str> = "shared through weak capture".into();
        vm.stack = vec![Value::String(text.clone())];
        let capture = vm.capture_upvalue(0);
        let weak = Arc::downgrade(&capture);
        drop(capture);
        let barrier = Arc::new(std::sync::Barrier::new(2));
        let worker_barrier = barrier.clone();
        let worker_text = text.clone();
        let worker = std::thread::spawn(move || {
            let observer = weak.upgrade().unwrap();
            worker_barrier.wait();
            worker_barrier.wait();
            let observer = observer.lock().unwrap();
            let UpvalueLocation::Closed(Value::String(value)) = &observer.location else {
                panic!("capture was not closed");
            };
            assert!(Arc::ptr_eq(value, &worker_text));
        });
        barrier.wait();
        vm.close_upvalues(0);
        assert!(vm.open_upvalues.is_empty());
        barrier.wait();
        worker.join().unwrap();
        assert_eq!(Arc::strong_count(&text), 2);
    }

    #[test]
    fn repeated_capture_shares_identity_and_releases_closed_registry_entries() {
        let mut vm = VM::new();
        vm.stack = vec![Value::I32(17), Value::I32(29)];
        let outer = vm.capture_upvalue(0);
        let inner = vm.capture_upvalue(1);
        let repeated = vm.capture_upvalue(1);
        assert!(Arc::ptr_eq(&inner, &repeated));
        assert_eq!(vm.open_upvalues.len(), 2);
        drop(repeated);
        vm.close_upvalues(1);
        assert_eq!(vm.open_upvalues.len(), 1);
        assert_eq!(Arc::strong_count(&inner), 1);
        assert!(matches!(inner.lock().unwrap().location,
            UpvalueLocation::Closed(Value::I32(29))));
        assert!(Arc::ptr_eq(&outer, &vm.capture_upvalue(0)));
        vm.close_upvalues(0);
        assert!(vm.open_upvalues.is_empty());
        assert_eq!(Arc::strong_count(&outer), 1);
    }

    #[test]
    fn unordered_and_closed_registry_entries_keep_capture_identity() {
        let mut vm = VM::new();
        let slots = [9, 1, 12, 4, 7];
        let entries: Vec<_> = slots.iter().map(|&slot| {
            Arc::new(Mutex::new(Upvalue {
                location: UpvalueLocation::Open(slot),
            }))
        }).collect();
        vm.open_upvalues = entries.clone();
        for (&slot, expected) in slots.iter().zip(&entries) {
            assert!(Arc::ptr_eq(&vm.capture_upvalue(slot), expected));
        }
        entries[2].lock().unwrap().location = UpvalueLocation::Closed(Value::I32(81));
        for (&slot, expected) in slots.iter().zip(&entries) {
            if slot != 12 {
                assert!(Arc::ptr_eq(&vm.capture_upvalue(slot), expected));
            }
        }
        let replacement = vm.capture_upvalue(12);
        assert!(!Arc::ptr_eq(&replacement, &entries[2]));
        assert!(Arc::ptr_eq(&replacement, &vm.capture_upvalue(12)));
    }

    #[test]
    fn out_of_order_capture_and_partial_close_preserve_unique_slots() {
        let mut vm = VM::new();
        let slots = [31, 2, 19, 0, 8, 27, 5, 14];
        let entries: Vec<_> = slots.iter().map(|&slot| vm.capture_upvalue(slot)).collect();
        for (&slot, expected) in slots.iter().zip(&entries) {
            assert!(Arc::ptr_eq(&vm.capture_upvalue(slot), expected));
        }
        vm.close_upvalues(15);
        for (&slot, expected) in slots.iter().zip(&entries) {
            if slot < 15 {
                assert!(Arc::ptr_eq(&vm.capture_upvalue(slot), expected));
            } else {
                assert!(matches!(expected.lock().unwrap().location,
                    UpvalueLocation::Closed(Value::Null)));
            }
        }
        assert_eq!(vm.open_upvalues.len(), 5);
    }

    #[test]
    #[ignore = "reference-capture scan control; not a whole-VM or Doom speedup"]
    fn loaded_reference_capture_scan_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;

        let mut vm = VM::new();
        let references: Vec<_> = (0..1024)
            .map(|slot| vm.capture_upvalue(slot))
            .collect();
        for slot in [0, 1023] {
            let mut samples = Vec::new();
            for round in 0..19 {
                let start = Instant::now();
                for _ in 0..2000 {
                    let captured = vm.capture_upvalue(black_box(slot));
                    assert!(Arc::ptr_eq(&captured, &references[slot]));
                    black_box(captured);
                }
                let elapsed = start.elapsed().as_secs_f64();
                if round >= 2 { samples.push(elapsed); }
            }
            samples.sort_by(f64::total_cmp);
            assert_eq!(vm.open_upvalues.len(), references.len());
            assert_eq!(Arc::strong_count(&references[slot]), 2);
            eprintln!("reference-capture open=1024 slot={slot} captures=2000 p50_ms={:.3} p95_ms={:.3}",
                samples[8] * 1000.0, samples[16] * 1000.0);
        }
    }

    #[test]
    fn closes_many_upvalues_in_one_pass_and_keeps_outer_slots() {
        let mut vm = VM::new();
        vm.stack = (0..20_000).map(Value::I32).collect();
        let upvalues: Vec<_> = (0..20_000)
            .map(|idx| {
                Arc::new(Mutex::new(Upvalue {
                    location: UpvalueLocation::Open(idx),
                }))
            })
            .collect();
        vm.open_upvalues = upvalues.clone();

        vm.close_upvalues(10_000);

        assert_eq!(vm.open_upvalues.len(), 10_000);
        assert!(matches!(upvalues[9_999].lock().unwrap().location, UpvalueLocation::Open(9_999)));
        assert!(matches!(upvalues[10_000].lock().unwrap().location, UpvalueLocation::Closed(Value::I32(10_000))));
        assert!(matches!(upvalues[19_999].lock().unwrap().location, UpvalueLocation::Closed(Value::I32(19_999))));
    }

    #[test]
    #[ignore = "paired new reference allocations; excludes VM setup and destruction"]
    fn paired_reference_allocation_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;

        fn original_capture(vm: &mut VM, slot: usize) -> Arc<Mutex<Upvalue>> {
            for entry in &vm.open_upvalues {
                if let UpvalueLocation::Open(index) = entry.lock().unwrap().location {
                    if index == slot { return entry.clone(); }
                }
            }
            let entry = Arc::new(Mutex::new(Upvalue {
                location: UpvalueLocation::Open(slot),
            }));
            vm.open_upvalues.push(entry.clone());
            entry
        }

        for descending in [false, true] {
            let mut samples = [Vec::new(), Vec::new()];
            for round in 0..19 {
                let order = if round % 2 == 0 { [0, 1] } else { [1, 0] };
                for path in order {
                    let mut machines: Vec<_> = (0..20).map(|_| VM::new()).collect();
                    let start = Instant::now();
                    for vm in &mut machines {
                        for index in 0..512 {
                            let slot = if descending { 511 - index } else { index };
                            let entry = if path == 0 {
                                original_capture(vm, black_box(slot))
                            } else {
                                vm.capture_upvalue_candidate(black_box(slot))
                            };
                            black_box(entry);
                        }
                    }
                    let elapsed = start.elapsed().as_secs_f64();
                    for vm in &mut machines {
                        assert_eq!(vm.open_upvalues.len(), 512);
                        for entry in &vm.open_upvalues {
                            assert_eq!(Arc::strong_count(entry), 1);
                        }
                        for slot in [0, 255, 511] {
                            let entry = vm.capture_upvalue(slot);
                            assert!(matches!(entry.lock().unwrap().location,
                                UpvalueLocation::Open(index) if index == slot));
                            assert_eq!(vm.open_upvalues.len(), 512);
                        }
                    }
                    if round >= 2 { samples[path].push(elapsed); }
                }
            }
            for sample in &mut samples { sample.sort_by(f64::total_cmp); }
            eprintln!("reference-allocation descending={descending} count=10240 original_ms={:.3} candidate_ms={:.3} original_over_candidate={:.3}",
                samples[0][8] * 1000.0, samples[1][8] * 1000.0,
                samples[0][8] / samples[1][8]);
        }
    }
}
