//! Central heap allocator + registry enabling VM snapshot / hot-reset.
//!
//! Every `Object` backing store is allocated through [`alloc`]. When tracking is
//! enabled (a reset-capable host — the test harness or `--serve`), each allocation
//! is recorded as a `Weak` so a reset can find and reclaim the whole post-snapshot
//! generation — **including cycles**, which `Arc` reference-counting alone cannot
//! free. When tracking is off (a normal `vybex file.js` run — the default), `alloc`
//! is a plain `Arc::new` behind one predictable branch, so non-reset runs stay
//! effectively zero-cost.
//!
//! The registry is thread-local: the VM and its host functions run on one thread,
//! and each test-harness worker / server request owns its thread, so a thread-local
//! cleanly scopes "this VM's objects" without threading a handle through the ~230
//! allocation sites (many in host code with no VM in scope).

use std::cell::{Cell, RefCell};
use std::sync::{Arc, Mutex, Weak};

use crate::value::{Object, ObjectKind, Value};

thread_local! {
    static TRACKING: Cell<bool> = const { Cell::new(false) };
    static HEAP: RefCell<Heap> = RefCell::new(Heap::default());
}

#[derive(Default)]
struct Heap {
    /// Every tracked allocation retains its position until reset. Dead handles
    /// are replaced in place so their Arc backing can be freed without shifting
    /// allocation-index snapshot boundaries.
    objects: Vec<Weak<Mutex<Object>>>,
    dead_sweep_cursor: usize,
}

const DEAD_SWEEP_INTERVAL: usize = 64;

#[inline]
fn retained_capacity_limit<T>(len: usize) -> usize {
    len.saturating_mul(4).max(65536 / std::mem::size_of::<T>().max(1))
}

impl Heap {
    fn sweep_dead_handles(&mut self) {
        let len = self.objects.len();
        if len == 0 {
            return;
        }
        // Recent short-lived objects should not wait for a full registry pass.
        for weak in self.objects.iter_mut().rev().take(DEAD_SWEEP_INTERVAL) {
            if weak.strong_count() == 0 {
                *weak = Weak::new();
            }
        }
        // The recent pass already covered this tail. Revisit only older
        // allocations, at most once per pass even for small registries.
        // Positions remain stable for allocation-index snapshot boundaries.
        let older_len = len.saturating_sub(DEAD_SWEEP_INTERVAL);
        for _ in 0..older_len.min(DEAD_SWEEP_INTERVAL * 2) {
            if self.dead_sweep_cursor >= older_len {
                self.dead_sweep_cursor = 0;
            }
            let weak = &mut self.objects[self.dead_sweep_cursor];
            if weak.strong_count() == 0 {
                *weak = Weak::new();
            }
            self.dead_sweep_cursor += 1;
        }
    }
}

/// THE object allocation choke point — all `Value::Object` backing store is created
/// here so a reset can account for every object (and break its cycles).
#[inline]
pub fn alloc(obj: Object) -> Arc<Mutex<Object>> {
    let arc = Arc::new(Mutex::new(obj));
    if TRACKING.with(Cell::get) {
        HEAP.with(|h| {
            let mut h = h.borrow_mut();
            if h.objects.len() % DEAD_SWEEP_INTERVAL == 0 {
                h.sweep_dead_handles();
            }
            h.objects.push(Arc::downgrade(&arc));
        });
    }
    arc
}

/// Enable allocation tracking. Call once, before boot, in a reset-capable host.
pub fn enable_tracking() {
    TRACKING.with(|tracking| tracking.set(true));
}

pub fn is_tracking() -> bool {
    TRACKING.with(Cell::get)
}

/// Number of still-live tracked objects (registry entries whose `Weak` still
/// upgrades). A leak/diagnostic aid: after a `restore`, this returns to the
/// post-snapshot baseline iff the whole script generation — cycles included —
/// was reclaimed. O(registry length); test/serve use only.
pub fn live_count() -> usize {
    HEAP.with(|h| {
        h.borrow()
            .objects
            .iter()
            .filter(|w| w.strong_count() > 0)
            .count()
    })
}

/// Saved contents of one baseline object, so a reset can undo script mutations to
/// it (e.g. `Object.prototype.x = 1`, pushing to a prelude array). Values are
/// shallow-cloned — they keep pointing at the SAME baseline objects, so restoring
/// re-wires the baseline graph and drops any script object a baseline object came
/// to reference.
struct ObjectContents {
    properties: indexmap::IndexMap<String, Value>,
    fields: Vec<Value>,
    type_id: usize,
    descriptor: Option<Arc<std::sync::Mutex<Object>>>,
    kind: ObjectKind,
}

impl ObjectContents {
    fn capture(o: &Object) -> Self {
        ObjectContents {
            properties: o.properties.clone(),
            fields: o.fields.clone(),
            type_id: o.type_id,
            descriptor: o.descriptor.clone(),
            kind: o.kind.clone(),
        }
    }
    fn apply(&self, o: &mut Object) {
        o.properties.clone_from(&self.properties);
        if o.fields.capacity() <= retained_capacity_limit::<Value>(self.fields.len()) {
            o.fields.clone_from(&self.fields);
        } else {
            o.fields = self.fields.clone();
        }
        o.type_id = self.type_id;
        o.descriptor.clone_from(&self.descriptor);
        match (&mut o.kind, &self.kind) {
            (ObjectKind::Array(current), ObjectKind::Array(saved)) => {
                // Reuse ordinary working sets, but release temporary growth
                // rather than pinning a large request's buffer across resets.
                let retained_limit = retained_capacity_limit::<Value>(saved.len());
                if current.capacity() <= retained_limit {
                    current.clone_from(saved);
                } else {
                    *current = saved.clone();
                }
            }
            _ => o.kind = self.kind.clone(),
        }
    }
}

/// A restorable baseline: the generation boundary plus the saved contents of every
/// object that existed at snapshot time (the prelude/boot heap, incl. the shared
/// prototypes if they were allocated through `alloc` before the snapshot).
pub struct HeapSnapshot {
    boundary: usize,
    baseline: Vec<(Weak<Mutex<Object>>, ObjectContents)>,
}

/// Capture the current heap as a restorable baseline. Call right after boot.
pub fn snapshot() -> HeapSnapshot {
    HEAP.with(|h| {
        let h = h.borrow();
        let boundary = h.objects.len();
        let baseline = h
            .objects
            .iter()
            .filter_map(|w| {
                let arc = w.upgrade()?;
                let o = arc.lock().ok()?;
                Some((w.clone(), ObjectContents::capture(&o)))
            })
            .collect();
        HeapSnapshot { boundary, baseline }
    })
}

/// Free every object allocated after `boundary` by CLEARING its contents (severing
/// cyclic references so refcounts collapse to 0), then drop the registry tail.
/// This is what reclaims the cyclic script garbage that `Arc` can't.
pub fn collect_since(boundary: usize) {
    HEAP.with(|h| {
        let mut h = h.borrow_mut();
        for weak in h.objects.iter().skip(boundary) {
            if let Some(arc) = weak.upgrade() {
                if let Ok(mut o) = arc.lock() {
                    o.properties.clear();
                    o.fields.clear();
                    o.type_id = 0;
                    o.descriptor = None;
                    o.kind = ObjectKind::Ordinary;
                }
            }
        }
        h.objects.truncate(boundary);
        let retained_limit = retained_capacity_limit::<Weak<Mutex<Object>>>(h.objects.len());
        if h.objects.capacity() > retained_limit {
            if h.objects.is_empty() {
                h.objects = Vec::new();
            } else {
                h.objects.shrink_to(retained_limit);
            }
        }
        h.dead_sweep_cursor = 0;
    });
}

/// Restore the heap to a snapshot: free the whole post-snapshot generation (script
/// objects + their cycles), then undo any mutations to baseline objects. Baseline
/// object *identity* is preserved (contents are overwritten in place), so live
/// references from anywhere stay valid.
pub fn restore(snap: &HeapSnapshot) {
    collect_since(snap.boundary);
    for (weak, contents) in &snap.baseline {
        if let Some(arc) = weak.upgrade() {
            if let Ok(mut o) = arc.lock() {
                contents.apply(&mut o);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dead_sweep_preserves_positions_and_visits_older_allocations() {
        for len in [0, 1, 63, 64, 65, 127, 128, 193, 1024] {
            let objects: Vec<_> = (0..len)
                .map(|_| Arc::new(Mutex::new(Object::new())))
                .collect();
            let mut heap = Heap {
                objects: objects.iter().map(Arc::downgrade).collect(),
                dead_sweep_cursor: len,
            };
            heap.sweep_dead_handles();
            assert_eq!(heap.objects.len(), len);
            for (weak, object) in heap.objects.iter().zip(&objects) {
                assert!(Weak::ptr_eq(weak, &Arc::downgrade(object)));
            }
            drop(objects);
            // A full older-region rotation must clear every dead handle;
            // the recent tail is cleared on the very first pass.
            for _ in 0..1 + len.div_ceil(DEAD_SWEEP_INTERVAL * 2) {
                heap.sweep_dead_handles();
            }
            assert_eq!(heap.objects.len(), len);
            assert!(heap.objects.iter().all(|weak| Weak::ptr_eq(weak, &Weak::new())));
        }
    }

    #[test]
    fn restore_releases_oversized_property_capacity_and_request_references() {
        std::thread::spawn(|| {
            enable_tracking();
            let baseline = alloc(Object::new());
            {
                let mut object = baseline.lock().unwrap();
                object.properties.insert("first".into(), Value::I32(1));
                object.properties.insert("second".into(), Value::I32(2));
            }
            let identity = Arc::as_ptr(&baseline);
            let saved = snapshot();
            let request_object = alloc(Object::new());
            let request_weak = Arc::downgrade(&request_object);
            let limit = retained_capacity_limit::<(String, Value)>(2);
            {
                let mut object = baseline.lock().unwrap();
                object.properties.reserve(limit + 1);
                assert!(object.properties.capacity() > limit);
                object.properties.shift_remove("first");
                object.properties.insert("request".into(), Value::Object(request_object.clone()));
            }
            drop(request_object);
            restore(&saved);
            assert_eq!(Arc::as_ptr(&baseline), identity);
            let object = baseline.lock().unwrap();
            assert!(object.properties.capacity() <= limit);
            assert_eq!(object.properties.keys().map(String::as_str).collect::<Vec<_>>(),
                ["first", "second"]);
            assert_eq!(object.properties["first"], Value::I32(1));
            assert_eq!(object.properties["second"], Value::I32(2));
            assert!(request_weak.upgrade().is_none());
        }).join().unwrap();
    }

    #[test]
    fn dead_handle_sweeps_preserve_nested_snapshots_and_live_identity() {
        std::thread::spawn(|| {
            enable_tracking();
            let base = alloc(Object::new());
            base.lock().unwrap().properties.insert("value".into(), Value::I32(7));
            let identity = Arc::as_ptr(&base);
            let first = snapshot();
            for _ in 0..128 { drop(alloc(Object::new())); }
            assert_eq!(snapshot().boundary, first.boundary + 128);
            let kept = alloc(Object::new());
            kept.lock().unwrap().properties.insert("value".into(), Value::I32(9));
            let second = snapshot();
            let post = alloc(Object::new());
            post.lock().unwrap().properties.insert("post".into(), Value::I32(1));
            let cycle = alloc(Object::new());
            cycle.lock().unwrap().properties.insert("self".into(), Value::Object(cycle.clone()));
            let weak_cycle = Arc::downgrade(&cycle);
            drop(cycle);
            base.lock().unwrap().properties.insert("value".into(), Value::I32(42));
            kept.lock().unwrap().properties.insert("value".into(), Value::I32(11));
            for _ in 0..1024 { drop(alloc(Object::new())); }
            assert!(weak_cycle.upgrade().is_some(), "sweeps must not collect live cycles");
            restore(&second);
            assert_eq!(Arc::as_ptr(&base), identity);
            assert_eq!(base.lock().unwrap().properties["value"], Value::I32(7));
            assert_eq!(kept.lock().unwrap().properties["value"], Value::I32(9));
            assert!(post.lock().unwrap().properties.is_empty());
            assert!(weak_cycle.upgrade().is_none());
            restore(&first);
            assert_eq!(Arc::as_ptr(&base), identity);
            assert_eq!(base.lock().unwrap().properties["value"], Value::I32(7));
            assert!(kept.lock().unwrap().properties.is_empty());
        }).join().unwrap();
    }

    #[test]
    fn dead_handle_sweeps_revisit_older_live_registrations() {
        std::thread::spawn(|| {
            enable_tracking();
            let older = alloc(Object::new());
            let weak = Arc::downgrade(&older);
            for _ in 0..256 { drop(alloc(Object::new())); }
            assert!(weak.upgrade().is_some());
            drop(older);
            for _ in 0..2048 { drop(alloc(Object::new())); }
            assert!(weak.upgrade().is_none());
            HEAP.with(|heap| {
                assert!(Weak::ptr_eq(&heap.borrow().objects[0], &Weak::new()));
            });
        }).join().unwrap();
    }

    #[test]
    fn restore_reuses_fields_and_preserves_saved_contents() {
        std::thread::spawn(|| {
            enable_tracking();
            let mut contents = Object::new_typed_with_fields(42, 128);
            contents.fields[0] = Value::I32(7);
            contents.fields[1] = Value::String(Arc::from("saved"));
            contents.properties.insert("a".into(), Value::I32(1));
            contents.properties.insert("b".into(), Value::I32(2));
            let object = alloc(contents);
            let snapshot = snapshot();
            let fields_address = object.lock().unwrap().fields.as_ptr() as usize;
            for _ in 0..3 {
                {
                    let mut object = object.lock().unwrap();
                    object.fields.truncate(1);
                    object.fields[0] = Value::Null;
                    object.properties.shift_remove("a");
                    object.properties.insert("extra".into(), Value::Undefined);
                    object.type_id = 0;
                    object.kind = ObjectKind::Array(vec![Value::I32(99)]);
                }
                restore(&snapshot);
                let object = object.lock().unwrap();
                assert_eq!(object.fields.as_ptr() as usize, fields_address);
                assert_eq!(object.fields.len(), 128);
                assert!(matches!(object.fields[0], Value::I32(7)));
                assert_eq!(object.fields[1].as_str(), "saved");
                assert!(object.fields[2..].iter().all(|v| matches!(v, Value::Null)));
                assert_eq!(object.type_id, 42);
                assert!(matches!(object.kind, ObjectKind::Ordinary));
                assert_eq!(
                    object.properties.keys().map(String::as_str).collect::<Vec<_>>(),
                    vec!["a", "b"]
                );
            }
        })
        .join()
        .unwrap();
    }

    #[test]
    fn untracked_allocations_skip_registry_borrow() {
        // A fresh thread starts with tracking disabled. Holding the registry
        // borrow proves the allocation fast path does not touch that registry.
        std::thread::spawn(|| {
            assert!(!is_tracking());
            HEAP.with(|heap| {
                let registry = heap.borrow_mut();
                let object = alloc(Object::new());
                assert_eq!(Arc::strong_count(&object), 1);
                assert!(registry.objects.is_empty());
            });
            assert_eq!(snapshot().boundary, 0);
        })
        .join()
        .unwrap();
    }

    #[test]
    fn enabling_tracking_preserves_allocation_boundary() {
        std::thread::spawn(|| {
            let untracked = alloc(Object::new());
            let weak_untracked = Arc::downgrade(&untracked);
            enable_tracking();
            assert!(is_tracking());
            let baseline = alloc(Object::new());
            let snap = snapshot();
            assert_eq!(snap.boundary, 1);
            assert_eq!(live_count(), 1);
            let generation = alloc(Object::new());
            generation
                .lock()
                .unwrap()
                .properties
                .insert("self".into(), Value::Object(generation.clone()));
            let weak_generation = Arc::downgrade(&generation);
            drop(generation);
            restore(&snap);
            assert!(weak_generation.upgrade().is_none());
            assert!(weak_untracked.upgrade().is_some());
            assert_eq!(live_count(), 1);
            assert_eq!(Arc::strong_count(&baseline), 1);
        })
        .join()
        .unwrap();
    }

    #[test]
    fn collect_frees_cyclic_garbage() {
        enable_tracking();
        let snap = snapshot();
        // Build a reference cycle a↔b, then drop the only strong roots.
        let a = alloc(Object::new());
        let b = alloc(Object::new());
        a.lock()
            .unwrap()
            .properties
            .insert("b".into(), Value::Object(b.clone()));
        b.lock()
            .unwrap()
            .properties
            .insert("a".into(), Value::Object(a.clone()));
        let (wa, wb) = (Arc::downgrade(&a), Arc::downgrade(&b));
        drop(a);
        drop(b);
        // Pure refcounting CANNOT reclaim this — the cycle keeps both alive.
        assert!(
            wa.upgrade().is_some() && wb.upgrade().is_some(),
            "cycle must survive under plain Arc refcounting"
        );
        // Reset frees the whole post-snapshot generation, cycle included.
        restore(&snap);
        assert!(
            wa.upgrade().is_none() && wb.upgrade().is_none(),
            "collect_since must break the cycle and free both objects"
        );
    }

    #[test]
    fn restore_undoes_baseline_mutation() {
        enable_tracking();
        let base = alloc(Object::new());
        let snap = snapshot(); // baseline: `base` with no properties
        base.lock()
            .unwrap()
            .properties
            .insert("x".into(), Value::I32(1)); // script mutates a baseline object
        assert!(base.lock().unwrap().properties.contains_key("x"));
        restore(&snap);
        assert!(
            !base.lock().unwrap().properties.contains_key("x"),
            "restore must roll a baseline object back to its snapshot contents"
        );
    }

    #[test]
    fn acyclic_garbage_frees_and_baseline_survives() {
        enable_tracking();
        let base = alloc(Object::new());
        let wbase = Arc::downgrade(&base);
        let snap = snapshot();
        let tmp = alloc(Object::new());
        let wtmp = Arc::downgrade(&tmp);
        drop(tmp);
        restore(&snap);
        assert!(wtmp.upgrade().is_none(), "script object must be freed");
        assert!(
            wbase.upgrade().is_some(),
            "baseline object must survive reset"
        );
    }
}

#[cfg(test)]
mod restore_array_reuse_tests {
    use super::*;

    #[test]
    fn restore_reuses_array_storage_and_preserves_references_and_bits() {
        let child = alloc(Object::new());
        let saved = Object::new_array(vec![
            Value::Object(child.clone()), Value::F32(f32::from_bits(0x7fc01234)),
            Value::F64(f64::from_bits(0x8000000000000000)), Value::TypedNull(42),
        ]);
        let contents = ObjectContents::capture(&saved);
        let mut current = Object::new_array(Vec::with_capacity(128));
        let ObjectKind::Array(values) = &current.kind else { unreachable!() };
        let pointer = values.as_ptr();
        let extra = alloc(Object::new());
        let weak = Arc::downgrade(&extra);
        if let ObjectKind::Array(values) = &mut current.kind {
            values.resize(64, Value::Object(extra.clone()));
        }
        drop(extra);
        contents.apply(&mut current);
        assert!(weak.upgrade().is_none(), "replaced and truncated lanes must release references");
        let ObjectKind::Array(values) = &mut current.kind else { unreachable!() };
        assert_eq!(values.as_ptr(), pointer);
        assert_eq!(values.len(), 4);
        let Value::Object(restored) = &values[0] else { unreachable!() };
        assert!(Arc::ptr_eq(restored, &child));
        let Value::F32(v) = values[1] else { unreachable!() };
        assert_eq!(v.to_bits(), 0x7fc01234);
        let Value::F64(v) = values[2] else { unreachable!() };
        assert_eq!(v.to_bits(), 0x8000000000000000);
        assert!(matches!(values[3], Value::TypedNull(42)));
        values.truncate(1);
        contents.apply(&mut current);
        let ObjectKind::Array(values) = &current.kind else { unreachable!() };
        assert_eq!(values.as_ptr(), pointer);
        assert_eq!(values.len(), 4);
        assert_eq!(current.properties["length"], Value::F64(4.0));
        current.kind = ObjectKind::Ordinary;
        contents.apply(&mut current);
        assert!(matches!(current.kind, ObjectKind::Array(ref values) if values.len() == 4));
        ObjectContents::capture(&Object::new()).apply(&mut current);
        assert!(matches!(current.kind, ObjectKind::Ordinary));
    }

    #[test]
    fn restore_releases_oversized_array_storage_and_removed_references() {
        for len in [0usize, 8, 2048] {
            let saved = Object::new_array(vec![Value::I32(7); len]);
            let contents = ObjectContents::capture(&saved);
            let limit = len.saturating_mul(4).max(65536 / std::mem::size_of::<Value>());
            for capacity in [limit, limit + 1] {
                let extra = alloc(Object::new());
                let weak = Arc::downgrade(&extra);
                let mut values = Vec::with_capacity(capacity);
                values.resize(3, Value::Object(extra.clone()));
                let pointer = values.as_ptr();
                let mut current = Object::new_array(values);
                drop(extra);
                contents.apply(&mut current);
                assert!(weak.upgrade().is_none());
                let ObjectKind::Array(values) = &current.kind else { unreachable!() };
                assert_eq!(values.len(), len);
                assert!(values.iter().all(|value| matches!(value, Value::I32(7))));
                if capacity == limit {
                    assert_eq!(values.as_ptr(), pointer);
                } else {
                    assert!(values.capacity() <= limit);
                    if len == 0 { assert_eq!(values.capacity(), 0); }
                }
                assert_eq!(current.properties["length"], Value::F64(len as f64));
            }
        }
    }

    fn previous(contents: &ObjectContents, current: &mut Object) {
        current.properties.clone_from(&contents.properties);
        current.fields.clone_from(&contents.fields);
        current.type_id = contents.type_id;
        current.descriptor.clone_from(&contents.descriptor);
        current.kind = contents.kind.clone();
    }

    #[test]
    #[ignore = "array restore allocation comparison, excludes VM reset orchestration"]
    fn array_restore_allocation_benchmark() {
        use std::{hint::black_box, time::Instant};
        for references in [false, true] {
            for len in [16usize, 1024, 32768] {
                let value = if references { Value::Object(alloc(Object::new())) } else { Value::I32(7) };
                let saved = Object::new_array(vec![value; len]);
                let contents = ObjectContents::capture(&saved);
                let rounds = (262144 / len).max(8);
                let mut samples = [Vec::new(), Vec::new()];
                for pair in 0..9 {
                    for step in 0..2 {
                        let index = (pair + step) % 2;
                        let mut current = Object::new_array(vec![Value::Null; len]);
                        contents.apply(&mut current);
                        let clock = Instant::now();
                        for _ in 0..rounds {
                            if index == 0 { contents.apply(black_box(&mut current)); }
                            else { previous(&contents, black_box(&mut current)); }
                            black_box(&current);
                        }
                        if pair >= 2 { samples[index].push(clock.elapsed().as_secs_f64()); }
                    }
                }
                for sample in &mut samples { sample.sort_by(f64::total_cmp); }
                eprintln!("array-restore references={references} len={len} current_ms={:.3} previous_ms={:.3} previous_over_current={:.3}",
                    samples[0][3] * 1000.0, samples[1][3] * 1000.0, samples[1][3] / samples[0][3]);
            }
        }
    }
}

#[cfg(test)]
mod registry_retention_tests {
    use super::*;

    #[test]
    fn restore_releases_oversized_registry_without_changing_snapshot_boundaries() {
        for baseline in [0usize, 16, 4096] {
            std::thread::spawn(move || {
                enable_tracking();
                let objects: Vec<_> = (0..baseline).map(|_| alloc(Object::new())).collect();
                if let Some(first) = objects.first() {
                    first.lock().unwrap().properties.insert("value".into(), Value::I32(7));
                }
                let snap = snapshot();
                assert_eq!(snap.boundary, baseline);
                let limit = retained_capacity_limit::<Weak<Mutex<Object>>>(baseline);
                HEAP.with(|heap| heap.borrow_mut().objects.reserve(limit + 1));
                let cycle = alloc(Object::new());
                cycle.lock().unwrap().properties.insert("self".into(), Value::Object(cycle.clone()));
                let weak_cycle = Arc::downgrade(&cycle);
                drop(cycle);
                if let Some(first) = objects.first() {
                    first.lock().unwrap().properties.insert("value".into(), Value::I32(99));
                }
                restore(&snap);
                assert!(weak_cycle.upgrade().is_none());
                assert_eq!(live_count(), baseline);
                HEAP.with(|heap| {
                    let heap = heap.borrow();
                    assert_eq!(heap.objects.len(), baseline);
                    assert!(heap.objects.capacity() <= limit);
                    if baseline == 0 { assert_eq!(heap.objects.capacity(), 0); }
                    for (weak, object) in heap.objects.iter().zip(&objects) {
                        assert!(Arc::ptr_eq(&weak.upgrade().unwrap(), object));
                    }
                });
                if let Some(first) = objects.first() {
                    assert_eq!(first.lock().unwrap().properties["value"], Value::I32(7));
                }
                // Reusing the same saved boundary after a capacity change
                // still collects only the next generation.
                let next = alloc(Object::new());
                next.lock().unwrap().properties.insert("generation".into(), Value::I32(1));
                restore(&snap);
                assert!(next.lock().unwrap().properties.is_empty());
                assert_eq!(snapshot().boundary, baseline);
            }).join().unwrap();
        }
    }

    #[test]
    fn ordinary_registry_working_set_retains_capacity() {
        std::thread::spawn(|| {
            enable_tracking();
            let baseline = alloc(Object::new());
            let snap = snapshot();
            for _ in 0..128 { drop(alloc(Object::new())); }
            let before = HEAP.with(|heap| heap.borrow().objects.capacity());
            restore(&snap);
            HEAP.with(|heap| {
                let heap = heap.borrow();
                assert_eq!(heap.objects.capacity(), before);
                assert_eq!(heap.objects.len(), snap.boundary);
                assert!(Arc::ptr_eq(&heap.objects[0].upgrade().unwrap(), &baseline));
            });
        }).join().unwrap();
    }
}

#[cfg(test)]
mod field_retention_tests {
    use super::*;

    #[test]
    fn restore_releases_oversized_fields_and_removed_references() {
        for saved_len in [0usize, 16, 4096] {
            std::thread::spawn(move || {
                enable_tracking();
                let baseline = alloc(Object::new());
                baseline.lock().unwrap().fields = (0..saved_len)
                    .map(|i| Value::I32(i as i32)).collect();
                let saved = snapshot();
                let limit = retained_capacity_limit::<Value>(saved_len);
                let temporary = alloc(Object::new());
                let weak = Arc::downgrade(&temporary);
                {
                    let mut object = baseline.lock().unwrap();
                    object.fields.reserve(limit + 1);
                    object.fields.push(Value::Object(temporary.clone()));
                }
                drop(temporary);
                restore(&saved);
                assert!(weak.upgrade().is_none());
                let object = baseline.lock().unwrap();
                assert_eq!(object.fields.len(), saved_len);
                assert!(object.fields.capacity() <= limit);
                if saved_len == 0 {
                    assert_eq!(object.fields.capacity(), 0);
                }
                for (i, value) in object.fields.iter().enumerate() {
                    assert_eq!(*value, Value::I32(i as i32));
                }
            }).join().unwrap();
        }
    }

    #[test]
    fn ordinary_field_capacity_and_saved_reference_identity_are_reused() {
        std::thread::spawn(|| {
            enable_tracking();
            let child = alloc(Object::new());
            let baseline = alloc(Object::new());
            {
                let mut object = baseline.lock().unwrap();
                object.fields = Vec::with_capacity(128);
                object.fields.push(Value::Object(child.clone()));
                object.fields.push(Value::F64(-0.0));
            }
            let saved = snapshot();
            let (pointer, capacity) = {
                let object = baseline.lock().unwrap();
                (object.fields.as_ptr(), object.fields.capacity())
            };
            baseline.lock().unwrap().fields[1] = Value::I32(99);
            restore(&saved);
            let object = baseline.lock().unwrap();
            assert_eq!(object.fields.as_ptr(), pointer);
            assert_eq!(object.fields.capacity(), capacity);
            match &object.fields[0] {
                Value::Object(restored) => assert!(Arc::ptr_eq(restored, &child)),
                _ => panic!("saved field reference changed kind"),
            }
            match object.fields[1] {
                Value::F64(value) => assert_eq!(value.to_bits(), (-0.0f64).to_bits()),
                _ => panic!("saved scalar field changed kind"),
            }
        }).join().unwrap();
    }
}
