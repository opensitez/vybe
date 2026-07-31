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

use std::cell::RefCell;
use std::sync::{Arc, Mutex, Weak};

use crate::value::{Object, ObjectKind, Value};

thread_local! {
    static HEAP: RefCell<Heap> = RefCell::new(Heap::default());
}

#[derive(Default)]
struct Heap {
    tracking: bool,
    /// Every object allocated while tracking, in allocation order. Entries whose
    /// object was already freed by refcount linger as dangling `Weak`s until the
    /// next `collect_since` truncates the tail.
    objects: Vec<Weak<Mutex<Object>>>,
}

/// THE object allocation choke point — all `Value::Object` backing store is created
/// here so a reset can account for every object (and break its cycles).
#[inline]
pub fn alloc(obj: Object) -> Arc<Mutex<Object>> {
    let arc = Arc::new(Mutex::new(obj));
    HEAP.with(|h| {
        let mut h = h.borrow_mut();
        if h.tracking {
            h.objects.push(Arc::downgrade(&arc));
        }
    });
    arc
}

/// Enable allocation tracking. Call once, before boot, in a reset-capable host.
pub fn enable_tracking() {
    HEAP.with(|h| h.borrow_mut().tracking = true);
}

pub fn is_tracking() -> bool {
    HEAP.with(|h| h.borrow().tracking)
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
    kind: ObjectKind,
}

impl ObjectContents {
    fn capture(o: &Object) -> Self {
        ObjectContents {
            properties: o.properties.clone(),
            fields: o.fields.clone(),
            type_id: o.type_id,
            kind: o.kind.clone(),
        }
    }
    fn apply(&self, o: &mut Object) {
        o.properties = self.properties.clone();
        o.fields = self.fields.clone();
        o.type_id = self.type_id;
        o.kind = self.kind.clone();
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
                    o.kind = ObjectKind::Ordinary;
                }
            }
        }
        h.objects.truncate(boundary);
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
