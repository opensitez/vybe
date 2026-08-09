//! VM-owned storage for host resources — the plugin side of [`crate::heap`].
//!
//! A plugin that holds anything on the running program's behalf — open
//! descriptors, response bodies, key material, cached constructors, event
//! listeners — asks for it here instead of declaring a `static`. The VM then
//! owns that memory the same way it owns the object heap, and drops it on
//! [`VM::reset_to`](crate::VM::reset_to) without the plugin taking part.
//!
//! That is the whole point. A `static` inside a platform crate is unreachable
//! from the VM by construction, so every one of them needed a hand-written
//! `reset()` that something had to remember to call. Each forgotten one is a
//! tenant boundary a reused VM silently crosses: the next program in the
//! worker inherits the previous program's open files, its HTTP bodies, its
//! keys. Storage that the VM owns cannot be forgotten, because nobody has to
//! remember it.
//!
//! ```ignore
//! // Before: process-global, invisible to the VM, needs a reset() somebody calls.
//! fn registry() -> &'static Mutex<Registry> {
//!     static R: OnceLock<Mutex<Registry>> = OnceLock::new();
//!     R.get_or_init(|| Mutex::new(Registry::default()))
//! }
//!
//! // After: VM-owned. Call sites are unchanged; the reset() is deleted.
//! fn registry() -> &'static Mutex<Registry> {
//!     vybe_runtime::resources::get::<Registry>()
//! }
//! ```
//!
//! **Store a type you own.** The key is `TypeId`, so a bare `HashMap<i32,
//! String>` is the SAME resource for every plugin that asks for one. Wrap it in
//! a named struct (deriving `Deref` by hand if the call sites want the inner
//! type) — the newtype is what makes the resource yours.
//!
//! **Tenant state only.** What a plugin builds during `init`/`finalize` is boot
//! state — host functions, prototypes, registry indices — and must survive a
//! reset, so it stays in a `static`. This store is for what accumulates while a
//! *program* runs.
//!
//! **Ids do not belong here.** A handle allocator (`next_id`) must keep
//! counting across a reset — reissuing an id a prior tenant still holds is how
//! a stale handle silently starts addressing another tenant's file. Keep the
//! counter in a plain `AtomicU64` static outside the resource, as
//! `platforms/node/src/fs.rs` does, so clearing tenant data can never rewind
//! it. Then every resource clears to `Default` and there is no per-type reset
//! rule left to get wrong.
//!
//! Thread-local, for the reason [`crate::heap`] gives: the VM and its host
//! functions run on one thread, and each worker / request owns its thread, so a
//! thread-local scopes "this VM's resources" without threading a handle through
//! call sites that have no VM in scope. A resource whose handle is moved to
//! another thread must NOT live here — it would get a different cell per thread
//! and split in silence.

use std::any::{TypeId, type_name};
use std::cell::RefCell;
use std::collections::HashMap;
use std::sync::Mutex;

thread_local! {
    static STORE: RefCell<HashMap<TypeId, Entry>> = RefCell::new(HashMap::new());
}

/// One stored resource, type-erased. `addr` points at a leaked
/// `Mutex<T>`; `clear` is the monomorphized function that knows how to put a
/// fresh `T` back into it.
struct Entry {
    addr: usize,
    clear: fn(usize),
    /// Type name, for [`live`] — diagnostics and leak-hunting only.
    name: &'static str,
}

/// This thread's `Mutex<T>`, created empty on first use.
///
/// The reference is `'static` because the cell is leaked, which is what lets a
/// call site keep the `&'static Mutex<T>` shape a `static` had — no churn, and
/// no lifetime threaded through host code. What leaks is the *cell*: one empty
/// `T` per type per thread, boot-sized and bounded by the number of resource
/// types. The tenant's data inside it is dropped on every reset.
pub fn get<T: Default + Send + 'static>() -> &'static Mutex<T> {
    STORE.with(|s| {
        let mut s = s.borrow_mut();
        let entry = s.entry(TypeId::of::<T>()).or_insert_with(|| {
            let cell: &'static Mutex<T> = Box::leak(Box::new(Mutex::new(T::default())));
            Entry {
                addr: cell as *const Mutex<T> as usize,
                clear: clear_one::<T>,
                name: type_name::<T>(),
            }
        });
        // SAFETY: `addr` was produced by `Box::leak` for exactly this `T` (the
        // map is keyed by `TypeId::of::<T>()`, so no other type can reach this
        // entry), the allocation is never freed, and the store is thread-local
        // so the cell is only ever handed to the thread that created it.
        unsafe { &*(entry.addr as *const Mutex<T>) }
    })
}

/// Put a fresh `T` in the cell, dropping the tenant's data.
///
/// Recovers a poisoned lock rather than skipping: a panic inside a host
/// function must not be able to make a resource permanently unclearable — that
/// would leave one tenant's data readable by the next, which is precisely the
/// boundary this module exists to hold.
fn clear_one<T: Default + Send + 'static>(addr: usize) {
    // SAFETY: as in `get` — `addr` is the leaked `Mutex<T>` registered under
    // `TypeId::of::<T>()`, and `clear_one::<T>` is only ever stored beside it.
    let cell = unsafe { &*(addr as *const Mutex<T>) };
    let mut guard = cell.lock().unwrap_or_else(|e| e.into_inner());
    *guard = T::default();
}

/// Drop every host resource on this thread. Called by
/// [`VM::reset_to`](crate::VM::reset_to) — after the plugins' own `reset`, so a
/// plugin closing sockets or connections still sees its table.
pub fn clear_all() {
    // Copy the list out and release the borrow BEFORE clearing: dropping a
    // resource runs its `Drop`, and a `Drop` that reaches back into the store
    // would panic on the outstanding borrow. Nothing does that today; it should
    // not become a landmine for the first resource that does.
    let entries: Vec<(usize, fn(usize))> =
        STORE.with(|s| s.borrow().values().map(|e| (e.addr, e.clear)).collect());
    for (addr, clear) in entries {
        clear(addr);
    }
}

/// Drop one resource without waiting for a reset — the program closed its last
/// socket, finished with the filesystem, dropped its keys. Nothing else is
/// touched.
pub fn release<T: Default + Send + 'static>() {
    // Borrow released before the clear — see [`clear_all`].
    let entry = STORE.with(|s| {
        s.borrow()
            .get(&TypeId::of::<T>())
            .map(|e| (e.addr, e.clear))
    });
    if let Some((addr, clear)) = entry {
        clear(addr);
    }
}

/// The type names of every resource this thread has handed out, for
/// diagnostics: what is a reused VM actually carrying?
pub fn live() -> Vec<&'static str> {
    STORE.with(|s| {
        let mut names: Vec<&'static str> = s.borrow().values().map(|e| e.name).collect();
        names.sort_unstable();
        names
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap as Map;

    #[derive(Default)]
    struct Table(Map<u32, String>);

    #[test]
    fn same_cell_every_time() {
        get::<Table>().lock().unwrap().0.insert(1, "a".into());
        assert_eq!(
            get::<Table>().lock().unwrap().0.get(&1).map(String::as_str),
            Some("a"),
            "a second get must reach the same resource, not a fresh one"
        );
    }

    #[test]
    fn clear_all_drops_tenant_data() {
        #[derive(Default)]
        struct Fds(Vec<i32>);
        get::<Fds>().lock().unwrap().0.push(7);
        clear_all();
        assert!(
            get::<Fds>().lock().unwrap().0.is_empty(),
            "reset must not hand the next tenant the previous one's data"
        );
    }

    #[test]
    fn release_touches_only_its_own_type() {
        #[derive(Default)]
        struct A(u32);
        #[derive(Default)]
        struct B(u32);
        get::<A>().lock().unwrap().0 = 1;
        get::<B>().lock().unwrap().0 = 2;
        release::<A>();
        assert_eq!(get::<A>().lock().unwrap().0, 0, "A must be released");
        assert_eq!(get::<B>().lock().unwrap().0, 2, "B must be untouched");
    }

    #[test]
    fn a_poisoned_resource_still_clears() {
        #[derive(Default)]
        struct Keys(Vec<u8>);
        get::<Keys>().lock().unwrap().0.push(0xAB);
        // A host function panics while holding the lock.
        let _ = std::panic::catch_unwind(|| {
            let _guard = get::<Keys>().lock().unwrap();
            panic!("host fn blew up mid-write");
        });
        assert!(get::<Keys>().is_poisoned());
        clear_all();
        let survived = get::<Keys>()
            .lock()
            .unwrap_or_else(|e| e.into_inner())
            .0
            .clone();
        assert!(
            survived.is_empty(),
            "a panic must not leave key material readable by the next tenant"
        );
    }
}
