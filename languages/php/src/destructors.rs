//! PHP object destruction: `__destruct` runs once, when the last program
//! reference to an object disappears, and at the end of the request for every
//! object still alive.
//!
//! Every object whose class has a callable destructor is registered when it is
//! constructed. Releases are reported by the emitted code — a rebound variable,
//! a displaced array element, a local dropped at return — and each registered
//! object the released value still contains is destructed once no program root
//! reaches it. The guest reads `__php_destructibles` before reporting a release
//! that is not itself a destructible object: it is an i31 while any registered
//! object is alive and null otherwise, so a program without destructible
//! objects never calls in.
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::HostContext;
use vybe_runtime::{Framework, Value};

/// The guest-visible flag global; see the module documentation.
pub const LIVE_FLAG: &str = "__php_destructibles";

struct Entry {
    object: Value,
    destructor: Value,
}

#[derive(Default)]
struct Registry {
    /// Creation order is destruction order at the end of the request.
    order: Vec<usize>,
    live: HashMap<usize, Entry>,
}

thread_local! {
    static REGISTRY: RefCell<Registry> = RefCell::new(Registry::default());
}

fn identity(object: &Arc<std::sync::Mutex<Object>>) -> usize {
    Arc::as_ptr(object) as usize
}

fn live_count() -> usize {
    REGISTRY.with(|registry| registry.borrow().live.len())
}

/// Drop the request's registry. The VM's reset rolls back the flag global.
pub fn reset() {
    REGISTRY.with(|registry| *registry.borrow_mut() = Registry::default());
}

fn publish_flag(ctx: &mut HostContext) {
    let flag = if live_count() > 0 { Value::I32(1) } else { Value::Null };
    ctx.set_global(LIVE_FLAG, flag);
}

/// Unregister `id` and hand back its entry; the caller runs the destructor.
fn take(id: usize) -> Option<Entry> {
    REGISTRY.with(|registry| {
        let mut registry = registry.borrow_mut();
        let entry = registry.live.remove(&id)?;
        registry.order.retain(|other| *other != id);
        Some(entry)
    })
}

fn is_registered(id: usize) -> bool {
    REGISTRY.with(|registry| registry.borrow().live.contains_key(&id))
}

/// Run one destructor. The entry is already unregistered, so a destructor that
/// releases values or re-enters this module cannot run it twice.
fn destruct(ctx: &mut HostContext, entry: Entry) {
    publish_flag(ctx);
    ctx.invoke_with_receiver(&entry.destructor, entry.object, &[]);
}

/// Destruct an object php frees now, then whatever it alone owned.
fn destruct_and_free(ctx: &mut HostContext, entry: Entry) {
    let object = entry.object.clone();
    destruct(ctx, entry);
    release(ctx, &object);
}

/// Identities of every object reachable from `roots`.
fn reachable_from(roots: &[Value]) -> HashSet<usize> {
    let mut seen = HashSet::new();
    let mut pending: Vec<Value> = roots.to_vec();
    while let Some(value) = pending.pop() {
        let Value::Object(object) = value else { continue };
        if !seen.insert(identity(&object)) {
            continue;
        }
        let Ok(data) = object.lock() else { continue };
        match &data.kind {
            ObjectKind::Array(items) => pending.extend(items.iter().cloned()),
            ObjectKind::Map(entries) => pending.extend(entries.values().cloned()),
            ObjectKind::Function(_) | ObjectKind::HostFunction(_) => {}
            _ => {
                pending.extend(data.fields.iter().cloned());
                pending.extend(data.properties.values().cloned());
            }
        }
    }
    seen
}

/// The registered objects `value` owns, in the order php frees them: an
/// object before its properties, an array's elements in order.
fn owned_candidates(value: &Value) -> Vec<usize> {
    let mut found = Vec::new();
    let mut seen = HashSet::new();
    let mut pending = vec![value.clone()];
    while let Some(value) = pending.pop() {
        let Value::Object(object) = value else { continue };
        let id = identity(&object);
        if !seen.insert(id) {
            continue;
        }
        if is_registered(id) {
            found.push(id);
        }
        let Ok(data) = object.lock() else { continue };
        let mut children: Vec<Value> = Vec::new();
        match &data.kind {
            ObjectKind::Array(items) => children.extend(items.iter().cloned()),
            ObjectKind::Map(entries) => children.extend(entries.values().cloned()),
            ObjectKind::Function(_) | ObjectKind::HostFunction(_) => {}
            _ => {
                children.extend(data.fields.iter().cloned());
                children.extend(data.properties.values().cloned());
            }
        }
        // Depth-first, first child first.
        pending.extend(children.into_iter().rev());
    }
    found
}

/// Destruct every registered object `released` owns that no program root
/// still reaches.
pub fn release(ctx: &mut HostContext, released: &Value) {
    if live_count() == 0 {
        return;
    }
    for id in owned_candidates(released) {
        let unreachable = REGISTRY.with(|registry| {
            registry.borrow().live.get(&id).map(|entry| entry.object.clone())
        });
        let Some(object) = unreachable else { continue };
        if ctx.is_reachable_from_program(&object) {
            continue;
        }
        if let Some(entry) = take(id) {
            destruct(ctx, entry);
        }
    }
    publish_flag(ctx);
}

/// The request-global variable store, if the program created one.
fn global_variable_store(ctx: &mut HostContext) -> Option<Arc<std::sync::Mutex<Object>>> {
    let Value::Object(global_this) = call_host(ctx, "ecma:globalThis", "get", &[])? else {
        return None;
    };
    let store = global_this.lock().ok()?.properties.get("__vybe_php_globals").cloned();
    match store {
        Some(Value::Object(store)) => Some(store),
        _ => None,
    }
}

/// How many references each object has from `roots`, counting each edge once
/// per distinct container.
fn reference_counts(roots: &[Value]) -> HashMap<usize, usize> {
    let mut counts: HashMap<usize, usize> = HashMap::new();
    let mut seen = HashSet::new();
    let mut pending: Vec<Value> = roots.to_vec();
    while let Some(value) = pending.pop() {
        let Value::Object(object) = value else { continue };
        let id = identity(&object);
        *counts.entry(id).or_default() += 1;
        if !seen.insert(id) {
            continue;
        }
        let Ok(data) = object.lock() else { continue };
        match &data.kind {
            ObjectKind::Array(items) => pending.extend(items.iter().cloned()),
            ObjectKind::Map(entries) => pending.extend(entries.values().cloned()),
            ObjectKind::Function(_) | ObjectKind::HostFunction(_) => {}
            _ => {
                pending.extend(data.fields.iter().cloned());
                pending.extend(data.properties.values().cloned());
            }
        }
    }
    counts
}

/// php's end-of-request pass. On an explicit exit, objects only the
/// still-running frames hold go first, as those frames unwind, in creation
/// order. Then, walking the global variables from the last defined to the
/// first, an object held by nothing but that variable is destructed and the
/// variable removed, repeated while that frees more. Finally every object
/// still alive is destructed in creation order, including objects created by
/// those destructors.
pub fn shutdown(ctx: &mut HostContext, explicit_exit: bool) {
    if live_count() == 0 {
        return;
    }
    if let Some(store) = global_variable_store(ctx)
        && explicit_exit
    {
        let variables: Vec<Value> = match store.lock() {
            Ok(store) => store.properties.values().cloned().collect(),
            Err(_) => Vec::new(),
        };
        let global = reachable_from(&variables);
        let frame_only: Vec<usize> = REGISTRY.with(|registry| {
            registry.borrow().order.iter().copied().filter(|id| !global.contains(id)).collect()
        });
        for id in frame_only {
            if let Some(entry) = take(id) {
                destruct_and_free(ctx, entry);
            }
        }
    }
    if let Some(store) = global_variable_store(ctx) {
        loop {
            let variables: Vec<(String, Value)> = match store.lock() {
                Ok(store) => store.properties.iter().map(|(k, v)| (k.clone(), v.clone())).collect(),
                Err(_) => break,
            };
            let roots: Vec<Value> = variables.iter().map(|(_, value)| value.clone()).collect();
            let counts = reference_counts(&roots);
            let mut freed = false;
            for (name, value) in variables.iter().rev() {
                let Value::Object(object) = value else { continue };
                let id = identity(object);
                if counts.get(&id) != Some(&1) || !is_registered(id) {
                    continue;
                }
                if let Ok(mut store) = store.lock() {
                    store.properties.shift_remove(name);
                }
                if let Some(entry) = take(id) {
                    destruct_and_free(ctx, entry);
                    freed = true;
                }
            }
            if !freed {
                break;
            }
        }
    }
    loop {
        let next = REGISTRY.with(|registry| registry.borrow().order.first().copied());
        let Some(id) = next else { break };
        if let Some(entry) = take(id) {
            destruct(ctx, entry);
        }
    }
    publish_flag(ctx);
}

/// Invoke a host function by name, if this VM registered it.
fn call_host(ctx: &mut HostContext, module: &str, name: &str, args: &[Value]) -> Option<Value> {
    let index = ctx.resolve_host_function_index(module, name)?;
    let mut function = Object::new();
    function.kind = ObjectKind::HostFunction(index);
    let function = Value::Object(vybe_runtime::heap::alloc(function));
    Some(ctx.invoke(&function, args))
}

/// Register every object in `value` whose Destructor slot is callable. For
/// instances a host operation builds without running a class constructor
/// (`unserialize`, fetching rows as objects).
fn track_tree(ctx: &mut HostContext, value: &Value) {
    let key = Value::String(Arc::from(vybe_ast::protocol_slot_key(
        vybe_ast::ProtocolSlot::Destructor,
    )));
    let mut seen = HashSet::new();
    let mut pending = vec![value.clone()];
    while let Some(value) = pending.pop() {
        let Value::Object(object) = &value else { continue };
        let id = identity(object);
        if !seen.insert(id) {
            continue;
        }
        let is_instance = {
            let Ok(data) = object.lock() else { continue };
            match &data.kind {
                ObjectKind::Array(items) => {
                    pending.extend(items.iter().cloned());
                    false
                }
                ObjectKind::Map(entries) => {
                    pending.extend(entries.values().cloned());
                    false
                }
                ObjectKind::Function(_) | ObjectKind::HostFunction(_) => false,
                _ => {
                    pending.extend(data.fields.iter().cloned());
                    pending.extend(data.properties.values().cloned());
                    true
                }
            }
        };
        if !is_instance || is_registered(id) {
            continue;
        }
        let Some(destructor) = call_host(ctx, "ecma:reflect", "get", &[value.clone(), key.clone()]) else {
            continue;
        };
        let callable = matches!(&destructor, Value::Object(function)
            if matches!(function.lock().map(|f| matches!(f.kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_))), Ok(true)));
        if callable {
            REGISTRY.with(|registry| {
                let mut registry = registry.borrow_mut();
                registry.order.push(id);
                registry.live.insert(id, Entry { object: value.clone(), destructor });
            });
        }
    }
}

pub fn register(fw: &mut Framework<'_>) {
    // `trackTree(value)` — see `track_tree`. Answers the flag value.
    fw.register_host_fn(
        "php:object",
        "trackTree",
        Box::new(|ctx, args| {
            if let Some(value) = args.first() {
                track_tree(ctx, value);
            }
            publish_flag(ctx);
            if live_count() > 0 { Value::I32(1) } else { Value::Null }
        }),
    );
    // `track(object, destructor)` — called once per construction of an
    // object whose Destructor slot is callable. Answers the flag value.
    fw.register_host_fn(
        "php:object",
        "track",
        Box::new(|ctx, args| {
            let [Value::Object(object), destructor, ..] = args else {
                return Value::Null;
            };
            let id = identity(object);
            REGISTRY.with(|registry| {
                let mut registry = registry.borrow_mut();
                if !registry.live.contains_key(&id) {
                    registry.order.push(id);
                    registry.live.insert(id, Entry {
                        object: Value::Object(object.clone()),
                        destructor: destructor.clone(),
                    });
                }
            });
            publish_flag(ctx);
            Value::I32(1)
        }),
    );
    // `finalize(old, destructor)` — a rebound variable's old value is an
    // object with a callable destructor.
    fw.register_host_fn(
        "php:object",
        "finalize",
        Box::new(|ctx, args| {
            let Some(old) = args.first() else { return Value::Null };
            if let Value::Object(object) = old
                && !is_registered(identity(object))
            {
                // Constructed by a path that does not register (for example
                // a host-created instance): finalize it on this release.
                if let Some(destructor @ Value::Object(_)) = args.get(1)
                    && !ctx.is_reachable_from_program(old)
                {
                    ctx.invoke_with_receiver(destructor, old.clone(), &[]);
                }
                return Value::Null;
            }
            release(ctx, old);
            Value::Null
        }),
    );
    // `release(value)` — a value that may own destructible objects stopped
    // being held by a variable, an array element or a returning frame.
    fw.register_host_fn(
        "php:object",
        "release",
        Box::new(|ctx, args| {
            if let Some(value) = args.first() {
                release(ctx, value);
            }
            Value::Null
        }),
    );
    // The termination pass, declared on the module as its shutdown finalizer.
    fw.register_host_fn(
        "php:object",
        "shutdown",
        Box::new(|ctx, args| {
            let explicit_exit = matches!(args.first(), Some(Value::I32(n)) if *n != 0);
            shutdown(ctx, explicit_exit);
            Value::Null
        }),
    );
}
