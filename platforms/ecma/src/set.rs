//! # `ecma:set` — ECMA-262 §24.2 Set
//!
//! Native Rust impls of `Set.prototype.*` + ES2025 set-algebra methods
//! (`union`, `intersection`, `difference`, `symmetricDifference`,
//! `isSubsetOf`, `isSupersetOf`, `isDisjointFrom`).
//!
//! Backing storage is `ObjectKind::Set(IndexSet<Value>)` — O(1) avg
//! add/has/delete while preserving insertion order for iteration.
//! Membership uses `SameValueZero` via `Value`'s `Hash + Eq`.
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_runtime/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, Param, VM, ValType};

static SET_ITERATOR_IDX: OnceLock<usize> = OnceLock::new();
static SET_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

fn invoke_prepared_or_direct(
    ctx: &mut HostContext,
    callback: &Value,
    prepared: &Option<crate::function::PreparedBoundCallback>,
    args: &[Value],
) -> Value {
    if let Some(prepared) = prepared {
        crate::function::invoke_prepared_bound_callback(ctx, prepared, args)
    } else {
        ctx.invoke(callback, args)
    }
}

/// %Set.prototype% (§24.2.3) — the ONE object every Set instance inherits
/// from. See `map::shared_map_prototype`; §24.2.4 is the same sentence for
/// Sets: "Set instances are ordinary objects that inherit properties from
/// %Set.prototype%".
pub fn shared_set_prototype() -> Value {
    let proto = SET_PROTOTYPE.get_or_init(|| {
        let mut obj = Object::new();
        obj.properties.reserve(3);
        obj.properties
            .insert("__proto__".into(), crate::object::shared_object_prototype());
        // §24.2.3.12 — `Set.prototype[%Symbol.toStringTag%]` is "Set",
        // { [[Writable]]: false, [[Enumerable]]: false, [[Configurable]]: true }.
        obj.properties
            .insert("@@toStringTag".into(), crate::keys::string_value("Set"));
        obj.properties.insert(
            "__nonenum".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                crate::keys::string_value("@@toStringTag"),
            ]))),
        );
        vybe_runtime::heap::alloc(obj)
    });
    Value::Object(proto.clone())
}

fn bound_iterator_method(
    receiver: &Arc<Mutex<Object>>,
    module: &str,
    name: &str,
    idx: usize,
) -> Value {
    let mut fn_obj = Object::new();
    fn_obj.properties.reserve(6);
    fn_obj.kind = ObjectKind::HostFunction(idx);
    fn_obj
        .properties
        .insert("__host_module".into(), crate::keys::string_value(module));
    fn_obj
        .properties
        .insert("__host_name".into(), crate::keys::string_value(name));
    fn_obj
        .properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    fn_obj.properties.insert(
        "__proto__".into(),
        crate::function::shared_function_prototype(),
    );
    fn_obj
        .properties
        .insert("name".into(), crate::keys::string_value(name));
    fn_obj.properties.insert(
        "__bound_args".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            Value::Object(receiver.clone()),
        ]))),
    );
    Value::Object(vybe_runtime::heap::alloc(fn_obj))
}

/// Build a fully-formed Set object from `values`, carrying the same
/// `__type` stamp, `size` property, and `[@@iterator]` binding as a
/// user-constructed `new Set()`. Host methods that return fresh Sets
/// (`union`, `intersection`, `difference`, `symmetricDifference`, …)
/// MUST route through here so their results are spec-iterable — a raw
/// `Object::new()` result has no iterator method and `[...result]`
/// yields nothing.
pub fn make_set(values: indexmap::IndexSet<Value>) -> Value {
    let mut obj = Object::new();
    obj.properties.reserve(if SET_ITERATOR_IDX.get().is_some() {
        4
    } else {
        2
    });
    obj.kind = ObjectKind::Set(values);
    // §24.2.3.9: `size` is an accessor on the PROTOTYPE — instances have none.
    obj.properties
        .insert("__proto__".into(), shared_set_prototype());
    // __type stamp: see comment on `ecma:map.new`. Without it the
    // TypeRegistry-driven `STRUCT_GET s "add"` lookup misses.
    obj.properties
        .insert("__type".into(), crate::keys::string_value("Set"));
    let set = vybe_runtime::heap::alloc(obj);
    if let Some(idx) = SET_ITERATOR_IDX.get() {
        let mut guard = set.lock().unwrap();
        guard.properties.insert(
            "iterator".into(),
            bound_iterator_method(&set, "ecma:set", "values", *idx),
        );
        // `@@iterator` under a string spelling — see the note in `map.rs`.
        guard.properties.insert(
            "__nonenum".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                crate::keys::string_value("iterator"),
            ]))),
        );
    }
    Value::Object(set)
}

fn new_set() -> Value {
    make_set(indexmap::IndexSet::new())
}

#[inline]
fn set_values_host_key() -> &'static (String, String) {
    static KEY: std::sync::OnceLock<(String, String)> = std::sync::OnceLock::new();
    KEY.get_or_init(|| ("ecma:set".to_string(), "values".to_string()))
}

fn set_values_from_iterable_arg(arg: Option<&Value>) -> indexmap::IndexSet<Value> {
    match arg {
        Some(Value::Object(src)) => {
            let source = src.lock().unwrap();
            if let ObjectKind::Array(items) = &source.kind {
                let mut set = indexmap::IndexSet::with_capacity(items.len());
                for item in items {
                    set.insert(item.clone());
                }
                set
            } else {
                indexmap::IndexSet::new()
            }
        }
        Some(Value::String(text)) => {
            let mut set = indexmap::IndexSet::with_capacity(text.len());
            for ch in text.chars() {
                set.insert(char_value(ch));
            }
            set
        }
        _ => indexmap::IndexSet::new(),
    }
}

fn new_set_from_iterable(args: &[Value]) -> Value {
    make_set(set_values_from_iterable_arg(args.first()))
}

#[inline]
fn object_arg(args: &[Value], idx: usize) -> Option<&Arc<Mutex<Object>>> {
    match args.get(idx) {
        Some(Value::Object(obj)) => Some(obj),
        _ => None,
    }
}

fn with_two_sets<R>(
    a: &Arc<Mutex<Object>>,
    b: &Arc<Mutex<Object>>,
    f: impl FnOnce(&indexmap::IndexSet<Value>, &indexmap::IndexSet<Value>) -> R,
) -> Option<R> {
    if Arc::ptr_eq(a, b) {
        let guard = a.lock().unwrap();
        if let ObjectKind::Set(ref s) = guard.kind {
            return Some(f(s, s));
        }
        return None;
    }
    let ag = a.lock().unwrap();
    let bg = b.lock().unwrap();
    match (&ag.kind, &bg.kind) {
        (ObjectKind::Set(avs), ObjectKind::Set(bvs)) => Some(f(avs, bvs)),
        _ => None,
    }
}

/// Declare an `ecma:set` function — same closure, plus the signature.
///
/// The §24.2.4 set-operation family is the fixed-arity part of this module:
/// each takes the receiver and exactly one other set. `add`/`delete`/`has`
/// tolerate a missing operand (real JS adds `undefined`), and `forEach` carries
/// an optional `thisArg`, so those stay undeclared for the reason spelled out
/// over `register` in `array.rs`.
fn set_fn(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(HostFnDecl::new("ecma:set", name, call).with_sig(FuncSig {
        name: name.to_string(),
        params: Param::unnamed_list(params),
        results,
    }));
}

/// A Set — an object reference, so `Any` rather than a resource handle.
fn set_t() -> ValType {
    ValType::Any
}

/// The receiver and one other set: `a.union(b)`, `a.isSubsetOf(b)`, …
fn set_pair(
    vm: &mut VM,
    name: &str,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    set_fn(vm, name, vec![set_t(), set_t()], results, call);
}

pub fn register(vm: &mut VM) {
    // `new Set(iterable?)` — per ECMA-262 §24.2.1.1 the constructor optionally
    // takes an iterable whose elements become Set members.
    vm.register_host_fn(
        "ecma:set",
        "new",
        Box::new(|_ctx, args| new_set_from_iterable(args)),
    );

    set_fn(
        vm,
        "fromIterable",
        vec![ValType::Any],
        vec![set_t()],
        Box::new(|_ctx, args| make_set(set_values_from_iterable_arg(args.first()))),
    );

    vm.register_host_fn(
        "ecma:set",
        "add",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(setobj)) = args.first() {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                {
                    let mut so = setobj.lock().unwrap();
                    if let ObjectKind::Set(ref mut s) = so.kind {
                        s.insert(v);
                    }
                }
                return Value::Object(setobj.clone());
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:set",
        "has",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(setobj)) = args.first() {
                let undefined = Value::Undefined;
                let v = args.get(1).unwrap_or(&undefined);
                let so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    return Value::Bool(s.contains(v));
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:set",
        "delete",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(setobj)) = args.first() {
                let undefined = Value::Undefined;
                let v = args.get(1).unwrap_or(&undefined);
                let mut so = setobj.lock().unwrap();
                let removed = if let ObjectKind::Set(ref mut s) = so.kind {
                    // `shift_remove` preserves insertion order of the
                    // remaining members per ECMA-262 §24.2.3.4.
                    s.shift_remove(v)
                } else {
                    false
                };
                return Value::Bool(removed);
            }
            Value::Bool(false)
        }),
    );

    set_fn(
        vm,
        "clear",
        vec![set_t()],
        vec![],
        Box::new(|_ctx, args| {
            if let Some(Value::Object(setobj)) = args.first() {
                let mut so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref mut s) = so.kind {
                    s.clear();
                }
            }
            Value::Null
        }),
    );

    set_fn(
        vm,
        "size",
        vec![set_t()],
        vec![ValType::I32],
        Box::new(|_ctx, args| {
            if let Some(Value::Object(setobj)) = args.first() {
                let so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    return Value::I32(s.len() as i32);
                }
            }
            Value::I32(0)
        }),
    );

    for name in &["values", "keys"] {
        vm.register_host_fn(
            "ecma:set",
            name,
            Box::new(|_ctx, args| {
                if let Some(Value::Object(setobj)) = args.first() {
                    let so = setobj.lock().unwrap();
                    if let ObjectKind::Set(ref s) = so.kind {
                        let mut snapshot = Vec::with_capacity(s.len());
                        snapshot.extend(s.iter().cloned());
                        return crate::array::make_array_iterator(snapshot);
                    }
                }
                crate::array::make_array_iterator(Vec::new())
            }),
        );
    }
    if let Some(idx) = vm.host_registry.get(set_values_host_key()).copied() {
        let _ = SET_ITERATOR_IDX.set(idx);
    }

    set_fn(
        vm,
        "entries",
        vec![set_t()],
        vec![ValType::Any],
        Box::new(|_ctx, args| {
            if let Some(Value::Object(setobj)) = args.first() {
                let so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    let mut pairs = Vec::with_capacity(s.len());
                    for v in s {
                        pairs.push(crate::array::make_pair_array(v.clone(), v.clone()));
                    }
                    return crate::array::make_array_iterator(pairs);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );

    // Set.prototype.forEach(callback) — callback receives (value,
    // value, set) — the key mirrors the value per §24.2.3.6.
    vm.register_host_fn(
        "ecma:set",
        "forEach",
        Box::new(|ctx, args| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let this_arg = args.get(2).cloned();
            let saved_this = this_arg.as_ref().map(|_| ctx.current_js_this());
            let prepared_callback = crate::function::prepare_bound_callback(callback);
            if let Some(Value::Object(setobj)) = args.first() {
                let snapshot: Vec<Value> = {
                    let so = setobj.lock().unwrap();
                    if let ObjectKind::Set(ref s) = so.kind {
                        let mut snapshot = Vec::with_capacity(s.len());
                        snapshot.extend(s.iter().cloned());
                        snapshot
                    } else {
                        Vec::new()
                    }
                };
                let receiver = Value::Object(setobj.clone());
                let mut invoke_args = [Value::Undefined, Value::Undefined, receiver];
                for v in snapshot {
                    invoke_args[0] = v.clone();
                    invoke_args[1] = v;
                    if let Some(this_arg) = &this_arg {
                        ctx.set_js_this(this_arg.clone());
                    }
                    invoke_prepared_or_direct(ctx, callback, &prepared_callback, &invoke_args);
                    if let Some(saved_this) = &saved_this {
                        ctx.set_js_this(saved_this.clone());
                    }
                }
            }
            Value::Undefined
        }),
    );

    // ── Set algebra (ES2025) ────────────────────────────────────────
    //
    // IndexSet gives us native `.union` / `.intersection` methods, but
    // we hand-roll here to preserve ECMA-262's insertion-order
    // semantics for the result: "iterate a first, then take b's
    // members that aren't in a" for `union`; "iterate a, keep those
    // also in b" for `intersection`; etc.

    set_pair(
        vm,
        "union",
        vec![set_t()],
        Box::new(|_ctx, args| {
            let mut out = indexmap::IndexSet::new();
            for arg_idx in 0..2 {
                if let Some(setobj) = object_arg(args, arg_idx) {
                    let so = setobj.lock().unwrap();
                    if let ObjectKind::Set(ref s) = so.kind {
                        out.reserve(s.len());
                        for v in s.iter() {
                            out.insert(v.clone());
                        }
                    }
                }
            }
            make_set(out)
        }),
    );

    set_pair(
        vm,
        "intersection",
        vec![set_t()],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if let Some(out) = with_two_sets(a, b, |avs, bvs| {
                    let mut out = indexmap::IndexSet::with_capacity(avs.len().min(bvs.len()));
                    for v in avs.iter() {
                        if bvs.contains(v) {
                            out.insert(v.clone());
                        }
                    }
                    out
                }) {
                    return make_set(out);
                }
            }
            new_set()
        }),
    );

    set_pair(
        vm,
        "difference",
        vec![set_t()],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if let Some(out) = with_two_sets(a, b, |avs, bvs| {
                    let mut out = indexmap::IndexSet::with_capacity(avs.len());
                    for v in avs.iter() {
                        if !bvs.contains(v) {
                            out.insert(v.clone());
                        }
                    }
                    out
                }) {
                    return make_set(out);
                }
            }
            new_set()
        }),
    );

    set_pair(
        vm,
        "symmetricDifference",
        vec![set_t()],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if let Some(out) = with_two_sets(a, b, |avs, bvs| {
                    let mut out = indexmap::IndexSet::with_capacity(avs.len() + bvs.len());
                    for v in avs.iter() {
                        if !bvs.contains(v) {
                            out.insert(v.clone());
                        }
                    }
                    for v in bvs.iter() {
                        if !avs.contains(v) {
                            out.insert(v.clone());
                        }
                    }
                    out
                }) {
                    return make_set(out);
                }
            }
            new_set()
        }),
    );

    set_pair(
        vm,
        "isSubsetOf",
        vec![ValType::I32],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if let Some(is_sub) = with_two_sets(a, b, |avs, bvs| {
                    avs.len() <= bvs.len() && avs.iter().all(|v| bvs.contains(v))
                }) {
                    return Value::I32(if is_sub { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }),
    );

    set_pair(
        vm,
        "isSupersetOf",
        vec![ValType::I32],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if let Some(is_super) = with_two_sets(a, b, |avs, bvs| {
                    avs.len() >= bvs.len() && bvs.iter().all(|v| avs.contains(v))
                }) {
                    return Value::I32(if is_super { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }),
    );

    set_pair(
        vm,
        "isDisjointFrom",
        vec![ValType::I32],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if let Some(disjoint) = with_two_sets(a, b, |avs, bvs| {
                    let (smaller, larger) = if avs.len() <= bvs.len() {
                        (avs, bvs)
                    } else {
                        (bvs, avs)
                    };
                    !smaller.iter().any(|v| larger.contains(v))
                }) {
                    return Value::I32(if disjoint { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }),
    );

    // .NET HashSet mutating set algebra — `UnionWith` / `IntersectWith` /
    // `ExceptWith` / `SymmetricExceptWith` modify the receiver in place.
    // Distinct from the immutable ES2025 `union` / `intersection` / etc.
    // which return a fresh Set. The ES variants are still registered above;
    // these mutate variants are the .NET-shape entry points.
    set_pair(
        vm,
        "unionWith",
        vec![],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if Arc::ptr_eq(a, b) {
                    return Value::Undefined;
                }
                let to_add: Vec<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        let mut values = Vec::with_capacity(bvs.len());
                        values.extend(bvs.iter().cloned());
                        values
                    } else {
                        Vec::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    for v in to_add {
                        avs.insert(v);
                    }
                }
            }
            Value::Undefined
        }),
    );

    set_pair(
        vm,
        "intersectWith",
        vec![],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if Arc::ptr_eq(a, b) {
                    return Value::Undefined;
                }
                let b_snapshot: indexmap::IndexSet<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        bvs.clone()
                    } else {
                        indexmap::IndexSet::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    avs.retain(|v| b_snapshot.contains(v));
                }
            }
            Value::Undefined
        }),
    );

    set_pair(
        vm,
        "exceptWith",
        vec![],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if Arc::ptr_eq(a, b) {
                    let mut alock = a.lock().unwrap();
                    if let ObjectKind::Set(ref mut avs) = alock.kind {
                        avs.clear();
                    }
                    return Value::Undefined;
                }
                let b_snapshot: indexmap::IndexSet<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        bvs.clone()
                    } else {
                        indexmap::IndexSet::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    avs.retain(|v| !b_snapshot.contains(v));
                }
            }
            Value::Undefined
        }),
    );

    set_pair(
        vm,
        "symmetricExceptWith",
        vec![],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if Arc::ptr_eq(a, b) {
                    let mut alock = a.lock().unwrap();
                    if let ObjectKind::Set(ref mut avs) = alock.kind {
                        avs.clear();
                    }
                    return Value::Undefined;
                }
                let b_snapshot: indexmap::IndexSet<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        bvs.clone()
                    } else {
                        indexmap::IndexSet::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    let existing: indexmap::IndexSet<Value> = avs
                        .iter()
                        .filter(|v| b_snapshot.contains(*v))
                        .cloned()
                        .collect();
                    avs.retain(|v| !b_snapshot.contains(v));
                    for v in b_snapshot {
                        if !existing.contains(&v) {
                            avs.insert(v);
                        }
                    }
                }
            }
            Value::Undefined
        }),
    );

    set_pair(
        vm,
        "overlaps",
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (object_arg(args, 0), object_arg(args, 1)) {
                if let Some(overlap) = with_two_sets(a, b, |avs, bvs| {
                    let (smaller, larger) = if avs.len() <= bvs.len() {
                        (avs, bvs)
                    } else {
                        (bvs, avs)
                    };
                    smaller.iter().any(|v| larger.contains(v))
                }) {
                    return Value::Bool(overlap);
                }
            }
            Value::Bool(false)
        }),
    );
}
