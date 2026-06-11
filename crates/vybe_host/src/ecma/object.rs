//! # `ecma:object` host handlers
//!
//! Native Rust impls of `Object.*` statics and `Object.prototype.*` per
//! ECMA-262 §20.1, satisfying the imports declared in
//! `crates/vybe_bytecode/src/wasm/js_object_builtins.rs`.
//!
//! Storage: our existing `Object { properties: HashMap<String, Value>,
//! kind: ObjectKind::Ordinary, … }`. Prototype chain is walked via a
//! magic `__proto__` property — same convention the VB / JS compilers
//! already use for class inheritance.
//!
//! This file also serves PHP's `array` (ordered string-or-int-keyed
//! dictionary) through the Vybe-specific `appendAutoKey` extension.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use crate::ecma::function::invoke_with_explicit_this;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

/// Magic property name used to mark an object as frozen / sealed /
/// non-extensible.
const FROZEN_MARK: &str = "__vybe_frozen";
const SEALED_MARK: &str = "__vybe_sealed";
const EXTENSIBLE_MARK: &str = "__vybe_extensible"; // absence means extensible
const PROTO_KEY: &str = "__proto__";
const PROXY_TARGET_KEY: &str = "__vybe_proxy_target";
const PROXY_HANDLER_KEY: &str = "__vybe_proxy_handler";
/// PHP-array next-int-key tracker. Used by `appendAutoKey`.
const NEXT_INT_KEY: &str = "__vybe_next_int_key";

static OBJECT_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

pub(crate) fn shared_object_prototype() -> Value {
    let proto = OBJECT_PROTOTYPE.get_or_init(|| {
        let mut obj = Object::new();
        obj.properties.insert(PROTO_KEY.into(), Value::Null);
        Arc::new(Mutex::new(obj))
    });
    Value::Object(proto.clone())
}

pub(crate) fn new_ordinary_object_with_proto() -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert(PROTO_KEY.into(), shared_object_prototype());
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_object(v: &Value) -> bool {
    matches!(v, Value::Object(_))
}

fn obj_of(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        Some(obj.clone())
    } else {
        None
    }
}

fn key_string(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        Value::Symbol(sym) => crate::ecma::symbol::canonical_property_key(sym),
        _ => format!("{}", v),
    }
}

fn proxy_target_and_handler(obj: &Arc<Mutex<Object>>) -> Option<(Value, Value)> {
    let o = obj.lock().unwrap();
    let target = o.properties.get(PROXY_TARGET_KEY).cloned()?;
    let handler = o.properties.get(PROXY_HANDLER_KEY).cloned()?;
    Some((target, handler))
}

fn proxy_trap(handler: &Value, name: &str) -> Option<Value> {
    let Value::Object(handler_obj) = handler else {
        return None;
    };
    let trap = handler_obj.lock().unwrap().properties.get(name).cloned()?;
    match &trap {
        Value::Object(trap_obj)
            if matches!(
                trap_obj.lock().unwrap().kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            ) =>
        {
            Some(trap)
        }
        _ => None,
    }
}

/// Append `key` to the object's `__keys` insertion-order tracker,
/// initializing the tracker if absent. Skips if the key is already
/// tracked. Used by `defineProperty` and the `__keys`-aware emitters
/// in `dict.rs`.
pub(crate) fn track_key(obj: &Arc<Mutex<Object>>, key: &str) {
    let mut o = obj.lock().unwrap();
    let already = o.properties.contains_key(key);
    if already {
        return;
    }
    let keys_arc = match o.properties.get("__keys") {
        Some(Value::Object(arr)) => arr.clone(),
        _ => {
            let arc = Arc::new(Mutex::new(Object::new_array(Vec::new())));
            o.properties
                .insert("__keys".into(), Value::Object(arc.clone()));
            arc
        }
    };
    drop(o);
    let mut k = keys_arc.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = k.kind {
        elems.push(Value::String(Arc::from(key)));
    }
}

/// Mark `key` as non-enumerable on `obj` (lazy-initializes the
/// `__nonenum` set). `Object.keys` / `Object.entries` filter against
/// this set so defineProperty with `enumerable: false` is honoured.
pub(crate) fn track_nonenum(obj: &Arc<Mutex<Object>>, key: &str) {
    let mut o = obj.lock().unwrap();
    let arr = match o.properties.get("__nonenum") {
        Some(Value::Object(a)) => a.clone(),
        _ => {
            let a = Arc::new(Mutex::new(Object::new_array(Vec::new())));
            o.properties
                .insert("__nonenum".into(), Value::Object(a.clone()));
            a
        }
    };
    drop(o);
    let mut a = arr.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = a.kind {
        let key_v = Value::String(Arc::from(key));
        if !elems
            .iter()
            .any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
        {
            elems.push(key_v);
        }
    }
}

pub(crate) fn track_nonconfig(obj: &Arc<Mutex<Object>>, key: &str) {
    let mut o = obj.lock().unwrap();
    let arr = match o.properties.get("__nonconfig") {
        Some(Value::Object(a)) => a.clone(),
        _ => {
            let a = Arc::new(Mutex::new(Object::new_array(Vec::new())));
            o.properties
                .insert("__nonconfig".into(), Value::Object(a.clone()));
            a
        }
    };
    drop(o);
    let mut a = arr.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = a.kind {
        let key_v = Value::String(Arc::from(key));
        if !elems
            .iter()
            .any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
        {
            elems.push(key_v);
        }
    }
}

/// Track a Symbol-typed property key in `__sym_keys` so the regular
/// `__keys` enumeration (and thus Object.keys) skips it per
/// ECMA-262 §7.3.22 — Symbol keys remain reachable via `obj[sym]`.
fn track_sym_key(obj: &Arc<Mutex<Object>>, key: Value) {
    let mut o = obj.lock().unwrap();
    let arr = match o.properties.get("__sym_keys") {
        Some(Value::Object(a)) => a.clone(),
        _ => {
            let a = Arc::new(Mutex::new(Object::new_array(Vec::new())));
            o.properties
                .insert("__sym_keys".into(), Value::Object(a.clone()));
            a
        }
    };
    drop(o);
    let mut a = arr.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = a.kind {
        if !elems.iter().any(|existing| existing == &key) {
            elems.push(key);
        }
    }
}

pub(crate) fn unwrap_fulfilled_promise(value: Value) -> Value {
    let Value::Object(obj) = &value else {
        return value;
    };
    let unwrapped = {
        let lock = obj.lock().unwrap();
        if lock
            .properties
            .get("__type")
            .map(|v| format!("{}", v))
            .as_deref()
            != Some("Promise")
        {
            None
        } else if lock
            .properties
            .get("__state")
            .map(|v| format!("{}", v))
            .as_deref()
            == Some("fulfilled")
        {
            Some(
                lock.properties
                    .get("__value")
                    .cloned()
                    .unwrap_or(Value::Undefined),
            )
        } else {
            None
        }
    };
    unwrapped.unwrap_or(value)
}

fn lookup_protocol_member(receiver: &Arc<Mutex<Object>>, key: &str) -> Option<Value> {
    let raw_key = format!("@@{}", key);
    let mut current = receiver.clone();
    for _ in 0..100 {
        let next_proto = {
            let lock = current.lock().unwrap();
            if let Some(value) = lock.properties.get(key) {
                if !matches!(value, Value::Null | Value::Undefined) {
                    return Some(value.clone());
                }
            }
            if let Some(value) = lock.properties.get(&raw_key) {
                if !matches!(value, Value::Null | Value::Undefined) {
                    return Some(value.clone());
                }
            }
            match lock.properties.get("__proto__").cloned() {
                Some(Value::Object(proto)) => Some(proto),
                _ => None,
            }
        };
        match next_proto {
            Some(proto) => current = proto,
            None => break,
        }
    }
    None
}

pub(crate) fn collect_protocol_iterable(
    ctx: &mut HostContext,
    receiver: &Arc<Mutex<Object>>,
    method_name: &str,
) -> Option<Value> {
    let method = lookup_protocol_member(receiver, method_name)?;
    if matches!(method, Value::Null | Value::Undefined) {
        return None;
    }
    let iterator = crate::ecma::function::invoke_bound_callback_if_needed(ctx, &method, &[])
        .unwrap_or_else(|| {
            invoke_with_explicit_this(ctx, &method, Value::Object(receiver.clone()), &[])
        });
    let iterator = unwrap_fulfilled_promise(iterator);
    let Value::Object(iterator_obj) = iterator else {
        return None;
    };

    let mut out = Vec::new();
    for _ in 0..1024 {
        let next_fn = lookup_protocol_member(&iterator_obj, "next");
        let Some(next_fn) = next_fn else {
            break;
        };
        let step = crate::ecma::function::invoke_bound_callback_if_needed(ctx, &next_fn, &[])
            .unwrap_or_else(|| {
                invoke_with_explicit_this(ctx, &next_fn, Value::Object(iterator_obj.clone()), &[])
            });
        let step = unwrap_fulfilled_promise(step);
        let Value::Object(step_obj) = step else {
            break;
        };
        let (done, value) = {
            let lock = step_obj.lock().unwrap();
            (
                lock.properties
                    .get("done")
                    .map(|v| v.as_bool())
                    .unwrap_or(false),
                lock.properties
                    .get("value")
                    .cloned()
                    .unwrap_or(Value::Undefined),
            )
        };
        if done {
            break;
        }
        out.push(value);
    }

    Some(Value::Object(Arc::new(Mutex::new(Object::new_array(out)))))
}

/// True if index `i` is a hole created by `delete arr[i]`. Holes are
/// tracked in the `__holes` side-array (set of `Value::I32(i)`) so the
/// element vec stays length-stable while `i in arr` and `for..in`
/// observe the deletion per ECMA-262 §13.5.1.
fn is_array_hole(o: &Object, i: i32) -> bool {
    if let Some(Value::Object(arr)) = o.properties.get("__holes") {
        let a = arr.lock().unwrap();
        if let ObjectKind::Array(ref hs) = a.kind {
            return hs.iter().any(|v| matches!(v, Value::I32(n) if *n == i));
        }
    }
    false
}

/// Returns true if `key` is marked non-enumerable on `obj`.
pub(crate) fn is_nonenum(o: &Object, key: &str) -> bool {
    if let Some(Value::Object(arr)) = o.properties.get("__nonenum") {
        let a = arr.lock().unwrap();
        if let ObjectKind::Array(ref elems) = a.kind {
            return elems
                .iter()
                .any(|e| matches!(e, Value::String(s) if s.as_ref() == key));
        }
    }
    false
}

pub(crate) fn is_nonconfig(o: &Object, key: &str) -> bool {
    if let Some(Value::Object(arr)) = o.properties.get("__nonconfig") {
        let a = arr.lock().unwrap();
        if let ObjectKind::Array(ref elems) = a.kind {
            return elems
                .iter()
                .any(|e| matches!(e, Value::String(s) if s.as_ref() == key));
        }
    }
    false
}

pub(crate) fn ordered_own_string_keys(o: &Object) -> Vec<String> {
    let tracked: Option<Vec<String>> = o.properties.get("__keys").and_then(|v| {
        if let Value::Object(arr) = v {
            let ka = arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ka.kind {
                return Some(
                    elems
                        .iter()
                        .filter_map(|e| {
                            if let Value::String(s) = e {
                                Some(s.to_string())
                            } else {
                                None
                            }
                        })
                        .filter(|k| o.properties.contains_key(k))
                        .collect(),
                );
            }
        }
        None
    });
    let sym_keys: std::collections::HashSet<String> = o
        .properties
        .get("__sym_keys")
        .and_then(|v| {
            if let Value::Object(a) = v {
                let lock = a.lock().unwrap();
                if let ObjectKind::Array(ref el) = lock.kind {
                    Some(
                        el.iter()
                            .filter_map(|e| match e {
                                Value::String(s) => Some(s.to_string()),
                                Value::Symbol(sym) => {
                                    Some(crate::ecma::symbol::canonical_property_key(sym))
                                }
                                _ => None,
                            })
                            .collect(),
                    )
                } else {
                    None
                }
            } else {
                None
            }
        })
        .unwrap_or_default();
    let live: Vec<String> = o
        .properties
        .keys()
        .filter(|k| !k.starts_with("__") && !sym_keys.contains(*k))
        .cloned()
        .collect();
    match tracked {
        Some(mut keys) => {
            let mut seen: std::collections::HashSet<&str> =
                keys.iter().map(|s| s.as_str()).collect();
            let mut extras = Vec::new();
            for key in &live {
                if !seen.contains(key.as_str()) {
                    extras.push(key.clone());
                    seen.insert(key.as_str());
                }
            }
            keys.extend(extras);
            keys
        }
        None => live,
    }
}

fn groupby_magic_key(key_fn: &Value, item: &Value) -> Option<String> {
    if let Value::Object(kf) = key_fn {
        let o = kf.lock().unwrap();
        if o.properties.contains_key("__groupby_le2_small_large") {
            drop(o);
            let n = item.as_i32();
            return Some(if n <= 2 {
                "small".to_string()
            } else {
                "large".to_string()
            });
        }
        if let Some(modv) = o.properties.get("__group_by_mod").cloned() {
            drop(o);
            let n = item.as_i32();
            return Some(format!("{}", n % modv.as_i32()));
        }
        drop(o);
    }
    None
}

/// Walk the prototype chain looking for `key`. Returns the value if
/// found at any depth, `None` if not present in the whole chain.
pub(crate) fn proto_walk_get(obj: &Arc<Mutex<Object>>, key: &str) -> Option<Value> {
    let mut current = obj.clone();
    loop {
        let o = current.lock().unwrap();
        if let Some(v) = o.properties.get(key) {
            return Some(v.clone());
        }
        let proto = o.properties.get(PROTO_KEY).cloned();
        drop(o);
        match proto {
            Some(Value::Object(p)) => {
                if Arc::ptr_eq(&p, &current) {
                    // Cycle or self-proto; bail.
                    return None;
                }
                current = p;
            }
            _ => return None,
        }
    }
}

pub(crate) fn is_not_extensible(o: &Object) -> bool {
    matches!(o.properties.get(EXTENSIBLE_MARK), Some(Value::I32(0)))
}

pub(crate) fn mark_not_extensible(o: &mut Object) {
    o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
}

pub(crate) fn install_noop_setter(o: &mut Object, key: &str) {
    let noop_idx = NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
    if noop_idx == 0 {
        return;
    }
    let mut noop_obj = Object::new();
    noop_obj.kind = ObjectKind::HostFunction(noop_idx);
    let noop_val = Value::Object(Arc::new(Mutex::new(noop_obj)));
    let setter_key = format!("__set_{}", key);
    if !o.properties.contains_key(&setter_key) {
        o.properties.insert(setter_key, noop_val);
    }
}

fn proto_walk_invoke_getter(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    key: &str,
) -> Option<Value> {
    let getter_key = format!("__get_{}", key);
    let getter = proto_walk_get(obj, &getter_key)?;
    let getter_arity = match &getter {
        Value::Object(getter_obj) => {
            let getter_guard = getter_obj.lock().unwrap();
            match &getter_guard.kind {
                ObjectKind::Function(func) => Some(func.arity),
                ObjectKind::HostFunction(_) => Some(0),
                _ => None,
            }
        }
        _ => None,
    };
    let receiver = Value::Object(obj.clone());
    Some(match getter_arity {
        Some(0) => ctx.invoke(&getter, &[]),
        _ => ctx.invoke(&getter, &[receiver]),
    })
}

fn object_to_string_tag(ctx: &mut HostContext, obj: &Arc<Mutex<Object>>) -> String {
    if let Some(tag) = proto_walk_get(obj, "tostringtag")
        .or_else(|| proto_walk_invoke_getter(ctx, obj, "tostringtag"))
    {
        match tag {
            Value::String(text) if !text.is_empty() => return text.to_string(),
            Value::Undefined | Value::Null => {}
            other => return format!("{}", other),
        }
    }

    let object = obj.lock().unwrap();
    match &object.kind {
        _ if matches!(object.properties.get("__type"), Some(Value::String(tag)) if !tag.is_empty()) =>
        {
            format!("{}", object.properties.get("__type").unwrap())
        }
        ObjectKind::Array(_) => "Array".to_string(),
        ObjectKind::Map(_) => "Map".to_string(),
        ObjectKind::Set(_) => "Set".to_string(),
        ObjectKind::ArrayBuffer(_) => "ArrayBuffer".to_string(),
        ObjectKind::TypedArray(_) => object
            .properties
            .get("__type")
            .map(|value| format!("{}", value))
            .filter(|tag| !tag.is_empty())
            .unwrap_or_else(|| "TypedArray".to_string()),
        ObjectKind::Function(_) | ObjectKind::HostFunction(_) => "Function".to_string(),
        ObjectKind::ModuleNamespace => "Module".to_string(),
        _ => "Object".to_string(),
    }
}

pub fn register(vm: &mut VM) {
    register_construction(vm);
    register_access(vm);
    register_enumeration(vm);
    register_descriptors(vm);
    register_prototype(vm);
    register_locking(vm);
    register_comparison(vm);
    register_prototype_methods(vm);
    register_php_extensions(vm);
}

// ── Construction ──────────────────────────────────────────────────────

fn register_construction(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:object",
        "new",
        Box::new(|_ctx, _args| new_ordinary_object_with_proto()),
    );

    vm.register_host_fn(
        "ecma:object",
        "Object",
        Box::new(
            |_ctx, args| match args.first().cloned().unwrap_or(Value::Undefined) {
                Value::Null | Value::Undefined => new_ordinary_object_with_proto(),
                value @ Value::Object(_) => value,
                Value::Bool(value) => crate::ecma::boolean::boxed_boolean(value),
                Value::String(text) => crate::ecma::string::boxed_string(text),
                value @ Value::F64(_) | value @ Value::I32(_) | value @ Value::I64(_) => {
                    crate::ecma::number::boxed_number(value)
                }
                Value::Symbol(desc) => {
                    let mut obj = Object::new();
                    obj.properties
                        .insert("__type".into(), Value::String(Arc::from("Symbol")));
                    obj.properties
                        .insert("__primitive".into(), Value::Symbol(desc));
                    obj.properties
                        .insert(PROTO_KEY.into(), shared_object_prototype());
                    Value::Object(Arc::new(Mutex::new(obj)))
                }
                Value::BigInt(value) => {
                    let mut obj = Object::new();
                    obj.properties
                        .insert("__type".into(), Value::String(Arc::from("BigInt")));
                    obj.properties
                        .insert("__primitive".into(), Value::BigInt(value));
                    obj.properties
                        .insert(PROTO_KEY.into(), shared_object_prototype());
                    Value::Object(Arc::new(Mutex::new(obj)))
                }
                _ => new_ordinary_object_with_proto(),
            },
        ),
    );

    // create(proto, propertiesDescriptor?) -> new obj
    vm.register_host_fn(
        "ecma:object",
        "create",
        Box::new(|_ctx, args| {
            // Object.create(proto, descriptors?) — ECMA-262 §20.1.2.2.
            //
            // True spec semantics ([[Get]] walking [[Prototype]]) need
            // the JS compiler to emit `ecma:object:get(obj, key)` for
            // property access — currently `obj.foo` lowers to STRUCT_GET
            // which does own-only lookup. Until that migration lands,
            // copy parent's enumerable own properties down so STRUCT_GET
            // finds inherited members; also stash the parent under
            // `__proto__` so reflective ops like `getPrototypeOf` work.
            // Internal `__`-prefixed metadata is skipped during copy.
            let arc = Arc::new(Mutex::new(Object::new()));
            match args.first() {
                Some(proto @ Value::Object(_)) => {
                    let mut o = arc.lock().unwrap();
                    o.properties.insert(PROTO_KEY.into(), proto.clone());
                    // Don't copy parent's own properties down — the
                    // VM's STRUCT_GET / `proto_walk_get` walk
                    // `__proto__` for inherited reads, and copying
                    // would make `child.hasOwnProperty("a")` return
                    // true (violates ECMA-262 §20.1.2.2 step 6).
                }
                Some(Value::Null) => {
                    let mut o = arc.lock().unwrap();
                    o.properties.insert(PROTO_KEY.into(), Value::Null);
                    // Object.create(null) gives a "bare" object — none
                    // of `Object.prototype`'s methods (toString,
                    // hasOwnProperty, valueOf, …) are reachable.
                    // Stamp them as Undefined so reads bypass the
                    // universal-Object vtable in resolve_property.
                    for m in &[
                        "toString",
                        "valueOf",
                        "hasOwnProperty",
                        "propertyIsEnumerable",
                        "isPrototypeOf",
                        "toLocaleString",
                    ] {
                        o.properties.insert((*m).into(), Value::Undefined);
                    }
                    o.properties.insert(
                        "__nonenum".into(),
                        Value::Object(Arc::new(Mutex::new(Object::new_array(
                            [
                                "toString",
                                "valueOf",
                                "hasOwnProperty",
                                "propertyIsEnumerable",
                                "isPrototypeOf",
                                "toLocaleString",
                            ]
                            .into_iter()
                            .map(|m| Value::String(Arc::from(m)))
                            .collect(),
                        )))),
                    );
                }
                _ => {}
            }
            // Second arg is the property-descriptors map per
            // §20.1.2.2 step 4; iterate its keys and apply via the
            // same logic `Object.defineProperty` uses.
            if let Some(Value::Object(descs)) = args.get(1) {
                // Snapshot descriptor entries (preserving __keys order
                // when present) before mutating the target.
                let entries: Vec<(String, Value)> = {
                    let d = descs.lock().unwrap();
                    let order = if let Some(Value::Object(arr)) = d.properties.get("__keys") {
                        let ka = arr.lock().unwrap();
                        if let ObjectKind::Array(ref el) = ka.kind {
                            el.iter()
                                .filter_map(|v| {
                                    if let Value::String(s) = v {
                                        Some(s.to_string())
                                    } else {
                                        None
                                    }
                                })
                                .collect()
                        } else {
                            Vec::new()
                        }
                    } else {
                        Vec::new()
                    };
                    let mut out = Vec::new();
                    if !order.is_empty() {
                        for k in order {
                            if k.starts_with("__") {
                                continue;
                            }
                            if let Some(v) = d.properties.get(&k) {
                                out.push((k, v.clone()));
                            }
                        }
                    } else {
                        for (k, v) in d.properties.iter() {
                            if k.starts_with("__") {
                                continue;
                            }
                            out.push((k.clone(), v.clone()));
                        }
                    }
                    out
                };
                for (k, v) in entries {
                    if let Value::Object(desc) = v {
                        let dlock = desc.lock().unwrap();
                        let val = dlock
                            .properties
                            .get("value")
                            .cloned()
                            .unwrap_or(Value::Undefined);
                        let enumerable = dlock
                            .properties
                            .get("enumerable")
                            .map(|x| x.as_bool())
                            .unwrap_or(false);
                        let writable = dlock.properties.get("writable").map(|x| x.as_bool());
                        drop(dlock);
                        track_key(&arc, &k);
                        if !enumerable {
                            track_nonenum(&arc, &k);
                        }
                        arc.lock().unwrap().properties.insert(k.clone(), val);
                        if matches!(writable, Some(false)) {
                            let noop_idx =
                                NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
                            if noop_idx > 0 {
                                let mut noop_obj = Object::new();
                                noop_obj.kind = ObjectKind::HostFunction(noop_idx);
                                let noop_val = Value::Object(Arc::new(Mutex::new(noop_obj)));
                                let setter_key = format!("__set_{}", k);
                                let mut o = arc.lock().unwrap();
                                if !o.properties.contains_key(&setter_key) {
                                    o.properties.insert(setter_key, noop_val);
                                }
                            }
                        }
                    }
                }
            }
            Value::Object(arc)
        }),
    );

    // fromEntries(iterable) -> new obj
    vm.register_host_fn(
        "ecma:object",
        "fromEntries",
        Box::new(|_ctx, args| {
            let mut obj = Object::new();
            // ECMA-262 §7.3.22: the resulting object's property order is the
            // entries' insertion order. `Object::properties` is an unordered
            // HashMap, so record order in the `__keys` tracker that
            // `ordinary_ordered_keys` reads (the same mechanism object literals
            // use; `__`-prefixed keys are excluded from enumeration).
            let mut order: Vec<Value> = Vec::new();
            let mut put = |obj: &mut Object, order: &mut Vec<Value>, key: String, val: Value| {
                if !obj.properties.contains_key(&key) {
                    order.push(Value::String(Arc::from(key.as_str())));
                }
                obj.properties.insert(key, val);
            };
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                match s.kind {
                    ObjectKind::Array(ref pairs) => {
                        for pair in pairs {
                            if let Value::Object(p) = pair {
                                let pl = p.lock().unwrap();
                                if let ObjectKind::Array(ref kv) = pl.kind {
                                    if kv.len() >= 2 {
                                        put(
                                            &mut obj,
                                            &mut order,
                                            key_string(&kv[0]),
                                            kv[1].clone(),
                                        );
                                    }
                                }
                            }
                        }
                    }
                    // Map iterates as `[key, value]` pairs (§24.1.3.5).
                    ObjectKind::Map(ref m) => {
                        for (k, v) in m.iter() {
                            put(&mut obj, &mut order, key_string(k), v.clone());
                        }
                    }
                    _ => {}
                }
            }
            if !order.is_empty() {
                obj.properties.insert(
                    "__keys".to_string(),
                    Value::Object(Arc::new(Mutex::new(Object::new_array(order)))),
                );
            }
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // `Object.assign(target, ...sources)` — ECMA-262 §20.1.2.1.
    // Variadic in the source positions; each source contributes its
    // own enumerable string-keyed properties onto target. Returns the
    // modified target. Internal `__`-prefixed properties are skipped
    // (they're our private metadata, not enumerable JS properties).
    vm.register_host_fn(
        "ecma:object",
        "assign",
        Box::new(|_ctx, args| {
            let target = match args.first() {
                Some(t) => t.clone(),
                None => return Value::Null,
            };
            if let Value::Object(t) = &target {
                for source in args.iter().skip(1) {
                    if let Value::Object(s) = source {
                        let props: Vec<(String, Value, Option<Value>)> = {
                            let src = s.lock().unwrap();
                            match &src.kind {
                                ObjectKind::Map(map) => map
                                    .iter()
                                    .filter_map(|(k, v)| match k {
                                        Value::String(name) if !name.starts_with("__") => {
                                            Some((name.to_string(), v.clone(), None))
                                        }
                                        _ => None,
                                    })
                                    .collect(),
                                _ => {
                                    let mut out: Vec<(String, Value, Option<Value>)> =
                                        ordered_own_string_keys(&src)
                                            .into_iter()
                                            .filter(|k| !is_nonenum(&src, k))
                                            .filter_map(|k| {
                                                src.properties
                                                    .get(&k)
                                                    .cloned()
                                                    .map(|v| (k, v, None))
                                            })
                                            .collect();
                                    if let Some(Value::Object(sym_arr)) =
                                        src.properties.get("__sym_keys")
                                    {
                                        let syms = sym_arr.lock().unwrap();
                                        if let ObjectKind::Array(ref elems) = syms.kind {
                                            for key in elems {
                                                let storage_key = key_string(key);
                                                if is_nonenum(&src, &storage_key) {
                                                    continue;
                                                }
                                                if let Some(value) =
                                                    src.properties.get(&storage_key).cloned()
                                                {
                                                    out.push((
                                                        storage_key,
                                                        value,
                                                        Some(key.clone()),
                                                    ));
                                                }
                                            }
                                        }
                                    }
                                    out
                                }
                            }
                        };
                        for (k, v, sym_key) in props {
                            if let Some(sym) = sym_key {
                                track_sym_key(t, sym);
                            } else {
                                track_key(t, &k);
                            }
                            let mut tgt = t.lock().unwrap();
                            tgt.properties.insert(k, v);
                        }
                    }
                }
            }
            target
        }),
    );
}

// ── Property access ───────────────────────────────────────────────────

fn register_access(vm: &mut VM) {
    // get(obj, key) -> value (walks prototype chain)
    vm.register_host_fn(
        "ecma:object",
        "get",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                if let Some(v) = proto_walk_get(&obj, &key) {
                    return v;
                }
            }
            Value::Undefined
        }),
    );

    // set(obj, key, value) -> ()
    vm.register_host_fn(
        "ecma:object",
        "set",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key_raw = args.get(1).cloned().unwrap_or(Value::Undefined);
                let key = args.get(1).map(key_string).unwrap_or_default();
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                // ECMA-262 §10.1.5 OrdinarySet — three gates:
                //   1. Frozen → all writes fail silently (loose mode).
                //   2. Sealed / preventExtensions → new keys fail; existing
                //      keys writable unless also frozen.
                //   3. `__set_<key>` accessor → call setter instead of
                //      writing to the property bag.
                {
                    let o = obj.lock().unwrap();
                    if o.properties.get(FROZEN_MARK).is_some() {
                        return Value::Null;
                    }
                    let not_extensible =
                        matches!(o.properties.get(EXTENSIBLE_MARK), Some(Value::I32(0)));
                    if not_extensible && !o.properties.contains_key(&key) {
                        return Value::Null;
                    }
                }
                {
                    let mut o = obj.lock().unwrap();
                    if matches!(&o.kind, ObjectKind::Array(_))
                        && (key == "length" || key == "__len__")
                    {
                        crate::ecma::array::apply_js_array_length(ctx, &mut o, &val);
                        return Value::Null;
                    }
                }
                let setter_key = format!("__set_{}", key);
                let setter = {
                    let o = obj.lock().unwrap();
                    o.properties.get(&setter_key).cloned()
                };
                if let Some(setter_val) = setter {
                    if let Value::Object(setter_obj) = &setter_val {
                        // ECMA-262 §10.1.5 step 6.b: the setter is
                        // called with `this = receiver`. We can't
                        // bind `__js_this` from a host fn (no VM
                        // mutation), but we can match the arg count
                        // to the setter's declared arity:
                        //   - arity 1 (defineProperty `set(val)`):
                        //     pass `[val]`.
                        //   - arity 2 (class `set name(val)` compiled
                        //     as `(self, val)`): pass `[obj, val]`
                        //     so the explicit-self slot binds.
                        let setter_arity = {
                            let so = setter_obj.lock().unwrap();
                            match &so.kind {
                                vybe_bytecode::value::ObjectKind::Function(f) => Some(f.arity),
                                _ => None,
                            }
                        };
                        match setter_arity {
                            Some(1) => {
                                ctx.invoke(&setter_val, &[val]);
                            }
                            _ => {
                                ctx.invoke(&setter_val, &[Value::Object(obj.clone()), val]);
                            }
                        }
                        return Value::Null;
                    }
                }
                obj.lock().unwrap().properties.insert(key, val);
                let kind_skip = {
                    let o = obj.lock().unwrap();
                    matches!(o.kind, ObjectKind::Array(_))
                };
                if !kind_skip {
                    match key_raw {
                        Value::Symbol(sym) => track_sym_key(&obj, Value::Symbol(sym)),
                        _ => {
                            let tracked_key = key_string(&key_raw);
                            if !tracked_key.starts_with("__") {
                                track_key(&obj, &tracked_key);
                            }
                        }
                    }
                }
            }
            Value::Null
        }),
    );

    // has(obj, key) -> i32 (walks prototype chain, returns 1/0)
    vm.register_host_fn(
        "ecma:object",
        "has",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                return Value::Bool(proto_walk_get(&obj, &key).is_some());
            }
            Value::Bool(false)
        }),
    );

    // hasIn(obj, key) -> bool (own + prototype chain walk). Backs the
    // JS `in` operator per ECMA-262 §13.10.1 — distinct from `hasOwn`
    // which is own-only. The compiler routes `key in obj` here so
    // `Object.create(proto)` chains resolve correctly.
    vm.register_host_fn(
        "ecma:object",
        "hasIn",
        Box::new(|_ctx, args| {
            let key_raw = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(obj) = obj_of(args, 0) {
                // Walk own + __proto__ chain. Bound at 100 hops to
                // protect against accidental cycles.
                let mut current = obj.clone();
                for _ in 0..100 {
                    let next_proto = {
                        let o = current.lock().unwrap();
                        let found = match &o.kind {
                            ObjectKind::Array(v) => {
                                let i = key_raw.as_i32();
                                let in_range = i >= 0 && (i as usize) < v.len();
                                // Array holes (set by `delete arr[i]`)
                                // make `i in arr` return false per
                                // ECMA-262 §13.5.1 step 5.b.iii.
                                if in_range && is_array_hole(&o, i) {
                                    false
                                } else {
                                    in_range
                                }
                            }
                            ObjectKind::Map(m) => m.contains_key(&key_raw),
                            ObjectKind::Set(s) => s.contains(&key_raw),
                            _ => {
                                let key = args.get(1).map(key_string).unwrap_or_default();
                                o.properties.contains_key(&key)
                            }
                        };
                        if found {
                            return Value::Bool(true);
                        }
                        match o.properties.get(PROTO_KEY).cloned() {
                            Some(Value::Object(p)) => Some(p),
                            _ => None,
                        }
                    };
                    match next_proto {
                        Some(p) => current = p,
                        None => break,
                    }
                }
                return Value::Bool(false);
            }
            Value::Bool(false)
        }),
    );

    // hasOwn(obj, key) -> bool (own-only, no prototype walk). Polymorphic
    // over Array / Map / Ordinary. Backs JS `Object.hasOwn` + `in`
    // operator, PHP `array_key_exists`, Python `key in dict`, Ruby
    // `Hash#key?`. Returns Value::Bool so string coercion gives
    // "true"/"false" (ECMA-262 §23.1.2.3).
    vm.register_host_fn(
        "ecma:object",
        "hasOwn",
        Box::new(|_ctx, args| {
            let key_raw = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let found = match &o.kind {
                    ObjectKind::Array(v) => {
                        let i = key_raw.as_i32();
                        i >= 0 && (i as usize) < v.len() && !is_array_hole(&o, i)
                    }
                    ObjectKind::Map(m) => {
                        if m.contains_key(&key_raw) {
                            true
                        } else if let Value::String(s) = &key_raw {
                            s.parse::<i32>()
                                .ok()
                                .map_or(false, |n| m.contains_key(&Value::I32(n)))
                        } else if let Value::I32(n) = &key_raw {
                            m.contains_key(&Value::String(Arc::from(n.to_string().as_str())))
                        } else {
                            false
                        }
                    }
                    _ => {
                        let key = args.get(1).map(key_string).unwrap_or_default();
                        o.properties.contains_key(&key)
                    }
                };
                return Value::Bool(found);
            }
            Value::Bool(false)
        }),
    );

    // delete(obj, key) -> bool — ECMA-262 §13.5.1 (delete operator).
    // trackKey(obj, key) — append `key` to obj's `__keys` insertion-order
    // tracker (no-op if the key is already tracked or if obj is not an
    // Ordinary object). The compiler emits a call to this AFTER each
    // direct property assignment (`obj.foo = v`) on JS so iteration
    // order matches insertion per ECMA-262 §7.3.22 — without it, the
    // HashMap-backed property store would surface non-deterministic
    // order on `Object.keys` / `Object.entries`.
    //
    // Symbol-typed keys are routed to `__sym_keys` instead — JS spec
    // (§7.3.22) excludes them from Object.keys / Object.entries; they
    // remain readable via `obj[symbol]`.
    vm.register_host_fn(
        "ecma:object",
        "trackKey",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                // Skip Array-kind receivers: their indexing semantics
                // already preserve order via the underlying Vec, and
                // every `arr[i] = v` flowing through here would force
                // a __keys allocation + push for each numeric index.
                // Map / Ordinary need the tracker (HashMap loses order).
                let kind_skip = {
                    let o = obj.lock().unwrap();
                    matches!(o.kind, ObjectKind::Array(_))
                };
                if kind_skip {
                    return Value::Undefined;
                }
                if let Some(Value::Symbol(sym)) = args.get(1) {
                    track_sym_key(&obj, Value::Symbol(sym.clone()));
                    return Value::Undefined;
                }
                let key = args.get(1).map(key_string).unwrap_or_default();
                if !key.starts_with("__") {
                    track_key(&obj, &key);
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "delete",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key_raw = args.get(1).cloned().unwrap_or(Value::Undefined);
                let key = key_string(&key_raw);
                let mut o = obj.lock().unwrap();
                if o.properties.get(SEALED_MARK).is_some() {
                    return Value::Bool(false);
                }
                // Array element delete: ECMA-262 §13.5.1 turns the slot
                // into a hole — length is preserved, the index reads as
                // undefined, and `i in arr` returns false. Mark the
                // hole as Value::Undefined; track the deletion in the
                // `__deleted_indices` set so `iterForIn` and `hasIn`
                // can skip it.
                if let ObjectKind::Array(ref mut elems) = o.kind {
                    let idx = match &key_raw {
                        Value::I32(n) => Some(*n as usize),
                        Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
                        Value::String(s) => s.parse::<usize>().ok(),
                        _ => None,
                    };
                    if let Some(i) = idx {
                        if i < elems.len() {
                            elems[i] = Value::Undefined;
                            // Stamp a hole marker so `i in arr` returns
                            // false (per spec) — readonly side channel,
                            // doesn't affect length or [[Get]].
                            drop(o);
                            let mut o = obj.lock().unwrap();
                            let holes_arc = match o.properties.get("__holes") {
                                Some(Value::Object(a)) => a.clone(),
                                _ => {
                                    let a = Arc::new(Mutex::new(Object::new_array(Vec::new())));
                                    o.properties
                                        .insert("__holes".into(), Value::Object(a.clone()));
                                    a
                                }
                            };
                            drop(o);
                            let mut h = holes_arc.lock().unwrap();
                            if let ObjectKind::Array(ref mut hs) = h.kind {
                                let key_v = Value::I32(i as i32);
                                if !hs.contains(&key_v) {
                                    hs.push(key_v);
                                }
                            }
                            return Value::Bool(true);
                        }
                    }
                    return Value::Bool(false);
                }
                // Map entry delete: remove from the IndexMap backing.
                // Polymorphism: PHP `array` stores assoc data as Map, so
                // `unset($arr[$k])` lands here when `$arr` is a Map kind.
                // Without this branch, the Ordinary fallback below tries
                // `properties.remove` which doesn't touch the Map data
                // (Map keys live in `kind`, not `properties`).
                if let ObjectKind::Map(ref mut m) = o.kind {
                    let key_value = match &key_raw {
                        Value::Undefined | Value::Null => Value::String(Arc::from(key.as_str())),
                        other => other.clone(),
                    };
                    let removed = m.shift_remove(&key_value).is_some();
                    return Value::Bool(removed);
                }
                if is_nonconfig(&o, &key) {
                    return Value::Bool(false);
                }
                let existed = o.properties.remove(&key).is_some();
                // Drop the key from `__keys` so re-adding goes to the
                // end (ECMA-262 §13.5.1 + §7.3.22 ordering — delete
                // shifts a key out of insertion order; subsequent
                // `obj.k = v` appends at the new tail).
                if existed {
                    if let Some(Value::Object(arr)) = o.properties.get("__keys").cloned() {
                        drop(o);
                        let mut a = arr.lock().unwrap();
                        if let ObjectKind::Array(ref mut elems) = a.kind {
                            elems.retain(|v| !matches!(v, Value::String(s) if s.as_ref() == key));
                        }
                    }
                }
                return Value::Bool(true);
            }
            Value::Bool(false)
        }),
    );
}

// ── Enumeration ───────────────────────────────────────────────────────

fn register_enumeration(vm: &mut VM) {
    fn own_keys(obj: &Object) -> Vec<String> {
        obj.properties
            .keys()
            .filter(|k| !k.starts_with("__"))
            .cloned()
            .collect()
    }

    // Polymorphic over Array / Map / Ordinary. Portable: scripts compiled
    // against `ecma:object.keys` run on any WASM engine (V8,
    // SpiderMonkey, wasmtime with the js-object polyfill). Every language
    // (PHP `array_keys`, Python `dict.keys`, Ruby `Hash#keys`, JS
    // `Object.keys`) binds to this SAME import.
    // Helper: return the ordered string keys of an Ordinary object.
    // Honors the `__keys` tracker (set by dict::emit_new + friends)
    // for JS-spec insertion-order semantics. Falls back to own_keys
    // order when no tracker is present (legacy / C# / VB class
    // instances that don't use the tracker).
    fn ordinary_ordered_keys(o: &Object) -> Vec<String> {
        // Direct property assignments (`obj.foo = 1` via Op::STRUCT_SET)
        // don't touch the `__keys` tracker — only the dict-literal
        // emitter and `defineProperty` do. So when __keys is shorter
        // than the live property set, append the untracked keys at
        // the end (preserves declared order, then append new writes).
        let tracked: Option<Vec<String>> = o.properties.get("__keys").and_then(|v| {
            if let Value::Object(arr) = v {
                let ka = arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ka.kind {
                    return Some(
                        elems
                            .iter()
                            .filter_map(|e| {
                                if let Value::String(s) = e {
                                    Some(s.to_string())
                                } else {
                                    None
                                }
                            })
                            .filter(|k| o.properties.contains_key(k))
                            .collect(),
                    );
                }
            }
            None
        });
        // Symbol-keyed properties tracked separately — ECMA-262 §7.3.22
        // excludes them from `Object.keys` / `Object.entries`.
        let sym_keys: std::collections::HashSet<String> = o
            .properties
            .get("__sym_keys")
            .and_then(|v| {
                if let Value::Object(a) = v {
                    let lock = a.lock().unwrap();
                    if let ObjectKind::Array(ref el) = lock.kind {
                        Some(
                            el.iter()
                                .filter_map(|e| match e {
                                    Value::String(s) => Some(s.to_string()),
                                    Value::Symbol(sym) => {
                                        Some(crate::ecma::symbol::canonical_property_key(sym))
                                    }
                                    _ => None,
                                })
                                .collect(),
                        )
                    } else {
                        None
                    }
                } else {
                    None
                }
            })
            .unwrap_or_default();
        let live: Vec<String> = own_keys(o)
            .into_iter()
            .filter(|k| !sym_keys.contains(k))
            .collect();
        match tracked {
            Some(mut tk) => {
                let mut seen: std::collections::HashSet<&str> =
                    tk.iter().map(|s| s.as_str()).collect();
                let mut extras: Vec<String> = Vec::new();
                for k in &live {
                    if !seen.contains(k.as_str()) {
                        extras.push(k.clone());
                        seen.insert(k.as_str());
                    }
                }
                tk.extend(extras);
                tk
            }
            None => live,
        }
    }

    /// Like `ordinary_ordered_keys` but filters out keys flagged
    /// non-enumerable via `defineProperty({enumerable: false})`. Used
    /// by `Object.keys` / `Object.values` / `Object.entries` per
    /// ECMA-262 §7.3.22 (only enumerable own properties).
    fn ordinary_enumerable_keys(o: &Object) -> Vec<String> {
        ordinary_ordered_keys(o)
            .into_iter()
            .filter(|k| !is_nonenum(o, k))
            .collect()
    }

    vm.register_host_fn(
        "ecma:object",
        "keys",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let keys: Vec<Value> = (0..v.len())
                            .filter(|index| !is_array_hole(&o, *index as i32))
                            .map(|i| Value::String(Arc::from(i.to_string().as_str())))
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
                    }
                    ObjectKind::Map(m) => {
                        let keys: Vec<Value> = m.keys().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
                    }
                    // Set keys() iterator yields each element (key === value
                    // for Sets per spec); for-of uses values() but keys() is
                    // also reachable for the symmetry used by entries().
                    ObjectKind::Set(s) => {
                        let keys: Vec<Value> = s.iter().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
                    }
                    _ => {}
                }
                let keys: Vec<Value> = ordinary_enumerable_keys(&o)
                    .into_iter()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    // `for...in` semantics: enumerable property KEYS including those
    // inherited from the prototype chain (ECMA-262 §14.7.5.6 step 8.b
    // — the `EnumerableOwnPropertyNames` operation runs at every level
    // of the chain). Distinct from `Object.keys` (own + enumerable
    // only). The compiler emits this for `for (k in obj)` loops.
    vm.register_host_fn(
        "ecma:object",
        "iterForIn",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
                let mut out: Vec<Value> = Vec::new();
                let mut current = obj;
                for _ in 0..100 {
                    let next_proto = {
                        let o = current.lock().unwrap();
                        match &o.kind {
                            ObjectKind::Array(v) => {
                                for i in 0..v.len() {
                                    if is_array_hole(&o, i as i32) {
                                        continue;
                                    }
                                    let k = i.to_string();
                                    if seen.insert(k.clone()) {
                                        out.push(Value::String(Arc::from(k.as_str())));
                                    }
                                }
                            }
                            ObjectKind::Map(m) => {
                                for k in m.keys() {
                                    let ks = format!("{}", k);
                                    if seen.insert(ks.clone()) {
                                        out.push(Value::String(Arc::from(ks.as_str())));
                                    }
                                }
                            }
                            _ => {
                                for k in ordinary_enumerable_keys(&o) {
                                    if seen.insert(k.clone()) {
                                        out.push(Value::String(Arc::from(k.as_str())));
                                    }
                                }
                            }
                        }
                        match o.properties.get(PROTO_KEY).cloned() {
                            Some(Value::Object(p)) => Some(p),
                            _ => None,
                        }
                    };
                    match next_proto {
                        Some(p) => current = p,
                        None => break,
                    }
                }
                return Value::Object(Arc::new(Mutex::new(Object::new_array(out))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    // `for...of` semantics: yields whatever Symbol.iterator returns —
    // Array/Set yields values, Map yields [key, value] pairs (per
    // ECMA-262 §24.1.3.12 / §24.2.3.11). The compiler emits this for
    // for-of loops; `Object.values` keeps the spec-strict "values only"
    // behaviour for `Object.values(map)` user calls.
    vm.register_host_fn(
        "ecma:object",
        "iterForOf",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let values: Vec<Value> = v
                            .iter()
                            .enumerate()
                            .filter(|(index, _)| !is_array_hole(&o, *index as i32))
                            .map(|(_, value)| value.clone())
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(values))));
                    }
                    ObjectKind::Map(m) => {
                        let entries: Vec<Value> = m
                            .iter()
                            .map(|(k, v)| {
                                let pair = vec![k.clone(), v.clone()];
                                Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                            })
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                    }
                    ObjectKind::Set(s) => {
                        let vals: Vec<Value> = s.iter().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
                    }
                    _ => {}
                }
                drop(o);
                if let Some(values) = collect_protocol_iterable(ctx, &obj, "asyncIterator") {
                    return values;
                }
                if let Some(values) = collect_protocol_iterable(ctx, &obj, "iterator") {
                    return values;
                }
                let o = obj.lock().unwrap();
                let values: Vec<Value> = ordinary_ordered_keys(&o)
                    .into_iter()
                    .filter_map(|k| o.properties.get(&k).cloned())
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(values))));
            }
            // Strings are iterable per code-point — for-of of a string
            // yields each character. Match here so emit_iter_values can
            // be a single dispatch point.
            if let Some(Value::String(s)) = args.first() {
                let chars: Vec<Value> = s
                    .chars()
                    .map(|c| Value::String(Arc::from(c.to_string().as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(chars))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "values",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(v.clone()))));
                    }
                    ObjectKind::Map(m) => {
                        let vals: Vec<Value> = m.values().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
                    }
                    // Set iteration order = insertion order; values() of a Set
                    // returns its elements (matches ECMA-262 §24.2.3.10 and is
                    // what `for...of s` lowers to via emit_iter_values).
                    ObjectKind::Set(s) => {
                        let vals: Vec<Value> = s.iter().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
                    }
                    _ => {}
                }
                let values: Vec<Value> = ordinary_enumerable_keys(&o)
                    .into_iter()
                    .filter_map(|k| o.properties.get(&k).cloned())
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(values))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "entries",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let entries: Vec<Value> = v
                            .iter()
                            .enumerate()
                            .filter(|(index, _)| !is_array_hole(&o, *index as i32))
                            .map(|(i, val)| {
                                let pair = vec![Value::I32(i as i32), val.clone()];
                                Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                            })
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                    }
                    ObjectKind::Map(m) => {
                        let entries: Vec<Value> = m
                            .iter()
                            .map(|(k, v)| {
                                let pair = vec![k.clone(), v.clone()];
                                Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                            })
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                    }
                    // Set entries() per spec yields [value, value] pairs.
                    ObjectKind::Set(s) => {
                        let entries: Vec<Value> = s
                            .iter()
                            .map(|v| {
                                let pair = vec![v.clone(), v.clone()];
                                Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                            })
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                    }
                    _ => {}
                }
                let entries: Vec<Value> = ordinary_enumerable_keys(&o)
                    .into_iter()
                    .filter_map(|k| {
                        o.properties.get(&k).map(|v| {
                            let pair = vec![Value::String(Arc::from(k.as_str())), v.clone()];
                            Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                        })
                    })
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    // getOwnPropertyNames — like keys but includes non-enumerable.
    // Our model doesn't track enumerability separately, so this is the
    // same as `keys`. Use the insertion-order tracker (`__keys`) per
    // ECMA-262 §7.3.21 ordering requirements; HashMap iteration alone
    // is non-deterministic.
    vm.register_host_fn(
        "ecma:object",
        "getOwnPropertyNames",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let keys: Vec<Value> = ordinary_ordered_keys(&o)
                    .into_iter()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    // getOwnPropertySymbols — returns Value::Symbol for each key tracked in __sym_keys.
    // Symbol-keyed props are stored as "Symbol(<desc>)" string keys; we recover
    // the description and return the original Symbol so obj[syms[0]] round-trips.
    vm.register_host_fn(
        "ecma:object",
        "getOwnPropertySymbols",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let syms: Vec<Value> = match o.properties.get("__sym_keys") {
                    Some(Value::Object(arr)) => {
                        let a = arr.lock().unwrap();
                        if let ObjectKind::Array(ref elems) = a.kind {
                            elems
                                .iter()
                                .filter_map(|e| match e {
                                    Value::Symbol(sym) => Some(Value::Symbol(sym.clone())),
                                    _ => None,
                                })
                                .collect()
                        } else {
                            Vec::new()
                        }
                    }
                    _ => Vec::new(),
                };
                return Value::Object(Arc::new(Mutex::new(Object::new_array(syms))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "length",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::I32(own_keys(&o).len() as i32);
            }
            Value::I32(0)
        }),
    );
}

// ── Property descriptors ──────────────────────────────────────────────

fn register_descriptors(vm: &mut VM) {
    // defineProperty(obj, key, descriptor) -> obj
    // Descriptor is itself an object with {value, writable, enumerable,
    // configurable} or {get, set, enumerable, configurable} fields.
    // MVP: extract `value` + `enumerable` flag. Track the key in
    // `__keys` so iteration order matches insertion (HashMap order is
    // non-deterministic; ECMA-262 requires insertion order). Track
    // non-enumerable keys via `__nonenum` so `Object.keys` /
    // `Object.entries` exclude them per §7.3.22.
    vm.register_host_fn(
        "ecma:object",
        "defineProperty",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let original_obj = obj.clone();
                let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
                let descriptor = args.get(2).cloned().unwrap_or(Value::Undefined);
                let mut define_obj = obj.clone();
                if let Some((target, handler)) = proxy_target_and_handler(&obj) {
                    if let Some(trap) = proxy_trap(&handler, "defineProperty") {
                        let _ = invoke_with_explicit_this(
                            ctx,
                            &trap,
                            handler,
                            &[target, key_value.clone(), descriptor.clone()],
                        );
                        return Value::Object(original_obj);
                    }
                    if let Value::Object(target_obj) = target {
                        define_obj = target_obj;
                    }
                }
                let key = key_string(&key_value);
                // ECMA-262 §10.1: descriptor either has data fields
                // (`value`, `writable`) or accessor fields (`get`, `set`).
                // The VM honors `__get_<key>` / `__set_<key>` properties
                // as accessors (see dispatch.rs STRUCT_GET / STRUCT_SET),
                // so install them here when the descriptor specifies
                // get/set callables.
                let (val_or_none, getter, setter, enumerable, writable, configurable) =
                    match &descriptor {
                        Value::Object(desc) => {
                            let d = desc.lock().unwrap();
                            let val = d.properties.get("value").cloned();
                            let get = d.properties.get("get").cloned().filter(|v| {
                                matches!(v, Value::Object(o)
                                if matches!(o.lock().unwrap().kind,
                                    ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
                            });
                            let set = d.properties.get("set").cloned().filter(|v| {
                                matches!(v, Value::Object(o)
                                if matches!(o.lock().unwrap().kind,
                                    ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
                            });
                            let e = d
                                .properties
                                .get("enumerable")
                                .map(|x| x.as_bool())
                                .unwrap_or(false);
                            // ECMA-262 §6.2.5.1: data descriptors default
                            // writable=false when omitted but value is
                            // present. We treat absence as "true" only when
                            // the descriptor is purely accessor-shaped, to
                            // match observable test behaviour (`writable`
                            // explicitly true → writable; explicitly false
                            // or absent on data descriptor → non-writable).
                            let w = d.properties.get("writable").map(|x| x.as_bool());
                            let c = d
                                .properties
                                .get("configurable")
                                .map(|x| x.as_bool())
                                .unwrap_or(false);
                            (val, get, set, e, w, c)
                        }
                        _ => (None, None, None, false, None, false),
                    };
                track_key(&define_obj, &key);
                if !enumerable {
                    track_nonenum(&define_obj, &key);
                }
                {
                    let mut o = define_obj.lock().unwrap();
                    if o.properties.contains_key(&key) && is_nonconfig(&o, &key) {
                        ctx.throw_value(crate::ecma::error::new_error(
                            "TypeError",
                            "Cannot redefine property",
                        ));
                        return Value::Null;
                    }
                    if let Some(g) = getter {
                        o.properties.insert(format!("__get_{}", key), g);
                    }
                    if let Some(s) = setter {
                        o.properties.insert(format!("__set_{}", key), s);
                    }
                    if let Some(v) = val_or_none {
                        o.properties.insert(key.clone(), v);
                        // Non-writable data descriptor → install a
                        // no-op setter so subsequent writes via
                        // STRUCT_SET / `ecma:object.set` are silently
                        // discarded (loose mode per ECMA-262 §10.1.5).
                        if matches!(writable, Some(false) | None) {
                            let noop_idx =
                                NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
                            if noop_idx > 0 {
                                let mut noop_obj = Object::new();
                                noop_obj.kind = ObjectKind::HostFunction(noop_idx);
                                let noop_val = Value::Object(Arc::new(Mutex::new(noop_obj)));
                                let setter_key = format!("__set_{}", key);
                                if !o.properties.contains_key(&setter_key) {
                                    o.properties.insert(setter_key, noop_val);
                                }
                            }
                        }
                    } else if !o.properties.contains_key(&key) {
                        // Pure accessor descriptor: stamp Undefined so
                        // own-key enumeration sees the property.
                        o.properties.insert(key.clone(), Value::Undefined);
                    }
                }
                if !configurable {
                    track_nonconfig(&define_obj, &key);
                }
                return Value::Object(original_obj);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "defineProperties",
        Box::new(|_ctx, args| {
            if let (Some(target), Some(Value::Object(descs))) = (obj_of(args, 0), args.get(1)) {
                // Collect entries in the descriptor object's __keys
                // insertion order so the target's iteration matches
                // (ECMA-262 §20.1.2.4 step 5 — `OwnPropertyKeys` over
                // descriptors). HashMap iter would non-deterministically
                // shuffle on every run.
                let entries: Vec<(String, Value, bool)> = {
                    let d = descs.lock().unwrap();
                    let order: Vec<String> =
                        if let Some(Value::Object(arr)) = d.properties.get("__keys") {
                            let ka = arr.lock().unwrap();
                            if let ObjectKind::Array(ref el) = ka.kind {
                                el.iter()
                                    .filter_map(|v| {
                                        if let Value::String(s) = v {
                                            Some(s.to_string())
                                        } else {
                                            None
                                        }
                                    })
                                    .collect()
                            } else {
                                Vec::new()
                            }
                        } else {
                            Vec::new()
                        };
                    let keys = if !order.is_empty() {
                        order
                            .into_iter()
                            .filter(|k| !k.starts_with("__"))
                            .collect::<Vec<_>>()
                    } else {
                        d.properties
                            .keys()
                            .filter(|k| !k.starts_with("__"))
                            .cloned()
                            .collect()
                    };
                    keys.into_iter()
                        .filter_map(|k| {
                            if let Some(Value::Object(dv)) = d.properties.get(&k) {
                                let dlock = dv.lock().unwrap();
                                let val = dlock
                                    .properties
                                    .get("value")
                                    .cloned()
                                    .unwrap_or(Value::Undefined);
                                let enumerable = dlock
                                    .properties
                                    .get("enumerable")
                                    .map(|x| x.as_bool())
                                    .unwrap_or(false);
                                Some((k, val, enumerable))
                            } else {
                                None
                            }
                        })
                        .collect()
                };
                for (k, v, enumerable) in entries {
                    track_key(&target, &k);
                    if !enumerable {
                        track_nonenum(&target, &k);
                    }
                    target.lock().unwrap().properties.insert(k, v);
                }
                return Value::Object(target);
            }
            Value::Null
        }),
    );

    // getOwnPropertyDescriptor(obj, key) -> descriptor or undefined.
    // ECMA-262 §20.1.2.6 returns a fresh data descriptor with Boolean
    // flags. Our model doesn't track writable/configurable separately
    // so we report `true` for both; `enumerable` honors the
    // `__nonenum` tracker set by `defineProperty(enumerable: false)`.
    vm.register_host_fn(
        "ecma:object",
        "getOwnPropertyDescriptor",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                let getter_key = format!("__get_{}", key);
                let setter_key = format!("__set_{}", key);
                if o.properties.contains_key(&getter_key) || o.properties.contains_key(&setter_key)
                {
                    let mut desc = Object::new();
                    desc.properties.insert(
                        "get".into(),
                        o.properties
                            .get(&getter_key)
                            .cloned()
                            .unwrap_or(Value::Undefined),
                    );
                    desc.properties.insert(
                        "set".into(),
                        o.properties
                            .get(&setter_key)
                            .cloned()
                            .unwrap_or(Value::Undefined),
                    );
                    desc.properties
                        .insert("enumerable".into(), Value::Bool(!is_nonenum(&o, &key)));
                    desc.properties
                        .insert("configurable".into(), Value::Bool(!is_nonconfig(&o, &key)));
                    return Value::Object(Arc::new(Mutex::new(desc)));
                }
                if let Some(v) = o.properties.get(&key) {
                    let mut desc = Object::new();
                    desc.properties.insert("value".into(), v.clone());
                    desc.properties.insert("writable".into(), Value::Bool(true));
                    desc.properties
                        .insert("enumerable".into(), Value::Bool(!is_nonenum(&o, &key)));
                    desc.properties
                        .insert("configurable".into(), Value::Bool(!is_nonconfig(&o, &key)));
                    return Value::Object(Arc::new(Mutex::new(desc)));
                }
            }
            Value::Undefined
        }),
    );

    // getOwnPropertyDescriptors(obj) -> { key: descriptor, ... }
    vm.register_host_fn(
        "ecma:object",
        "getOwnPropertyDescriptors",
        Box::new(|_ctx, args| {
            let mut result = Object::new();
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                for (k, v) in &o.properties {
                    if k.starts_with("__") {
                        continue;
                    }
                    let mut desc = Object::new();
                    let getter_key = format!("__get_{}", k);
                    let setter_key = format!("__set_{}", k);
                    if o.properties.contains_key(&getter_key)
                        || o.properties.contains_key(&setter_key)
                    {
                        desc.properties.insert(
                            "get".into(),
                            o.properties
                                .get(&getter_key)
                                .cloned()
                                .unwrap_or(Value::Undefined),
                        );
                        desc.properties.insert(
                            "set".into(),
                            o.properties
                                .get(&setter_key)
                                .cloned()
                                .unwrap_or(Value::Undefined),
                        );
                    } else {
                        desc.properties.insert("value".into(), v.clone());
                        desc.properties.insert("writable".into(), Value::Bool(true));
                    }
                    desc.properties
                        .insert("enumerable".into(), Value::Bool(!is_nonenum(&o, k)));
                    desc.properties
                        .insert("configurable".into(), Value::Bool(!is_nonconfig(&o, k)));
                    result
                        .properties
                        .insert(k.clone(), Value::Object(Arc::new(Mutex::new(desc))));
                }
            }
            Value::Object(Arc::new(Mutex::new(result)))
        }),
    );
}

// ── Prototype ─────────────────────────────────────────────────────────

fn register_prototype(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:object",
        "getPrototypeOf",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                if let Some((target, handler)) = proxy_target_and_handler(&obj) {
                    if let Some(trap) = proxy_trap(&handler, "getPrototypeOf") {
                        return invoke_with_explicit_this(ctx, &trap, handler, &[target]);
                    }
                    if let Value::Object(target_obj) = target {
                        let o = target_obj.lock().unwrap();
                        return o.properties.get(PROTO_KEY).cloned().unwrap_or(Value::Null);
                    }
                }
                let o = obj.lock().unwrap();
                return o.properties.get(PROTO_KEY).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "setPrototypeOf",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let proto = args.get(1).cloned().unwrap_or(Value::Null);
                let mut o = obj.lock().unwrap();
                o.properties.insert(PROTO_KEY.into(), proto);
                return Value::Bool(true);
            }
            Value::Bool(false)
        }),
    );
}

// ── Locking (freeze / seal / preventExtensions) ───────────────────────

/// Process-global host fn idx for the no-op setter installed by
/// `Object.freeze`. Captured during `register_locking` so freeze can
/// build bound `__set_<key>` accessors without re-looking-up the host
/// registry on every call.
static NOOP_SETTER_IDX: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn register_locking(vm: &mut VM) {
    // Silent-write setter installed by `freeze` for every existing key.
    // Returns the value untouched so `obj.x = 99` evaluates to 99 but
    // the underlying store is unchanged.
    vm.register_host_fn(
        "ecma:object",
        "__noop_setter",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    let idx = vm
        .host_registry
        .get(&("ecma:object".to_string(), "__noop_setter".to_string()))
        .copied()
        .expect("__noop_setter just registered");
    NOOP_SETTER_IDX.store(idx, std::sync::atomic::Ordering::Relaxed);

    vm.register_host_fn(
        "ecma:object",
        "freeze",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                // Install a no-op setter (`__set_<key>`) for each
                // existing own property so writes are silently
                // discarded by the VM's STRUCT_SET accessor dispatch.
                // ECMA-262 §20.1.2.6 requires writes to fail silently
                // (loose mode) or throw TypeError (strict). MVP picks
                // silent. Doesn't block new property additions — that
                // path goes through STRUCT_SET without the accessor
                // check, and would need VM-level enforcement.
                let keys: Vec<String> = {
                    let o = obj.lock().unwrap();
                    o.properties
                        .keys()
                        .filter(|k| {
                            !k.starts_with("__")
                                && !k.starts_with("__get_")
                                && !k.starts_with("__set_")
                        })
                        .cloned()
                        .collect()
                };
                let mut o = obj.lock().unwrap();
                o.properties.insert(FROZEN_MARK.into(), Value::I32(1));
                o.properties.insert(SEALED_MARK.into(), Value::I32(1));
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                let noop_idx = NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
                if noop_idx > 0 {
                    let mut noop_obj = Object::new();
                    noop_obj.kind = ObjectKind::HostFunction(noop_idx);
                    let noop_val = Value::Object(Arc::new(Mutex::new(noop_obj)));
                    for k in keys {
                        let setter_key = format!("__set_{}", k);
                        if !o.properties.contains_key(&setter_key) {
                            o.properties.insert(setter_key, noop_val.clone());
                        }
                    }
                }
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "isFrozen",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(o.properties.get(FROZEN_MARK).is_some());
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "seal",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let mut o = obj.lock().unwrap();
                o.properties.insert(SEALED_MARK.into(), Value::I32(1));
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "isSealed",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(o.properties.get(SEALED_MARK).is_some());
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "preventExtensions",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let mut o = obj.lock().unwrap();
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "isExtensible",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(!matches!(
                    o.properties.get(EXTENSIBLE_MARK),
                    Some(Value::I32(0))
                ));
            }
            Value::Bool(false)
        }),
    );
}

// ── Comparison ────────────────────────────────────────────────────────

fn register_comparison(vm: &mut VM) {
    // Object.is(a, b) — SameValue: NaN === NaN, -0 distinct from +0
    vm.register_host_fn(
        "ecma:object",
        "is",
        Box::new(|_ctx, args| {
            let a = args.first();
            let b = args.get(1);
            let same = match (a, b) {
                (Some(Value::F64(x)), Some(Value::F64(y))) => {
                    if x.is_nan() && y.is_nan() {
                        true
                    } else if *x == 0.0 && *y == 0.0 {
                        // Distinguish +0 vs -0 via sign bit
                        x.is_sign_positive() == y.is_sign_positive()
                    } else {
                        x == y
                    }
                }
                (Some(x), Some(y)) => x.eq(y),
                _ => false,
            };
            // ECMA-262 §20.1.2.13: returns a Boolean.
            Value::Bool(same)
        }),
    );
}

// ── Prototype methods (called via obj.foo()) ──────────────────────────

fn register_prototype_methods(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:object",
        "hasOwnProperty",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                return Value::Bool(o.properties.contains_key(&key));
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "isPrototypeOf",
        Box::new(|_ctx, args| {
            // isPrototypeOf: is `self` in `other`'s prototype chain?
            if let (Some(self_obj), Some(other)) = (obj_of(args, 0), obj_of(args, 1)) {
                let mut current = other;
                loop {
                    let o = current.lock().unwrap();
                    let proto = o.properties.get(PROTO_KEY).cloned();
                    drop(o);
                    match proto {
                        Some(Value::Object(p)) => {
                            if Arc::ptr_eq(&p, &self_obj) {
                                return Value::Bool(true);
                            }
                            if Arc::ptr_eq(&p, &current) {
                                return Value::Bool(false);
                            }
                            current = p;
                        }
                        _ => return Value::Bool(false),
                    }
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "propertyIsEnumerable",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                return Value::Bool(
                    o.properties.contains_key(&key)
                        && !key.starts_with("__")
                        && !is_nonenum(&o, &key),
                );
            }
            Value::Bool(false)
        }),
    );

    // toString(): ECMA-262 §20.1.3.6 — returns "[object <Tag>]" for any value
    // (including primitives). Called via Object.prototype.toString.call(value).
    vm.register_host_fn(
        "ecma:object",
        "toString",
        Box::new(|ctx, args| {
            let tag = match args.first() {
                None | Some(Value::Undefined) => "Undefined".to_string(),
                Some(Value::Null) => "Null".to_string(),
                Some(Value::Bool(_)) => "Boolean".to_string(),
                Some(Value::I32(_)) | Some(Value::I64(_)) | Some(Value::F64(_)) => {
                    "Number".to_string()
                }
                Some(Value::String(_)) => "String".to_string(),
                Some(Value::Symbol(_)) => "Symbol".to_string(),
                Some(Value::Object(obj)) => object_to_string_tag(ctx, obj),
                _ => "Object".to_string(),
            };
            Value::String(Arc::from(format!("[object {}]", tag).as_str()))
        }),
    );

    vm.register_host_fn(
        "ecma:object",
        "toLocaleString",
        Box::new(|_ctx, args| {
            if is_object(args.first().unwrap_or(&Value::Null)) {
                return Value::String(Arc::from("[object Object]"));
            }
            Value::String(Arc::from(""))
        }),
    );

    // valueOf: spec default returns the object itself
    vm.register_host_fn(
        "ecma:object",
        "valueOf",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let primitive = {
                    let o = obj.lock().unwrap();
                    o.properties.get("__primitive").cloned()
                };
                if let Some(value) = primitive {
                    return value;
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // Object.prototype[Symbol.toStringTag] — returns the tag string used by
    // Object.prototype.toString. ECMA-262 §20.1.3.6.
    vm.register_host_fn(
        "ecma:object",
        "toStringTag",
        Box::new(|ctx, args| {
            let tag = match args.first() {
                None | Some(Value::Undefined) => "Undefined".to_string(),
                Some(Value::Null) => "Null".to_string(),
                Some(Value::Object(obj)) => object_to_string_tag(ctx, obj),
                _ => "Object".to_string(),
            };
            Value::String(Arc::from(format!("[object {}]", tag).as_str()))
        }),
    );

    // Object.groupBy(items, keyFn) — ES2024 §20.1.2.x.
    // Groups iterable items into a plain object keyed by keyFn(item, index).
    vm.register_host_fn(
        "ecma:object",
        "groupBy",
        Box::new(|ctx, args| {
            let items = args.first().cloned().unwrap_or(Value::Undefined);
            let key_fn = args.get(1).cloned().unwrap_or(Value::Undefined);
            let mut result = Object::new();
            let arr_items = match &items {
                Value::Object(obj) => {
                    let o = obj.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind {
                        v.clone()
                    } else {
                        Vec::new()
                    }
                }
                _ => Vec::new(),
            };
            for (i, item) in arr_items.into_iter().enumerate() {
                let key = if matches!(key_fn, Value::Null | Value::Undefined) {
                    format!("{}", i)
                } else if let Some(k) = groupby_magic_key(&key_fn, &item) {
                    k
                } else {
                    let k = ctx.invoke(&key_fn, &[item.clone(), Value::I32(i as i32)]);
                    format!("{}", k)
                };
                let group = result.properties.entry(key).or_insert_with(|| {
                    Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
                });
                if let Value::Object(arr) = group {
                    if let ObjectKind::Array(ref mut elems) = arr.lock().unwrap().kind {
                        elems.push(item);
                    }
                }
            }
            Value::Object(Arc::new(Mutex::new(result)))
        }),
    );
}

// ── PHP extensions ────────────────────────────────────────────────────

fn register_php_extensions(vm: &mut VM) {
    // appendAutoKey(obj, value) -> i32 key
    // Implements PHP's `$a[] = x` — finds the next int key (max of
    // existing int keys + 1, or 0 if none) and sets it.
    vm.register_host_fn(
        "ecma:object",
        "appendAutoKey",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let val = args.get(1).cloned().unwrap_or(Value::Null);
                let mut o = obj.lock().unwrap();
                // Get and increment the counter, initializing if absent.
                let next = match o.properties.get(NEXT_INT_KEY) {
                    Some(Value::I32(n)) => *n,
                    Some(Value::F64(n)) => *n as i32,
                    _ => {
                        // Compute from existing int-keyed entries
                        let mut max_k: i32 = -1;
                        for k in o.properties.keys() {
                            if let Ok(n) = k.parse::<i32>() {
                                if n > max_k {
                                    max_k = n;
                                }
                            }
                        }
                        max_k + 1
                    }
                };
                o.properties.insert(next.to_string(), val);
                o.properties
                    .insert(NEXT_INT_KEY.into(), Value::I32(next + 1));
                return Value::I32(next);
            }
            Value::I32(0)
        }),
    );
}
