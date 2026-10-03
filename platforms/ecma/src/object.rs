//! # `ecma:object` host handlers
//!
//! Native Rust impls of `Object.*` statics and `Object.prototype.*` per
//! ECMA-262 §20.1, satisfying the imports declared in
//! `crates/vybe_runtime/src/wasm/js_object_builtins.rs`.
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

use crate::function::invoke_with_explicit_this;
use crate::typedarray::typed_array_values_snapshot;
use indexmap::IndexMap;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, Param, VM, ValType};

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

/// Declare an `ecma:object` function — same closure, plus the signature.
///
/// No resource binding: a JS object is an ordinary reference, not a handle the
/// host mints and drops, so `own`/`borrow` would claim a lifetime that does not
/// exist here.
fn object_fn(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new(
            "ecma:object",
            name,
            crate::perf::wrap("ecma:object", name, call),
        )
        .with_sig(FuncSig {
            name: name.to_string(),
            params: Param::unnamed_list(params),
            results,
        }),
    );
}

/// Any JS value — the Component Model has no `any`-of-object type and these
/// handlers accept whatever `obj_of` can unwrap.
fn obj_t() -> ValType {
    ValType::Any
}

/// `Object.isFrozen(o)`, `Object.values(o)`, … — one operand, no optional tail.
fn object_unary(
    vm: &mut VM,
    name: &str,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    object_fn(vm, name, vec![obj_t()], results, call);
}

/// Magic property name used to mark an object as frozen / sealed /
/// non-extensible.
const FROZEN_MARK: &str = "__vybe_frozen";
const SEALED_MARK: &str = "__vybe_sealed";
const EXTENSIBLE_MARK: &str = "__vybe_extensible"; // absence means extensible
const PROTO_KEY: &str = "__proto__";
/// The custom-descriptor slot a GC struct carries (set by the VM in
/// `dispatch.rs`). Named here rather than imported because the runtime keeps
/// it private.
const DESCRIPTOR_SLOT: &str = "__descriptor";
const NULL_PROTO_MARK: &str = "__vybe_null_proto";
const ACCESSOR_SETTER_ACTIVE_MARK: &str = "__vybe_accessor_setter_active";
const PROXY_TARGET_KEY: &str = "__vybe_proxy_target";
const PROXY_HANDLER_KEY: &str = "__vybe_proxy_handler";
/// PHP-array next-int-key tracker. Used by `appendAutoKey`.
const NEXT_INT_KEY: &str = "__vybe_next_int_key";

static OBJECT_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

pub fn shared_object_prototype() -> Value {
    let proto = OBJECT_PROTOTYPE.get_or_init(|| {
        let mut obj = Object::new();
        obj.properties.reserve(1);
        obj.properties.insert(PROTO_KEY.into(), Value::Null);
        vybe_runtime::heap::alloc(obj)
    });
    Value::Object(proto.clone())
}

pub fn new_ordinary_object_with_proto() -> Value {
    let mut obj = Object::new();
    obj.properties.reserve(1);
    obj.properties
        .insert(PROTO_KEY.into(), shared_object_prototype());
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// Resolve a value's `[[Prototype]]` — ECMA-262 OrdinaryGetPrototypeOf and
/// the exotic-object / primitive-wrapper variants.
///
/// The VM creates arrays and ordinary objects WITHOUT a materialized
/// `__proto__` (it stays WASM-pure — a bare GC struct has no JS identity).
/// The canonical [[Prototype]] is therefore resolved here, by kind, so
/// every reflective op (`getPrototypeOf`, `instanceof`, `constructor`,
/// inherited-property lookup) sees the one shared prototype object.
///
/// An *explicitly present* `__proto__` always wins — including `null` from
/// `Object.create(null)` / `Object.setPrototypeOf(o, null)` — so the
/// fallback only fires when the slot was never written.
pub fn js_prototype_of(value: &Value) -> Value {
    match value {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            if o.properties.contains_key(NULL_PROTO_MARK) {
                return Value::Null;
            }
            if let Some(explicit) = o.properties.get(PROTO_KEY) {
                return explicit.clone();
            }
            implicit_object_prototype(obj, &o)
        }
        Value::String(_) => crate::string::shared_string_prototype(),
        Value::Bool(_) => crate::boolean::shared_boolean_prototype(),
        Value::F64(_) | Value::I32(_) | Value::I64(_) => crate::number::shared_number_prototype(),
        _ => Value::Null,
    }
}

pub(crate) fn implicit_object_prototype(obj: &Arc<Mutex<Object>>, o: &Object) -> Value {
    if o.properties.contains_key(NULL_PROTO_MARK) {
        return Value::Null;
    }
    if let Some(Value::Object(desc)) = o.properties.get(DESCRIPTOR_SLOT) {
        if !Arc::ptr_eq(obj, desc) {
            if let Ok(d) = desc.try_lock() {
                if let Some(p @ Value::Object(_)) = d.fields.first() {
                    return p.clone();
                }
            }
        }
    }
    match &o.kind {
        ObjectKind::Array(_) => crate::array::shared_array_prototype(),
        _ => shared_object_prototype(),
    }
}

fn object_prototype_of_locked(obj: &Arc<Mutex<Object>>) -> Value {
    let o = obj.lock().unwrap();
    if o.properties.contains_key(NULL_PROTO_MARK) {
        return Value::Null;
    }
    if let Some(explicit) = o.properties.get(PROTO_KEY) {
        return explicit.clone();
    }
    implicit_object_prototype(obj, &o)
}

fn fast_prototype_of(value: &Value) -> Value {
    match value {
        Value::Object(obj) => object_prototype_of_locked(obj),
        _ => js_prototype_of(value),
    }
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

fn to_object_for_object_static(
    ctx: &mut HostContext,
    value: &Value,
    message: &str,
) -> Option<Value> {
    match value {
        Value::Null | Value::Undefined => {
            ctx.throw_value(crate::error::new_error(ctx, "TypeError", message));
            None
        }
        value @ Value::Object(_) => Some(value.clone()),
        Value::Bool(value) => Some(crate::boolean::boxed_boolean(*value)),
        Value::String(text) => Some(crate::string::boxed_string(text.clone())),
        value @ Value::F64(_) | value @ Value::I32(_) | value @ Value::I64(_) => {
            Some(crate::number::boxed_number(value.clone()))
        }
        Value::Symbol(desc) => {
            let mut obj = Object::new();
            obj.properties.reserve(3);
            obj.properties
                .insert("__type".into(), crate::keys::string_value("Symbol"));
            obj.properties
                .insert("__primitive".into(), Value::Symbol(desc.clone()));
            obj.properties
                .insert(PROTO_KEY.into(), shared_object_prototype());
            Some(Value::Object(vybe_runtime::heap::alloc(obj)))
        }
        Value::BigInt(value) => {
            let mut obj = Object::new();
            obj.properties.reserve(3);
            obj.properties
                .insert("__type".into(), crate::keys::string_value("BigInt"));
            obj.properties
                .insert("__primitive".into(), Value::BigInt(value.clone()));
            obj.properties
                .insert(PROTO_KEY.into(), shared_object_prototype());
            Some(Value::Object(vybe_runtime::heap::alloc(obj)))
        }
        _ => Some(new_ordinary_object_with_proto()),
    }
}

fn key_string(v: &Value) -> String {
    crate::keys::property_key_string(v)
}

fn with_property_key<R>(value: &Value, f: impl FnOnce(&str) -> R) -> R {
    crate::keys::with_property_key(value, f)
}

fn getter_property_key(key: &str) -> String {
    crate::keys::getter_property_key(key)
}

fn setter_property_key(key: &str) -> String {
    crate::keys::setter_property_key(key)
}

#[inline]
fn has_getter_key(o: &Object, key: &str) -> bool {
    crate::keys::with_getter_property_key(key, |getter_key| o.properties.contains_key(getter_key))
}

fn has_any_getter(o: &Object) -> bool {
    o.properties.keys().any(|key| key.starts_with("__get_"))
}

fn index_key_value(index: usize) -> Value {
    crate::keys::with_index_key(index, crate::keys::string_value)
}

fn array_elements(value: &Value) -> Vec<Value> {
    if let Value::Object(obj) = value {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(values) = &o.kind {
            return values.clone();
        }
    }
    Vec::new()
}

fn descriptor_bool(desc: &Value, key: &str, default: bool) -> bool {
    if let Value::Object(obj) = desc {
        return obj
            .lock()
            .unwrap()
            .properties
            .get(key)
            .map(|v| v.as_bool())
            .unwrap_or(default);
    }
    default
}

fn array_index_key(key: &str) -> Option<u32> {
    crate::keys::canonical_array_index_key(key)
}

fn sort_array_indices_first(keys: &mut Vec<String>) {
    if !keys.iter().any(|key| array_index_key(key).is_some()) {
        return;
    }
    let mut indexed: Vec<(u32, String)> = Vec::with_capacity(keys.len());
    let mut rest = Vec::with_capacity(keys.len());
    for key in keys.drain(..) {
        if let Some(index) = array_index_key(&key) {
            indexed.push((index, key));
        } else {
            rest.push(key);
        }
    }
    indexed.sort_by_key(|(index, _)| *index);
    keys.extend(indexed.into_iter().map(|(_, key)| key));
    keys.extend(rest);
}

fn enumerable_assign_keys(source: &Arc<Mutex<Object>>) -> Vec<(String, Option<Value>)> {
    let src = source.lock().unwrap();
    let mut out: Vec<(String, Option<Value>)> = Vec::with_capacity(src.properties.len());
    let symbol_storage_keys: Option<std::collections::HashSet<String>> =
        src.properties.get("__sym_keys").and_then(|v| {
            if let Value::Object(a) = v {
                let lock = a.lock().unwrap();
                if let ObjectKind::Array(ref elems) = lock.kind {
                    let mut keys = std::collections::HashSet::with_capacity(elems.len());
                    for elem in elems {
                        keys.insert(key_string(elem));
                    }
                    Some(keys)
                } else {
                    None
                }
            } else {
                None
            }
        });
    match &src.kind {
        ObjectKind::Array(values) => {
            out.reserve(values.len());
            let holes = array_hole_set(&src);
            for index in 0..values.len() {
                if !cached_array_hole_contains(&holes, index) {
                    out.push((
                        crate::keys::with_index_key(index, |key| key.to_owned()),
                        None,
                    ));
                }
            }
            for key in descriptor_own_keys(&src)
                .into_iter()
                .filter(|key| !is_nonenum(&src, key))
                .filter(
                    |key| !matches!(symbol_storage_keys.as_ref(), Some(keys) if keys.contains(key)),
                )
            {
                if array_index_key(&key).is_none() {
                    out.push((key, None));
                }
            }
        }
        ObjectKind::TypedArray(ta) => {
            let live = crate::typedarray::ta_live_length(ta);
            out.reserve(live);
            for index in 0..live {
                out.push((
                    crate::keys::with_index_key(index, |key| key.to_owned()),
                    None,
                ));
            }
            for key in descriptor_own_keys(&src)
                .into_iter()
                .filter(|key| !is_nonenum(&src, key))
                .filter(
                    |key| !matches!(symbol_storage_keys.as_ref(), Some(keys) if keys.contains(key)),
                )
            {
                if array_index_key(&key).is_none() {
                    out.push((key, None));
                }
            }
        }
        ObjectKind::Map(_) | ObjectKind::Set(_) => {
            for key in descriptor_own_keys(&src)
                .into_iter()
                .filter(|key| !is_nonenum(&src, key))
                .filter(
                    |key| !matches!(symbol_storage_keys.as_ref(), Some(keys) if keys.contains(key)),
                )
            {
                out.push((key, None));
            }
        }
        _ => {
            for key in descriptor_own_keys(&src)
                .into_iter()
                .filter(|key| !is_nonenum(&src, key))
                .filter(
                    |key| !matches!(symbol_storage_keys.as_ref(), Some(keys) if keys.contains(key)),
                )
            {
                out.push((key, None));
            }
        }
    }
    if let Some(Value::Object(sym_arr)) = src.properties.get("__sym_keys") {
        let syms = sym_arr.lock().unwrap();
        if let ObjectKind::Array(ref elems) = syms.kind {
            for key in elems {
                let storage_key = key_string(key);
                if !is_nonenum(&src, &storage_key) && src.properties.contains_key(&storage_key) {
                    out.push((storage_key, Some(key.clone())));
                }
            }
        }
    }
    out
}

fn assign_source_get(ctx: &mut HostContext, source: &Arc<Mutex<Object>>, key: &str) -> Value {
    {
        let src = source.lock().unwrap();
        if let ObjectKind::Array(values) = &src.kind {
            if let Some(index) = array_index_key(key) {
                if !is_array_hole(&src, index as i32) {
                    if let Some(value) = values.get(index as usize) {
                        return value.clone();
                    }
                }
            }
        }
        if let ObjectKind::TypedArray(ta) = &src.kind {
            if let Some(index) = array_index_key(key) {
                if (index as usize) < crate::typedarray::ta_live_length(ta) {
                    return crate::typedarray::read_element(ta, index as usize);
                }
            }
        }
        if let Some(value) = src.properties.get(key) {
            if !has_getter_key(&src, key) {
                return value.clone();
            }
        }
    }
    proto_walk_invoke_getter(ctx, source, key).unwrap_or(Value::Undefined)
}

fn assign_strict_set(
    ctx: &mut HostContext,
    target: &Arc<Mutex<Object>>,
    key: &str,
    value: Value,
    symbol_key: Option<Value>,
) -> bool {
    {
        let tgt = target.lock().unwrap();
        if is_not_extensible(&tgt) {
            let exists = if tgt.properties.contains_key(key) {
                true
            } else {
                crate::keys::with_getter_property_key(key, |getter_key| {
                    tgt.properties.contains_key(getter_key)
                }) || crate::keys::with_setter_property_key(key, |setter_key| {
                    tgt.properties.contains_key(setter_key)
                })
            };
            if !exists {
                drop(tgt);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot add property, object is not extensible",
                ));
                return false;
            }
        }
    }

    {
        let mut tgt = target.lock().unwrap();
        if matches!(&tgt.kind, ObjectKind::Array(_)) && (key == "length" || key == "__len__") {
            crate::array::apply_js_array_length(ctx, &mut tgt, &value);
            return true;
        }
    }

    let setter = crate::keys::with_setter_property_key(key, |setter_key| {
        {
            let tgt = target.lock().unwrap();
            tgt.properties.get(setter_key).cloned()
        }
        .or_else(|| proto_walk_get(target, setter_key))
    });

    if let Some(setter_val) = setter {
        if let Value::Object(setter_obj) = &setter_val {
            let is_noop_setter = {
                let so = setter_obj.lock().unwrap();
                matches!(
                    so.kind,
                    ObjectKind::HostFunction(idx)
                        if idx == NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed)
                )
            };
            if is_noop_setter {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot assign to read only property",
                ));
                return false;
            }
            // ⛔ ONE RECEIVER, PASSED ONCE. The arity guess below decided
            // whether to pass the receiver, but under `ReceiverAbi::Parameter`
            // the invoke ALSO supplies one, so an arity-2 setter got it twice:
            // `this` took the supplied receiver and the VALUE parameter took
            // the target. Measured: `Object.assign(t, {a: 5})` left `t._got`
            // undefined. `invoke_with_receiver` is correct under BOTH bindings,
            // so the arity question disappears rather than being answered.
            ctx.invoke_with_receiver(&setter_val, Value::Object(target.clone()), &[value]);
            return true;
        }
    }

    let has_getter_without_setter = crate::keys::with_getter_property_key(key, |getter_key| {
        let own_getter = {
            let tgt = target.lock().unwrap();
            tgt.properties.contains_key(getter_key)
        };
        own_getter || proto_walk_has(target, getter_key)
    });
    if has_getter_without_setter {
        ctx.throw_value(crate::error::new_error(
            ctx,
            "TypeError",
            "Cannot set property which has only a getter",
        ));
        return false;
    }

    if let Some(sym) = symbol_key {
        let sym_keys = {
            let mut tgt = target.lock().unwrap();
            if tgt.properties.get(FROZEN_MARK).is_some() {
                drop(tgt);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot assign to read only property of frozen object",
                ));
                return false;
            }
            let sym_keys = match tgt.properties.get("__sym_keys") {
                Some(Value::Object(existing)) => existing.clone(),
                _ => {
                    let created = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
                    tgt.properties
                        .insert("__sym_keys".into(), Value::Object(created.clone()));
                    created
                }
            };
            tgt.properties.insert(key.to_string(), value);
            sym_keys
        };
        let mut sym_keys_guard = sym_keys.lock().unwrap();
        if let ObjectKind::Array(ref mut elems) = sym_keys_guard.kind {
            if !elems.iter().any(|existing| existing == &sym) {
                elems.push(sym);
            }
        }
    } else {
        let keys_to_append = {
            let mut tgt = target.lock().unwrap();
            if tgt.properties.get(FROZEN_MARK).is_some() {
                drop(tgt);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot assign to read only property of frozen object",
                ));
                return false;
            }
            let keys_to_append = prepare_track_key_in(&mut tgt, key);
            tgt.properties.insert(key.to_string(), value);
            keys_to_append
        };
        if let Some(keys_arc) = keys_to_append {
            append_tracked_key(keys_arc, key);
        }
    }
    true
}

fn has_own_property_key(value: &Value, key_raw: &Value) -> Option<bool> {
    match value {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(values) => {
                    if matches!(key_raw, Value::String(text) if text.as_ref() == "length") {
                        return Some(true);
                    }
                    if let Some(index) = crate::keys::canonical_array_index_value(key_raw) {
                        return Some(
                            (index as usize) < values.len() && !is_array_hole(&o, index as i32),
                        );
                    }
                }
                ObjectKind::TypedArray(ta) => {
                    if let Some(index) = crate::keys::canonical_array_index_value(key_raw) {
                        return Some((index as usize) < crate::typedarray::ta_live_length(ta));
                    }
                }
                ObjectKind::Map(_) | ObjectKind::Set(_) => {
                    if matches!(key_raw, Value::String(text) if text.as_ref() == "size") {
                        return Some(false);
                    }
                }
                _ => {}
            }
            with_property_key(key_raw, |key| {
                let visible = !key.starts_with("__")
                    || matches!(key_raw, Value::String(text) if text.as_ref() == "__proto__");
                let present = o.properties.contains_key(key)
                    || crate::keys::with_getter_property_key(key, |getter_key| {
                        o.properties.contains_key(getter_key)
                    })
                    || crate::keys::with_setter_property_key(key, |setter_key| {
                        o.properties.contains_key(setter_key)
                    });
                Some(visible && present)
            })
        }
        Value::String(text) => {
            if matches!(key_raw, Value::String(key) if key.as_ref() == "length") {
                return Some(true);
            }
            if let Some(index) = crate::keys::canonical_array_index_value(key_raw) {
                let len = if text.is_ascii() {
                    text.len()
                } else {
                    text.chars().count()
                };
                return Some((index as usize) < len);
            }
            Some(false)
        }
        Value::Null | Value::Undefined => None,
        _ => Some(false),
    }
}

pub fn proxy_target_and_handler(obj: &Arc<Mutex<Object>>) -> Option<(Value, Value)> {
    let o = obj.lock().unwrap();
    let target = o.properties.get(PROXY_TARGET_KEY).cloned()?;
    let handler = o.properties.get(PROXY_HANDLER_KEY).cloned()?;
    Some((target, handler))
}

pub fn proxy_trap(handler: &Value, name: &str) -> Option<Value> {
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
pub fn track_key(obj: &Arc<Mutex<Object>>, key: &str) {
    let mut o = obj.lock().unwrap();
    let already = o.properties.contains_key(key);
    if already {
        return;
    }
    let keys_arc = match o.properties.get("__keys") {
        Some(Value::Object(arr)) => arr.clone(),
        _ => {
            let arc = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            o.properties
                .insert("__keys".into(), Value::Object(arc.clone()));
            arc
        }
    };
    drop(o);
    let mut k = keys_arc.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = k.kind {
        elems.push(crate::keys::string_value(key));
    }
}

pub fn prepare_track_key_in(o: &mut Object, key: &str) -> Option<Arc<Mutex<Object>>> {
    if o.properties.contains_key(key) {
        return None;
    }
    let keys_arc = match o.properties.get("__keys") {
        Some(Value::Object(arr)) => arr.clone(),
        _ => {
            let arc = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            o.properties
                .insert("__keys".into(), Value::Object(arc.clone()));
            arc
        }
    };
    Some(keys_arc)
}

pub fn append_tracked_key(keys_arc: Arc<Mutex<Object>>, key: &str) {
    let mut k = keys_arc.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = k.kind {
        elems.push(crate::keys::string_value(key));
    }
}

/// Mark `key` as non-enumerable on `obj` (lazy-initializes the
/// `__nonenum` set). `Object.keys` / `Object.entries` filter against
/// this set so defineProperty with `enumerable: false` is honoured.
pub fn track_nonenum(obj: &Arc<Mutex<Object>>, key: &str) {
    let mut o = obj.lock().unwrap();
    let arr = match o.properties.get("__nonenum") {
        Some(Value::Object(a)) => a.clone(),
        _ => {
            let a = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            o.properties
                .insert("__nonenum".into(), Value::Object(a.clone()));
            a
        }
    };
    drop(o);
    let mut a = arr.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = a.kind {
        if !elems
            .iter()
            .any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
        {
            elems.push(crate::keys::string_value(key));
        }
    }
}

/// [`track_nonenum`] for a caller that already holds the object BORROWED
/// rather than by `Arc` — the size-cache and iterator-method writers run
/// inside `&mut Object` and cannot re-lock. Same `__nonenum` set, so both
/// forms answer one question.
pub fn track_nonenum_in(o: &mut Object, key: &str) {
    let arr = match o.properties.get("__nonenum") {
        Some(Value::Object(a)) => a.clone(),
        _ => {
            let a = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            o.properties
                .insert("__nonenum".into(), Value::Object(a.clone()));
            a
        }
    };
    let mut a = arr.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = a.kind {
        if !elems
            .iter()
            .any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
        {
            elems.push(crate::keys::string_value(key));
        }
    }
}

pub fn track_nonconfig(obj: &Arc<Mutex<Object>>, key: &str) {
    let mut o = obj.lock().unwrap();
    let arr = match o.properties.get("__nonconfig") {
        Some(Value::Object(a)) => a.clone(),
        _ => {
            let a = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            o.properties
                .insert("__nonconfig".into(), Value::Object(a.clone()));
            a
        }
    };
    drop(o);
    let mut a = arr.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = a.kind {
        if !elems
            .iter()
            .any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
        {
            elems.push(crate::keys::string_value(key));
        }
    }
}

pub fn track_nonconfig_in(o: &mut Object, key: &str) {
    let arr = match o.properties.get("__nonconfig") {
        Some(Value::Object(a)) => a.clone(),
        _ => {
            let a = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            o.properties
                .insert("__nonconfig".into(), Value::Object(a.clone()));
            a
        }
    };
    let mut a = arr.lock().unwrap();
    if let ObjectKind::Array(ref mut elems) = a.kind {
        if !elems
            .iter()
            .any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
        {
            elems.push(crate::keys::string_value(key));
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
            let a = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
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

pub fn unwrap_fulfilled_promise(value: Value) -> Value {
    let Value::Object(obj) = &value else {
        return value;
    };
    let unwrapped = {
        let lock = obj.lock().unwrap();
        if !matches!(
            lock.properties.get("__type"),
            Some(Value::String(tag)) if tag.as_ref() == "Promise"
        ) {
            None
        } else if matches!(
            lock.properties.get("__state"),
            Some(Value::String(state)) if state.as_ref() == "fulfilled"
        ) {
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
    crate::keys::with_prefixed_property_key("@@", key, |raw_key| {
        crate::keys::with_prefixed_property_key("Symbol.", key, |symbol_key| {
            crate::keys::with_wrapped_property_key("Symbol(@@", key, ")", |symbol_paren_key| {
                let mut current = receiver.clone();
                for _ in 0..100 {
                    let next_proto = {
                        let lock = current.lock().unwrap();
                        for check_key in [key, raw_key, symbol_key, symbol_paren_key] {
                            if let Some(value) = lock.properties.get(check_key) {
                                if !matches!(value, Value::Null | Value::Undefined) {
                                    return Some(value.clone());
                                }
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
            })
        })
    })
}

fn call_iterator_if_generator(
    ctx: &mut HostContext,
    receiver: &Arc<Mutex<Object>>,
    method_name: &str,
) -> Option<Value> {
    let method = lookup_protocol_member(receiver, method_name)?;
    if matches!(method, Value::Null | Value::Undefined) {
        return None;
    }
    let iterator = crate::function::invoke_bound_callback_if_needed(ctx, &method, &[])
        .unwrap_or_else(|| {
            invoke_with_explicit_this(ctx, &method, Value::Object(receiver.clone()), &[])
        });
    let iterator = match await_or_reject(iterator) {
        Ok(value) => value,
        Err(_) => return None,
    };
    if let Value::Object(ref obj) = iterator {
        let o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::Continuation(_)) {
            return Some(iterator.clone());
        }
    }
    None
}

pub fn collect_protocol_iterable(
    ctx: &mut HostContext,
    receiver: &Arc<Mutex<Object>>,
    method_name: &str,
) -> Option<Value> {
    collect_protocol_iterable_result(ctx, receiver, method_name)?.ok()
}

pub fn collect_protocol_iterable_result(
    ctx: &mut HostContext,
    receiver: &Arc<Mutex<Object>>,
    method_name: &str,
) -> Option<Result<Value, Value>> {
    let method = lookup_protocol_member(receiver, method_name)?;
    if matches!(method, Value::Null | Value::Undefined) {
        return None;
    }
    let iterator = crate::function::invoke_bound_callback_if_needed(ctx, &method, &[])
        .unwrap_or_else(|| {
            invoke_with_explicit_this(ctx, &method, Value::Object(receiver.clone()), &[])
        });
    let iterator = match await_or_reject(iterator) {
        Ok(value) => value,
        Err(reason) => return Some(Err(reason)),
    };
    let Value::Object(iterator_obj) = iterator else {
        return None;
    };

    let mut out = Vec::with_capacity(8);
    let next_fn = lookup_protocol_member(&iterator_obj, "next");
    let Some(next_fn) = next_fn else {
        return Some(Ok(Value::Object(vybe_runtime::heap::alloc(
            Object::new_array(out),
        ))));
    };
    for _ in 0..1024 {
        let step = if let Some(result) =
            crate::function::try_invoke_bound_callback_if_needed(ctx, &next_fn, &[])
        {
            match result {
                Ok(value) => value,
                Err(reason) => return Some(Err(reason)),
            }
        } else {
            match crate::function::try_invoke_with_explicit_this(
                ctx,
                &next_fn,
                Value::Object(iterator_obj.clone()),
                &[],
            ) {
                Ok(value) => value,
                Err(reason) => return Some(Err(reason)),
            }
        };
        let step = match await_or_reject(step) {
            Ok(value) => value,
            Err(reason) => return Some(Err(reason)),
        };
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
            if let Err(reason) = await_or_reject(value) {
                return Some(Err(reason));
            }
            break;
        }
        out.push(value);
    }

    Some(Ok(Value::Object(vybe_runtime::heap::alloc(
        Object::new_array(out),
    ))))
}

fn await_or_reject(value: Value) -> Result<Value, Value> {
    if let Value::Object(obj) = &value {
        let lock = obj.lock().unwrap();
        let is_promise = matches!(
            lock.properties.get("__type"),
            Some(Value::String(tag)) if tag.as_ref() == "Promise"
        );
        if is_promise {
            let state = match lock.properties.get("__state") {
                Some(Value::String(state)) => state.as_ref(),
                _ => "",
            };
            let settled = lock
                .properties
                .get("__value")
                .cloned()
                .unwrap_or(Value::Undefined);
            if state == "rejected" {
                return Err(settled);
            }
            if state == "fulfilled" {
                return Ok(settled);
            }
        }
    }
    Ok(unwrap_fulfilled_promise(value))
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

fn array_hole_set(o: &Object) -> Option<std::collections::HashSet<i32>> {
    let Some(Value::Object(arr)) = o.properties.get("__holes") else {
        return None;
    };
    let mut holes = std::collections::HashSet::new();
    let a = arr.lock().unwrap();
    if let ObjectKind::Array(ref hs) = a.kind {
        holes.reserve(hs.len());
        for v in hs {
            if let Value::I32(n) = v {
                holes.insert(*n);
            }
        }
    }
    if holes.is_empty() { None } else { Some(holes) }
}

#[inline]
fn cached_array_hole_contains(
    holes: &Option<std::collections::HashSet<i32>>,
    index: usize,
) -> bool {
    match holes {
        Some(holes) => holes.contains(&(index as i32)),
        None => false,
    }
}

/// Returns true if `key` is marked non-enumerable on `obj`.
pub fn is_nonenum(o: &Object, key: &str) -> bool {
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

pub fn is_nonconfig(o: &Object, key: &str) -> bool {
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

pub fn ordered_own_string_keys(o: &Object) -> Vec<String> {
    if let Some(keys) = plain_own_string_keys(o) {
        return keys;
    }
    let tracked: Option<Vec<String>> = o.properties.get("__keys").and_then(|v| {
        if let Value::Object(arr) = v {
            let ka = arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ka.kind {
                let mut keys = Vec::with_capacity(elems.len());
                for e in elems {
                    if let Value::String(s) = e {
                        if o.properties.contains_key(s.as_ref()) {
                            keys.push(s.to_string());
                        }
                    }
                }
                return Some(keys);
            }
        }
        None
    });
    let sym_keys: Option<std::collections::HashSet<String>> =
        o.properties.get("__sym_keys").and_then(|v| {
            if let Value::Object(a) = v {
                let lock = a.lock().unwrap();
                if let ObjectKind::Array(ref el) = lock.kind {
                    let mut keys = std::collections::HashSet::with_capacity(el.len());
                    for e in el {
                        match e {
                            Value::String(s) => {
                                keys.insert(s.to_string());
                            }
                            Value::Symbol(sym) => {
                                keys.insert(crate::symbol::canonical_property_key(sym));
                            }
                            _ => {}
                        }
                    }
                    Some(keys)
                } else {
                    None
                }
            } else {
                None
            }
        });
    let mut live = Vec::with_capacity(o.properties.len());
    for key in o.properties.keys() {
        let is_symbol_key = matches!(sym_keys.as_ref(), Some(keys) if keys.contains(key));
        if !key.starts_with("__") && !is_symbol_key {
            live.push(key.clone());
        }
    }
    match tracked {
        Some(mut keys) => {
            let mut seen: std::collections::HashSet<&str> =
                std::collections::HashSet::with_capacity(keys.len());
            for key in &keys {
                seen.insert(key.as_str());
            }
            let mut extras = Vec::with_capacity(live.len().saturating_sub(keys.len()));
            for key in &live {
                if !seen.contains(key.as_str()) {
                    extras.push(key.clone());
                    seen.insert(key.as_str());
                }
            }
            keys.extend(extras);
            sort_array_indices_first(&mut keys);
            keys
        }
        None => {
            let mut keys = live;
            sort_array_indices_first(&mut keys);
            keys
        }
    }
}

fn has_plain_ordered_data_properties(o: &Object) -> bool {
    // Match the existing plain-key boundary, but require ordinary data
    // storage and insertion-order keys so no descriptor/proxy/accessor or
    // numeric sorting behavior is bypassed. Check before building pairs.
    if !matches!(o.kind, ObjectKind::Ordinary)
        || o.type_id != 0
        || !o.fields.is_empty()
        || o.properties.keys().any(|key| {
            key.starts_with("__") || crate::keys::canonical_array_index_key(key).is_some()
        })
    {
        return false;
    }
    true
}

fn plain_ordered_data_entries(o: &Object) -> Option<Vec<Value>> {
    if !has_plain_ordered_data_properties(o) {
        return None;
    }
    let mut entries = Vec::with_capacity(o.properties.len());
    for (key, value) in &o.properties {
        entries.push(crate::array::make_pair_array(
            crate::keys::string_value(key),
            value.clone(),
        ));
    }
    Some(entries)
}

fn plain_own_string_keys(o: &Object) -> Option<Vec<String>> {
    let mut keys = Vec::with_capacity(o.properties.len());
    for key in o.properties.keys() {
        if key.starts_with("__") {
            return None;
        }
        keys.push(key.clone());
    }
    sort_array_indices_first(&mut keys);
    Some(keys)
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
            return Some((n % modv.as_i32()).to_string());
        }
        drop(o);
    }
    None
}

fn groupby_magic_key_callable(key_fn: &Value) -> bool {
    let Value::Object(kf) = key_fn else {
        return false;
    };
    let o = kf.lock().unwrap();
    matches!(o.kind, ObjectKind::Ordinary)
        && (o.properties.contains_key("__groupby_le2_small_large")
            || o.properties.contains_key("__group_by_mod"))
}

fn is_callable_value(value: &Value) -> bool {
    match value {
        Value::Object(obj) => {
            matches!(
                obj.lock().unwrap().kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            ) || groupby_magic_key_callable(value)
        }
        _ => false,
    }
}

fn throw_type_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "TypeError", message));
    Value::Undefined
}

fn collect_groupby_items(
    ctx: &mut HostContext,
    items: &Value,
    message: &str,
) -> Option<Vec<Value>> {
    match items {
        Value::Null
        | Value::Undefined
        | Value::Bool(_)
        | Value::I32(_)
        | Value::I64(_)
        | Value::F32(_)
        | Value::F64(_)
        | Value::Symbol(_)
        | Value::BigInt(_) => {
            let _ = throw_type_error(ctx, message);
            None
        }
        _ => match crate::iterator::try_materialize_iterable_values(ctx, items, false) {
            Ok(values) => Some(values),
            Err(error) => {
                ctx.throw_value(error);
                None
            }
        },
    }
}

/// Walk the prototype chain looking for `key`. Returns the value if
/// found at any depth, `None` if not present in the whole chain.
pub fn proto_walk_get(obj: &Arc<Mutex<Object>>, key: &str) -> Option<Value> {
    let mut current = obj.clone();
    loop {
        let o = current.lock().unwrap();
        if let Some(v) = o.properties.get(key) {
            return Some(v.clone());
        }
        let proto = o
            .properties
            .get(PROTO_KEY)
            .cloned()
            .unwrap_or_else(|| implicit_object_prototype(&current, &o));
        drop(o);
        // An *absent* `__proto__` means the VM created this object bare
        // (WASM-pure) — resolve its [[Prototype]] by kind so inherited
        // members (`[].map`, `arr[Symbol.iterator]`, `{}.hasOwnProperty`)
        // are found on the canonical shared prototype. The shared
        // prototypes carry explicit `__proto__` links up to
        // %Object.prototype% (whose `__proto__` is null), so the walk
        // always terminates. An *explicit* `null` (Object.create(null))
        // ends the chain immediately.
        match proto {
            Value::Object(p) => {
                if Arc::ptr_eq(&p, &current) {
                    return None;
                }
                current = p;
            }
            _ => return None,
        }
    }
}

/// Walk the prototype chain looking only for the presence of `key`.
/// This is the same lookup shape as `proto_walk_get`, but avoids cloning
/// the found value for Reflect.has / `in` / proxy has paths.
pub fn proto_walk_has(obj: &Arc<Mutex<Object>>, key: &str) -> bool {
    let mut current = obj.clone();
    loop {
        let o = current.lock().unwrap();
        if o.properties.contains_key(key) {
            return true;
        }
        let proto = o
            .properties
            .get(PROTO_KEY)
            .cloned()
            .unwrap_or_else(|| implicit_object_prototype(&current, &o));
        drop(o);
        match proto {
            Value::Object(p) => {
                if Arc::ptr_eq(&p, &current) {
                    return false;
                }
                current = p;
            }
            _ => return false,
        }
    }
}

// Assignment searches for a setter before rejecting a getter-only property.
// Collect getter presence during that same search instead of locking and
// traversing the whole prototype chain a second time on the common miss.
fn proto_walk_assignment_accessors(
    obj: &Arc<Mutex<Object>>,
    setter_key: &str,
    getter_key: &str,
) -> (Option<Value>, bool) {
    let mut current = obj.clone();
    let mut has_getter = false;
    loop {
        let object = current.lock().unwrap();
        if let Some(setter) = object.properties.get(setter_key) {
            return (Some(setter.clone()), has_getter);
        }
        has_getter |= object.properties.contains_key(getter_key);
        let prototype = object
            .properties
            .get(PROTO_KEY)
            .cloned()
            .unwrap_or_else(|| implicit_object_prototype(&current, &object));
        drop(object);
        match prototype {
            Value::Object(next) if !Arc::ptr_eq(&next, &current) => current = next,
            _ => return (None, has_getter),
        }
    }
}

pub fn is_not_extensible(o: &Object) -> bool {
    matches!(o.properties.get(EXTENSIBLE_MARK), Some(Value::I32(0)))
}

pub fn mark_not_extensible(o: &mut Object) {
    o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
}

pub fn install_noop_setter(o: &mut Object, key: &str) {
    let noop_idx = NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
    if noop_idx == 0 {
        return;
    }
    let mut noop_obj = Object::new();
    noop_obj.properties.reserve(1);
    noop_obj.kind = ObjectKind::HostFunction(noop_idx);
    let noop_val = Value::Object(vybe_runtime::heap::alloc(noop_obj));
    let setter_key = setter_property_key(key);
    if !o.properties.contains_key(&setter_key) {
        o.properties.insert(setter_key, noop_val);
    }
}

fn proto_walk_invoke_getter(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    key: &str,
) -> Option<Value> {
    let getter =
        crate::keys::with_getter_property_key(key, |getter_key| proto_walk_get(obj, getter_key))?;
    Some(invoke_getter_value(ctx, obj, getter))
}

fn invoke_getter_value(ctx: &mut HostContext, obj: &Arc<Mutex<Object>>, getter: Value) -> Value {
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
    if ctx.receiver_is_parameter() {
        let previous = ctx.current_js_this();
        ctx.set_js_this(receiver);
        let out = ctx.invoke(&getter, &[]);
        ctx.set_js_this(previous);
        return out;
    }
    match getter_arity {
        Some(0) => ctx.invoke(&getter, &[]),
        _ => ctx.invoke(&getter, &[receiver]),
    }
}

fn proto_walk_get_or_invoke_getter(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    key: &str,
) -> Option<Value> {
    proto_walk_get_or_invoke_getter_from(ctx, obj, obj.clone(), key)
}

// Continue a lookup after its receiver has already been checked, while
// retaining that original receiver for an inherited accessor invocation.
fn proto_walk_get_or_invoke_getter_from(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    first: Arc<Mutex<Object>>,
    key: &str,
) -> Option<Value> {
    crate::keys::with_getter_property_key(key, |getter_key| {
        let mut current = first;
        loop {
            let (value, getter, proto) = {
                let o = current.lock().unwrap();
                let value = o.properties.get(key).cloned();
                let getter = if value.is_none() {
                    o.properties.get(getter_key).cloned()
                } else {
                    None
                };
                let proto = if value.is_none() && getter.is_none() {
                    o.properties
                        .get(PROTO_KEY)
                        .cloned()
                        .unwrap_or_else(|| implicit_object_prototype(&current, &o))
                } else {
                    Value::Null
                };
                (value, getter, proto)
            };
            if let Some(value) = value {
                return Some(value);
            }
            if let Some(getter) = getter {
                return Some(invoke_getter_value(ctx, obj, getter));
            }
            match proto {
                Value::Object(p) => {
                    if Arc::ptr_eq(&p, &current) {
                        return None;
                    }
                    current = p;
                }
                _ => return None,
            }
        }
    })
}

fn object_to_string_tag_value(ctx: &mut HostContext, obj: &Arc<Mutex<Object>>) -> Value {
    if let Some(tag) = proto_walk_get_or_invoke_getter(ctx, obj, "tostringtag") {
        match tag {
            Value::String(text) if !text.is_empty() => {
                return crate::keys::object_tag_value(text.as_ref());
            }
            Value::Undefined | Value::Null => {}
            other => {
                let tag = crate::keys::value_display_cow(&other);
                return crate::keys::object_tag_value(tag.as_ref());
            }
        }
    }

    let object = obj.lock().unwrap();
    let type_tag = object.properties.get("__type");
    match &object.kind {
        _ if matches!(type_tag, Some(Value::String(tag)) if !tag.is_empty()) => match type_tag {
            Some(Value::String(tag)) => crate::keys::object_tag_value(tag.as_ref()),
            _ => crate::keys::object_tag_value("Object"),
        },
        ObjectKind::Array(_) => crate::keys::object_tag_value("Array"),
        ObjectKind::Map(_) => crate::keys::object_tag_value("Map"),
        ObjectKind::Set(_) => crate::keys::object_tag_value("Set"),
        ObjectKind::ArrayBuffer(_) => crate::keys::object_tag_value("ArrayBuffer"),
        ObjectKind::TypedArray(_) => match type_tag {
            Some(Value::String(tag)) if !tag.is_empty() => {
                crate::keys::object_tag_value(tag.as_ref())
            }
            _ => crate::keys::object_tag_value("TypedArray"),
        },
        ObjectKind::Function(_) | ObjectKind::HostFunction(_) => {
            crate::keys::object_tag_value("Function")
        }
        ObjectKind::ModuleNamespace => crate::keys::object_tag_value("Module"),
        _ => crate::keys::object_tag_value("Object"),
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

    // §20.1.3: %Object.prototype% carries its intrinsics as OWN callable
    // values so borrowed-call forms work —
    // `Object.prototype.hasOwnProperty.call(o, k)` etc.
    if let Value::Object(proto) = shared_object_prototype() {
        let mut p = proto.lock().unwrap();
        for (name, registry_name) in [
            // Borrowed-call values are the RAW intrinsics — a receiver's
            // own override must not shadow an explicitly borrowed
            // Object.prototype method.
            ("hasOwnProperty", "hasOwnPropertyIntrinsic"),
            ("propertyIsEnumerable", "propertyIsEnumerable"),
            ("isPrototypeOf", "isPrototypeOf"),
            ("toString", "toString"),
            ("valueOf", "valueOf"),
            ("toLocaleString", "toLocaleString"),
        ] {
            if p.properties.contains_key(name) {
                continue;
            }
            let Some(&idx) = vm
                .host_registry
                .get(&("ecma:object".to_string(), registry_name.to_string()))
            else {
                continue;
            };
            let mut f = Object::new();
            f.properties.reserve(2);
            f.kind = ObjectKind::HostFunction(idx);
            f.properties
                .insert("name".into(), crate::keys::string_value(name));
            f.properties
                .insert("__vybe_method_receiver".into(), Value::Bool(true));
            p.properties
                .insert(name.into(), Value::Object(vybe_runtime::heap::alloc(f)));
        }
    }
}

// ── Construction ──────────────────────────────────────────────────────

fn register_construction(vm: &mut VM) {
    object_fn(
        vm,
        "new",
        vec![],
        vec![obj_t()],
        Box::new(|_ctx, _args| new_ordinary_object_with_proto()),
    );

    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "Object",
        Box::new(
            |_ctx, args| match args.first().cloned().unwrap_or(Value::Undefined) {
                Value::Null | Value::Undefined => new_ordinary_object_with_proto(),
                value @ Value::Object(_) => value,
                Value::Bool(value) => crate::boolean::boxed_boolean(value),
                Value::String(text) => crate::string::boxed_string(text),
                value @ Value::F64(_) | value @ Value::I32(_) | value @ Value::I64(_) => {
                    crate::number::boxed_number(value)
                }
                Value::Symbol(desc) => {
                    let mut obj = Object::new();
                    obj.properties
                        .insert("__type".into(), crate::keys::string_value("Symbol"));
                    obj.properties
                        .insert("__primitive".into(), Value::Symbol(desc));
                    obj.properties
                        .insert(PROTO_KEY.into(), shared_object_prototype());
                    Value::Object(vybe_runtime::heap::alloc(obj))
                }
                Value::BigInt(value) => {
                    let mut obj = Object::new();
                    obj.properties
                        .insert("__type".into(), crate::keys::string_value("BigInt"));
                    obj.properties
                        .insert("__primitive".into(), Value::BigInt(value));
                    obj.properties
                        .insert(PROTO_KEY.into(), shared_object_prototype());
                    Value::Object(vybe_runtime::heap::alloc(obj))
                }
                _ => new_ordinary_object_with_proto(),
            },
        ),
    );

    // create(proto, propertiesDescriptor?) -> new obj
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "create",
        Box::new(|ctx, args| {
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
            let arc = vybe_runtime::heap::alloc(Object::new());
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
                    // Object.create(null) gives a "bare" object — none of
                    // `Object.prototype`'s methods are reachable. The
                    // explicit `__proto__: Null` marker is what
                    // `resolve_property` keys off to skip the universal
                    // Object vtable; no placeholder own-properties (they
                    // made `hasOwnProperty.call(o, "toString")` lie).
                    o.properties.insert(PROTO_KEY.into(), Value::Null);
                    o.properties
                        .insert(NULL_PROTO_MARK.into(), Value::Bool(true));
                }
                Some(_) => {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "Object prototype may only be an Object or null",
                    ));
                    return Value::Undefined;
                }
                _ => {}
            }
            // Second arg is the property-descriptors map per
            // §20.1.2.2 step 4; iterate its keys and apply via the
            // same logic `Object.defineProperty` uses.
            if let Some(Value::Object(descs)) = args.get(1) {
                // Snapshot descriptor entries (preserving __keys order
                // when present) before mutating the target.
                let entries: Vec<(String, Value, Option<Value>)> = {
                    let d = descs.lock().unwrap();
                    let order = descriptor_own_keys(&d);
                    let sym_values: std::collections::HashMap<String, Value> = d
                        .properties
                        .get("__sym_keys")
                        .and_then(|v| {
                            if let Value::Object(arr) = v {
                                let ka = arr.lock().unwrap();
                                if let ObjectKind::Array(ref elems) = ka.kind {
                                    Some(
                                        elems
                                            .iter()
                                            .map(|key| (key_string(key), key.clone()))
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
                    let mut out = Vec::with_capacity(if order.is_empty() {
                        d.properties.len()
                    } else {
                        order.len()
                    });
                    if !order.is_empty() {
                        for k in order {
                            if k.starts_with("__") {
                                continue;
                            }
                            if let Some(v) = d.properties.get(&k) {
                                let sym = sym_values.get(&k).cloned();
                                out.push((k, v.clone(), sym));
                            }
                        }
                    } else {
                        for (k, v) in d.properties.iter() {
                            if k.starts_with("__") {
                                continue;
                            }
                            let sym = sym_values.get(k).cloned();
                            out.push((k.clone(), v.clone(), sym));
                        }
                    }
                    out
                };
                for (k, v, sym_key) in entries {
                    let Value::Object(desc) = v else {
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "TypeError",
                            "Property descriptor must be an object",
                        ));
                        return Value::Undefined;
                    };
                    let dlock = desc.lock().unwrap();
                    let has_value = dlock.properties.contains_key("value");
                    let has_writable = dlock.properties.contains_key("writable");
                    let has_get = dlock.properties.contains_key("get");
                    let has_set = dlock.properties.contains_key("set");
                    if (has_value || has_writable) && (has_get || has_set) {
                        drop(dlock);
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "TypeError",
                            "Invalid property descriptor",
                        ));
                        return Value::Undefined;
                    }
                    let val = dlock.properties.get("value").cloned();
                    let getter = dlock.properties.get("get").cloned().filter(|v| {
                        matches!(v, Value::Object(o)
                            if matches!(o.lock().unwrap().kind,
                                ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
                    });
                    let setter = dlock.properties.get("set").cloned().filter(|v| {
                        matches!(v, Value::Object(o)
                            if matches!(o.lock().unwrap().kind,
                                ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
                    });
                    let enumerable = dlock
                        .properties
                        .get("enumerable")
                        .map(|x| x.as_bool())
                        .unwrap_or(false);
                    let configurable = dlock
                        .properties
                        .get("configurable")
                        .map(|x| x.as_bool())
                        .unwrap_or(false);
                    let writable = dlock.properties.get("writable").map(|x| x.as_bool());
                    drop(dlock);

                    if let Some(sym) = sym_key {
                        track_sym_key(&arc, sym);
                    } else {
                        track_key(&arc, &k);
                    }
                    if !enumerable {
                        track_nonenum(&arc, &k);
                    }
                    if !configurable {
                        track_nonconfig(&arc, &k);
                    }
                    let mut o = arc.lock().unwrap();
                    let getter_key = getter_property_key(&k);
                    let setter_key = setter_property_key(&k);
                    if let Some(g) = getter {
                        o.properties.insert(getter_key.clone(), g);
                    }
                    if let Some(s) = setter {
                        o.properties.insert(setter_key.clone(), s);
                    }
                    if let Some(v) = val {
                        o.properties.shift_remove(&getter_key);
                        o.properties.shift_remove(&setter_key);
                        o.properties.insert(k.clone(), v);
                        if matches!(writable, Some(false) | None) {
                            install_noop_setter(&mut o, &k);
                        }
                    } else if has_get || has_set {
                        o.properties.insert(k.clone(), Value::Undefined);
                    } else if has_writable {
                        o.properties.insert(k.clone(), Value::Undefined);
                        if matches!(writable, Some(false) | None) {
                            install_noop_setter(&mut o, &k);
                        }
                    }
                }
            }
            Value::Object(arc)
        }),
    );

    // fromEntries(iterable) -> new obj
    object_unary(
        vm,
        "fromEntries",
        vec![obj_t()],
        Box::new(|ctx, args| {
            let mut obj = Object::new();
            // ECMA-262 §7.3.22: the resulting object's property order is the
            // entries' insertion order. `Object::properties` is an unordered
            // HashMap, so record order in the `__keys` tracker that
            // `ordinary_ordered_keys` reads (the same mechanism object literals
            // use; `__`-prefixed keys are excluded from enumeration).
            let mut order: Vec<Value> = Vec::new();
            let put = |obj: &mut Object, order: &mut Vec<Value>, key: String, val: Value| {
                if !obj.properties.contains_key(&key) {
                    order.push(crate::keys::string_value(&key));
                }
                obj.properties.insert(key, val);
            };
            let Some(source) = args.first() else {
                return throw_type_error(ctx, "undefined is not iterable");
            };
            let dense_array_pairs = match source {
                Value::Object(source_obj) => {
                    let source_lock = source_obj.lock().unwrap();
                    if !source_lock.properties.contains_key("__holes")
                        && !source_lock.properties.contains_key("iterator")
                        && !has_getter_key(&source_lock, "iterator")
                    {
                        if let ObjectKind::Array(values) = &source_lock.kind {
                            Some(values.clone())
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                }
                _ => None,
            };
            let pairs = match dense_array_pairs {
                Some(values) => values,
                None => {
                    match crate::iterator::try_materialize_iterable_values(ctx, source, false) {
                        Ok(values) => values,
                        Err(error) => {
                            ctx.throw_value(error);
                            return Value::Undefined;
                        }
                    }
                }
            };
            order.reserve(pairs.len());
            for pair in pairs {
                let Value::Object(pair_obj) = pair else {
                    return throw_type_error(ctx, "Iterator value is not an entry object");
                };
                let entry = {
                    let p = pair_obj.lock().unwrap();
                    match &p.kind {
                        ObjectKind::Array(kv) if kv.len() >= 2 => {
                            Some((kv[0].clone(), kv[1].clone()))
                        }
                        _ => {
                            let key = p.properties.get("0").cloned();
                            let value = p.properties.get("1").cloned();
                            key.zip(value)
                        }
                    }
                };
                let Some((key, value)) = entry else {
                    return throw_type_error(ctx, "Iterator value is not an entry object");
                };
                put(&mut obj, &mut order, key_string(&key), value);
            }
            if !order.is_empty() {
                obj.properties.insert(
                    "__keys".to_string(),
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(order))),
                );
            }
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // `Object.assign(target, ...sources)` — ECMA-262 §20.1.2.1.
    // Variadic in the source positions; each source contributes its
    // own enumerable string-keyed properties onto target. Returns the
    // modified target. Internal `__`-prefixed properties are skipped
    // (they're our private metadata, not enumerable JS properties).
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "assign",
        Box::new(|ctx, args| {
            let Some(raw_target) = args.first() else {
                return throw_type_error(ctx, "Cannot convert undefined or null to object");
            };
            let Some(target) = to_object_for_object_static(
                ctx,
                raw_target,
                "Cannot convert undefined or null to object",
            ) else {
                return Value::Undefined;
            };
            let Value::Object(target_obj) = &target else {
                return target;
            };

            for source in args.iter().skip(1) {
                if matches!(source, Value::Null | Value::Undefined) {
                    continue;
                }
                let Some(source_value) =
                    to_object_for_object_static(ctx, source, "Cannot convert source to object")
                else {
                    return Value::Undefined;
                };
                if let Some(proxy_keys) = crate::proxy::own_keys_dispatch(ctx, &source_value) {
                    for key_value in array_elements(&proxy_keys) {
                        if !matches!(key_value, Value::String(_)) {
                            continue;
                        }
                        let desc = crate::proxy::get_own_property_descriptor_dispatch(
                            ctx,
                            &source_value,
                            &key_value,
                        )
                        .unwrap_or(Value::Undefined);
                        if !matches!(desc, Value::Object(_))
                            || !descriptor_bool(&desc, "enumerable", false)
                        {
                            continue;
                        }
                        let key = key_string(&key_value);
                        let value = crate::proxy::get_dispatch(ctx, &source_value, &key_value);
                        if !assign_strict_set(ctx, target_obj, &key, value, None) {
                            return Value::Undefined;
                        }
                    }
                    continue;
                }
                let Value::Object(source_obj) = source_value else {
                    continue;
                };
                let assign_keys = enumerable_assign_keys(&source_obj);
                let direct_values = {
                    let src = source_obj.lock().unwrap();
                    let mut values = Vec::with_capacity(assign_keys.len());
                    let mut direct = true;
                    let no_getters = !has_any_getter(&src);
                    for (key, _) in &assign_keys {
                        if !no_getters {
                            direct = false;
                            break;
                        }
                        let value = match &src.kind {
                            ObjectKind::Array(values) => {
                                if let Some(index) = array_index_key(key) {
                                    values.get(index as usize).cloned()
                                } else {
                                    src.properties.get(key).cloned()
                                }
                            }
                            ObjectKind::TypedArray(ta) => {
                                if let Some(index) = array_index_key(key) {
                                    Some(crate::typedarray::read_element(ta, index as usize))
                                } else {
                                    src.properties.get(key).cloned()
                                }
                            }
                            _ => src.properties.get(key).cloned(),
                        };
                        let Some(value) = value else {
                            direct = false;
                            break;
                        };
                        values.push(value);
                    }
                    if direct { Some(values) } else { None }
                };
                if let Some(values) = direct_values {
                    for ((key, symbol_key), value) in assign_keys.into_iter().zip(values) {
                        if !assign_strict_set(ctx, target_obj, &key, value, symbol_key) {
                            return Value::Undefined;
                        }
                    }
                } else {
                    for (key, symbol_key) in assign_keys {
                        let value = assign_source_get(ctx, &source_obj, &key);
                        if !assign_strict_set(ctx, target_obj, &key, value, symbol_key) {
                            return Value::Undefined;
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
    // get(obj, key) -> value — §7.3.2 GetV: full [[Get]], walking the
    // prototype chain AND invoking accessor getters (a destructuring
    // read like `const {msg} = o` lands here and must fire `get msg()`).
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "get",
        Box::new(|ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let undefined = Value::Undefined;
                let key_raw = args.get(1).unwrap_or(&undefined);
                // ECMA-262 §10.4.5.4 [[Get]] on an Integer-Indexed exotic
                // object — the mirror of the §10.4.5.5 arm in `set` below.
                // A CANONICAL NUMERIC key reads the typed buffer; an invalid
                // index answers `undefined` WITHOUT walking the prototype
                // chain. Without this arm, a TypedArray subscript through the
                // dynamic path proto-walked and answered `undefined` even for
                // a valid index.
                let string_miss = {
                    let o = obj.lock().unwrap();
                    if let ObjectKind::Array(ref values) = o.kind {
                        if matches!(key_raw, Value::String(text) if text.as_ref() == "length") {
                            return Value::I32(values.len() as i32);
                        }
                        if let Some(index) = crate::keys::canonical_array_index_value(key_raw) {
                            let index = index as usize;
                            if index < values.len() && !crate::array::is_array_hole(&o, index) {
                                return values[index].clone();
                            }
                        }
                    }
                    if let ObjectKind::TypedArray(ref ta) = o.kind {
                        let canonical = crate::keys::canonical_numeric_index_value(key_raw);
                        if let Some(n) = canonical {
                            let valid = n.fract() == 0.0
                                && n >= 0.0
                                && !(n == 0.0 && n.is_sign_negative())
                                && (n as usize) < crate::typedarray::ta_live_length(ta)
                                && n.is_finite();
                            if valid {
                                return crate::typedarray::read_element(ta, n as usize);
                            }
                            return Value::Undefined;
                        }
                    }
                    // Ordinary string-key hits need no key conversion or
                    // second lock. Exotic indexed reads above take priority.
                    if let Value::String(key) = key_raw {
                        if let Some(value) = o.properties.get(key.as_ref()) {
                            return value.clone();
                        }
                        let getter =
                            crate::keys::with_getter_property_key(key.as_ref(), |getter_key| {
                                o.properties.get(getter_key).cloned()
                            });
                        let prototype = if getter.is_none() {
                            o.properties
                                .get(PROTO_KEY)
                                .cloned()
                                .unwrap_or_else(|| implicit_object_prototype(&obj, &o))
                        } else {
                            Value::Null
                        };
                        Some((getter, prototype))
                    } else {
                        None
                    }
                };
                if let (Value::String(key), Some((getter, prototype))) = (key_raw, string_miss) {
                    if let Some(getter) = getter {
                        return invoke_getter_value(ctx, &obj, getter);
                    }
                    return match prototype {
                        Value::Object(next) if !Arc::ptr_eq(&next, &obj) => {
                            proto_walk_get_or_invoke_getter_from(ctx, &obj, next, key.as_ref())
                                .unwrap_or(Value::Undefined)
                        }
                        _ => Value::Undefined,
                    };
                }
                return with_property_key(key_raw, |key| {
                    if let Some(v) = proto_walk_get_or_invoke_getter(ctx, &obj, key) {
                        v
                    } else {
                        Value::Undefined
                    }
                });
            }
            Value::Undefined
        }),
    );

    // set(obj, key, value) -> ()
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "set",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let undefined = Value::Undefined;
                let key_raw = args.get(1).unwrap_or(&undefined);
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                // ECMA-262 §10.4.5.5 [[Set]] on an Integer-Indexed exotic
                // object: a CANONICAL NUMERIC key never reaches OrdinarySet.
                // A valid index stores through the typed buffer
                // (§10.4.5.16 — truncate/wrap/clamp per element kind); an
                // invalid one (out of bounds, fractional, negative, -0,
                // non-finite) is IGNORED — no expando property, no throw.
                // Before this arm existed, `u[1] = 5` landed in the property
                // bag while reads hit the typed buffer, so every indexed
                // write silently read back 0.
                {
                    let o = obj.lock().unwrap();
                    if let ObjectKind::TypedArray(ref ta) = o.kind {
                        let canonical = crate::keys::canonical_numeric_index_value(key_raw);
                        if let Some(n) = canonical {
                            let valid = n.fract() == 0.0
                                && n >= 0.0
                                && !(n == 0.0 && n.is_sign_negative())
                                && (n as usize) < crate::typedarray::ta_live_length(ta)
                                && n.is_finite();
                            if valid {
                                crate::typedarray::write_element(ta, n as usize, &val);
                            }
                            return Value::Null;
                        }
                    }
                }
                return with_property_key(key_raw, |key| {
                    crate::keys::with_getter_property_key(key, |getter_key| {
                        crate::keys::with_setter_property_key(key, |setter_key| {
                            // ECMA-262 §10.1.5 OrdinarySet — three gates:
                            //   1. Frozen → writes fail: silently in loose mode,
                            //      TypeError in strict (§13.15.2, caller passes the
                            //      optional 4th `strict` arg).
                            //   2. Sealed / preventExtensions → new keys fail; existing
                            //      keys writable unless also frozen.
                            //   3. `__set_<key>` accessor → call setter instead of
                            //      writing to the property bag.
                            let strict =
                                args.get(3).map(crate::boolean::to_boolean).unwrap_or(false);
                            let (is_error_type, own_setter, own_getter) = {
                                let mut o = obj.lock().unwrap();
                                let not_extensible = matches!(
                                    o.properties.get(EXTENSIBLE_MARK),
                                    Some(Value::I32(0))
                                );
                                let exists = if o.properties.contains_key(key) {
                                    true
                                } else {
                                    o.properties.contains_key(getter_key)
                                        || o.properties.contains_key(setter_key)
                                };
                                if not_extensible && !exists {
                                    drop(o);
                                    if strict {
                                        ctx.throw_value(crate::error::new_error(
                                            ctx,
                                            "TypeError",
                                            "Cannot add property, object is not extensible",
                                        ));
                                        return Value::Undefined;
                                    }
                                    return Value::Null;
                                }
                                if matches!(&o.kind, ObjectKind::Array(_))
                                    && (key == "length" || key == "__len__")
                                {
                                    crate::array::apply_js_array_length(ctx, &mut o, &val);
                                    return Value::Null;
                                }
                                (
                                    o.properties.contains_key("__exception_type"),
                                    o.properties.get(setter_key).cloned(),
                                    o.properties.contains_key(getter_key),
                                )
                            };
                            // ECMA-262 §20.5.2.1: Error.prototype.name is a data property,
                            // not an accessor. But some Error instances may have spurious
                            // __set_name setters from the type system. Ignore them for
                            // Error types to allow `e.name = "CustomError"` to work.
                            let skip_setter = is_error_type && (key == "name" || key == "message");

                            let (setter, has_getter) = if skip_setter {
                                (None, false)
                            } else if own_setter.is_some() {
                                (own_setter, own_getter)
                            } else {
                                let (setter, inherited_getter) =
                                    proto_walk_assignment_accessors(&obj, setter_key, getter_key);
                                (setter, own_getter || inherited_getter)
                            };
                            let has_getter_without_setter = setter.is_none() && has_getter;
                            if has_getter_without_setter {
                                if strict {
                                    ctx.throw_value(crate::error::new_error(
                                        ctx,
                                        "TypeError",
                                        "Cannot set property which has only a getter",
                                    ));
                                    return Value::Undefined;
                                }
                                return Value::Null;
                            }
                            if let Some(setter_val) = setter {
                                if let Value::Object(setter_obj) = &setter_val {
                                    // ECMA-262 §10.1.5 step 6.b: the setter is
                                    // called with `this = receiver`. A host fn cannot
                                    // mutate the VM, but it can match the arg count to
                                    // the setter's declared arity:
                                    //   - arity 1 (defineProperty `set(val)`):
                                    //     pass `[val]`.
                                    //   - arity 2 (class `set name(val)` compiled
                                    //     as `(self, val)`): pass `[obj, val]`
                                    //     so the explicit-self slot binds.
                                    let (setter_arity, is_noop_setter) = {
                                        let so = setter_obj.lock().unwrap();
                                        let arity = match &so.kind {
                                            vybe_runtime::value::ObjectKind::Function(f) => {
                                                Some(f.arity)
                                            }
                                            _ => None,
                                        };
                                        let noop = matches!(
                                            so.kind,
                                            vybe_runtime::value::ObjectKind::HostFunction(idx)
                                                if idx == NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed)
                                        );
                                        (arity, noop)
                                    };
                                    if is_noop_setter {
                                        let accessor_setter_active = {
                                            let o = obj.lock().unwrap();
                                            o.properties.get(ACCESSOR_SETTER_ACTIVE_MARK).is_some()
                                                || is_accessor_backing_slot_write(&o, key)
                                        };
                                        if !accessor_setter_active && strict {
                                            ctx.throw_value(crate::error::new_error(
                                                ctx,
                                                "TypeError",
                                                "Cannot assign to read only property",
                                            ));
                                            return Value::Undefined;
                                        }
                                        if !accessor_setter_active {
                                            return Value::Null;
                                        }
                                    } else {
                                        {
                                            let mut o = obj.lock().unwrap();
                                            o.properties.insert(
                                                ACCESSOR_SETTER_ACTIVE_MARK.into(),
                                                Value::Bool(true),
                                            );
                                        }
                                        // ECMA-262 §10.1.5 step 6.b: the setter is called
                                        // with `this = receiver`. WHICH convention it wants
                                        // is now DECLARED, not inferred: `classes.rs`
                                        // stamps its `__set_x(self, v)` accessors with the
                                        // receiver-first call tag, and anything that does
                                        // not declare it — a `defineProperty` `set(v)` —
                                        // takes its receiver from the call.
                                        //
                                        // This used to guess from `setter_arity`, and the
                                        // two shapes are indistinguishable that way: an
                                        // arity-1 setter that wants a receiver and one that
                                        // does not are the same signature. The guess handed
                                        // `set(v)` no receiver at all, so `this.x = v` wrote
                                        // into nowhere and the assignment READ BACK
                                        // CORRECTLY off the plain property — which is why
                                        // `element.textContent = x` rendered nothing.
                                        // §10.1.5 step 6.b is unconditional: the setter is
                                        // called with `this = Receiver`. So BOTH branches
                                        // bind it; they differ only in whether the callee's
                                        // own signature ALSO takes the receiver as an
                                        // argument, which is an internal compilation
                                        // detail the standard knows nothing about.
                                        //
                                        // `classes.rs` compiles `__set_x(self, v)` and a JS
                                        // `defineProperty` setter is `set(v)`. Same wasm
                                        // signature class, opposite conventions — which is
                                        // why this used to GUESS from arity and hand the
                                        // ambient form no receiver at all, so `this.x = v`
                                        // wrote nowhere and read back off the plain
                                        // property.
                                        let receiver = Value::Object(obj.clone());
                                        // Does the callee's own signature take the receiver
                                        // as a parameter? A declared call tag answers it
                                        // outright; otherwise the arity does, and it is
                                        // reliable for the shape because the compiler emits
                                        // exactly two: `(self, v)` and `(v)`.
                                        //
                                        // What the arity CANNOT tell you — and what this
                                        // used to get wrong — is that a `(v)` setter still
                                        // needs `this`. Both branches bind it now, because
                                        // §10.1.5 step 6.b is unconditional; only the
                                        // argument list differs.
                                        // ⛔ UNDER `ReceiverAbi::Parameter` THE VM'S OWN
                                        // INVOKE SUPPLIES THE RECEIVER, so passing it here
                                        // as well hands the setter TWO — `this` becomes the
                                        // prepended one and `v` becomes the receiver.
                                        // Measured: `defineProperty(o,"v",{set})` then
                                        // `o.v = 9` wrote nothing, and an object-literal
                                        // getter threw "Cannot read properties of
                                        // undefined". Take the binding branch instead: it
                                        // sets the receiver and passes only the value,
                                        // which is exactly what the parameter ABI wants.
                                        // ⛔ NO TAG, NO ARITY GUESS — BOTH BRANCHES AGREE
                                        // NOW. `invoke_with_receiver` binds the receiver
                                        // channel and passes only the value, and the VM
                                        // fills the receiver slot in one place, so it is
                                        // right whether or not the setter declares the
                                        // receiver as a parameter. The tag and the
                                        // `arity >= 2` guess existed only to choose
                                        // between two spellings of the same call; there is
                                        // one spelling now. Left in place upstream, no
                                        // longer consulted here.
                                        // No call tag: `takes_receiver` is now set
                                        // wherever the tag is, so the signature answers it.
                                        let receiver_first = setter_arity.is_none_or(|a| a >= 2);
                                        if receiver_first {
                                            // The receiver is an ARGUMENT here, so the
                                            // callee already has it and the ambient binding
                                            // is not merely redundant — rebinding the global
                                            // around the call disturbs an enclosing method's
                                            // own `this` in languages that read it there,
                                            // which turned a PHP `$this->n++` inside a
                                            // method into NaN.
                                            // ⛔ `invoke_with_receiver`, NOT `invoke` with
                                            // the receiver in the argument list: under
                                            // `Parameter` a plain `invoke` PREPENDS one
                                            // too, and the setter then took the prepended
                                            // value as `this` and the receiver as its
                                            // value parameter. One receiver, passed once.
                                            ctx.invoke_with_receiver(&setter_val, receiver, &[val]);
                                        } else {
                                            // Bind `this` around the SAME invoke path the
                                            // arity branch used. `invoke_with_explicit_this`
                                            // additionally re-resolves the callee (proxy
                                            // apply trap, bound-args unwrapping), and for a
                                            // closure-wrapped object-literal accessor that
                                            // resolution lands somewhere the setter body is
                                            // never reached — the assignment silently did
                                            // nothing. Binding the receiver is the part
                                            // §10.1.5 step 6.b requires; re-resolving the
                                            // callee is not.
                                            // §20.2.3.3's own [[Call]] path: it binds
                                            // `this` AND resolves a bound function or proxy
                                            // apply trap, both of which a `defineProperty`
                                            // setter can be. `set_js_this` + a bare invoke
                                            // binds the receiver but skips that resolution,
                                            // and the setter then never runs.
                                            crate::function::invoke_with_explicit_this(
                                                ctx,
                                                &setter_val,
                                                receiver,
                                                &[val],
                                            );
                                        }
                                        obj.lock()
                                            .unwrap()
                                            .properties
                                            .shift_remove(ACCESSOR_SETTER_ACTIVE_MARK);
                                        return Value::Null;
                                    }
                                }
                            }
                            let symbol_to_track = match key_raw {
                                Value::Symbol(sym) => Some(Value::Symbol(sym.clone())),
                                _ => None,
                            };
                            let (symbol_track, keys_to_append) = {
                                let mut o = obj.lock().unwrap();
                                if o.properties.get(FROZEN_MARK).is_some()
                                    && o.properties.get(ACCESSOR_SETTER_ACTIVE_MARK).is_none()
                                    && !is_accessor_backing_slot_write(&o, key)
                                {
                                    drop(o);
                                    if strict {
                                        ctx.throw_value(crate::error::new_error(
                                            ctx,
                                            "TypeError",
                                            "Cannot assign to read only property of frozen object",
                                        ));
                                        return Value::Undefined;
                                    }
                                    return Value::Null;
                                }
                                if let ObjectKind::Array(values) = &mut o.kind {
                                    if let Some(index) = array_index_key(key) {
                                        if (index as usize) < values.len() {
                                            values[index as usize] = val;
                                            return Value::Null;
                                        }
                                    }
                                }
                                let kind_skip = matches!(o.kind, ObjectKind::Array(_));
                                let keys_to_append = if !kind_skip
                                    && symbol_to_track.is_none()
                                    && !key.starts_with("__")
                                {
                                    prepare_track_key_in(&mut o, key)
                                } else {
                                    None
                                };
                                // For typed objects (Error etc.), also update the fields Vec.
                                // Error types have "message" at field index 0.
                                if o.type_id > 0 && key == "message" && !o.fields.is_empty() {
                                    o.fields[0] = val.clone();
                                }
                                if let Some(slot) = o.properties.get_mut(key) {
                                    *slot = val;
                                } else {
                                    o.properties.insert(key.to_owned(), val);
                                }
                                let symbol_track = if kind_skip { None } else { symbol_to_track };
                                (symbol_track, keys_to_append)
                            };
                            if let Some(sym) = symbol_track {
                                track_sym_key(&obj, sym);
                            }
                            if let Some(keys_arc) = keys_to_append {
                                append_tracked_key(keys_arc, key);
                            }
                            Value::Null
                        })
                    })
                });
            }
            Value::Null
        }),
    );

    // has(obj, key) -> i32 (walks prototype chain, returns 1/0)
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "has",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let undefined = Value::Undefined;
                let key_raw = args.get(1).unwrap_or(&undefined);
                {
                    let o = obj.lock().unwrap();
                    if let ObjectKind::Array(ref values) = o.kind {
                        if matches!(key_raw, Value::String(text) if text.as_ref() == "length") {
                            return Value::Bool(true);
                        }
                        if let Some(index) = crate::keys::canonical_array_index_value(key_raw) {
                            let index = index as usize;
                            if index < values.len() && !crate::array::is_array_hole(&o, index) {
                                return Value::Bool(true);
                            }
                        }
                    }
                    if let ObjectKind::TypedArray(ref ta) = o.kind {
                        if let Some(index) = crate::keys::canonical_array_index_value(key_raw) {
                            return Value::Bool(
                                (index as usize) < crate::typedarray::ta_live_length(ta),
                            );
                        }
                    }
                    // Share the exotic-check lock for ordinary string-key
                    // own hits. No value clone or key allocation is needed.
                    if let Value::String(key) = key_raw {
                        if o.properties.contains_key(key.as_ref()) {
                            return Value::Bool(true);
                        }
                    }
                }
                return with_property_key(key_raw, |key| {
                    // The chain walker includes the receiver's own map;
                    // probing it separately only adds a redundant lock.
                    Value::Bool(proto_walk_has(&obj, key))
                });
            }
            Value::Bool(false)
        }),
    );

    // hasIn(obj, key) -> bool (own + prototype chain walk). Backs the
    // JS `in` operator per ECMA-262 §13.10.1 — distinct from `hasOwn`
    // which is own-only. The compiler routes `key in obj` here so
    // `Object.create(proto)` chains resolve correctly.
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "hasIn",
        Box::new(|_ctx, args| {
            let undefined = Value::Undefined;
            let key_raw = args.get(1).unwrap_or(&undefined);
            if let Some(obj) = obj_of(args, 0) {
                let array_index = crate::keys::canonical_array_index_value(key_raw);
                return with_property_key(key_raw, |property_key| {
                    // Walk own + __proto__ chain. Bound at 100 hops to
                    // protect against accidental cycles.
                    let mut current = obj.clone();
                    for _ in 0..100 {
                        let next_proto = {
                            let o = current.lock().unwrap();
                            let found = match &o.kind {
                                ObjectKind::Array(v) => {
                                    let in_range = match array_index {
                                        Some(i) => (i as usize) < v.len(),
                                        None => false,
                                    };
                                    // Array holes (set by `delete arr[i]`)
                                    // make `i in arr` return false per
                                    // ECMA-262 §13.5.1 step 5.b.iii.
                                    match array_index {
                                        Some(i) if in_range && is_array_hole(&o, i as i32) => false,
                                        _ => in_range,
                                    }
                                }
                                ObjectKind::Map(m) => m.contains_key(key_raw),
                                ObjectKind::Set(s) => s.contains(key_raw),
                                _ => o.properties.contains_key(property_key),
                            };
                            if found {
                                return Value::Bool(true);
                            }
                            // §13.5.1: `in` walks the prototype chain. A bare
                            // VM-created object has no explicit `__proto__`, so
                            // resolve its [[Prototype]] by kind (Object/Array
                            // prototype) — otherwise `"toString" in {}` would
                            // miss the inherited method. NOTE: must use the
                            // already-held guard `o` here — calling the locking
                            // `js_prototype_of` on `current` would re-lock the
                            // same Mutex and deadlock.
                            match o.properties.get(PROTO_KEY).cloned() {
                                Some(Value::Object(p)) => Some(p),
                                Some(_) => None, // explicit null proto → chain ends
                                None => match &o.kind {
                                    ObjectKind::Array(_) => {
                                        match crate::array::shared_array_prototype() {
                                            Value::Object(p) => Some(p),
                                            _ => None,
                                        }
                                    }
                                    _ => match shared_object_prototype() {
                                        Value::Object(p) => Some(p),
                                        _ => None,
                                    },
                                },
                            }
                        };
                        match next_proto {
                            Some(p) => current = p,
                            None => break,
                        }
                    }
                    Value::Bool(false)
                });
            }
            Value::Bool(false)
        }),
    );

    // hasOwn(obj, key) -> bool (own-only, no prototype walk). Polymorphic
    // over Array / Map / Ordinary. Backs JS `Object.hasOwn` + `in`
    // operator, PHP `array_key_exists`, Python `key in dict`, Ruby
    // `Hash#key?`. Returns Value::Bool so string coercion gives
    // "true"/"false" (ECMA-262 §23.1.2.3).
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "hasOwn",
        Box::new(|ctx, args| {
            let undefined = Value::Undefined;
            let key_raw = args.get(1).unwrap_or(&undefined);
            let target = args.first().unwrap_or(&undefined);
            if matches!(target, Value::Null | Value::Undefined) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot convert undefined or null to object",
                ));
                return Value::Undefined;
            }
            if let Some(desc) =
                crate::proxy::get_own_property_descriptor_dispatch(ctx, target, key_raw)
            {
                return Value::Bool(matches!(desc, Value::Object(_)));
            }
            Value::Bool(has_own_property_key(target, key_raw).unwrap_or(false))
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
    crate::perf::register_host_fn(
        vm,
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
                let undefined = Value::Undefined;
                let key_value = args.get(1).unwrap_or(&undefined);
                with_property_key(key_value, |key| {
                    if !key.starts_with("__") {
                        track_key(&obj, key);
                    }
                });
            }
            Value::Undefined
        }),
    );

    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "delete",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let undefined = Value::Undefined;
                let key_raw = args.get(1).unwrap_or(&undefined);
                let key = key_string(key_raw);
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
                    let idx = match key_raw {
                        Value::I32(n) => Some(*n as usize),
                        Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
                        Value::String(s) => crate::keys::non_negative_integer_index_key(s),
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
                                    let a =
                                        vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
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
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let idx = match key_raw {
                        Value::I32(n) if *n >= 0 => Some(*n as usize),
                        Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
                        Value::String(s) => crate::keys::non_negative_integer_index_key(s),
                        _ => None,
                    };
                    if matches!(idx, Some(i) if i < crate::typedarray::ta_live_length(ta)) {
                        drop(o);
                        _ctx.throw_value(crate::error::new_error(
                            _ctx,
                            "TypeError",
                            "Cannot delete typed array indexed property",
                        ));
                        return Value::Undefined;
                    }
                }
                // Map entry delete: remove from the IndexMap backing.
                // Polymorphism: PHP `array` stores assoc data as Map, so
                // `unset($arr[$k])` lands here when `$arr` is a Map kind.
                // Without this branch, the Ordinary fallback below tries
                // `properties.remove` which doesn't touch the Map data
                // (Map keys live in `kind`, not `properties`).
                if let ObjectKind::Map(ref mut m) = o.kind {
                    let key_value = match key_raw {
                        Value::Undefined | Value::Null => crate::keys::string_value(&key),
                        other => other.clone(),
                    };
                    let removed = match key_raw {
                        Value::Undefined | Value::Null => m.shift_remove(&key_value).is_some(),
                        other => m.shift_remove(other).is_some(),
                    };
                    return Value::Bool(removed);
                }
                if is_nonconfig(&o, &key) {
                    return Value::Bool(false);
                }
                // ⛔ AN ACCESSOR IS NOT STORED UNDER ITS OWN NAME. A getter/
                // setter lives as `__get_<k>` / `__set_<k>`, so removing only
                // `k` left the accessor intact and `delete o.g` was a silent
                // no-op — `o.g` kept returning 7 where ECMA-262 §13.5.1 says
                // the whole property goes. Measured against node, which
                // returns `undefined`.
                let removed_accessor = crate::keys::with_getter_property_key(&key, |getter_key| {
                    o.properties.shift_remove(getter_key).is_some()
                }) | crate::keys::with_setter_property_key(
                    &key,
                    |setter_key| o.properties.shift_remove(setter_key).is_some(),
                );
                let existed = o.properties.shift_remove(&key).is_some() | removed_accessor;
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
        let mut keys = Vec::with_capacity(obj.properties.len());
        for key in obj.properties.keys() {
            if !key.starts_with("__") {
                keys.push(key.clone());
            }
        }
        keys
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
        if let Some(keys) = plain_own_string_keys(o) {
            return keys;
        }
        // Direct property assignments (`obj.foo = 1` via Op::STRUCT_SET)
        // don't touch the `__keys` tracker — only the dict-literal
        // emitter and `defineProperty` do. So when __keys is shorter
        // than the live property set, append the untracked keys at
        // the end (preserves declared order, then append new writes).
        let tracked: Option<Vec<String>> = o.properties.get("__keys").and_then(|v| {
            if let Value::Object(arr) = v {
                let ka = arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ka.kind {
                    let mut keys = Vec::with_capacity(elems.len());
                    for e in elems {
                        if let Value::String(s) = e {
                            if o.properties.contains_key(s.as_ref()) {
                                keys.push(s.to_string());
                            }
                        }
                    }
                    return Some(keys);
                }
            }
            None
        });
        // Symbol-keyed properties tracked separately — ECMA-262 §7.3.22
        // excludes them from `Object.keys` / `Object.entries`.
        let sym_keys: Option<std::collections::HashSet<String>> =
            o.properties.get("__sym_keys").and_then(|v| {
                if let Value::Object(a) = v {
                    let lock = a.lock().unwrap();
                    if let ObjectKind::Array(ref el) = lock.kind {
                        let mut keys = std::collections::HashSet::with_capacity(el.len());
                        for e in el {
                            match e {
                                Value::String(s) => {
                                    keys.insert(s.to_string());
                                }
                                Value::Symbol(sym) => {
                                    keys.insert(crate::symbol::canonical_property_key(sym));
                                }
                                _ => {}
                            }
                        }
                        return Some(keys);
                    }
                }
                None
            });
        let mut live = Vec::with_capacity(o.properties.len());
        for key in o.properties.keys() {
            let is_symbol_key = matches!(sym_keys.as_ref(), Some(keys) if keys.contains(key));
            if !key.starts_with("__") && !is_symbol_key {
                live.push(key.clone());
            }
        }
        match tracked {
            Some(mut tk) => {
                let mut seen: std::collections::HashSet<&str> =
                    std::collections::HashSet::with_capacity(tk.len());
                for key in &tk {
                    seen.insert(key.as_str());
                }
                let mut extras: Vec<String> =
                    Vec::with_capacity(live.len().saturating_sub(tk.len()));
                for k in &live {
                    if !seen.contains(k.as_str()) {
                        extras.push(k.clone());
                        seen.insert(k.as_str());
                    }
                }
                tk.extend(extras);
                sort_array_indices_first(&mut tk);
                tk
            }
            None => {
                let mut keys = live;
                sort_array_indices_first(&mut keys);
                keys
            }
        }
    }

    /// Like `ordinary_ordered_keys` but filters out keys flagged
    /// non-enumerable via `defineProperty({enumerable: false})`. Used
    /// by `Object.keys` / `Object.values` / `Object.entries` per
    /// ECMA-262 §7.3.22 (only enumerable own properties).
    fn ordinary_enumerable_keys(o: &Object) -> Vec<String> {
        let mut keys = descriptor_own_keys(o);
        keys.retain(|k| !is_nonenum(o, k));
        keys
    }

    object_unary(
        vm,
        "keys",
        vec![obj_t()],
        Box::new(|ctx, args| {
            let Some(raw_value) = args.first() else {
                return throw_type_error(ctx, "Cannot convert undefined or null to object");
            };
            let Some(value) = to_object_for_object_static(
                ctx,
                raw_value,
                "Cannot convert undefined or null to object",
            ) else {
                return Value::Undefined;
            };
            // §20.1.2.17 Object.keys routes through [[OwnPropertyKeys]] —
            // for proxy exotic objects that is the ownKeys trap.
            if let Some(keys) = crate::proxy::own_keys_dispatch(ctx, &value) {
                let elements = array_elements(&keys);
                let mut filtered = Vec::with_capacity(elements.len());
                for key in elements {
                    if !matches!(&key, Value::String(_)) {
                        continue;
                    }
                    let enumerable = {
                        let desc =
                            crate::proxy::get_own_property_descriptor_dispatch(ctx, &value, &key)
                                .unwrap_or(Value::Undefined);
                        matches!(desc, Value::Object(_))
                            && descriptor_bool(&desc, "enumerable", false)
                    };
                    if enumerable {
                        filtered.push(key);
                    }
                }
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(filtered)));
            }
            if let Value::Object(obj) = value {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let mut keys = Vec::with_capacity(v.len());
                        let holes = array_hole_set(&o);
                        for index in 0..v.len() {
                            if !cached_array_hole_contains(&holes, index) {
                                keys.push(index_key_value(index));
                            }
                        }
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
                    }
                    ObjectKind::Map(m) => {
                        let mut keys = Vec::with_capacity(m.len());
                        keys.extend(m.keys().cloned());
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
                    }
                    // Set keys() iterator yields each element (key === value
                    // for Sets per spec); for-of uses values() but keys() is
                    // also reachable for the symmetry used by entries().
                    ObjectKind::Set(s) => {
                        let mut keys = Vec::with_capacity(s.len());
                        keys.extend(s.iter().cloned());
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
                    }
                    _ => {}
                }
                if has_plain_ordered_data_properties(&o) {
                    let keys = o
                        .properties
                        .keys()
                        .map(|key| crate::keys::string_value(key))
                        .collect();
                    return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
                }
                let enumerable = ordinary_enumerable_keys(&o);
                let mut keys = Vec::with_capacity(enumerable.len());
                for key in enumerable {
                    keys.push(crate::keys::string_value(&key));
                }
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
        }),
    );

    // `for...in` semantics: enumerable property KEYS including those
    // inherited from the prototype chain (ECMA-262 §14.7.5.6 step 8.b
    // — the `EnumerableOwnPropertyNames` operation runs at every level
    // of the chain). Distinct from `Object.keys` (own + enumerable
    // only). The compiler emits this for `for (k in obj)` loops.
    object_unary(
        vm,
        "iterForIn",
        vec![obj_t()],
        Box::new(|ctx, args| {
            if let Some(value) = args.first() {
                if let Some(keys) = crate::proxy::own_keys_dispatch(ctx, value) {
                    let elements = array_elements(&keys);
                    let mut filtered = Vec::with_capacity(elements.len());
                    for key in elements {
                        if !matches!(&key, Value::String(_)) {
                            continue;
                        }
                        let enumerable = {
                            let desc = crate::proxy::get_own_property_descriptor_dispatch(
                                ctx, value, &key,
                            )
                            .unwrap_or(Value::Undefined);
                            matches!(desc, Value::Object(_))
                                && descriptor_bool(&desc, "enumerable", false)
                        };
                        if enumerable {
                            filtered.push(key);
                        }
                    }
                    return Value::Object(vybe_runtime::heap::alloc(Object::new_array(filtered)));
                }
            }
            if let Some(obj) = obj_of(args, 0) {
                let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
                let mut out: Vec<Value> = Vec::new();
                let mut current = obj;
                for _ in 0..100 {
                    let next_proto = {
                        let o = current.lock().unwrap();
                        match &o.kind {
                            ObjectKind::Array(v) => {
                                seen.reserve(v.len());
                                out.reserve(v.len());
                                let holes = array_hole_set(&o);
                                for i in 0..v.len() {
                                    if cached_array_hole_contains(&holes, i) {
                                        continue;
                                    }
                                    crate::keys::with_index_key(i, |k| {
                                        if seen.insert(k.to_owned()) {
                                            out.push(crate::keys::string_value(k));
                                        }
                                    });
                                }
                            }
                            ObjectKind::Map(m) => {
                                seen.reserve(m.len());
                                out.reserve(m.len());
                                for k in m.keys() {
                                    crate::keys::with_property_key(k, |ks| {
                                        if seen.insert(ks.to_owned()) {
                                            out.push(crate::keys::string_value(ks));
                                        }
                                    });
                                }
                            }
                            ObjectKind::TypedArray(ta) => {
                                let live = crate::typedarray::ta_live_length(ta);
                                seen.reserve(live);
                                out.reserve(live);
                                for i in 0..live {
                                    crate::keys::with_index_key(i, |k| {
                                        if seen.insert(k.to_owned()) {
                                            out.push(crate::keys::string_value(k));
                                        }
                                    });
                                }
                            }
                            _ => {
                                let keys = ordinary_enumerable_keys(&o);
                                seen.reserve(keys.len());
                                out.reserve(keys.len());
                                for k in keys {
                                    if seen.insert(k.clone()) {
                                        out.push(crate::keys::string_value(&k));
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
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(out)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
        }),
    );

    // `for...of` semantics: yields whatever Symbol.iterator returns —
    // Array/Set yields values, Map yields [key, value] pairs (per
    // ECMA-262 §24.1.3.12 / §24.2.3.11). The compiler emits this for
    // for-of loops; `Object.values` keeps the spec-strict "values only"
    // behaviour for `Object.values(map)` user calls.
    object_unary(
        vm,
        "iterForOf",
        vec![obj_t()],
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let mut values = Vec::with_capacity(v.len());
                        let holes = array_hole_set(&o);
                        for (index, value) in v.iter().enumerate() {
                            values.push(if cached_array_hole_contains(&holes, index) {
                                Value::Undefined
                            } else {
                                value.clone()
                            });
                        }
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                    }
                    ObjectKind::Map(m) => {
                        let mut entries = Vec::with_capacity(m.len());
                        for (k, v) in m {
                            entries.push(crate::array::make_pair_array(k.clone(), v.clone()));
                        }
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                            entries,
                        )));
                    }
                    ObjectKind::Set(s) => {
                        let mut vals = Vec::with_capacity(s.len());
                        vals.extend(s.iter().cloned());
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(vals)));
                    }
                    ObjectKind::TypedArray(ta) => {
                        let len = crate::typedarray::ta_live_length(ta);
                        let vals = typed_array_values_snapshot(ta, len);
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(vals)));
                    }
                    _ => {}
                }
                drop(o);
                // Try iterator protocol. If the iterator() call returns a
                // generator Continuation, pass it through so the bytecode
                // caller can drain via stack-switching (host can't resume).
                if let Some(cont) = call_iterator_if_generator(ctx, &obj, "asyncIterator")
                    .or_else(|| call_iterator_if_generator(ctx, &obj, "iterator"))
                {
                    return cont;
                }
                if let Some(values) = collect_protocol_iterable(ctx, &obj, "asyncIterator") {
                    return values;
                }
                if let Some(values) = collect_protocol_iterable(ctx, &obj, "iterator") {
                    return values;
                }
                let o = obj.lock().unwrap();
                if let Some(len_val) = o.properties.get("length") {
                    let len = len_val.as_f64().max(0.0) as usize;
                    let mut values = Vec::with_capacity(len);
                    for i in 0..len {
                        values.push(crate::keys::with_index_key(i, |key| {
                            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
                        }));
                    }
                    return Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                }
                let values: Vec<Value> = ordinary_ordered_keys(&o)
                    .into_iter()
                    .filter_map(|k| o.properties.get(&k).cloned())
                    .collect();
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
            }
            // Strings are iterable per code-point — for-of of a string
            // yields each character. Match here so emit_iter_values can
            // be a single dispatch point.
            if let Some(Value::String(s)) = args.first() {
                let mut chars = Vec::with_capacity(s.len());
                for ch in s.chars() {
                    chars.push(char_value(ch));
                }
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(chars)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
        }),
    );

    object_unary(
        vm,
        "values",
        vec![obj_t()],
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let mut values = Vec::with_capacity(v.len());
                        let holes = array_hole_set(&o);
                        for (index, value) in v.iter().enumerate() {
                            if !cached_array_hole_contains(&holes, index) {
                                values.push(value.clone());
                            }
                        }
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                    }
                    ObjectKind::Map(m) => {
                        let mut vals = Vec::with_capacity(m.len());
                        vals.extend(m.values().cloned());
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(vals)));
                    }
                    // Set iteration order = insertion order; values() of a Set
                    // returns its elements (matches ECMA-262 §24.2.3.10 and is
                    // what `for...of s` lowers to via emit_iter_values).
                    ObjectKind::Set(s) => {
                        let mut vals = Vec::with_capacity(s.len());
                        vals.extend(s.iter().cloned());
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(vals)));
                    }
                    _ => {}
                }
                if has_plain_ordered_data_properties(&o) {
                    let values = o.properties.values().cloned().collect();
                    return Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                }
                let keys = ordinary_enumerable_keys(&o);
                let mut values = Vec::with_capacity(keys.len());
                if !has_any_getter(&o) {
                    let mut direct_data = true;
                    for k in &keys {
                        match o.properties.get(k) {
                            Some(value) => values.push(value.clone()),
                            None => {
                                direct_data = false;
                                values.clear();
                                break;
                            }
                        }
                    }
                    if direct_data {
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
                    }
                }
                drop(o);
                for k in keys {
                    values.push(assign_source_get(ctx, &obj, &k));
                }
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
        }),
    );

    object_unary(
        vm,
        "entries",
        vec![obj_t()],
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let mut entries = Vec::with_capacity(v.len());
                        let holes = array_hole_set(&o);
                        for (i, val) in v.iter().enumerate() {
                            if !cached_array_hole_contains(&holes, i) {
                                entries.push(crate::array::make_pair_array(
                                    Value::I32(i as i32),
                                    val.clone(),
                                ));
                            }
                        }
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                            entries,
                        )));
                    }
                    ObjectKind::Map(m) => {
                        let mut entries = Vec::with_capacity(m.len());
                        for (k, v) in m {
                            entries.push(crate::array::make_pair_array(k.clone(), v.clone()));
                        }
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                            entries,
                        )));
                    }
                    // Set entries() per spec yields [value, value] pairs.
                    ObjectKind::Set(s) => {
                        let mut entries = Vec::with_capacity(s.len());
                        for v in s {
                            entries.push(crate::array::make_pair_array(v.clone(), v.clone()));
                        }
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                            entries,
                        )));
                    }
                    _ => {}
                }
                if let Some(entries) = plain_ordered_data_entries(&o) {
                    return Value::Object(vybe_runtime::heap::alloc(Object::new_array(entries)));
                }
                let keys = ordinary_enumerable_keys(&o);
                let mut entries = Vec::with_capacity(keys.len());
                if !has_any_getter(&o) {
                    let mut direct_data = true;
                    for k in &keys {
                        match o.properties.get(k) {
                            Some(value) => entries.push(crate::array::make_pair_array(
                                crate::keys::string_value(k),
                                value.clone(),
                            )),
                            None => {
                                direct_data = false;
                                entries.clear();
                                break;
                            }
                        }
                    }
                    if direct_data {
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                            entries,
                        )));
                    }
                }
                drop(o);
                for k in keys {
                    let v = assign_source_get(ctx, &obj, &k);
                    entries.push(crate::array::make_pair_array(
                        crate::keys::string_value(&k),
                        v,
                    ));
                }
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(entries)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
        }),
    );

    // getOwnPropertyNames — like keys but includes non-enumerable.
    // Our model doesn't track enumerability separately, so this is the
    // same as `keys`. Use the insertion-order tracker (`__keys`) per
    // ECMA-262 §7.3.21 ordering requirements; HashMap iteration alone
    // is non-deterministic.
    object_unary(
        vm,
        "getOwnPropertyNames",
        vec![obj_t()],
        Box::new(|ctx, args| {
            // §20.1.2.10 routes through [[OwnPropertyKeys]] — the ownKeys
            // trap for proxy exotic objects.
            if let Some(value) = args.first() {
                if let Some(keys) = crate::proxy::own_keys_dispatch(ctx, value) {
                    return keys;
                }
            }
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let names = ordinary_ordered_keys(&o);
                let mut keys = Vec::with_capacity(names.len());
                for k in names {
                    keys.push(crate::keys::string_value(&k));
                }
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
        }),
    );

    // getOwnPropertySymbols — returns Value::Symbol for each key tracked in __sym_keys.
    // Symbol-keyed props are stored as "Symbol(<desc>)" string keys; we recover
    // the description and return the original Symbol so obj[syms[0]] round-trips.
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "getOwnPropertySymbols",
        Box::new(|ctx, args| {
            if let Some(value) = args.first() {
                if let Some(keys) = crate::proxy::own_keys_dispatch(ctx, value) {
                    let elements = array_elements(&keys);
                    let mut syms = Vec::with_capacity(elements.len());
                    for key in elements {
                        if matches!(&key, Value::Symbol(_))
                            || matches!(&key, Value::String(s) if s.starts_with("Symbol("))
                        {
                            syms.push(key);
                        }
                    }
                    return Value::Object(vybe_runtime::heap::alloc(Object::new_array(syms)));
                }
            }
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let syms: Vec<Value> = match o.properties.get("__sym_keys") {
                    Some(Value::Object(arr)) => {
                        let a = arr.lock().unwrap();
                        if let ObjectKind::Array(ref elems) = a.kind {
                            let mut syms = Vec::with_capacity(elems.len());
                            for e in elems {
                                if let Value::Symbol(sym) = e {
                                    syms.push(Value::Symbol(sym.clone()));
                                }
                            }
                            syms
                        } else {
                            Vec::new()
                        }
                    }
                    _ => Vec::new(),
                };
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(syms)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
        }),
    );

    object_unary(
        vm,
        "length",
        vec![ValType::I32],
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
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "defineProperty",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let original_obj = obj.clone();
                let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
                let descriptor = args.get(2).cloned().unwrap_or(Value::Undefined);
                let mut define_obj = obj.clone();
                if proxy_target_and_handler(&obj).is_some() {
                    if let Some(result) = crate::proxy::define_property_dispatch(
                        ctx,
                        &obj,
                        key_value.clone(),
                        descriptor.clone(),
                    ) {
                        if matches!(result, Value::Undefined) {
                            return Value::Undefined;
                        }
                        if crate::boolean::to_boolean(&result) {
                            return Value::Object(original_obj);
                        }
                    }
                    if let Some((target, _)) = proxy_target_and_handler(&obj) {
                        if let Value::Object(target_obj) = target {
                            define_obj = target_obj;
                        }
                    } else {
                        return Value::Object(original_obj);
                    }
                }
                let key = key_string(&key_value);
                let getter_key = getter_property_key(&key);
                let setter_key = setter_property_key(&key);
                // §10.1.6.3: a NEW key on a non-extensible object is
                // rejected — Object.defineProperty surfaces that as
                // TypeError (§20.1.2.4; Reflect's form returns false).
                {
                    let o = define_obj.lock().unwrap();
                    let exists = o.properties.contains_key(&key)
                        || o.properties.contains_key(&getter_key)
                        || o.properties.contains_key(&setter_key);
                    if !exists && is_not_extensible(&o) {
                        drop(o);
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "TypeError",
                            "Cannot define property, object is not extensible",
                        ));
                        return Value::Undefined;
                    }
                }
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
                            let has_value = d.properties.contains_key("value");
                            let has_writable = d.properties.contains_key("writable");
                            let has_get = d.properties.contains_key("get");
                            let has_set = d.properties.contains_key("set");
                            if (has_value || has_writable) && (has_get || has_set) {
                                drop(d);
                                ctx.throw_value(crate::error::new_error(
                                    ctx,
                                    "TypeError",
                                    "Invalid property descriptor",
                                ));
                                return Value::Undefined;
                            }
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
                {
                    let mut o = define_obj.lock().unwrap();
                    if matches!(o.kind, ObjectKind::Array(_)) && key == "length" {
                        if let Some(v) = val_or_none.as_ref() {
                            crate::array::apply_js_array_length(ctx, &mut o, v);
                        }
                        if matches!(writable, Some(false)) {
                            o.properties
                                .insert("__array_length_readonly".into(), Value::Bool(true));
                        }
                        return Value::Object(original_obj);
                    }
                }
                if matches!(key_value, Value::Symbol(_)) {
                    track_sym_key(&define_obj, key_value.clone());
                } else {
                    track_key(&define_obj, &key);
                }
                if !enumerable {
                    track_nonenum(&define_obj, &key);
                }
                {
                    let mut o = define_obj.lock().unwrap();
                    if o.properties.contains_key(&key) && is_nonconfig(&o, &key) {
                        let current_value = o.properties.get(&key).cloned();
                        let value_changes = val_or_none
                            .as_ref()
                            .zip(current_value.as_ref())
                            .is_some_and(|(next, current)| next != current);
                        let tries_configurable = configurable;
                        let writable_false_transition = matches!(writable, Some(false));
                        if value_changes || tries_configurable {
                            ctx.throw_value(crate::error::new_error(
                                ctx,
                                "TypeError",
                                "Cannot redefine property",
                            ));
                            return Value::Null;
                        }
                        if writable_false_transition {
                            let noop_idx =
                                NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
                            if noop_idx > 0 {
                                let mut noop_obj = Object::new();
                                noop_obj.kind = ObjectKind::HostFunction(noop_idx);
                                let noop_val = Value::Object(vybe_runtime::heap::alloc(noop_obj));
                                let setter_key = setter_property_key(&key);
                                if !o.properties.contains_key(&setter_key) {
                                    o.properties.insert(setter_key, noop_val);
                                }
                            }
                        }
                        return Value::Object(original_obj);
                    }
                    let getter_key = getter_property_key(&key);
                    let setter_key = setter_property_key(&key);
                    if let Some(g) = getter {
                        o.properties.insert(getter_key.clone(), g);
                    }
                    if let Some(s) = setter {
                        o.properties.insert(setter_key.clone(), s);
                    }
                    if let Some(v) = val_or_none {
                        // §10.1.6.3: converting an accessor property to a data
                        // property clears [[Get]]/[[Set]]. Required, not
                        // cosmetic — a property read checks `__get_<key>`
                        // BEFORE the data slot, so a stale accessor left here
                        // would keep shadowing the value being written. Class
                        // getters bind `__get_<name>` onto the instance, so
                        // `defineProperty(this, "value", { value: … })` in a
                        // subclass constructor hits exactly this case.
                        o.properties.shift_remove(&getter_key);
                        o.properties.shift_remove(&setter_key);
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
                                let noop_val = Value::Object(vybe_runtime::heap::alloc(noop_obj));
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
            ctx.throw_value(crate::error::new_error(
                ctx,
                "TypeError",
                "Object.defineProperty called on non-object",
            ));
            Value::Undefined
        }),
    );

    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "defineProperties",
        Box::new(|_ctx, args| {
            if let (Some(target), Some(Value::Object(descs))) = (obj_of(args, 0), args.get(1)) {
                // Collect entries in the descriptor object's __keys
                // insertion order so the target's iteration matches
                // (ECMA-262 §20.1.2.4 step 5 — `OwnPropertyKeys` over
                // descriptors). HashMap iter would non-deterministically
                // shuffle on every run.
                let entries: Vec<(String, Value)> = {
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
                        .filter_map(|k| d.properties.get(&k).cloned().map(|v| (k, v)))
                        .collect()
                };
                for (k, desc_value) in entries {
                    let Value::Object(desc) = desc_value else {
                        continue;
                    };
                    let dlock = desc.lock().unwrap();
                    let val = dlock.properties.get("value").cloned();
                    let getter = dlock.properties.get("get").cloned();
                    let setter = dlock.properties.get("set").cloned();
                    let enumerable = dlock
                        .properties
                        .get("enumerable")
                        .map(|x| x.as_bool())
                        .unwrap_or(false);
                    let writable = dlock.properties.get("writable").map(|x| x.as_bool());
                    let configurable = dlock
                        .properties
                        .get("configurable")
                        .map(|x| x.as_bool())
                        .unwrap_or(false);
                    drop(dlock);
                    track_key(&target, &k);
                    if !enumerable {
                        track_nonenum(&target, &k);
                    }
                    {
                        let mut o = target.lock().unwrap();
                        let getter_key = getter_property_key(&k);
                        let setter_key = setter_property_key(&k);
                        if let Some(g) = getter {
                            o.properties.insert(getter_key.clone(), g);
                        }
                        if let Some(s) = setter {
                            o.properties.insert(setter_key.clone(), s);
                        }
                        if let Some(v) = val {
                            o.properties.shift_remove(&getter_key);
                            o.properties.shift_remove(&setter_key);
                            o.properties.insert(k.clone(), v);
                            if matches!(writable, Some(false) | None) {
                                let noop_idx =
                                    NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
                                if noop_idx > 0 {
                                    let mut noop_obj = Object::new();
                                    noop_obj.kind = ObjectKind::HostFunction(noop_idx);
                                    let noop_val =
                                        Value::Object(vybe_runtime::heap::alloc(noop_obj));
                                    o.properties.insert(setter_key, noop_val);
                                }
                            }
                        } else if !o.properties.contains_key(&k) {
                            o.properties.insert(k.clone(), Value::Undefined);
                        }
                    }
                    if !configurable {
                        track_nonconfig(&target, &k);
                    }
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
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "getOwnPropertyDescriptor",
        Box::new(|ctx, args| {
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            if matches!(target, Value::Null | Value::Undefined) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot convert undefined or null to object",
                ));
                return Value::Undefined;
            }
            if let Some(obj) = obj_of(args, 0) {
                let key_value = args.get(1).unwrap_or(&undefined);
                // §10.5.5: proxies answer via their trap (with the
                // non-configurable invariant enforced) or their target.
                if let Some(proxy_desc) =
                    crate::proxy::get_own_property_descriptor_dispatch(ctx, target, key_value)
                {
                    return proxy_desc;
                }
                return with_property_key(key_value, |key| own_property_descriptor(&obj, key));
            }
            if matches!(target, Value::String(_)) {
                let undefined = Value::Undefined;
                let key_value = args.get(1).unwrap_or(&undefined);
                return with_property_key(key_value, |key| string_length_descriptor(&target, key));
            }
            Value::Undefined
        }),
    );

    // getOwnPropertyDescriptors(obj) -> { key: descriptor, ... }
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "getOwnPropertyDescriptors",
        Box::new(|ctx, args| {
            if matches!(args.first(), Some(Value::Null | Value::Undefined) | None) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot convert undefined or null to object",
                ));
                return Value::Undefined;
            }
            let result = vybe_runtime::heap::alloc(Object::new());
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let keys = descriptor_own_keys(&o);
                drop(o);
                result.lock().unwrap().properties.reserve(keys.len());
                for k in keys {
                    let desc = own_property_descriptor(&obj, &k);
                    if matches!(desc, Value::Undefined) {
                        continue;
                    }
                    if k.starts_with("Symbol(") {
                        result.lock().unwrap().properties.insert(k, desc);
                    } else {
                        track_key(&result, &k);
                        result.lock().unwrap().properties.insert(k, desc);
                    }
                }
            } else if let Some(value @ Value::String(_)) = args.first() {
                let desc = string_length_descriptor(value, "length");
                if !matches!(desc, Value::Undefined) {
                    track_key(&result, "length");
                    result
                        .lock()
                        .unwrap()
                        .properties
                        .insert("length".into(), desc);
                }
            }
            Value::Object(result)
        }),
    );
}

/// §20.1.2.8 core for an ORDINARY object: fresh accessor/data descriptor,
/// or Undefined when `key` is not an own property. Proxy callers resolve
/// their target first and pass it here.
pub fn own_property_descriptor(obj: &Arc<Mutex<Object>>, key: &str) -> Value {
    let o = obj.lock().unwrap();
    if let ObjectKind::Array(values) = &o.kind {
        if key == "length" {
            let mut desc = Object::new();
            desc.properties.reserve(4);
            desc.properties
                .insert("value".into(), Value::I32(values.len() as i32));
            desc.properties.insert("writable".into(), Value::Bool(true));
            desc.properties
                .insert("enumerable".into(), Value::Bool(false));
            desc.properties
                .insert("configurable".into(), Value::Bool(false));
            return Value::Object(vybe_runtime::heap::alloc(desc));
        }
        if let Some(index) = crate::keys::non_negative_integer_index_key(key) {
            if index < values.len() && !is_array_hole(&o, index as i32) {
                let mut desc = Object::new();
                desc.properties.reserve(4);
                desc.properties
                    .insert("value".into(), values[index].clone());
                desc.properties.insert("writable".into(), Value::Bool(true));
                desc.properties
                    .insert("enumerable".into(), Value::Bool(true));
                desc.properties
                    .insert("configurable".into(), Value::Bool(true));
                return Value::Object(vybe_runtime::heap::alloc(desc));
            }
        }
    }
    // Discriminating data vs accessor in the accessor convention:
    //   - `__get_<key>` present ⇒ ACCESSOR (real getters always install
    //     it, even when a plain placeholder entry coexists).
    //   - plain entry without `__get_` ⇒ DATA; a noop `__set_<key>`
    //     alongside it is the freeze/defineProperty non-writable guard
    //     (writable=false), NOT an accessor.
    //   - `__set_` only, no plain entry ⇒ setter-only accessor.
    crate::keys::with_getter_property_key(key, |getter_key| {
        crate::keys::with_setter_property_key(key, |setter_key| {
            if !o.properties.contains_key(getter_key) {
                if let Some(v) = o.properties.get(key) {
                    let mut desc = Object::new();
                    desc.properties.reserve(4);
                    desc.properties.insert("value".into(), v.clone());
                    let function_metadata = matches!(o.kind, ObjectKind::Function(_))
                        && (key == "name" || key == "length");
                    desc.properties.insert(
                        "writable".into(),
                        Value::Bool(!function_metadata && !o.properties.contains_key(setter_key)),
                    );
                    desc.properties
                        .insert("enumerable".into(), Value::Bool(!is_nonenum(&o, key)));
                    desc.properties
                        .insert("configurable".into(), Value::Bool(!is_nonconfig(&o, key)));
                    return Value::Object(vybe_runtime::heap::alloc(desc));
                }
            }
            if o.properties.contains_key(getter_key) || o.properties.contains_key(setter_key) {
                let mut desc = Object::new();
                desc.properties.reserve(4);
                desc.properties.insert(
                    "get".into(),
                    o.properties
                        .get(getter_key)
                        .cloned()
                        .unwrap_or(Value::Undefined),
                );
                desc.properties.insert(
                    "set".into(),
                    o.properties
                        .get(setter_key)
                        .cloned()
                        .unwrap_or(Value::Undefined),
                );
                desc.properties
                    .insert("enumerable".into(), Value::Bool(!is_nonenum(&o, key)));
                desc.properties
                    .insert("configurable".into(), Value::Bool(!is_nonconfig(&o, key)));
                return Value::Object(vybe_runtime::heap::alloc(desc));
            }
            Value::Undefined
        })
    })
}

fn string_length_descriptor(value: &Value, key: &str) -> Value {
    if key != "length" {
        return Value::Undefined;
    }
    let Value::String(text) = value else {
        return Value::Undefined;
    };
    let mut desc = Object::new();
    desc.properties.reserve(4);
    desc.properties.insert(
        "value".into(),
        Value::I32(if text.is_ascii() {
            text.len() as i32
        } else {
            text.chars().count() as i32
        }),
    );
    desc.properties
        .insert("writable".into(), Value::Bool(false));
    desc.properties
        .insert("enumerable".into(), Value::Bool(false));
    desc.properties
        .insert("configurable".into(), Value::Bool(false));
    Value::Object(vybe_runtime::heap::alloc(desc))
}

pub fn descriptor_own_keys(o: &Object) -> Vec<String> {
    if matches!(&o.kind, ObjectKind::Ordinary) {
        if let Some(keys) = plain_own_string_keys(o) {
            return keys;
        }
        if !o.properties.contains_key("__sym_keys")
            && !o
                .properties
                .keys()
                .any(|key| key.starts_with("__get_") || key.starts_with("__set_"))
        {
            return ordered_own_string_keys(o);
        }
    }
    let base_capacity = match &o.kind {
        ObjectKind::Array(values) => values.len().saturating_add(1),
        ObjectKind::TypedArray(ta) => crate::typedarray::ta_live_length(ta),
        _ => 0,
    };
    let mut keys = Vec::with_capacity(base_capacity.saturating_add(o.properties.len()));
    match &o.kind {
        ObjectKind::Array(values) => {
            let holes = array_hole_set(o);
            for index in 0..values.len() {
                if !cached_array_hole_contains(&holes, index) {
                    keys.push(crate::keys::with_index_key(index, |key| key.to_owned()));
                }
            }
            keys.push("length".to_string());
        }
        ObjectKind::TypedArray(ta) => {
            for index in 0..crate::typedarray::ta_live_length(ta) {
                keys.push(crate::keys::with_index_key(index, |key| key.to_owned()));
            }
        }
        _ => {}
    }
    keys.extend(ordered_own_string_keys(o));
    for key in o.properties.keys() {
        if let Some(name) = key.strip_prefix("__get_") {
            keys.push(name.to_string());
        } else if let Some(name) = key.strip_prefix("__set_") {
            keys.push(name.to_string());
        }
    }
    if let Some(Value::Object(sym_arr)) = o.properties.get("__sym_keys") {
        let syms = sym_arr.lock().unwrap();
        if let ObjectKind::Array(ref elems) = syms.kind {
            for key in elems {
                keys.push(key_string(key));
            }
        }
    }
    let mut seen = std::collections::HashSet::new();
    keys.into_iter()
        .filter(|key| !key.starts_with("__") && seen.insert(key.clone()))
        .collect()
}

fn is_noop_setter_value(value: &Value) -> bool {
    let Value::Object(obj) = value else {
        return false;
    };
    let noop_idx = NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
    noop_idx > 0
        && matches!(obj.lock().unwrap().kind, ObjectKind::HostFunction(idx) if idx == noop_idx)
}

pub fn is_data_property_writable(o: &Object, key: &str) -> bool {
    if o.properties.contains_key(FROZEN_MARK) {
        return false;
    }
    if crate::keys::with_getter_property_key(key, |getter_key| {
        o.properties.contains_key(getter_key)
    }) {
        return false;
    }
    match &o.kind {
        ObjectKind::Array(values) => {
            if key == "length" {
                return !o.properties.contains_key("__array_length_readonly")
                    && !o.properties.contains_key(FROZEN_MARK);
            }
            if let Some(index) = array_index_key(key) {
                return (index as usize) < values.len() && !o.properties.contains_key(FROZEN_MARK);
            }
        }
        ObjectKind::TypedArray(ta) => {
            if let Some(index) = array_index_key(key) {
                return (index as usize) < crate::typedarray::ta_live_length(ta);
            }
        }
        _ => {}
    }
    crate::keys::with_setter_property_key(
        key,
        |setter_key| !matches!(o.properties.get(setter_key), Some(setter) if is_noop_setter_value(setter)),
    )
}

fn is_accessor_backing_slot_write(o: &Object, key: &str) -> bool {
    key.starts_with('_') && o.properties.keys().any(|name| name.starts_with("__set_"))
}

fn is_effectively_sealed(o: &Object) -> bool {
    if !is_not_extensible(o) {
        return false;
    }
    descriptor_own_keys(o)
        .into_iter()
        .all(|key| is_nonconfig(o, &key))
}

fn is_effectively_frozen(o: &Object) -> bool {
    if !is_effectively_sealed(o) {
        return false;
    }
    descriptor_own_keys(o)
        .into_iter()
        .all(|key| !is_data_property_writable(o, &key))
}

// ── Prototype ─────────────────────────────────────────────────────────
//
// `js_prototype_of` above answers the ORDINARY internal method (§10.1.1) and
// is deliberately left untouched — it sits on the hot path of every
// inherited-property read, `instanceof` and `isPrototypeOf`. Proxy dispatch
// and the §10.5.1/§10.5.2 invariants layer ABOVE it, in the two functions
// below, so `Object.getPrototypeOf`, `Reflect.getPrototypeOf`, the
// `__proto__` accessor and the `ecma:proxy` entry points share ONE
// implementation instead of four partial ones that disagree.

/// `SameValue` restricted to what a `[[Prototype]]` can hold: an object
/// (compared by identity) or null.
fn same_prototype(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
        (Value::Null, Value::Null) => true,
        _ => false,
    }
}

/// `IsExtensible(O)` — §7.2.5. Seal and freeze both imply
/// `[[PreventExtensions]]` (§7.3.15), so they answer false too.
pub fn value_is_extensible(value: &Value) -> bool {
    match value {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            let sealed = o.properties.contains_key("__vybe_sealed")
                || o.properties.contains_key("__vybe_frozen");
            !sealed && !is_not_extensible(&o)
        }
        _ => false,
    }
}

/// §10.1.1 / §10.5.1 `[[GetPrototypeOf]]` — proxy-aware and invariant
/// enforcing. `None` means a TypeError has been thrown and the caller must
/// return immediately.
pub fn get_prototype_of(ctx: &mut HostContext, value: &Value) -> Option<Value> {
    let Some(proxy) = crate::proxy::is_proxy(value) else {
        return Some(fast_prototype_of(value));
    };
    if crate::proxy::proxy_is_revoked(&proxy) {
        throw_type_error(
            ctx,
            "Cannot perform 'getPrototypeOf' on a proxy that has been revoked",
        );
        return None;
    }
    let Some((target, handler)) = proxy_target_and_handler(&proxy) else {
        return Some(fast_prototype_of(value));
    };
    // No trap: forward to the target — which may itself be a proxy.
    let Some(trap) = proxy_trap(&handler, "getPrototypeOf") else {
        return get_prototype_of(ctx, &target);
    };
    let result = invoke_with_explicit_this(ctx, &trap, handler, &[target.clone()]);
    if !matches!(result, Value::Object(_) | Value::Null) {
        throw_type_error(
            ctx,
            "'getPrototypeOf' on proxy: trap returned neither object nor null",
        );
        return None;
    }
    // §10.5.1 step 7: an EXTENSIBLE target pins nothing — any object-or-null
    // the trap returns is legal. Only a non-extensible target forces the
    // trap to agree with the target's real prototype.
    if value_is_extensible(&target) {
        return Some(result);
    }
    let actual = get_prototype_of(ctx, &target)?;
    if !same_prototype(&result, &actual) {
        throw_type_error(
            ctx,
            "'getPrototypeOf' on proxy: proxy target is non-extensible but the trap did not return its actual prototype",
        );
        return None;
    }
    Some(result)
}

/// §10.1.2 / §10.5.2 `[[SetPrototypeOf]]` — answers the SUCCESS FLAG.
/// `Object.setPrototypeOf` turns a `false` into a TypeError; `Reflect
/// .setPrototypeOf` hands it back as-is. `None` means a TypeError has
/// already been thrown.
pub fn set_prototype_of(ctx: &mut HostContext, value: &Value, proto: &Value) -> Option<bool> {
    let Value::Object(obj) = value else {
        return Some(false);
    };
    if let Some(proxy) = crate::proxy::is_proxy(value) {
        if crate::proxy::proxy_is_revoked(&proxy) {
            throw_type_error(
                ctx,
                "Cannot perform 'setPrototypeOf' on a proxy that has been revoked",
            );
            return None;
        }
        if let Some((target, handler)) = proxy_target_and_handler(&proxy) {
            let Some(trap) = proxy_trap(&handler, "setPrototypeOf") else {
                // No trap: the TARGET's prototype changes, never the shell's.
                return set_prototype_of(ctx, &target, proto);
            };
            let result =
                invoke_with_explicit_this(ctx, &trap, handler, &[target.clone(), proto.clone()]);
            if !crate::boolean::to_boolean(&result) {
                return Some(false);
            }
            // §10.5.2 step 11: a trap may claim success on a non-extensible
            // target only if it actually left the prototype where the caller
            // asked for it.
            if value_is_extensible(&target) {
                return Some(true);
            }
            let actual = get_prototype_of(ctx, &target)?;
            if !same_prototype(proto, &actual) {
                throw_type_error(
                    ctx,
                    "'setPrototypeOf' on proxy: trap returned truish for a non-extensible target",
                );
                return None;
            }
            return Some(true);
        }
    }

    // §10.1.2 OrdinarySetPrototypeOf.
    let current = fast_prototype_of(value);
    if same_prototype(&current, proto) {
        return Some(true);
    }
    if !value_is_extensible(value) {
        return Some(false);
    }
    // Step 8: walk the new prototype's chain looking for `value` itself.
    // The walk STOPS at a proxy — the spec deliberately refuses to run user
    // trap code during the cycle check, so a proxy link ends the search
    // rather than being followed through.
    let mut p = proto.clone();
    for _ in 0..1000 {
        let Value::Object(p_obj) = &p else {
            break;
        };
        if Arc::ptr_eq(p_obj, obj) {
            return Some(false);
        }
        if let Some((Value::Object(proxy_target), _)) = proxy_target_and_handler(p_obj) {
            if Arc::ptr_eq(&proxy_target, obj) {
                return Some(false);
            }
        }
        let next = if crate::proxy::is_proxy(&p).is_some() {
            get_prototype_of(ctx, &p)?
        } else if let Value::Object(p_obj) = &p {
            object_prototype_of_locked(p_obj)
        } else {
            fast_prototype_of(&p)
        };
        // A root prototype can resolve to itself; that is the end of the
        // chain, not a cycle in `value`.
        if same_prototype(&next, &p) {
            break;
        }
        p = next;
    }

    let mut o = obj.lock().unwrap();
    o.properties.insert(PROTO_KEY.into(), proto.clone());
    if matches!(proto, Value::Null) {
        o.properties
            .insert(NULL_PROTO_MARK.into(), Value::Bool(true));
    } else {
        o.properties.shift_remove(NULL_PROTO_MARK);
    }
    Some(true)
}

fn register_prototype(vm: &mut VM) {
    object_unary(
        vm,
        "getPrototypeOf",
        vec![obj_t()],
        Box::new(|ctx, args| {
            // Primitives are coerced (§20.1.2.12 ToObject), not rejected —
            // only `Reflect.getPrototypeOf` throws on a non-object.
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            get_prototype_of(ctx, &value).unwrap_or(Value::Undefined)
        }),
    );

    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "setPrototypeOf",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let value = Value::Object(obj);
                let proto = args.get(1).cloned().unwrap_or(Value::Null);
                if !matches!(proto, Value::Object(_) | Value::Null) {
                    throw_type_error(ctx, "Object prototype may only be an Object or null");
                    return Value::Undefined;
                }
                // §20.1.2.22 step 4: `Object.setPrototypeOf` is the caller
                // that turns a false success flag into a TypeError.
                // `Reflect.setPrototypeOf` returns it instead.
                match set_prototype_of(ctx, &value, &proto) {
                    None => Value::Undefined,
                    Some(true) => value,
                    Some(false) => {
                        throw_type_error(ctx, "Cannot set prototype of this object");
                        Value::Undefined
                    }
                }
            } else {
                Value::Bool(false)
            }
        }),
    );
}

// ── Locking (freeze / seal / preventExtensions) ───────────────────────

/// Process-global host fn idx for the no-op setter installed by
/// `Object.freeze`. Captured during `register_locking` so freeze can
/// build bound `__set_<key>` accessors without re-looking-up the host
/// registry on every call.
static NOOP_SETTER_IDX: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

#[inline]
fn noop_setter_host_key() -> &'static (String, String) {
    static KEY: std::sync::OnceLock<(String, String)> = std::sync::OnceLock::new();
    KEY.get_or_init(|| ("ecma:object".to_string(), "__noop_setter".to_string()))
}

fn register_locking(vm: &mut VM) {
    // Silent-write setter installed by `freeze` for every existing key.
    // Returns the value untouched so `obj.x = 99` evaluates to 99 but
    // the underlying store is unchanged.
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "__noop_setter",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    let idx = vm
        .host_registry
        .get(noop_setter_host_key())
        .copied()
        .expect("__noop_setter just registered");
    NOOP_SETTER_IDX.store(idx, std::sync::atomic::Ordering::Relaxed);

    object_unary(
        vm,
        "freeze",
        vec![obj_t()],
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                if let Some((target, handler)) = proxy_target_and_handler(&obj) {
                    if crate::proxy::proxy_is_revoked(&obj) {
                        throw_type_error(
                            ctx,
                            "Cannot perform 'preventExtensions' on a proxy that has been revoked",
                        );
                        return Value::Undefined;
                    }
                    if let Some(trap) = proxy_trap(&handler, "preventExtensions") {
                        let success = crate::boolean::to_boolean(&invoke_with_explicit_this(
                            ctx,
                            &trap,
                            handler,
                            &[target.clone()],
                        ));
                        if !success || value_is_extensible(&target) {
                            throw_type_error(
                                ctx,
                                "Proxy preventExtensions trap did not make target non-extensible",
                            );
                            return Value::Undefined;
                        }
                    } else if let Value::Object(target_obj) = &target {
                        mark_not_extensible(&mut target_obj.lock().unwrap());
                    }
                }
                // §10.4.5 integer-indexed exotic: a typed array WITH
                // elements cannot be frozen — its indices can never be
                // made non-writable — so freeze throws TypeError.
                {
                    let o = obj.lock().unwrap();
                    if let ObjectKind::TypedArray(ta) = &o.kind {
                        if crate::typedarray::ta_live_length(ta) > 0 {
                            drop(o);
                            ctx.throw_value(crate::error::new_error(
                                ctx,
                                "TypeError",
                                "Cannot freeze array buffer views with elements",
                            ));
                            return Value::Undefined;
                        }
                    }
                }
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
                    descriptor_own_keys(&o)
                };
                let mut o = obj.lock().unwrap();
                o.properties.insert(FROZEN_MARK.into(), Value::I32(1));
                o.properties.insert(SEALED_MARK.into(), Value::I32(1));
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                let noop_idx = NOOP_SETTER_IDX.load(std::sync::atomic::Ordering::Relaxed);
                if noop_idx > 0 {
                    let mut noop_obj = Object::new();
                    noop_obj.kind = ObjectKind::HostFunction(noop_idx);
                    let noop_val = Value::Object(vybe_runtime::heap::alloc(noop_obj));
                    for k in &keys {
                        if is_accessor_backing_slot_write(&o, k) {
                            continue;
                        }
                        let setter_key = setter_property_key(k);
                        if !o.properties.contains_key(&setter_key) {
                            o.properties.insert(setter_key, noop_val.clone());
                        }
                    }
                }
                drop(o);
                for k in keys {
                    track_nonconfig(&obj, &k);
                }
                return Value::Object(obj);
            }
            // §20.1.2.7 step 1 (ES2015+): non-object → return it unchanged.
            args.first().cloned().unwrap_or(Value::Undefined)
        }),
    );

    object_unary(
        vm,
        "isFrozen",
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(is_effectively_frozen(&o));
            }
            Value::Bool(true)
        }),
    );

    object_unary(
        vm,
        "seal",
        vec![obj_t()],
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                if let Some((target, handler)) = proxy_target_and_handler(&obj) {
                    if crate::proxy::proxy_is_revoked(&obj) {
                        throw_type_error(
                            ctx,
                            "Cannot perform 'preventExtensions' on a proxy that has been revoked",
                        );
                        return Value::Undefined;
                    }
                    if let Some(trap) = proxy_trap(&handler, "preventExtensions") {
                        let success = crate::boolean::to_boolean(&invoke_with_explicit_this(
                            ctx,
                            &trap,
                            handler,
                            &[target.clone()],
                        ));
                        if !success || value_is_extensible(&target) {
                            throw_type_error(
                                ctx,
                                "Proxy preventExtensions trap did not make target non-extensible",
                            );
                            return Value::Undefined;
                        }
                    } else if let Value::Object(target_obj) = &target {
                        mark_not_extensible(&mut target_obj.lock().unwrap());
                    }
                }
                let keys = {
                    let o = obj.lock().unwrap();
                    descriptor_own_keys(&o)
                };
                {
                    let mut o = obj.lock().unwrap();
                    o.properties.insert(SEALED_MARK.into(), Value::I32(1));
                    o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                }
                for k in keys {
                    track_nonconfig(&obj, &k);
                }
                return Value::Object(obj);
            }
            // §20.1.2.20 step 1 (ES2015+): non-object → return it unchanged.
            args.first().cloned().unwrap_or(Value::Undefined)
        }),
    );

    object_unary(
        vm,
        "isSealed",
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(is_effectively_sealed(&o));
            }
            Value::Bool(true)
        }),
    );

    object_unary(
        vm,
        "preventExtensions",
        vec![obj_t()],
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                // §10.5.4: trap when present, otherwise the TARGET
                // becomes non-extensible — never the proxy shell.
                if let Some((target, handler)) = proxy_target_and_handler(&obj) {
                    if crate::proxy::proxy_is_revoked(&obj) {
                        throw_type_error(
                            ctx,
                            "Cannot perform 'preventExtensions' on a proxy that has been revoked",
                        );
                        return Value::Undefined;
                    }
                    if let Some(trap) = proxy_trap(&handler, "preventExtensions") {
                        let success = crate::boolean::to_boolean(&invoke_with_explicit_this(
                            ctx,
                            &trap,
                            handler,
                            &[target.clone()],
                        ));
                        if !success {
                            throw_type_error(ctx, "Proxy preventExtensions trap returned false");
                            return Value::Undefined;
                        }
                        if value_is_extensible(&target) {
                            throw_type_error(
                                ctx,
                                "Proxy preventExtensions trap returned true but target is still extensible",
                            );
                            return Value::Undefined;
                        }
                        return Value::Object(obj);
                    }
                    if let Value::Object(t) = &target {
                        mark_not_extensible(&mut t.lock().unwrap());
                    }
                    return Value::Object(obj);
                }
                let mut o = obj.lock().unwrap();
                mark_not_extensible(&mut o);
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }),
    );

    object_unary(
        vm,
        "isExtensible",
        vec![ValType::Bool],
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                // §10.5.3: trap when present, otherwise the TARGET's
                // extensibility answers.
                if let Some((target, handler)) = proxy_target_and_handler(&obj) {
                    if crate::proxy::proxy_is_revoked(&obj) {
                        throw_type_error(
                            ctx,
                            "Cannot perform 'isExtensible' on a proxy that has been revoked",
                        );
                        return Value::Undefined;
                    }
                    let target_extensible = value_is_extensible(&target);
                    if let Some(trap) = proxy_trap(&handler, "isExtensible") {
                        let result = invoke_with_explicit_this(ctx, &trap, handler, &[target]);
                        let reported = crate::boolean::to_boolean(&result);
                        if reported != target_extensible {
                            throw_type_error(
                                ctx,
                                "Proxy isExtensible trap result does not match target",
                            );
                            return Value::Undefined;
                        }
                        return Value::Bool(reported);
                    }
                    return Value::Bool(target_extensible);
                }
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
    crate::perf::register_host_fn(
        vm,
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

/// §20.1.3: a user-defined method on the receiver (or its prototype
/// chain) SHADOWS the Object.prototype intrinsic. Compile-time routed
/// intrinsics call this first so overrides win.
fn user_method_override(obj: &Arc<Mutex<Object>>, name: &str) -> Option<Value> {
    let mut current = Some(obj.clone());
    let mut guard = 0;
    while let Some(cur) = current {
        guard += 1;
        if guard > 10_000 {
            break;
        }
        let (prop, proto) = {
            let o = cur.lock().unwrap();
            (
                o.properties.get(name).cloned(),
                o.properties.get(PROTO_KEY).cloned(),
            )
        };
        // Only USER-compiled functions shadow the intrinsic — a
        // HostFunction found on the chain IS the intrinsic (re-invoking
        // it would recurse forever).
        if let Some(Value::Object(f)) = &prop {
            if matches!(f.lock().unwrap().kind, ObjectKind::Function(_)) {
                return prop;
            }
        }
        current = match proto {
            Some(Value::Object(p)) => Some(p),
            _ => None,
        };
    }
    None
}

/// §20.1.3.2 raw intrinsic (no override dispatch) — the value installed
/// on %Object.prototype% for borrowed-call forms.
fn has_own_property_intrinsic(args: &[Value]) -> Value {
    let undefined = Value::Undefined;
    let target = args.first().unwrap_or(&undefined);
    let key = args.get(1).unwrap_or(&undefined);
    Value::Bool(has_own_property_key(target, key).unwrap_or(false))
}

fn register_prototype_methods(vm: &mut VM) {
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "hasOwnPropertyIntrinsic",
        Box::new(|_ctx, args| has_own_property_intrinsic(args)),
    );
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "hasOwnProperty",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                if obj.lock().unwrap().properties.contains_key(NULL_PROTO_MARK) {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "hasOwnProperty is not a function",
                    ));
                    return Value::Undefined;
                }
                if let Some(f) = user_method_override(&obj, "hasOwnProperty") {
                    return invoke_with_explicit_this(
                        ctx,
                        &f,
                        Value::Object(obj),
                        args.get(1..).unwrap_or(&[]),
                    );
                }
                let target = Value::Object(obj);
                let undefined = Value::Undefined;
                let key = args.get(1).unwrap_or(&undefined);
                return Value::Bool(has_own_property_key(&target, key).unwrap_or(false));
            }
            Value::Bool(false)
        }),
    );

    // Receiver first (`__vybe_method_receiver`), then the object being tested.
    object_fn(
        vm,
        "isPrototypeOf",
        vec![obj_t(), obj_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            // isPrototypeOf: is `self` in `other`'s prototype chain?
            // Resolve each link via `js_prototype_of` (the same resolver
            // `getPrototypeOf` uses) rather than reading `__proto__` directly:
            // VM-created plain objects/arrays carry no explicit `__proto__`,
            // their `[[Prototype]]` is resolved by kind to the shared
            // prototype singleton.
            if let (Some(self_obj), Some(other)) = (obj_of(args, 0), obj_of(args, 1)) {
                let mut current = Value::Object(other);
                loop {
                    let next = match &current {
                        Value::Object(current_obj) => object_prototype_of_locked(current_obj),
                        _ => fast_prototype_of(&current),
                    };
                    match next {
                        Value::Object(p) => {
                            if Arc::ptr_eq(&p, &self_obj) {
                                return Value::Bool(true);
                            }
                            // Fixed point (e.g. a root prototype whose
                            // kind-resolved proto is itself) — not found.
                            if let Value::Object(c) = &current {
                                if Arc::ptr_eq(&p, c) {
                                    return Value::Bool(false);
                                }
                            }
                            current = Value::Object(p);
                        }
                        _ => return Value::Bool(false),
                    }
                }
            }
            Value::Bool(false)
        }),
    );

    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "propertyIsEnumerable",
        Box::new(|ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                if let Some(f) = user_method_override(&obj, "propertyIsEnumerable") {
                    return invoke_with_explicit_this(
                        ctx,
                        &f,
                        Value::Object(obj),
                        args.get(1..).unwrap_or(&[]),
                    );
                }
                let undefined = Value::Undefined;
                let key_value = args.get(1).unwrap_or(&undefined);
                let o = obj.lock().unwrap();
                // §20.1.3.4: array index elements are own enumerable
                // properties (they live in ObjectKind::Array, not the
                // property map); `length` is own but non-enumerable
                // (§10.4.2).
                return with_property_key(key_value, |key| {
                    if let ObjectKind::Array(ref elems) = o.kind {
                        if key == "length" {
                            return Value::Bool(false);
                        }
                        if let Some(idx) = crate::keys::non_negative_integer_index_key(key) {
                            if idx < elems.len() {
                                return Value::Bool(true);
                            }
                        }
                    }
                    Value::Bool(
                        o.properties.contains_key(key)
                            && !key.starts_with("__")
                            && !is_nonenum(&o, key),
                    )
                });
            }
            // §10.4.3 string exotics: char indices are own enumerable.
            if let (Some(Value::String(s)), Some(key)) = (args.first(), args.get(1)) {
                return with_property_key(key, |key| {
                    if let Some(idx) = crate::keys::non_negative_integer_index_key(key) {
                        let len = if s.is_ascii() {
                            s.len()
                        } else {
                            s.chars().count()
                        };
                        return Value::Bool(idx < len);
                    }
                    Value::Bool(false)
                });
            }
            Value::Bool(false)
        }),
    );

    // toString(): ECMA-262 §20.1.3.6 — returns "[object <Tag>]" for any value
    // (including primitives). Called via Object.prototype.toString.call(value).
    object_unary(
        vm,
        "toString",
        vec![ValType::String],
        Box::new(|ctx, args| match args.first() {
            None | Some(Value::Undefined) => crate::keys::object_tag_value("Undefined"),
            Some(Value::Null) => crate::keys::object_tag_value("Null"),
            Some(Value::Bool(_)) => crate::keys::object_tag_value("Boolean"),
            Some(Value::I32(_)) | Some(Value::I64(_)) | Some(Value::F64(_)) => {
                crate::keys::object_tag_value("Number")
            }
            Some(Value::BigInt(_)) => crate::keys::object_tag_value("BigInt"),
            Some(Value::String(_)) => crate::keys::object_tag_value("String"),
            Some(Value::Symbol(_)) => crate::keys::object_tag_value("Symbol"),
            Some(Value::Object(obj)) => object_to_string_tag_value(ctx, obj),
            _ => crate::keys::object_tag_value("Object"),
        }),
    );

    object_unary(
        vm,
        "toLocaleString",
        vec![ValType::String],
        Box::new(|_ctx, args| {
            if is_object(args.first().unwrap_or(&Value::Null)) {
                return crate::keys::string_value("[object Object]");
            }
            crate::keys::string_value("")
        }),
    );

    // valueOf: spec default returns the object itself
    object_unary(
        vm,
        "valueOf",
        vec![obj_t()],
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
    object_unary(
        vm,
        "toStringTag",
        vec![ValType::String],
        Box::new(|ctx, args| match args.first() {
            None | Some(Value::Undefined) => crate::keys::object_tag_value("Undefined"),
            Some(Value::Null) => crate::keys::object_tag_value("Null"),
            Some(Value::Object(obj)) => object_to_string_tag_value(ctx, obj),
            _ => crate::keys::object_tag_value("Object"),
        }),
    );

    // Object.groupBy(items, keyFn) — ES2024 §20.1.2.x.
    // Groups iterable items into a plain object keyed by keyFn(item, index).
    crate::perf::register_host_fn(
        vm,
        "ecma:object",
        "groupBy",
        Box::new(|ctx, args| {
            let items = args.first().cloned().unwrap_or(Value::Undefined);
            let key_fn = args.get(1).cloned().unwrap_or(Value::Undefined);
            if !is_callable_value(&key_fn) {
                return throw_type_error(ctx, "Object.groupBy callback is not callable");
            }
            let prepared_key_fn = crate::function::prepare_bound_callback(&key_fn);
            let Some(arr_items) =
                collect_groupby_items(ctx, &items, "Object.groupBy argument is not iterable")
            else {
                return Value::Undefined;
            };
            let mut groups: IndexMap<String, Vec<Value>> = IndexMap::with_capacity(arr_items.len());
            for (i, item) in arr_items.into_iter().enumerate() {
                let key = if let Some(k) = groupby_magic_key(&key_fn, &item) {
                    k
                } else {
                    let invoke_args = [item.clone(), Value::I32(i as i32)];
                    let key_value = if let Some(prepared) = &prepared_key_fn {
                        crate::function::invoke_prepared_bound_callback(ctx, prepared, &invoke_args)
                    } else {
                        ctx.invoke(&key_fn, &invoke_args)
                    };
                    match key_value {
                        Value::Symbol(_) => {
                            return throw_type_error(
                                ctx,
                                "Cannot convert a Symbol value to a property key",
                            );
                        }
                        other => crate::keys::with_property_key(&other, |key| key.to_owned()),
                    }
                };
                groups
                    .entry(key)
                    .or_insert_with(|| Vec::with_capacity(4))
                    .push(item);
            }
            let mut result_obj = Object::new();
            result_obj.properties.insert(PROTO_KEY.into(), Value::Null);
            let mut order = Vec::with_capacity(groups.len());
            for (key, values) in groups {
                order.push(crate::keys::string_value(&key));
                result_obj.properties.insert(
                    key,
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values))),
                );
            }
            if !order.is_empty() {
                result_obj.properties.insert(
                    "__keys".to_string(),
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(order))),
                );
            }
            Value::Object(vybe_runtime::heap::alloc(result_obj))
        }),
    );
}

// ── PHP extensions ────────────────────────────────────────────────────

fn register_php_extensions(vm: &mut VM) {
    crate::perf::register_host_fn(
        vm,
        "php:reflect",
        "newInstanceArgs",
        Box::new(|ctx, args| {
            let Some(Value::String(class_name)) = args.first() else {
                return Value::Undefined;
            };
            let global_name = class_name.replace('\\', ".");
            let mut constructor = ctx.get_global(&global_name);
            if matches!(constructor, Value::Undefined | Value::Null) {
                if let Value::Object(stack) = ctx.get_global("__php_autoload_stack") {
                    let callbacks = {
                        let guard = stack.lock().unwrap();
                        match &guard.kind {
                            ObjectKind::Array(items) => items.clone(),
                            _ => Vec::new(),
                        }
                    };
                    for callback in callbacks {
                        ctx.invoke(&callback, &[Value::String(class_name.clone())]);
                        constructor = ctx.get_global(&global_name);
                        if !matches!(constructor, Value::Undefined | Value::Null) {
                            break;
                        }
                    }
                }
            }
            let positional = match args.get(1) {
                Some(Value::Object(array)) => {
                    let guard = array.lock().unwrap();
                    match &guard.kind {
                        ObjectKind::Array(items) => items.clone(),
                        ObjectKind::Map(items) => {
                            let mut values = Vec::with_capacity(items.len());
                            values.extend(items.values().cloned());
                            values
                        }
                        _ => {
                            let mut indexed = Vec::with_capacity(guard.properties.len());
                            for (key, value) in &guard.properties {
                                if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                                    indexed.push((i, value.clone()));
                                }
                            }
                            indexed.sort_by_key(|(i, _)| *i);
                            let mut values = Vec::with_capacity(indexed.len());
                            for (_, value) in indexed {
                                values.push(value);
                            }
                            values
                        }
                    }
                }
                _ => Vec::new(),
            };
            if matches!(constructor, Value::Undefined | Value::Null) {
                return Value::Undefined;
            }
            ctx.invoke(&constructor, &positional)
        }),
    );

    // PHP runs __destruct when the last program reference disappears. The
    // compiler calls this after rebinding a source variable; compiler spill
    // locals are excluded by HostContext's root walk.
    crate::perf::register_host_fn(
        vm,
        "php:object",
        "finalizeUnreachable",
        Box::new(|ctx, args| {
            let Some(old) = args.first() else {
                return Value::Null;
            };
            let Some(Value::Object(dtor)) = args.get(1) else {
                return Value::Null;
            };
            if !matches!(
                dtor.lock().unwrap().kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            ) {
                return Value::Null;
            }
            if !ctx.is_reachable_from_program(old) {
                ctx.invoke_with_receiver(&args[1], old.clone(), &[]);
            }
            Value::Null
        }),
    );

    // appendAutoKey(obj, value) -> i32 key
    // Implements PHP's `$a[] = x` — finds the next int key (max of
    // existing int keys + 1, or 0 if none) and sets it.
    crate::perf::register_host_fn(
        vm,
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
                            if let Some(n) = crate::keys::non_negative_i32_key(k) {
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
