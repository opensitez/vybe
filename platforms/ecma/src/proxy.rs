use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

const PROXY_TAG: &str = "__vybe_js_proxy";
const PROXY_TARGET: &str = "__vybe_proxy_target";
const PROXY_HANDLER: &str = "__vybe_proxy_handler";
const PROXY_REVOKED: &str = "__vybe_proxy_revoked";
const PROXY_SET_TRAP_DEPTH: &str = "__vybe_proxy_set_trap_depth";
const PROXY_APPLY_INLINE_ARG_LIMIT: usize = 8;

#[inline]
fn proxy_apply_host_key() -> &'static (String, String) {
    static KEY: std::sync::OnceLock<(String, String)> = std::sync::OnceLock::new();
    KEY.get_or_init(|| ("ecma:proxy".to_string(), "apply".to_string()))
}

#[inline]
fn proxy_call_host_key() -> &'static (String, String) {
    static KEY: std::sync::OnceLock<(String, String)> = std::sync::OnceLock::new();
    KEY.get_or_init(|| ("ecma:proxy".to_string(), "call".to_string()))
}

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

#[inline]
fn owned_string_value(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

fn new_proxy(target: Value, handler: Value, call_idx: Option<usize>) -> Value {
    let callable = call_idx.filter(|_| is_callable(&target));
    let mut obj = Object::new();
    obj.properties
        .reserve(if callable.is_some() { 11 } else { 4 });
    obj.properties.insert(PROXY_TAG.into(), Value::I32(1));
    obj.properties.insert(PROXY_TARGET.into(), target.clone());
    obj.properties.insert(PROXY_HANDLER.into(), handler);
    obj.properties
        .insert(PROXY_REVOKED.into(), Value::Bool(false));
    if let Some(idx) = callable {
        obj.kind = ObjectKind::HostFunction(idx);
        obj.properties.insert(
            "__host_module".into(),
            crate::keys::string_value("ecma:proxy"),
        );
        obj.properties
            .insert("__host_name".into(), crate::keys::string_value("call"));
        obj.properties
            .insert("__host_idx".into(), Value::F64(idx as f64));
        obj.properties
            .insert("__vybe_proxy_callable".into(), Value::Bool(true));
        obj.properties.insert(
            "__proto__".into(),
            crate::function::shared_function_prototype(),
        );
        if let Value::Object(target_obj) = &target {
            let target_locked = target_obj.lock().unwrap();
            if let Some(name) = target_locked.properties.get("name").cloned() {
                obj.properties.insert("name".into(), name);
            }
            if let Some(length) = target_locked.properties.get("length").cloned() {
                obj.properties.insert("length".into(), length);
            }
            if let Some(prototype) = target_locked.properties.get("prototype").cloned() {
                obj.properties.insert("prototype".into(), prototype);
            }
        }
    }
    let proxy = vybe_runtime::heap::alloc(obj);
    if callable.is_some() {
        proxy.lock().unwrap().properties.insert(
            "__bound_args".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                Value::Object(proxy.clone()),
                Value::Undefined,
            ]))),
        );
    }
    Value::Object(proxy)
}

// Per-operation snapshot, never a cache: handler changes and revocation
// are observed afresh before each operation, with no lock held across traps.
enum ProxyDispatchState {
    Ordinary,
    Revoked,
    Active { target: Value, handler: Value },
}

fn proxy_dispatch_state(value: &Value) -> ProxyDispatchState {
    let Value::Object(object) = value else {
        return ProxyDispatchState::Ordinary;
    };
    let object = object.lock().unwrap();
    if !object.properties.contains_key(PROXY_TAG)
        && !(object.properties.contains_key(PROXY_TARGET)
            && object.properties.contains_key(PROXY_HANDLER))
    {
        return ProxyDispatchState::Ordinary;
    }
    if matches!(object.properties.get(PROXY_REVOKED), Some(Value::Bool(true))) {
        return ProxyDispatchState::Revoked;
    }
    ProxyDispatchState::Active {
        target: object.properties.get(PROXY_TARGET).cloned().unwrap_or(Value::Undefined),
        handler: object.properties.get(PROXY_HANDLER).cloned().unwrap_or(Value::Undefined),
    }
}

pub fn is_proxy(v: &Value) -> Option<Arc<Mutex<Object>>> {
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        if o.properties.contains_key(PROXY_TAG)
            || (o.properties.contains_key(PROXY_TARGET) && o.properties.contains_key(PROXY_HANDLER))
        {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

pub fn proxy_is_revoked(proxy: &Arc<Mutex<Object>>) -> bool {
    matches!(
        proxy.lock().unwrap().properties.get(PROXY_REVOKED),
        Some(Value::Bool(true))
    )
}

fn get_trap(handler: &Value, name: &str) -> Option<Value> {
    if let Value::Object(h) = handler {
        let h = h.lock().unwrap();
        h.properties.get(name).cloned()
    } else {
        None
    }
}

fn get_target_and_handler(proxy: &Arc<Mutex<Object>>) -> (Value, Value) {
    let object = proxy.lock().unwrap();
    (
        object
            .properties
            .get(PROXY_TARGET)
            .cloned()
            .unwrap_or(Value::Undefined),
        object
            .properties
            .get(PROXY_HANDLER)
            .cloned()
            .unwrap_or(Value::Undefined),
    )
}

fn target_get(target: &Value, key: &str) -> Value {
    match target {
        Value::Object(obj) => {
            // Arrays keep elements in ObjectKind::Array, not properties —
            // mirror ARRAY_GET semantics for the no-trap/non-proxy path.
            {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(v) = &o.kind {
                    if key == "length" {
                        return Value::F64(v.len() as f64);
                    }
                    if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                        return v.get(i).cloned().unwrap_or(Value::Undefined);
                    }
                }
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    if key == "length" {
                        return Value::F64(crate::typedarray::ta_live_length(ta) as f64);
                    }
                    if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                        return crate::typedarray::read_element(ta, i);
                    }
                }
            }
            crate::object::proto_walk_get(obj, key).unwrap_or(Value::Undefined)
        }
        Value::String(s) => {
            if key == "length" {
                return Value::F64(s.chars().count() as f64);
            }
            if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                return s.chars().nth(i).map(char_value).unwrap_or(Value::Undefined);
            }
            Value::Undefined
        }
        _ => Value::Undefined,
    }
}

fn target_set(ctx: &mut HostContext, target: &Value, key_value: Value, key: &str, val: Value) {
    target_set_with_receiver(ctx, target, key_value, key, val, target.clone());
}

fn target_set_with_receiver(
    ctx: &mut HostContext,
    target: &Value,
    key_value: Value,
    key: &str,
    val: Value,
    receiver: Value,
) {
    if let Value::Object(obj) = target {
        let forwards_to_proto = {
            let o = obj.lock().unwrap();
            let ordinary = !matches!(o.kind, ObjectKind::Array(_) | ObjectKind::TypedArray(_));
            drop(o);
            ordinary
                && !in_proxy_set_trap(ctx)
                && !target_own_property_exists(target, key)
                && !matches!(crate::object::js_prototype_of(target), Value::Null)
        };
        if forwards_to_proto {
            let proto = crate::object::js_prototype_of(target);
            if let Some(proto_proxy) = is_proxy(&proto) {
                if proxy_is_revoked(&proto_proxy) {
                    throw_revoked(ctx);
                    return;
                }
                let (proto_target, handler) = get_target_and_handler(&proto_proxy);
                if let Some(trap) = get_trap(&handler, "set") {
                    if is_callable(&trap) {
                        let _ = call_set_trap(
                            ctx,
                            &handler,
                            &trap,
                            &[proto_target, key_value, val, receiver],
                        );
                        return;
                    }
                }
            }
        }
        {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(v) = &mut o.kind {
                if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                    if i >= v.len() {
                        v.resize(i + 1, Value::Undefined);
                    }
                    v[i] = val;
                    return;
                }
                if key == "length" {
                    let new_len = val.as_f64() as usize;
                    v.resize(new_len, Value::Undefined);
                    return;
                }
            }
        }
        obj.lock().unwrap().properties.insert(key.to_string(), val);
        if !key.starts_with("__") {
            crate::object::track_key(obj, key);
        }
    }
}

fn key_string(value: &Value) -> String {
    crate::keys::property_key_string(value)
}

fn with_key_string<R>(value: &Value, f: impl FnOnce(&str) -> R) -> R {
    crate::keys::with_property_key(value, f)
}

fn getter_key(key: &str) -> String {
    crate::keys::getter_property_key(key)
}

fn setter_key(key: &str) -> String {
    crate::keys::setter_property_key(key)
}

fn make_type_error(ctx: &HostContext, message: &str) -> Value {
    crate::error::new_error(ctx, "TypeError", message)
}

fn throw_revoked(ctx: &mut HostContext) -> Value {
    ctx.throw_value(make_type_error(
        ctx,
        "Cannot perform operation on a revoked proxy",
    ));
    Value::Undefined
}

fn apply_dispatch(ctx: &mut HostContext, args: &[Value]) -> Value {
    apply_dispatch_inner(ctx, args, None)
}

fn apply_dispatch_inner(ctx: &mut HostContext, args: &[Value], raw_args: Option<&[Value]>) -> Value {
    let proxy_obj = match args.first().and_then(is_proxy) {
        Some(p) => p,
        None => {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            // Proxy.revocable's revoke function (§28.2.2.1) is modelled as an
            // object carrying __revoke_target.
            if let Value::Object(obj) = &target {
                let revoke_target = obj
                    .lock()
                    .unwrap()
                    .properties
                    .get("__revoke_target")
                    .cloned();
                if let Some(Value::Object(proxy_obj)) = revoke_target {
                    proxy_obj
                        .lock()
                        .unwrap()
                        .properties
                        .insert(PROXY_REVOKED.into(), Value::Bool(true));
                    return Value::Undefined;
                }
            }
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(invoke_args) = raw_args {
                return crate::function::invoke_with_explicit_this(ctx, &target, this_arg, invoke_args);
            }
            let mut inline: [Value; PROXY_APPLY_INLINE_ARG_LIMIT] =
                std::array::from_fn(|_| Value::Undefined);
            if let Some(inline_len) = args
                .get(2)
                .and_then(|value| array_values_inline(value, &mut inline))
            {
                return crate::function::invoke_with_explicit_this(
                    ctx,
                    &target,
                    this_arg,
                    &inline[..inline_len],
                );
            }
            let invoke_args = args.get(2).map(array_values).unwrap_or_default();
            return crate::function::invoke_with_explicit_this(
                ctx,
                &target,
                this_arg,
                &invoke_args,
            );
        }
    };
    if proxy_is_revoked(&proxy_obj) {
        return throw_revoked(ctx);
    }
    let (target, handler) = get_target_and_handler(&proxy_obj);
    if !is_callable(&target) {
        ctx.throw_value(make_type_error(ctx, "Proxy target is not callable"));
        return Value::Undefined;
    }
    let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
    if let Some(trap) = get_trap(&handler, "apply") {
        if is_callable(&trap) {
            let args_list = if let Some(invoke_args) = raw_args {
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(invoke_args.to_vec())))
            } else {
                args.get(2).cloned().unwrap_or_else(|| {
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
                })
            };
            return call_trap(ctx, &handler, &trap, &[target, this_arg, args_list]);
        }
    }
    if let Some(invoke_args) = raw_args {
        return crate::function::invoke_with_explicit_this(ctx, &target, this_arg, invoke_args);
    }
    let mut inline: [Value; PROXY_APPLY_INLINE_ARG_LIMIT] =
        std::array::from_fn(|_| Value::Undefined);
    if let Some(inline_len) = args
        .get(2)
        .and_then(|value| array_values_inline(value, &mut inline))
    {
        return crate::function::invoke_with_explicit_this(
            ctx,
            &target,
            this_arg,
            &inline[..inline_len],
        );
    }
    let invoke_args = args.get(2).map(array_values).unwrap_or_default();
    crate::function::invoke_with_explicit_this(ctx, &target, this_arg, &invoke_args)
}

pub fn get_dispatch(ctx: &mut HostContext, value: &Value, key_value: &Value) -> Value {
    let (target, handler) = match proxy_dispatch_state(value) {
        ProxyDispatchState::Ordinary => {
            return with_key_string(key_value, |key| target_get(value, key));
        }
        ProxyDispatchState::Revoked => return throw_revoked(ctx),
        ProxyDispatchState::Active { target, handler } => (target, handler),
    };
    with_key_string(key_value, |key| {
        if let Some(trap) = get_trap(&handler, "get") {
            if is_callable(&trap) {
                let receiver = value.clone();
                let result = call_trap(
                    ctx,
                    &handler,
                    &trap,
                    &[target.clone(), key_value.clone(), receiver],
                );
                if let Some((actual, writable)) = target_nonconfig_data_value(&target, key) {
                    if !writable && result != actual {
                        ctx.throw_value(make_type_error(
                            ctx,
                            "Proxy get trap violated non-configurable property invariant",
                        ));
                        return Value::Undefined;
                    }
                }
                return result;
            }
        }
        if key == "__proto__" {
            return crate::object::get_prototype_of(ctx, value).unwrap_or(Value::Undefined);
        }
        target_get(&target, key)
    })
}

fn proxy_call_parts(args: &[Value]) -> (Value, Value, &[Value]) {
    let proxy = args.first().cloned().unwrap_or(Value::Undefined);
    let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
    (proxy, this_arg, args.get(2..).unwrap_or_default())
}

fn callable_proxy_dispatch(ctx: &mut HostContext, args: &[Value]) -> Value {
    let (proxy, this_arg, invoke_args) = proxy_call_parts(args);
    apply_dispatch_inner(ctx, &[proxy, this_arg], Some(invoke_args))
}

fn is_callable(value: &Value) -> bool {
    matches!(value, Value::Object(obj)
        if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
}

#[cfg(test)]
mod call_arguments_speedup_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    #[test]
    fn raw_call_parts_preserve_receiver_order_and_argument_identity() {
        for count in [0, 1, 8, 9, 32] {
            let obj = vybe_runtime::heap::alloc(Object::new());
            let mut args = vec![Value::I32(11), Value::I32(22)];
            args.extend((0..count).map(Value::I32));
            args.push(Value::Object(obj.clone()));
            let (proxy, receiver, raw) = proxy_call_parts(&args);
            assert_eq!(proxy.as_i32(), 11);
            assert_eq!(receiver.as_i32(), 22);
            assert_eq!(raw.len(), count as usize + 1);
            assert!(std::ptr::eq(raw.as_ptr(), args[2..].as_ptr()));
            for (i, value) in raw[..count as usize].iter().enumerate() {
                assert_eq!(value.as_i32(), i as i32);
            }
            assert!(matches!(raw.last(), Some(Value::Object(actual)) if Arc::ptr_eq(actual, &obj)));
        }
        let (proxy, receiver, raw) = proxy_call_parts(&[]);
        assert!(matches!(proxy, Value::Undefined));
        assert!(matches!(receiver, Value::Undefined));
        assert!(raw.is_empty());
        let (_, receiver, raw) = proxy_call_parts(&[Value::I32(11)]);
        assert!(matches!(receiver, Value::Undefined));
        assert!(raw.is_empty());
        let (_, _, raw) = proxy_call_parts(&[Value::I32(11), Value::I32(22)]);
        assert!(raw.is_empty());
    }

    #[test]
    #[ignore = "native call preparation microbenchmark, not full Proxy dispatch timing"]
    fn native_raw_call_preparation_microbenchmark() {
        for count in [0, 4, 8, 9, 32] {
            let mut args = vec![Value::I32(11), Value::I32(22)];
            args.extend((0..count).map(Value::I32));
            let mut old = Duration::ZERO;
            let mut new = Duration::ZERO;
            for batch in 0..10 {
                for old_path in if batch % 2 == 0 { [true, false] } else { [false, true] } {
                    let start = Instant::now();
                    for _ in 0..1000 {
                        let args = black_box(&args);
                        if old_path {
                            let proxy = args.first().cloned().unwrap_or(Value::Undefined);
                            let receiver = args.get(1).cloned().unwrap_or(Value::Undefined);
                            let list = Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                                args.iter().skip(2).cloned().collect(),
                            )));
                            black_box((proxy, receiver, list));
                        } else {
                            black_box(proxy_call_parts(args));
                        }
                    }
                    if old_path { old += start.elapsed(); } else { new += start.elapsed(); }
                }
            }
            eprintln!("native callable Proxy argument preparation, args={count}, calls=10000: array-object={old:?}, raw-slice={new:?}");
        }
    }
}

fn is_constructor(value: &Value) -> bool {
    let Value::Object(obj) = value else {
        return false;
    };
    let o = obj.lock().unwrap();
    if !matches!(
        o.kind,
        ObjectKind::Function(_) | ObjectKind::HostFunction(_)
    ) {
        return false;
    }
    !matches!(o.properties.get("__vybe_non_ctor"), Some(Value::Bool(true)))
}

fn call_trap(ctx: &mut HostContext, handler: &Value, trap: &Value, args: &[Value]) -> Value {
    crate::function::invoke_with_explicit_this(ctx, trap, handler.clone(), args)
}

fn try_call_trap(
    ctx: &mut HostContext,
    handler: &Value,
    trap: &Value,
    args: &[Value],
) -> Result<Value, Value> {
    crate::function::try_invoke_with_explicit_this(ctx, trap, handler.clone(), args)
}

fn call_set_trap(ctx: &mut HostContext, handler: &Value, trap: &Value, args: &[Value]) -> Value {
    match try_call_set_trap(ctx, handler, trap, args) {
        Ok(value) => value,
        Err(thrown) => { ctx.throw_value(thrown); Value::Undefined }
    }
}

fn try_call_set_trap(ctx: &mut HostContext, handler: &Value, trap: &Value, args: &[Value]) -> Result<Value, Value> {
    let previous = ctx.get_global(PROXY_SET_TRAP_DEPTH);
    let depth = previous.as_f64() as i32;
    ctx.set_global(PROXY_SET_TRAP_DEPTH, Value::I32(depth + 1));
    let result = try_call_trap(ctx, handler, trap, args);
    ctx.set_global(PROXY_SET_TRAP_DEPTH, previous);
    result
}

// Shared by the registered setter and inherited proxy [[Set]]. The fourth
// argument retains the original receiver when the proxy is a prototype.
pub(crate) fn set_dispatch(ctx: &mut HostContext, args: &[Value]) -> Value {
    let value = args.first().cloned().unwrap_or(Value::Undefined);
    let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
    let val = args.get(2).cloned().unwrap_or(Value::Undefined);
    let receiver = args.get(3).cloned().unwrap_or_else(|| value.clone());
    let Some(proxy_obj) = is_proxy(&value) else {
        return with_key_string(&key_value, |key| {
            crate::reflect::reflect_set(ctx, &value, key, val, receiver)
        });
    };
    if proxy_is_revoked(&proxy_obj) {
        return throw_revoked(ctx);
    }
    let (target, handler) = get_target_and_handler(&proxy_obj);
    let key = key_string(&key_value);
    if let Some(trap) = get_trap(&handler, "set") {
        if !matches!(trap, Value::Undefined | Value::Null) {
            if !is_callable(&trap) {
                ctx.throw_value(make_type_error(ctx, "Proxy set trap is not callable"));
                return Value::Undefined;
            }
            let result = match try_call_set_trap(ctx, &handler, &trap, &[
                target.clone(), key_value, val.clone(), receiver,
            ]) {
                Ok(result) => result,
                Err(thrown) => { ctx.throw_value(thrown); return Value::Undefined; }
            };
            let success = crate::boolean::to_boolean(&result);
            if success {
                if let Some((actual, writable)) = target_nonconfig_data_value(&target, &key) {
                    if !writable && val != actual {
                        ctx.throw_value(make_type_error(ctx,
                            "Proxy set trap violated non-configurable property invariant"));
                        return Value::Undefined;
                    }
                }
            }
            return Value::Bool(success);
        }
    }
    let receiver_is_proxy = matches!(&receiver, Value::Object(object) if Arc::ptr_eq(object, &proxy_obj));
    if key == "__proto__" {
        return match crate::object::set_prototype_of(ctx, &receiver, &val) {
            None => Value::Undefined,
            Some(success) => Value::Bool(success),
        };
    }
    if receiver_is_proxy && !crate::object::value_is_extensible(&target) && !target_has(&target, &key) {
        return Value::Bool(false);
    }
    if receiver_is_proxy && get_trap(&handler, "defineProperty").is_some() {
        let mut desc = Object::new();
        desc.properties.insert("value".into(), val.clone());
        desc.properties.insert("writable".into(), Value::Bool(true));
        desc.properties.insert("enumerable".into(), Value::Bool(true));
        desc.properties.insert("configurable".into(), Value::Bool(true));
        if let Some(result) = define_property_dispatch(ctx, &proxy_obj, key_value.clone(),
            Value::Object(vybe_runtime::heap::alloc(desc))) {
            return Value::Bool(crate::boolean::to_boolean(&result));
        }
    }
    let follow_receiver = if receiver_is_proxy { target.clone() } else { receiver };
    crate::reflect::reflect_set(ctx, &target, &key, val, follow_receiver)
}

fn in_proxy_set_trap(ctx: &HostContext) -> bool {
    ctx.get_global(PROXY_SET_TRAP_DEPTH).as_f64() > 0.0
}

fn target_has(target: &Value, key: &str) -> bool {
    match target {
        Value::Object(obj) => {
            {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(v) = &o.kind {
                    if key == "length" {
                        return true;
                    }
                    if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                        return i < v.len() && !crate::array::is_array_hole(&o, i);
                    }
                }
            }
            crate::object::proto_walk_has(obj, key)
        }
        _ => false,
    }
}

fn target_own_property_exists(target: &Value, key: &str) -> bool {
    let Value::Object(obj) = target else {
        return false;
    };
    let o = obj.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(values) => {
            if key == "length" {
                return true;
            }
            if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                return i < values.len() && !crate::array::is_array_hole(&o, i);
            }
        }
        ObjectKind::TypedArray(ta) => {
            if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                return i < crate::typedarray::ta_live_length(ta);
            }
        }
        _ => {}
    }
    o.properties.contains_key(key)
        || crate::keys::with_getter_property_key(key, |getter| o.properties.contains_key(getter))
        || crate::keys::with_setter_property_key(key, |setter| o.properties.contains_key(setter))
}

fn target_nonconfig_data_value(target: &Value, key: &str) -> Option<(Value, bool)> {
    let Value::Object(obj) = target else {
        return None;
    };
    let o = obj.lock().unwrap();
    if !crate::object::is_nonconfig(&o, key) {
        return None;
    }
    if let ObjectKind::Array(values) = &o.kind {
        if key == "length" {
            return Some((Value::I32(values.len() as i32), true));
        }
        if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
            if i < values.len() && !crate::array::is_array_hole(&o, i) {
                return Some((values[i].clone(), true));
            }
        }
    }
    if crate::keys::with_getter_property_key(key, |getter| o.properties.contains_key(getter)) {
        return None;
    }
    let value = o.properties.get(key).cloned()?;
    let writable = crate::object::is_data_property_writable(&o, key);
    Some((value, writable))
}

fn target_delete(target: &Value, key: &str) -> bool {
    if let Value::Object(obj) = target {
        let mut o = obj.lock().unwrap();
        if crate::object::is_nonconfig(&o, key) {
            return false;
        }
        if let ObjectKind::Array(values) = &mut o.kind {
            if let Some(i) = crate::keys::non_negative_integer_index_key(key) {
                if i < values.len() {
                    values[i] = Value::Undefined;
                    crate::array::mark_array_hole(&mut o, i);
                    return true;
                }
                return true;
            }
        }
        o.properties.shift_remove(key);
    }
    true
}

fn target_own_keys(target: &Value) -> Value {
    if let Value::Object(obj) = target {
        let o = obj.lock().unwrap();
        let own_keys = crate::object::ordered_own_string_keys(&o);
        let mut keys = Vec::with_capacity(own_keys.len());
        for key in own_keys {
            keys.push(crate::keys::string_value(key.as_str()));
        }
        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
    }
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
}

fn array_values(value: &Value) -> Vec<Value> {
    if let Value::Object(obj) = value {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(values) = &o.kind {
            return values.clone();
        }
        if let Some(length) = o
            .properties
            .get("length")
            .map(|v| v.as_i32().max(0) as usize)
        {
            let mut values = Vec::with_capacity(length);
            for i in 0..length {
                values.push(crate::keys::with_index_key(i, |key| {
                    o.properties.get(key).cloned().unwrap_or(Value::Undefined)
                }));
            }
            return values;
        }
    }
    Vec::new()
}

fn array_values_inline<const N: usize>(value: &Value, inline: &mut [Value; N]) -> Option<usize> {
    let Value::Object(obj) = value else {
        return Some(0);
    };
    let o = obj.lock().unwrap();
    if let ObjectKind::Array(values) = &o.kind {
        if values.len() > inline.len() {
            return None;
        }
        for (index, value) in values.iter().enumerate() {
            inline[index] = value.clone();
        }
        return Some(values.len());
    }
    let length = o
        .properties
        .get("length")
        .map(|v| v.as_i32().max(0) as usize)
        .unwrap_or(0);
    if length > inline.len() {
        return None;
    }
    for (index, slot) in inline.iter_mut().enumerate().take(length) {
        *slot = crate::keys::with_index_key(index, |key| {
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        });
    }
    Some(length)
}

fn desc_bool(desc: &Value, key: &str, default: bool) -> bool {
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

fn desc_value(desc: &Value) -> Option<Value> {
    if let Value::Object(obj) = desc {
        return obj.lock().unwrap().properties.get("value").cloned();
    }
    None
}

fn target_is_not_extensible(target: &Value) -> bool {
    !crate::object::value_is_extensible(target)
}

fn target_descriptor(target: &Value, key: &str) -> Value {
    match target {
        Value::Object(obj) => crate::object::own_property_descriptor(obj, key),
        _ => Value::Undefined,
    }
}

fn validate_proxy_descriptor_result(
    ctx: &mut HostContext,
    target: &Value,
    key: &str,
    target_desc: &Value,
    result: &Value,
) -> bool {
    if matches!(result, Value::Undefined) {
        if !matches!(target_desc, Value::Undefined)
            && (!desc_bool(target_desc, "configurable", true) || target_is_not_extensible(target))
        {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy getOwnPropertyDescriptor trap violated target invariant",
            ));
            return false;
        }
        return true;
    }
    if !matches!(result, Value::Object(_)) {
        ctx.throw_value(make_type_error(
            ctx,
            "Proxy getOwnPropertyDescriptor trap must return an object or undefined",
        ));
        return false;
    }
    if matches!(target_desc, Value::Undefined) {
        if target_is_not_extensible(target) || !desc_bool(result, "configurable", false) {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy getOwnPropertyDescriptor trap reported incompatible descriptor",
            ));
            return false;
        }
        return true;
    }
    if !desc_bool(target_desc, "configurable", true) {
        if desc_bool(result, "configurable", false) {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy getOwnPropertyDescriptor trap reported non-configurable target as configurable",
            ));
            return false;
        }
        if let (Some(actual), Some(reported)) = (desc_value(target_desc), desc_value(result)) {
            if actual != reported {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy getOwnPropertyDescriptor trap changed a non-configurable value",
                ));
                return false;
            }
        }
    }
    let _ = key;
    true
}

pub fn get_own_property_descriptor_dispatch(
    ctx: &mut HostContext,
    value: &Value,
    key_value: &Value,
) -> Option<Value> {
    let proxy_obj = is_proxy(value)?;
    if proxy_is_revoked(&proxy_obj) {
        return Some(throw_revoked(ctx));
    }
    let (target, handler) = get_target_and_handler(&proxy_obj);
    let key = key_string(key_value);
    let target_desc = target_descriptor(&target, &key);
    if let Some(trap) = get_trap(&handler, "getOwnPropertyDescriptor") {
        if is_callable(&trap) {
            let result = call_trap(ctx, &handler, &trap, &[target.clone(), key_value.clone()]);
            if !validate_proxy_descriptor_result(ctx, &target, &key, &target_desc, &result) {
                return Some(Value::Undefined);
            }
            return Some(result);
        }
    }
    Some(target_desc)
}

pub fn define_property_dispatch(
    ctx: &mut HostContext,
    proxy_obj: &Arc<Mutex<Object>>,
    key_value: Value,
    descriptor: Value,
) -> Option<Value> {
    if proxy_is_revoked(proxy_obj) {
        return Some(throw_revoked(ctx));
    }
    let (target, handler) = get_target_and_handler(proxy_obj);
    let key = key_string(&key_value);
    let target_desc = target_descriptor(&target, &key);
    if let Some(trap) = get_trap(&handler, "defineProperty") {
        if is_callable(&trap) {
            let result = match try_call_trap(
                ctx,
                &handler,
                &trap,
                &[target.clone(), key_value, descriptor.clone()],
            ) {
                Ok(result) => result,
                Err(thrown) => {
                    ctx.throw_value(thrown);
                    return Some(Value::Undefined);
                }
            };
            if !crate::boolean::to_boolean(&result) {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy defineProperty trap returned false",
                ));
                return Some(Value::Undefined);
            }
            let requested_nonconfig = !desc_bool(&descriptor, "configurable", false);
            if matches!(target_desc, Value::Undefined) {
                let next_target_desc = target_descriptor(&target, &key);
                if target_is_not_extensible(&target)
                    || (requested_nonconfig
                        && (matches!(next_target_desc, Value::Undefined)
                            || desc_bool(&next_target_desc, "configurable", true)))
                {
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Proxy defineProperty trap violated target invariant",
                    ));
                    return Some(Value::Undefined);
                }
            } else if !desc_bool(&target_desc, "configurable", true) {
                if desc_bool(&descriptor, "configurable", false) {
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Proxy defineProperty trap cannot make property configurable",
                    ));
                    return Some(Value::Undefined);
                }
                if let (Some(actual), Some(next)) =
                    (desc_value(&target_desc), desc_value(&descriptor))
                {
                    if actual != next && !desc_bool(&target_desc, "writable", true) {
                        ctx.throw_value(make_type_error(
                            ctx,
                            "Proxy defineProperty trap cannot change read-only property",
                        ));
                        return Some(Value::Undefined);
                    }
                }
            }
            return Some(Value::Bool(true));
        }
    }
    None
}

/// [[OwnPropertyKeys]] for proxy exotic objects — ECMA-262 §10.5.11.
/// Returns None when `value` is not a proxy so callers (Object.keys etc.)
/// keep their ordinary-object path.
fn own_keys_key_text(value: &Value) -> Option<std::borrow::Cow<'_, str>> {
    match value {
        Value::String(text) => Some(std::borrow::Cow::Borrowed(text.as_ref())),
        Value::Symbol(_) => Some(std::borrow::Cow::Owned(key_string(value))),
        _ => None,
    }
}

pub fn own_keys_dispatch(ctx: &mut HostContext, value: &Value) -> Option<Value> {
    let (target, handler) = match proxy_dispatch_state(value) {
        ProxyDispatchState::Ordinary => return None,
        ProxyDispatchState::Revoked => return Some(throw_revoked(ctx)),
        ProxyDispatchState::Active { target, handler } => (target, handler),
    };
    let result = if let Some(trap) = get_trap(&handler, "ownKeys") {
        if is_callable(&trap) {
            call_trap(ctx, &handler, &trap, std::slice::from_ref(&target))
        } else {
            target_own_keys(&target)
        }
    } else {
        target_own_keys(&target)
    };
    let keys = array_values(&result);
    if !matches!(result, Value::Object(_)) {
        ctx.throw_value(make_type_error(
            ctx,
            "Proxy ownKeys trap must return an object",
        ));
        return Some(Value::Undefined);
    }
    let mut seen = std::collections::HashSet::with_capacity(keys.len());
    for key in &keys {
        let Some(text) = own_keys_key_text(key) else {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy ownKeys trap keys must be strings or symbols",
            ));
            return Some(Value::Undefined);
        };
        if !seen.insert(text) {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy ownKeys trap returned duplicate keys",
            ));
            return Some(Value::Undefined);
        }
    }
    if let Value::Object(target_obj) = &target {
        let (target_keys, ordinary_missing_nonconfig) = {
            let o = target_obj.lock().unwrap();
            let target_keys = crate::object::descriptor_own_keys(&o);
            // Only native ordinary string-key shapes: exotic descriptors
            // and symbol storage retain the existing descriptor fallback.
            let ordinary = matches!(o.kind, ObjectKind::Ordinary)
                && !o.properties.contains_key(PROXY_TAG)
                && !o.properties.contains_key(PROXY_TARGET)
                && !o.properties.contains_key(PROXY_HANDLER)
                && !o.properties.contains_key("__sym_keys");
            let missing = ordinary.then(|| target_keys.iter().any(|key| {
                !seen.contains(key.as_str()) && crate::object::is_nonconfig(&o, key)
            }));
            (target_keys, missing)
        };
        let omitted_nonconfig = match ordinary_missing_nonconfig {
            Some(missing) => missing,
            None => target_keys.iter().any(|key| {
                let desc = crate::object::own_property_descriptor(target_obj, key);
                !desc_bool(&desc, "configurable", true) && !seen.contains(key.as_str())
            }),
        };
        if omitted_nonconfig {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy ownKeys trap omitted a non-configurable key",
            ));
            return Some(Value::Undefined);
        }
        if target_is_not_extensible(&target) {
            for key in &target_keys {
                if !seen.contains(key.as_str()) {
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Proxy ownKeys trap omitted a key from non-extensible target",
                    ));
                    return Some(Value::Undefined);
                }
            }
            if seen.len() != target_keys.len() {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy ownKeys trap added a key to a non-extensible target",
                ));
                return Some(Value::Undefined);
            }
        }
    }
    Some(result)
}

/// [[Construct]] dispatch — ECMA-262 §10.5.13 for proxy exotic objects,
/// ordinary construct for everything else. Shared by `ecma:proxy.construct`
/// and `ecma:reflect.construct` (§28.1.2 routes through [[Construct]]).
pub fn construct_dispatch(ctx: &mut HostContext, constructor: &Value, args_list: &Value) -> Value {
    construct_dispatch_with_new_target(ctx, constructor, args_list, None)
}

/// §28.1.2 step 2: newTarget defaults to the constructor itself; when
/// given, it is what `new.target` observes inside the ctor chain. The
/// binding rides the `__js_new_target` calling-convention global.
pub fn construct_dispatch_with_new_target(
    ctx: &mut HostContext,
    constructor: &Value,
    args_list: &Value,
    new_target: Option<Value>,
) -> Value {
    let effective_nt = new_target.unwrap_or_else(|| {
        let current = ctx.get_global("__js_new_target");
        if matches!(current, Value::Undefined | Value::Null) {
            constructor.clone()
        } else {
            current
        }
    });
    let previous_nt = ctx.get_global("__js_new_target");
    ctx.set_global("__js_new_target", effective_nt);
    let result = construct_dispatch_inner(ctx, constructor, args_list);
    ctx.set_global("__js_new_target", previous_nt);
    result
}

fn construct_dispatch_inner(
    ctx: &mut HostContext,
    constructor: &Value,
    args_list: &Value,
) -> Value {
    // ⛔ §13.3.5.1 EvaluateNew: *"If IsConstructor(constructor) is false, throw
    // a TypeError exception."* A PRIMITIVE is never a constructor, and without
    // this `new 5` silently produced an ordinary object — no throw at all,
    // where node reports `5 is not a constructor`.
    //
    // Deliberately narrower than `is_constructor`, which additionally requires
    // `ObjectKind::Function | HostFunction`: an object-shaped callee that this
    // predicate would reject already throws further down (`new {}` and
    // `new (()=>{})` both answer TypeError today), so widening the guard here
    // would only risk refusing a class whose runtime shape this predicate does
    // not anticipate. Primitives are the whole measured gap.
    if !matches!(constructor, Value::Object(_)) {
        let message = format!("{} is not a constructor", constructor);
        let err = make_type_error(ctx, &message);
        ctx.throw_value(err);
        return Value::Undefined;
    }
    if let Some(proxy_obj) = is_proxy(constructor) {
        if proxy_is_revoked(&proxy_obj) {
            return throw_revoked(ctx);
        }
        let (target, handler) = get_target_and_handler(&proxy_obj);
        if !is_constructor(&target) {
            ctx.throw_value(make_type_error(ctx, "Proxy target is not a constructor"));
            return Value::Undefined;
        }
        if let Some(trap) = get_trap(&handler, "construct") {
            if is_callable(&trap) {
                let new_target = ctx.get_global("__js_new_target");
                let new_target = if matches!(new_target, Value::Undefined | Value::Null) {
                    constructor.clone()
                } else {
                    new_target
                };
                let result = match try_call_trap(
                    ctx,
                    &handler,
                    &trap,
                    &[target.clone(), args_list.clone(), new_target],
                ) {
                    Ok(result) => result,
                    Err(thrown) => {
                        ctx.throw_value(thrown);
                        return Value::Undefined;
                    }
                };
                if matches!(result, Value::Object(_)) {
                    return result;
                }
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy construct trap must return an object",
                ));
                return Value::Undefined;
            }
        }
        // No trap: construct the target (recurses for proxy-of-proxy).
        return construct_dispatch(ctx, &target, args_list);
    }

    // Built-in Error constructor used as a value (`const T = TypeError;
    // new T(msg, { cause })`). The canonical `__ctor_<Name>` anchor is an inert
    // object (not callable), so construct the Error here — same shape as the
    // compiler's `emit_exception_new_finalize` (ECMA-262 §20.5.1 / §20.5.8.1).
    if let Value::Object(target_obj) = constructor {
        let err_name = {
            let locked = target_obj.lock().unwrap();
            match locked.properties.get("__error_ctor_name") {
                Some(Value::String(n)) => Some(n.clone()),
                _ => None,
            }
        };
        if let Some(err_name) = err_name {
            let args = array_values(args_list);
            let mut err = Object::new();
            let name_str = err_name.to_string();
            err.properties
                .insert("name".into(), Value::String(err_name.clone()));
            err.properties
                .insert("__type".into(), Value::String(err_name.clone()));
            err.properties
                .insert("__exception_type".into(), Value::String(err_name));
            err.properties.insert(
                "message".into(),
                match args.first() {
                    Some(Value::String(s)) => Value::String(s.clone()),
                    Some(Value::Undefined) | None => crate::keys::string_value(""),
                    Some(other) => owned_string_value(crate::keys::value_display_string(other)),
                },
            );
            // AggregateError takes the iterable of errors as the first arg and
            // the message as the second; everyone else is (message, options).
            let options_index = if name_str == "AggregateError" { 2 } else { 1 };
            if let Some(Value::Object(opts)) = args.get(options_index) {
                if let Some(cause) = opts.lock().unwrap().properties.get("cause").cloned() {
                    err.properties.insert("cause".into(), cause);
                }
            }
            return Value::Object(vybe_runtime::heap::alloc(err));
        }
    }

    // §7.2.4 IsConstructor: arrows / shorthand methods / generator
    // expressions carry no [[Construct]] — the compiler marks them with
    // `__vybe_non_ctor`. `new` on them is a TypeError.
    if let Value::Object(target_obj) = constructor {
        let (non_ctor, fn_name) = {
            let locked = target_obj.lock().unwrap();
            (
                matches!(
                    locked.properties.get("__vybe_non_ctor"),
                    Some(Value::Bool(true))
                ),
                match locked.properties.get("name") {
                    Some(Value::String(n)) if !n.is_empty() => n.to_string(),
                    _ => "anonymous".to_string(),
                },
            )
        };
        if non_ctor {
            ctx.throw_value(make_type_error(
                ctx,
                &format!("{} is not a constructor", fn_name),
            ));
            return Value::Undefined;
        }
    }

    let mut this_value = Object::new();
    let new_target = ctx.get_global("__js_new_target");
    let prototype_source = if matches!(new_target, Value::Undefined | Value::Null) {
        constructor
    } else {
        &new_target
    };
    if let Value::Object(target_obj) = prototype_source {
        if let Some(proto) = target_obj
            .lock()
            .unwrap()
            .properties
            .get("prototype")
            .cloned()
        {
            this_value.properties.insert("__proto__".into(), proto);
        }
    }
    let this_obj = Value::Object(vybe_runtime::heap::alloc(this_value));
    let mut inline: [Value; 4] = std::array::from_fn(|_| Value::Undefined);
    let result = if let Some(inline_len) = array_values_inline(args_list, &mut inline) {
        crate::function::invoke_with_explicit_this(
            ctx,
            constructor,
            this_obj.clone(),
            &inline[..inline_len],
        )
    } else {
        let args = array_values(args_list);
        crate::function::invoke_with_explicit_this(ctx, constructor, this_obj.clone(), &args)
    };
    if matches!(result, Value::Object(_)) {
        result
    } else {
        this_obj
    }
}

#[cfg(test)]
mod get_dispatch_speedup_tests {
    use super::*;
    use std::time::{Duration, Instant};

    fn object(object: Object) -> Value {
        Value::Object(Arc::new(Mutex::new(object)))
    }

    fn function(vm: &VM, namespace: &str, name: &str) -> Value {
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(namespace, name).unwrap());
        object(function)
    }

    fn target(key: &str) -> Value {
        let mut target = Object::new();
        target.properties.insert(key.into(), Value::I32(7));
        target.properties.insert("__proto__".into(), Value::Null);
        object(target)
    }

    #[test]
    fn state_snapshot_preserves_identity_and_rechecks_revocation() {
        let target = target("field");
        let handler = object(Object::new());
        let proxy = new_proxy(target.clone(), handler.clone(), None);
        let ProxyDispatchState::Active { target: actual_target, handler: actual_handler } =
            proxy_dispatch_state(&proxy) else { panic!("active proxy"); };
        let (Value::Object(actual_target), Value::Object(target)) = (actual_target, target) else { unreachable!() };
        let (Value::Object(actual_handler), Value::Object(handler)) = (actual_handler, handler) else { unreachable!() };
        assert!(Arc::ptr_eq(&actual_target, &target));
        assert!(Arc::ptr_eq(&actual_handler, &handler));
        let Value::Object(proxy_object) = &proxy else { unreachable!() };
        proxy_object.lock().unwrap().properties.insert(PROXY_REVOKED.into(), Value::Bool(true));
        assert!(matches!(proxy_dispatch_state(&proxy), ProxyDispatchState::Revoked));
        assert!(matches!(proxy_dispatch_state(&Value::I32(1)), ProxyDispatchState::Ordinary));
        assert!(matches!(proxy_dispatch_state(&object(Object::new())), ProxyDispatchState::Ordinary));
    }

    #[test]
    fn registered_get_releases_locks_and_observes_trap_replacement() {
        let mut vm = VM::new();
        register(&mut vm);
        let target = target("field");
        let Value::Object(target_object) = &target else { unreachable!() };
        let target_object = target_object.clone();
        let handler = Arc::new(Mutex::new(Object::new()));
        let callback_handler = handler.clone();
        let callback_target = target_object.clone();
        vm.register_host_fn("ecma:test-proxy", "getTrap", Box::new(move |_ctx, args| {
            assert!(matches!(args.first(), Some(Value::Object(actual)) if Arc::ptr_eq(actual, &callback_target)));
            assert!(matches!(args.get(1), Some(Value::String(key)) if key.as_ref() == "field"));
            let Some(Value::Object(receiver)) = args.get(2) else { panic!("proxy receiver"); };
            assert!(receiver.try_lock().is_ok(), "proxy lock must not span trap");
            callback_target.try_lock().unwrap().properties.insert("field".into(), Value::I32(9));
            callback_handler.try_lock().unwrap().properties.shift_remove("get");
            Value::I32(77)
        }));
        handler.lock().unwrap().properties.insert("get".into(), function(&vm, "ecma:test-proxy", "getTrap"));
        let proxy = new_proxy(target, Value::Object(handler), None);
        let get = function(&vm, "ecma:proxy", "get");
        let args = [proxy, crate::keys::string_value("field")];
        assert!(matches!(vm.invoke(&get, &args).unwrap(), Value::I32(77)));
        assert!(matches!(vm.invoke(&get, &args).unwrap(), Value::I32(9)));
    }

    #[test]
    fn registered_get_allows_current_trap_to_revoke_and_rejects_next_read() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:test-proxy", "revokeDuringGet", Box::new(|_ctx, args| {
            let Some(Value::Object(receiver)) = args.get(2) else { panic!("proxy receiver"); };
            receiver.lock().unwrap().properties.insert(PROXY_REVOKED.into(), Value::Bool(true));
            Value::I32(55)
        }));
        let mut handler = Object::new();
        handler.properties.insert("get".into(), function(&vm, "ecma:test-proxy", "revokeDuringGet"));
        let proxy = new_proxy(target("field"), object(handler), None);
        let get = function(&vm, "ecma:proxy", "get");
        let args = [proxy, crate::keys::string_value("field")];
        assert!(matches!(vm.invoke(&get, &args).unwrap(), Value::I32(55)));
        assert!(vm.invoke(&get, &args).is_err());
    }

    #[test]
    fn registered_get_retains_nonconfigurable_readonly_invariant() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:test-proxy", "wrongGet", Box::new(|_ctx, _args| Value::I32(99)));
        let target = target("field");
        let Value::Object(target_object) = &target else { unreachable!() };
        crate::object::install_noop_setter(&mut target_object.lock().unwrap(), "field");
        crate::object::track_nonconfig(target_object, "field");
        let mut handler = Object::new();
        handler.properties.insert("get".into(), function(&vm, "ecma:test-proxy", "wrongGet"));
        let proxy = new_proxy(target, object(handler), None);
        let get = function(&vm, "ecma:proxy", "get");
        assert!(vm.invoke(&get, &[proxy, crate::keys::string_value("field")]).is_err());
    }

    #[test]
    #[ignore = "paired registered Proxy.get benchmark; not a conformance gate"]
    fn registered_proxy_get_benchmark() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:test-proxy", "legacyGet", Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let key = args.get(1).cloned().unwrap_or(Value::Undefined);
            legacy_get_dispatch(ctx, &value, &key)
        }));
        vm.register_host_fn("ecma:test-proxy", "constantTrap", Box::new(|_ctx, _args| Value::I32(7)));
        let mut trap_handler = Object::new();
        trap_handler.properties.insert("get".into(), function(&vm, "ecma:test-proxy", "constantTrap"));
        let trapped = new_proxy(target("field"), object(trap_handler), None);
        let long_key = "long_property_name_".repeat(8);
        let old = function(&vm, "ecma:test-proxy", "legacyGet");
        let new = function(&vm, "ecma:proxy", "get");
        for (label, value, key) in [
            ("ordinary", target("field"), "field"),
            ("proxy-short-key", new_proxy(target("field"), object(Object::new()), None), "field"),
            ("proxy-long-key", new_proxy(target(&long_key), object(Object::new()), None), long_key.as_str()),
            ("proxy-get-trap", trapped, "field"),
        ] {
            let args = [value, crate::keys::string_value(key)];
            for function in [&old, &new] {
                assert!(matches!(vm.invoke(function, &args).unwrap(), Value::I32(7)));
                for _ in 0..5_000 { std::hint::black_box(vm.invoke(function, &args).unwrap()); }
            }
            let mut old_time = Duration::ZERO;
            let mut new_time = Duration::ZERO;
            for sample in 0..6 {
                for old_path in [sample % 2 == 0, sample % 2 != 0] {
                    let function = if old_path { &old } else { &new };
                    let start = Instant::now();
                    for _ in 0..30_000 { std::hint::black_box(vm.invoke(function, &args).unwrap()); }
                    if old_path { old_time += start.elapsed(); } else { new_time += start.elapsed(); }
                }
            }
            eprintln!("{label}: old={old_time:?} new={new_time:?} ratio={:.3}x",
                old_time.as_secs_f64() / new_time.as_secs_f64());
        }
    }

fn legacy_get_dispatch(ctx: &mut HostContext, value: &Value, key_value: &Value) -> Value {
    let proxy_obj = match is_proxy(value) {
        Some(p) => p,
        None => {
            return with_key_string(key_value, |key| target_get(value, key));
        }
    };
    if proxy_is_revoked(&proxy_obj) {
        return throw_revoked(ctx);
    }
    let (target, handler) = get_target_and_handler(&proxy_obj);
    let key = key_string(key_value);
    if let Some(trap) = get_trap(&handler, "get") {
        if is_callable(&trap) {
            let receiver = value.clone();
            let result = call_trap(
                ctx,
                &handler,
                &trap,
                &[target.clone(), key_value.clone(), receiver],
            );
            if let Some((actual, writable)) = target_nonconfig_data_value(&target, &key) {
                if !writable && result != actual {
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Proxy get trap violated non-configurable property invariant",
                    ));
                    return Value::Undefined;
                }
            }
            return result;
        }
    }
    if key == "__proto__" {
        return crate::object::get_prototype_of(ctx, value).unwrap_or(Value::Undefined);
    }
    target_get(&target, &key)
}
}

#[cfg(test)]
mod own_keys_speedup_tests {
    use super::*;
    use std::time::{Duration, Instant};

    #[test]
    fn duplicate_tracking_borrows_strings_and_retains_symbol_conversion() {
        let text = Arc::<str>::from("immutable_long_key_".repeat(16));
        let key = Value::String(text.clone());
        let Some(std::borrow::Cow::Borrowed(borrowed)) = own_keys_key_text(&key) else {
            panic!("string keys should borrow their backing");
        };
        assert_eq!(borrowed.as_ptr(), text.as_ptr());
        assert_eq!(borrowed, text.as_ref());
        let symbol = Value::Symbol(Arc::from("symbol-key"));
        let Some(std::borrow::Cow::Owned(converted)) = own_keys_key_text(&symbol) else {
            panic!("symbol fallback should retain canonical conversion");
        };
        assert_eq!(converted, key_string(&symbol));
        for invalid in [Value::Undefined, Value::Null, Value::I32(1), Value::Bool(true)] {
            assert!(own_keys_key_text(&invalid).is_none());
        }
    }

    #[test]
    #[ignore = "paired registered borrowed-key comparison against descriptor-optimized ownKeys"]
    fn registered_borrowed_own_keys_benchmark() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:test-ownkeys", "preBorrowedOwnKeys", Box::new(|ctx, args| {
            let value = args.first().unwrap_or(&Value::Undefined);
            pre_borrowed_own_keys_dispatch(ctx, value).unwrap_or_else(|| target_own_keys(value))
        }));
        let old = function(&vm, "ecma:test-ownkeys", "preBorrowedOwnKeys");
        let new = function(&vm, "ecma:proxy", "ownKeys");
        for count in [16, 256, 1024, 4096] {
            let names: Vec<_> = (0..count).map(|index| format!("property_{index}")).collect();
            let mut target = Object::new();
            for name in &names { target.properties.insert(name.clone(), Value::I32(1)); }
            let keys = object(Object::new_array(names.iter().map(|name| crate::keys::string_value(name)).collect()));
            let (proxy, _) = setup(&mut vm, object(target), keys);
            let args = [proxy];
            for function in [&old, &new] {
                let Value::Object(result) = vm.invoke(function, &args).unwrap() else { panic!("key array"); };
                let result = result.lock().unwrap();
                let ObjectKind::Array(values) = &result.kind else { panic!("key array"); };
                assert_eq!(values.len(), count);
                for (value, name) in values.iter().zip(&names) {
                    assert!(matches!(value, Value::String(text) if text.as_ref() == name));
                }
            }
            let iterations = (16_384 / count).max(8);
            let mut old_time = Duration::ZERO;
            let mut new_time = Duration::ZERO;
            for sample in 0..8 {
                for old_path in [sample % 2 == 0, sample % 2 != 0] {
                    let function = if old_path { &old } else { &new };
                    let start = Instant::now();
                    for _ in 0..iterations { std::hint::black_box(vm.invoke(function, &args).unwrap()); }
                    if old_path { old_time += start.elapsed(); } else { new_time += start.elapsed(); }
                }
            }
            eprintln!("borrowed-ownKeys count={count}: owned={old_time:?} borrowed={new_time:?} ratio={:.3}x",
                old_time.as_secs_f64() / new_time.as_secs_f64());
        }
    }

fn pre_borrowed_own_keys_dispatch(ctx: &mut HostContext, value: &Value) -> Option<Value> {
    let (target, handler) = match proxy_dispatch_state(value) {
        ProxyDispatchState::Ordinary => return None,
        ProxyDispatchState::Revoked => return Some(throw_revoked(ctx)),
        ProxyDispatchState::Active { target, handler } => (target, handler),
    };
    let result = if let Some(trap) = get_trap(&handler, "ownKeys") {
        if is_callable(&trap) {
            call_trap(ctx, &handler, &trap, std::slice::from_ref(&target))
        } else {
            target_own_keys(&target)
        }
    } else {
        target_own_keys(&target)
    };
    let keys = array_values(&result);
    if !matches!(result, Value::Object(_)) {
        ctx.throw_value(make_type_error(
            ctx,
            "Proxy ownKeys trap must return an object",
        ));
        return Some(Value::Undefined);
    }
    let mut seen = std::collections::HashSet::with_capacity(keys.len());
    for key in &keys {
        if !matches!(key, Value::String(_) | Value::Symbol(_)) {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy ownKeys trap keys must be strings or symbols",
            ));
            return Some(Value::Undefined);
        }
        let text = key_string(key);
        if !seen.insert(text) {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy ownKeys trap returned duplicate keys",
            ));
            return Some(Value::Undefined);
        }
    }
    if let Value::Object(target_obj) = &target {
        let (target_keys, ordinary_missing_nonconfig) = {
            let o = target_obj.lock().unwrap();
            let target_keys = crate::object::descriptor_own_keys(&o);
            let ordinary = matches!(o.kind, ObjectKind::Ordinary)
                && !o.properties.contains_key(PROXY_TAG)
                && !o.properties.contains_key(PROXY_TARGET)
                && !o.properties.contains_key(PROXY_HANDLER)
                && !o.properties.contains_key("__sym_keys");
            let missing = ordinary.then(|| target_keys.iter().any(|key| {
                !seen.contains(key) && crate::object::is_nonconfig(&o, key)
            }));
            (target_keys, missing)
        };
        let omitted_nonconfig = match ordinary_missing_nonconfig {
            Some(missing) => missing,
            None => target_keys.iter().any(|key| {
                let desc = crate::object::own_property_descriptor(target_obj, key);
                !desc_bool(&desc, "configurable", true) && !seen.contains(key)
            }),
        };
        if omitted_nonconfig {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy ownKeys trap omitted a non-configurable key",
            ));
            return Some(Value::Undefined);
        }
        if target_is_not_extensible(&target) {
            for key in &target_keys {
                if !seen.contains(key) {
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Proxy ownKeys trap omitted a key from non-extensible target",
                    ));
                    return Some(Value::Undefined);
                }
            }
            if seen.len() != target_keys.len() {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy ownKeys trap added a key to a non-extensible target",
                ));
                return Some(Value::Undefined);
            }
        }
    }
    Some(result)
}

    fn object(object: Object) -> Value {
        Value::Object(Arc::new(Mutex::new(object)))
    }

    fn function(vm: &VM, namespace: &str, name: &str) -> Value {
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(namespace, name).unwrap());
        object(function)
    }

    fn list(keys: &[&str]) -> Value {
        object(Object::new_array(keys.iter().map(|key| crate::keys::string_value(key)).collect()))
    }

    fn setup(vm: &mut VM, target: Value, result: Value) -> (Value, Arc<Mutex<Object>>) {
        let Value::Object(result_object) = result else { panic!("array result"); };
        let callback_result = result_object.clone();
        vm.register_host_fn("ecma:test-ownkeys", "trap", Box::new(move |_ctx, _args| {
            Value::Object(callback_result.clone())
        }));
        let mut handler = Object::new();
        handler.properties.insert("ownKeys".into(), function(vm, "ecma:test-ownkeys", "trap"));
        (new_proxy(target, object(handler), None), result_object)
    }

    fn replace_keys(result: &Arc<Mutex<Object>>, keys: &[&str]) {
        let mut result = result.lock().unwrap();
        let ObjectKind::Array(values) = &mut result.kind else { unreachable!() };
        *values = keys.iter().map(|key| crate::keys::string_value(key)).collect();
    }

    #[test]
    fn ordinary_metadata_gates_omissions_and_preserves_result_identity() {
        let mut vm = VM::new();
        register(&mut vm);
        let mut target = Object::new();
        target.properties.insert("fixed".into(), Value::I32(1));
        target.properties.insert("optional".into(), Value::I32(2));
        let target = Arc::new(Mutex::new(target));
        crate::object::track_nonconfig(&target, "fixed");
        let (proxy, result) = setup(&mut vm, Value::Object(target), list(&["fixed"]));
        let own_keys = function(&vm, "ecma:proxy", "ownKeys");
        let actual = vm.invoke(&own_keys, &[proxy.clone()]).unwrap();
        assert!(matches!(actual, Value::Object(actual) if Arc::ptr_eq(&actual, &result)));
        replace_keys(&result, &["optional"]);
        assert!(vm.invoke(&own_keys, &[proxy]).is_err());
    }

    #[test]
    fn nonextensible_target_requires_exact_key_set() {
        let mut vm = VM::new();
        register(&mut vm);
        let mut target = Object::new();
        target.properties.insert("field".into(), Value::I32(1));
        crate::object::mark_not_extensible(&mut target);
        let (proxy, result) = setup(&mut vm, object(target), list(&["field"]));
        let own_keys = function(&vm, "ecma:proxy", "ownKeys");
        assert!(vm.invoke(&own_keys, &[proxy.clone()]).is_ok());
        replace_keys(&result, &[]);
        assert!(vm.invoke(&own_keys, &[proxy.clone()]).is_err());
        replace_keys(&result, &["field", "extra"]);
        assert!(vm.invoke(&own_keys, &[proxy]).is_err());
    }

    #[test]
    fn accessor_metadata_is_checked_without_running_getter() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:test-ownkeys", "getter", Box::new(|_ctx, _args| {
            panic!("ownKeys must not invoke accessors")
        }));
        let mut target = Object::new();
        target.properties.insert("__get_accessor".into(), function(&vm, "ecma:test-ownkeys", "getter"));
        let target = Arc::new(Mutex::new(target));
        crate::object::track_nonconfig(&target, "accessor");
        let (proxy, result) = setup(&mut vm, Value::Object(target), list(&["accessor"]));
        let own_keys = function(&vm, "ecma:proxy", "ownKeys");
        assert!(vm.invoke(&own_keys, &[proxy.clone()]).is_ok());
        replace_keys(&result, &[]);
        assert!(vm.invoke(&own_keys, &[proxy]).is_err());
    }

    #[test]
    fn target_metadata_is_read_after_trap_mutation() {
        let mut vm = VM::new();
        register(&mut vm);
        let mut target = Object::new();
        target.properties.insert("field".into(), Value::I32(1));
        let target = Arc::new(Mutex::new(target));
        let callback_target = target.clone();
        vm.register_host_fn("ecma:test-ownkeys", "mutatingTrap", Box::new(move |_ctx, _args| {
            crate::object::track_nonconfig(&callback_target, "field");
            list(&[])
        }));
        let mut handler = Object::new();
        handler.properties.insert("ownKeys".into(), function(&vm, "ecma:test-ownkeys", "mutatingTrap"));
        let proxy = new_proxy(Value::Object(target), object(handler), None);
        let own_keys = function(&vm, "ecma:proxy", "ownKeys");
        assert!(vm.invoke(&own_keys, &[proxy]).is_err());
    }

    #[test]
    fn exotic_array_length_keeps_descriptor_fallback() {
        let mut vm = VM::new();
        register(&mut vm);
        let (proxy, result) = setup(&mut vm, object(Object::new_array(vec![Value::I32(1)])), list(&["0", "length"]));
        let own_keys = function(&vm, "ecma:proxy", "ownKeys");
        assert!(vm.invoke(&own_keys, &[proxy.clone()]).is_ok());
        replace_keys(&result, &["0"]);
        assert!(vm.invoke(&own_keys, &[proxy]).is_err());
    }

    #[test]
    fn invalid_and_duplicate_trap_keys_are_rejected() {
        let mut vm = VM::new();
        register(&mut vm);
        let (proxy, result) = setup(&mut vm, object(Object::new()), list(&["same", "same"]));
        let own_keys = function(&vm, "ecma:proxy", "ownKeys");
        assert!(vm.invoke(&own_keys, &[proxy.clone()]).is_err());
        result.lock().unwrap().kind = ObjectKind::Array(vec![Value::I32(1)]);
        assert!(vm.invoke(&own_keys, &[proxy]).is_err());
    }

    #[test]
    #[ignore = "paired registered Proxy.ownKeys benchmark; not a conformance gate"]
    fn registered_own_keys_benchmark() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:test-ownkeys", "legacyOwnKeys", Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            legacy_own_keys_dispatch(ctx, &value).unwrap_or_else(|| target_own_keys(&value))
        }));
        let old = function(&vm, "ecma:test-ownkeys", "legacyOwnKeys");
        let new = function(&vm, "ecma:proxy", "ownKeys");
        for count in [16, 256, 1024] {
            let names: Vec<_> = (0..count).map(|index| format!("field_{index}")).collect();
            let mut target = Object::new();
            for (index, name) in names.iter().enumerate() {
                target.properties.insert(name.clone(), Value::I32(index as i32));
            }
            let result = object(Object::new_array(names.iter().map(|name| crate::keys::string_value(name)).collect()));
            let (proxy, _) = setup(&mut vm, object(target), result);
            let args = [proxy];
            for function in [&old, &new] {
                let Value::Object(result) = vm.invoke(function, &args).unwrap() else { panic!("keys"); };
                let result = result.lock().unwrap();
                let ObjectKind::Array(values) = &result.kind else { panic!("key array"); };
                assert_eq!(values.len(), count);
                for (value, name) in values.iter().zip(&names) {
                    assert!(matches!(value, Value::String(key) if key.as_ref() == name));
                }
            }
            let iterations = (16_384 / count).max(16);
            let mut old_time = Duration::ZERO;
            let mut new_time = Duration::ZERO;
            for sample in 0..6 {
                for old_path in [sample % 2 == 0, sample % 2 != 0] {
                    let function = if old_path { &old } else { &new };
                    let start = Instant::now();
                    for _ in 0..iterations { std::hint::black_box(vm.invoke(function, &args).unwrap()); }
                    if old_path { old_time += start.elapsed(); } else { new_time += start.elapsed(); }
                }
            }
            eprintln!("ordinary-ownKeys count={count}: old={old_time:?} new={new_time:?} ratio={:.3}x",
                old_time.as_secs_f64() / new_time.as_secs_f64());
        }
    }

fn legacy_own_keys_dispatch(ctx: &mut HostContext, value: &Value) -> Option<Value> {
    let proxy_obj = is_proxy(value)?;
    if proxy_is_revoked(&proxy_obj) {
        return Some(throw_revoked(ctx));
    }
    let (target, handler) = get_target_and_handler(&proxy_obj);
    let result = if let Some(trap) = get_trap(&handler, "ownKeys") {
        if is_callable(&trap) {
            call_trap(ctx, &handler, &trap, std::slice::from_ref(&target))
        } else {
            target_own_keys(&target)
        }
    } else {
        target_own_keys(&target)
    };
    let keys = array_values(&result);
    if !matches!(result, Value::Object(_)) {
        ctx.throw_value(make_type_error(
            ctx,
            "Proxy ownKeys trap must return an object",
        ));
        return Some(Value::Undefined);
    }
    let mut seen = std::collections::HashSet::new();
    for key in &keys {
        let text = key_string(key);
        if !seen.insert(text) {
            ctx.throw_value(make_type_error(
                ctx,
                "Proxy ownKeys trap returned duplicate keys",
            ));
            return Some(Value::Undefined);
        }
    }
    if let Value::Object(target_obj) = &target {
        let target_keys = {
            let o = target_obj.lock().unwrap();
            crate::object::descriptor_own_keys(&o)
        };
        for key in &target_keys {
            let desc = crate::object::own_property_descriptor(target_obj, key);
            if !desc_bool(&desc, "configurable", true) && !seen.contains(key) {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy ownKeys trap omitted a non-configurable key",
                ));
                return Some(Value::Undefined);
            }
        }
        if target_is_not_extensible(&target) {
            for key in &target_keys {
                if !seen.contains(key) {
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Proxy ownKeys trap omitted a key from non-extensible target",
                    ));
                    return Some(Value::Undefined);
                }
            }
        }
    }
    Some(result)
}
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:proxy",
        "apply",
        Box::new(|ctx: &mut HostContext, args: &[Value]| apply_dispatch(ctx, args)),
    );
    let proxy_apply_idx = vm.host_registry.get(proxy_apply_host_key()).copied();
    vm.register_host_fn(
        "ecma:proxy",
        "call",
        Box::new(|ctx: &mut HostContext, args: &[Value]| callable_proxy_dispatch(ctx, args)),
    );
    let proxy_call_idx = vm.host_registry.get(proxy_call_host_key()).copied();

    vm.register_host_fn(
        "ecma:proxy",
        "new",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let handler = args.get(1).cloned().unwrap_or(Value::Undefined);
            new_proxy(target, handler, proxy_call_idx)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "get",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let value = args.first().unwrap_or(&Value::Undefined);
            let key_value = args.get(1).unwrap_or(&Value::Undefined);
            get_dispatch(ctx, value, key_value)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "set",
        Box::new(set_dispatch),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "has",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let target = args.first().cloned().unwrap_or(Value::Undefined);
                    let undefined = Value::Undefined;
                    let key_ref = args.get(1).unwrap_or(&undefined);
                    return with_key_string(key_ref, |key| Value::Bool(target_has(&target, key)));
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let (target, handler) = get_target_and_handler(&proxy_obj);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            let key = key_string(&key_value);
            if let Some(trap) = get_trap(&handler, "has") {
                if is_callable(&trap) {
                    // §10.5.7 step 6: ToBoolean(trap result), not === true.
                    let result = call_trap(ctx, &handler, &trap, &[target.clone(), key_value]);
                    let present = crate::boolean::to_boolean(&result);
                    if !present
                        && (target_nonconfig_data_value(&target, &key).is_some()
                            || (!crate::object::value_is_extensible(&target)
                                && target_own_property_exists(&target, &key)))
                    {
                        ctx.throw_value(make_type_error(
                            ctx,
                            "Proxy has trap violated target property invariant",
                        ));
                        return Value::Undefined;
                    }
                    return Value::Bool(present);
                }
            }
            Value::Bool(target_has(&target, &key))
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "deleteProperty",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let target = args.first().cloned().unwrap_or(Value::Undefined);
                    let undefined = Value::Undefined;
                    let key_ref = args.get(1).unwrap_or(&undefined);
                    return with_key_string(key_ref, |key| {
                        Value::Bool(target_delete(&target, key))
                    });
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let (target, handler) = get_target_and_handler(&proxy_obj);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            let key = key_string(&key_value);
            if let Some(trap) = get_trap(&handler, "deleteProperty") {
                if is_callable(&trap) {
                    // §10.5.10 step 6: ToBoolean(trap result), not === true.
                    let result = call_trap(ctx, &handler, &trap, &[target.clone(), key_value]);
                    let success = crate::boolean::to_boolean(&result);
                    if success && target_nonconfig_data_value(&target, &key).is_some() {
                        ctx.throw_value(make_type_error(
                            ctx,
                            "Proxy deleteProperty trap violated non-configurable property invariant",
                        ));
                        return Value::Undefined;
                    }
                    return Value::Bool(success);
                }
            }
            Value::Bool(target_delete(&target, &key))
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "ownKeys",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let value = args.first().unwrap_or(&Value::Undefined);
            own_keys_dispatch(ctx, value).unwrap_or_else(|| target_own_keys(value))
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "getOwnPropertyDescriptor",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(desc) = get_own_property_descriptor_dispatch(ctx, &value, &key_value) {
                return desc;
            }
            let Value::Object(obj) = value else {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Reflect.getOwnPropertyDescriptor called on non-object",
                ));
                return Value::Undefined;
            };
            with_key_string(&key_value, |key| {
                crate::object::own_property_descriptor(&obj, key)
            })
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "defineProperty",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            let descriptor = args.get(2).cloned().unwrap_or(Value::Undefined);
            if let Some(proxy_obj) = is_proxy(&value) {
                return define_property_dispatch(ctx, &proxy_obj, key_value, descriptor)
                    .unwrap_or(Value::Bool(false));
            }
            let Value::Object(obj) = value else {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Reflect.defineProperty called on non-object",
                ));
                return Value::Undefined;
            };
            let key = key_string(&key_value);
            let mut o = obj.lock().unwrap();
            let exists = o.properties.contains_key(&key)
                || o.properties.contains_key(&getter_key(&key))
                || o.properties.contains_key(&setter_key(&key));
            if !exists && crate::object::is_not_extensible(&o) {
                return Value::Bool(false);
            }
            let Value::Object(desc_obj) = descriptor else {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Property descriptor must be an object",
                ));
                return Value::Undefined;
            };
            let desc = desc_obj.lock().unwrap();
            let has_value = desc.properties.contains_key("value");
            let has_writable = desc.properties.contains_key("writable");
            let has_get = desc.properties.contains_key("get");
            let has_set = desc.properties.contains_key("set");
            if (has_value || has_writable) && (has_get || has_set) {
                ctx.throw_value(make_type_error(ctx, "Invalid property descriptor"));
                return Value::Undefined;
            }
            if let Some(getter) = desc.properties.get("get").cloned() {
                o.properties.insert(getter_key(&key), getter);
                o.properties.shift_remove(&key);
            }
            if let Some(setter) = desc.properties.get("set").cloned() {
                o.properties.insert(setter_key(&key), setter);
                o.properties.shift_remove(&key);
            }
            if has_value || (!has_get && !has_set) {
                let value = desc
                    .properties
                    .get("value")
                    .cloned()
                    .unwrap_or(Value::Undefined);
                o.properties.insert(key.clone(), value);
            }
            let enumerable = desc
                .properties
                .get("enumerable")
                .map(crate::boolean::to_boolean)
                .unwrap_or(false);
            let writable = desc
                .properties
                .get("writable")
                .map(crate::boolean::to_boolean)
                .unwrap_or(false);
            let configurable = desc
                .properties
                .get("configurable")
                .map(crate::boolean::to_boolean)
                .unwrap_or(false);
            drop(desc);
            drop(o);
            crate::object::track_key(&obj, &key);
            if !enumerable {
                crate::object::track_nonenum(&obj, &key);
            }
            if !writable {
                crate::object::install_noop_setter(&mut obj.lock().unwrap(), &key);
            }
            if !configurable {
                crate::object::track_nonconfig(&obj, &key);
            }
            Value::Bool(true)
        }),
    );

    // Both delegate to the one [[GetPrototypeOf]] / [[SetPrototypeOf]]
    // implementation in `object.rs` — these used to answer a constant
    // (`null` / `true`) regardless of target, trap or invariant.
    vm.register_host_fn(
        "ecma:proxy",
        "getPrototypeOf",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            crate::object::get_prototype_of(ctx, &value).unwrap_or(Value::Undefined)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "setPrototypeOf",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let proto = args.get(1).cloned().unwrap_or(Value::Null);
            match crate::object::set_prototype_of(ctx, &value, &proto) {
                None => Value::Undefined,
                Some(success) => Value::Bool(success),
            }
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "isExtensible",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let value = args.first().cloned().unwrap_or(Value::Undefined);
                    if let Value::Object(_) = &value {
                        return Value::Bool(crate::object::value_is_extensible(&value));
                    }
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Reflect.isExtensible called on non-object",
                    ));
                    return Value::Undefined;
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let (target, handler) = get_target_and_handler(&proxy_obj);
            let target_extensible = crate::object::value_is_extensible(&target);
            let reported = if let Some(trap) = get_trap(&handler, "isExtensible") {
                if is_callable(&trap) {
                    crate::boolean::to_boolean(&call_trap(ctx, &handler, &trap, &[target.clone()]))
                } else {
                    target_extensible
                }
            } else {
                target_extensible
            };
            if reported != target_extensible {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy isExtensible trap result does not match target",
                ));
                return Value::Undefined;
            }
            Value::Bool(reported)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "preventExtensions",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let value = args.first().cloned().unwrap_or(Value::Undefined);
                    if let Value::Object(obj) = value {
                        crate::object::mark_not_extensible(&mut obj.lock().unwrap());
                        return Value::Bool(true);
                    }
                    ctx.throw_value(make_type_error(
                        ctx,
                        "Reflect.preventExtensions called on non-object",
                    ));
                    return Value::Undefined;
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let (target, handler) = get_target_and_handler(&proxy_obj);
            let success = if let Some(trap) = get_trap(&handler, "preventExtensions") {
                if is_callable(&trap) {
                    crate::boolean::to_boolean(&call_trap(ctx, &handler, &trap, &[target.clone()]))
                } else {
                    if let Value::Object(target_obj) = &target {
                        crate::object::mark_not_extensible(&mut target_obj.lock().unwrap());
                    }
                    true
                }
            } else {
                if let Value::Object(target_obj) = &target {
                    crate::object::mark_not_extensible(&mut target_obj.lock().unwrap());
                }
                true
            };
            if success && crate::object::value_is_extensible(&target) {
                ctx.throw_value(make_type_error(
                    ctx,
                    "Proxy preventExtensions trap returned true but target is still extensible",
                ));
                return Value::Undefined;
            }
            Value::Bool(success)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "construct",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_value = args.first().cloned().unwrap_or(Value::Undefined);
            let args_list = args.get(1).cloned().unwrap_or_else(|| {
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
            });
            construct_dispatch(ctx, &proxy_value, &args_list)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "revocable",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let handler = args.get(1).cloned().unwrap_or(Value::Undefined);
            let proxy = new_proxy(target, handler, proxy_call_idx);
            let proxy_clone = proxy.clone();
            // revoke is represented as an object with __revoke_target pointing to the proxy
            let mut revoke_obj = Object::new();
            revoke_obj.properties.reserve(7);
            if let Some(idx) = proxy_apply_idx {
                revoke_obj.kind = ObjectKind::HostFunction(idx);
                revoke_obj.properties.insert(
                    "__host_module".into(),
                    crate::keys::string_value("ecma:proxy"),
                );
                revoke_obj
                    .properties
                    .insert("__host_name".into(), crate::keys::string_value("apply"));
                revoke_obj
                    .properties
                    .insert("__host_idx".into(), Value::F64(idx as f64));
                revoke_obj.properties.insert(
                    "__proto__".into(),
                    crate::function::shared_function_prototype(),
                );
                revoke_obj
                    .properties
                    .insert("name".into(), crate::keys::string_value("revoke"));
                revoke_obj
                    .properties
                    .insert("length".into(), Value::F64(0.0));
            }
            revoke_obj
                .properties
                .insert("__revoke_target".into(), proxy_clone);
            let revoke_arc = vybe_runtime::heap::alloc(revoke_obj);
            if proxy_apply_idx.is_some() {
                revoke_arc.lock().unwrap().properties.insert(
                    "__bound_args".into(),
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                        Value::Object(revoke_arc.clone()),
                    ]))),
                );
            }
            let revoke = Value::Object(revoke_arc);
            let mut result = Object::new();
            result.properties.insert("proxy".into(), proxy);
            result.properties.insert("revoke".into(), revoke);
            Value::Object(vybe_runtime::heap::alloc(result))
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "callRevoke",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(revoke_obj)) = args.first() {
                let ro = revoke_obj.lock().unwrap();
                if let Some(Value::Object(proxy_obj)) = ro.properties.get("__revoke_target") {
                    proxy_obj
                        .lock()
                        .unwrap()
                        .properties
                        .insert(PROXY_REVOKED.into(), Value::Bool(true));
                }
            }
            Value::Undefined
        }),
    );
}
