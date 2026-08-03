use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

const PROXY_TAG: &str = "__vybe_js_proxy";
const PROXY_TARGET: &str = "__vybe_proxy_target";
const PROXY_HANDLER: &str = "__vybe_proxy_handler";
const PROXY_REVOKED: &str = "__vybe_proxy_revoked";

fn new_proxy(target: Value, handler: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(PROXY_TAG.into(), Value::I32(1));
    obj.properties.insert(PROXY_TARGET.into(), target);
    obj.properties.insert(PROXY_HANDLER.into(), handler);
    obj.properties
        .insert(PROXY_REVOKED.into(), Value::Bool(false));
    Value::Object(vybe_runtime::heap::alloc(obj))
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

fn get_target(proxy: &Arc<Mutex<Object>>) -> Value {
    proxy
        .lock()
        .unwrap()
        .properties
        .get(PROXY_TARGET)
        .cloned()
        .unwrap_or(Value::Undefined)
}

fn get_handler(proxy: &Arc<Mutex<Object>>) -> Value {
    proxy
        .lock()
        .unwrap()
        .properties
        .get(PROXY_HANDLER)
        .cloned()
        .unwrap_or(Value::Undefined)
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
                    if let Ok(i) = key.parse::<usize>() {
                        return v.get(i).cloned().unwrap_or(Value::Undefined);
                    }
                }
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    if key == "length" {
                        return Value::F64(crate::typedarray::ta_live_length(ta) as f64);
                    }
                    if let Ok(i) = key.parse::<usize>() {
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
            if let Ok(i) = key.parse::<usize>() {
                return s
                    .chars()
                    .nth(i)
                    .map(|c| Value::String(Arc::from(c.to_string().as_str())))
                    .unwrap_or(Value::Undefined);
            }
            Value::Undefined
        }
        _ => Value::Undefined,
    }
}

fn target_set(target: &Value, key: &str, val: Value) {
    if let Value::Object(obj) = target {
        {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(v) = &mut o.kind {
                if let Ok(i) = key.parse::<usize>() {
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
    match value {
        Value::String(text) => text.to_string(),
        // Symbol-keyed properties are stored under the Display form
        // ("Symbol(desc)") — same convention as ecma:reflect/object.
        _ => format!("{}", value),
    }
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

fn is_callable(value: &Value) -> bool {
    matches!(value, Value::Object(obj)
        if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
}

fn call_trap(ctx: &mut HostContext, handler: &Value, trap: &Value, args: &[Value]) -> Value {
    crate::function::invoke_with_explicit_this(ctx, trap, handler.clone(), args)
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
                    if let Ok(i) = key.parse::<usize>() {
                        return i < v.len();
                    }
                }
            }
            crate::object::proto_walk_get(obj, key).is_some()
        }
        _ => false,
    }
}

fn target_delete(target: &Value, key: &str) -> bool {
    if let Value::Object(obj) = target {
        let mut o = obj.lock().unwrap();
        if crate::object::is_nonconfig(&o, key) {
            return false;
        }
        o.properties.shift_remove(key);
    }
    true
}

fn target_own_keys(target: &Value) -> Value {
    if let Value::Object(obj) = target {
        let o = obj.lock().unwrap();
        let keys = crate::object::ordered_own_string_keys(&o)
            .into_iter()
            .map(|key| Value::String(Arc::from(key.as_str())))
            .collect();
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
    }
    Vec::new()
}

/// [[OwnPropertyKeys]] for proxy exotic objects — ECMA-262 §10.5.11.
/// Returns None when `value` is not a proxy so callers (Object.keys etc.)
/// keep their ordinary-object path.
pub fn own_keys_dispatch(ctx: &mut HostContext, value: &Value) -> Option<Value> {
    let proxy_obj = is_proxy(value)?;
    if proxy_is_revoked(&proxy_obj) {
        return Some(throw_revoked(ctx));
    }
    let handler = get_handler(&proxy_obj);
    let target = get_target(&proxy_obj);
    if let Some(trap) = get_trap(&handler, "ownKeys") {
        if is_callable(&trap) {
            return Some(call_trap(
                ctx,
                &handler,
                &trap,
                std::slice::from_ref(&target),
            ));
        }
    }
    Some(target_own_keys(&target))
}

/// [[Construct]] dispatch — ECMA-262 §10.5.13 for proxy exotic objects,
/// ordinary construct for everything else. Shared by `ecma:proxy.construct`
/// and `ecma:reflect.construct` (§28.1.2 routes through [[Construct]]).
pub fn construct_dispatch(
    ctx: &mut HostContext,
    constructor: &Value,
    args_list: &Value,
) -> Value {
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
    let effective_nt = new_target.unwrap_or_else(|| constructor.clone());
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
    if let Some(proxy_obj) = is_proxy(constructor) {
        if proxy_is_revoked(&proxy_obj) {
            return throw_revoked(ctx);
        }
        let handler = get_handler(&proxy_obj);
        let target = get_target(&proxy_obj);
        if let Some(trap) = get_trap(&handler, "construct") {
            if is_callable(&trap) {
                let result = call_trap(
                    ctx,
                    &handler,
                    &trap,
                    &[target.clone(), args_list.clone(), constructor.clone()],
                );
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
                    Some(Value::Undefined) | None => Value::String(Arc::from("")),
                    Some(other) => Value::String(Arc::from(format!("{other}").as_str())),
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
    if let Value::Object(target_obj) = constructor {
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
    let result = crate::function::invoke_with_explicit_this(
        ctx,
        constructor,
        this_obj.clone(),
        &array_values(args_list),
    );
    if matches!(result, Value::Object(_)) {
        result
    } else {
        this_obj
    }
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:proxy",
        "new",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let handler = args.get(1).cloned().unwrap_or(Value::Undefined);
            new_proxy(target, handler)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "get",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let key = args.get(1).map(key_string).unwrap_or_default();
                    return target_get(args.first().unwrap_or(&Value::Undefined), &key);
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            let key = key_string(&key_value);
            if let Some(trap) = get_trap(&handler, "get") {
                if is_callable(&trap) {
                    let receiver = args.first().cloned().unwrap_or(Value::Undefined);
                    return call_trap(ctx, &handler, &trap, &[target, key_value, receiver]);
                }
            }
            // With no `get` trap, [[Get]] forwards to the target with the
            // PROXY as receiver. `__proto__` is an accessor on
            // %Object.prototype% whose getter is [[GetPrototypeOf]](receiver)
            // — so it lands back on the proxy and its `getPrototypeOf` trap.
            // Reading the target's raw slot instead skips the trap entirely.
            if key == "__proto__" {
                let receiver = args.first().cloned().unwrap_or(Value::Undefined);
                return crate::object::get_prototype_of(ctx, &receiver)
                    .unwrap_or(Value::Undefined);
            }
            target_get(&target, &key)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "set",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let target = args.first().cloned().unwrap_or(Value::Undefined);
                    let key = args.get(1).map(key_string).unwrap_or_default();
                    let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                    target_set(&target, &key, val);
                    return Value::Bool(true);
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            let key = key_string(&key_value);
            let val = args.get(2).cloned().unwrap_or(Value::Undefined);
            if let Some(trap) = get_trap(&handler, "set") {
                if is_callable(&trap) {
                    let receiver = args.first().cloned().unwrap_or(Value::Undefined);
                    // §10.5.9 step 6: ToBoolean(trap result), not === true.
                    let result =
                        call_trap(ctx, &handler, &trap, &[target, key_value, val, receiver]);
                    return Value::Bool(crate::boolean::to_boolean(&result));
                }
            }
            // Mirror of the `get` path: with no `set` trap, assigning
            // `__proto__` is %Object.prototype%'s setter, i.e.
            // [[SetPrototypeOf]](receiver) — which is where a
            // `setPrototypeOf` trap gets its chance to run.
            if key == "__proto__" {
                let receiver = args.first().cloned().unwrap_or(Value::Undefined);
                return match crate::object::set_prototype_of(ctx, &receiver, &val) {
                    None => Value::Undefined,
                    Some(success) => Value::Bool(success),
                };
            }
            target_set(&target, &key, val);
            Value::Bool(true)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "has",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let target = args.first().cloned().unwrap_or(Value::Undefined);
                    let key = args.get(1).map(key_string).unwrap_or_default();
                    return Value::Bool(target_has(&target, &key));
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            let key = key_string(&key_value);
            if let Some(trap) = get_trap(&handler, "has") {
                if is_callable(&trap) {
                    // §10.5.7 step 6: ToBoolean(trap result), not === true.
                    let result = call_trap(ctx, &handler, &trap, &[target, key_value]);
                    return Value::Bool(crate::boolean::to_boolean(&result));
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
                    let key = args.get(1).map(key_string).unwrap_or_default();
                    return Value::Bool(target_delete(&target, &key));
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return throw_revoked(ctx);
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            let key = key_string(&key_value);
            if let Some(trap) = get_trap(&handler, "deleteProperty") {
                if is_callable(&trap) {
                    // §10.5.10 step 6: ToBoolean(trap result), not === true.
                    let result = call_trap(ctx, &handler, &trap, &[target, key_value]);
                    return Value::Bool(crate::boolean::to_boolean(&result));
                }
            }
            Value::Bool(target_delete(&target, &key))
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "ownKeys",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            own_keys_dispatch(ctx, &value).unwrap_or_else(|| target_own_keys(&value))
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Bool(false),
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Bool(false);
            }
            Value::Bool(true)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "preventExtensions",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Bool(false),
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Bool(false);
            }
            Value::Bool(true)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "apply",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    let target = args.first().cloned().unwrap_or(Value::Undefined);
                    // Proxy.revocable's revoke function (§28.2.2.1) is
                    // modelled as an object carrying __revoke_target.
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
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            let args_list = args.get(2).cloned().unwrap_or_else(|| {
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
            });
            if let Some(trap) = get_trap(&handler, "apply") {
                if is_callable(&trap) {
                    return call_trap(ctx, &handler, &trap, &[target, this_arg, args_list]);
                }
            }
            let invoke_args = array_values(&args_list);
            crate::function::invoke_with_explicit_this(ctx, &target, this_arg, &invoke_args)
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let handler = args.get(1).cloned().unwrap_or(Value::Undefined);
            let proxy = new_proxy(target, handler);
            let proxy_clone = proxy.clone();
            // revoke is represented as an object with __revoke_target pointing to the proxy
            let mut revoke_obj = Object::new();
            revoke_obj
                .properties
                .insert("__revoke_target".into(), proxy_clone);
            let revoke = Value::Object(vybe_runtime::heap::alloc(revoke_obj));
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
