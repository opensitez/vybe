use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

const PROXY_TAG: &str = "__vybe_js_proxy";
const PROXY_TARGET: &str = "__vybe_proxy_target";
const PROXY_HANDLER: &str = "__vybe_proxy_handler";
const PROXY_REVOKED: &str = "__vybe_proxy_revoked";
const PROXY_SET_TRAP_DEPTH: &str = "__vybe_proxy_set_trap_depth";

fn new_proxy(target: Value, handler: Value, call_idx: Option<usize>) -> Value {
    let callable = call_idx.filter(|_| is_callable(&target));
    let mut obj = Object::new();
    obj.properties.insert(PROXY_TAG.into(), Value::I32(1));
    obj.properties.insert(PROXY_TARGET.into(), target.clone());
    obj.properties.insert(PROXY_HANDLER.into(), handler);
    obj.properties
        .insert(PROXY_REVOKED.into(), Value::Bool(false));
    if let Some(idx) = callable {
        obj.kind = ObjectKind::HostFunction(idx);
        obj.properties.insert(
            "__host_module".into(),
            Value::String(Arc::from("ecma:proxy")),
        );
        obj.properties
            .insert("__host_name".into(), Value::String(Arc::from("call")));
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
                let handler = get_handler(&proto_proxy);
                let proto_target = get_target(&proto_proxy);
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

fn apply_dispatch(ctx: &mut HostContext, args: &[Value]) -> Value {
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
    if !is_callable(&target) {
        ctx.throw_value(make_type_error(ctx, "Proxy target is not callable"));
        return Value::Undefined;
    }
    let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
    let args_list = args
        .get(2)
        .cloned()
        .unwrap_or_else(|| Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))));
    if let Some(trap) = get_trap(&handler, "apply") {
        if is_callable(&trap) {
            return call_trap(ctx, &handler, &trap, &[target, this_arg, args_list]);
        }
    }
    let invoke_args = array_values(&args_list);
    crate::function::invoke_with_explicit_this(ctx, &target, this_arg, &invoke_args)
}

pub fn get_dispatch(ctx: &mut HostContext, value: &Value, key_value: &Value) -> Value {
    let proxy_obj = match is_proxy(value) {
        Some(p) => p,
        None => {
            let key = key_string(key_value);
            return target_get(value, &key);
        }
    };
    if proxy_is_revoked(&proxy_obj) {
        return throw_revoked(ctx);
    }
    let handler = get_handler(&proxy_obj);
    let target = get_target(&proxy_obj);
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

fn callable_proxy_dispatch(ctx: &mut HostContext, args: &[Value]) -> Value {
    let proxy = args.first().cloned().unwrap_or(Value::Undefined);
    let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
    let args_list = Value::Object(vybe_runtime::heap::alloc(Object::new_array(
        args.iter().skip(2).cloned().collect(),
    )));
    apply_dispatch(ctx, &[proxy, this_arg, args_list])
}

fn is_callable(value: &Value) -> bool {
    matches!(value, Value::Object(obj)
        if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
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
    let previous = ctx.get_global(PROXY_SET_TRAP_DEPTH);
    let depth = previous.as_f64() as i32;
    ctx.set_global(PROXY_SET_TRAP_DEPTH, Value::I32(depth + 1));
    let result = call_trap(ctx, handler, trap, args);
    ctx.set_global(PROXY_SET_TRAP_DEPTH, previous);
    result
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
                    if let Ok(i) = key.parse::<usize>() {
                        return i < v.len() && !crate::array::is_array_hole(&o, i);
                    }
                }
            }
            crate::object::proto_walk_get(obj, key).is_some()
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
            if let Ok(i) = key.parse::<usize>() {
                return i < values.len() && !crate::array::is_array_hole(&o, i);
            }
        }
        ObjectKind::TypedArray(ta) => {
            if let Ok(i) = key.parse::<usize>() {
                return i < crate::typedarray::ta_live_length(ta);
            }
        }
        _ => {}
    }
    o.properties.contains_key(key)
        || o.properties.contains_key(&format!("__get_{}", key))
        || o.properties.contains_key(&format!("__set_{}", key))
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
        if let Ok(i) = key.parse::<usize>() {
            if i < values.len() && !crate::array::is_array_hole(&o, i) {
                return Some((values[i].clone(), true));
            }
        }
    }
    if o.properties.contains_key(&format!("__get_{}", key)) {
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
            if let Ok(i) = key.parse::<usize>() {
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
        if let Some(length) = o
            .properties
            .get("length")
            .map(|v| v.as_i32().max(0) as usize)
        {
            return (0..length)
                .map(|i| {
                    o.properties
                        .get(&i.to_string())
                        .cloned()
                        .unwrap_or(Value::Undefined)
                })
                .collect();
        }
    }
    Vec::new()
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
    let handler = get_handler(&proxy_obj);
    let target = get_target(&proxy_obj);
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
    let handler = get_handler(proxy_obj);
    let target = get_target(proxy_obj);
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
pub fn own_keys_dispatch(ctx: &mut HostContext, value: &Value) -> Option<Value> {
    let proxy_obj = is_proxy(value)?;
    if proxy_is_revoked(&proxy_obj) {
        return Some(throw_revoked(ctx));
    }
    let handler = get_handler(&proxy_obj);
    let target = get_target(&proxy_obj);
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
        let handler = get_handler(&proxy_obj);
        let target = get_target(&proxy_obj);
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
        "apply",
        Box::new(|ctx: &mut HostContext, args: &[Value]| apply_dispatch(ctx, args)),
    );
    let proxy_apply_idx = vm
        .host_registry
        .get(&("ecma:proxy".to_string(), "apply".to_string()))
        .copied();
    vm.register_host_fn(
        "ecma:proxy",
        "call",
        Box::new(|ctx: &mut HostContext, args: &[Value]| callable_proxy_dispatch(ctx, args)),
    );
    let proxy_call_idx = vm
        .host_registry
        .get(&("ecma:proxy".to_string(), "call".to_string()))
        .copied();

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
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
            get_dispatch(ctx, &value, &key_value)
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
                    let key_value = args.get(1).cloned().unwrap_or(Value::Undefined);
                    target_set(ctx, &target, key_value, &key, val);
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
                    let result = call_set_trap(
                        ctx,
                        &handler,
                        &trap,
                        &[target.clone(), key_value, val.clone(), receiver],
                    );
                    let success = crate::boolean::to_boolean(&result);
                    if success {
                        if let Some((actual, writable)) = target_nonconfig_data_value(&target, &key)
                        {
                            if !writable && val != actual {
                                ctx.throw_value(make_type_error(
                                    ctx,
                                    "Proxy set trap violated non-configurable property invariant",
                                ));
                                return Value::Undefined;
                            }
                        }
                    }
                    return Value::Bool(success);
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
            if !crate::object::value_is_extensible(&target) && !target_has(&target, &key) {
                return Value::Bool(false);
            }
            if get_trap(&handler, "defineProperty").is_some() {
                let mut desc = Object::new();
                desc.properties.insert("value".into(), val.clone());
                desc.properties.insert("writable".into(), Value::Bool(true));
                desc.properties
                    .insert("enumerable".into(), Value::Bool(true));
                desc.properties
                    .insert("configurable".into(), Value::Bool(true));
                if let Some(result) = define_property_dispatch(
                    ctx,
                    &proxy_obj,
                    key_value.clone(),
                    Value::Object(vybe_runtime::heap::alloc(desc)),
                ) {
                    return Value::Bool(crate::boolean::to_boolean(&result));
                }
            }
            target_set(ctx, &target, key_value, &key, val);
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
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            own_keys_dispatch(ctx, &value).unwrap_or_else(|| target_own_keys(&value))
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
            crate::object::own_property_descriptor(&obj, &key_string(&key_value))
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
                || o.properties.contains_key(&format!("__get_{}", key))
                || o.properties.contains_key(&format!("__set_{}", key));
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
                o.properties.insert(format!("__get_{}", key), getter);
                o.properties.shift_remove(&key);
            }
            if let Some(setter) = desc.properties.get("set").cloned() {
                o.properties.insert(format!("__set_{}", key), setter);
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
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
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
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
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
            if let Some(idx) = proxy_apply_idx {
                revoke_obj.kind = ObjectKind::HostFunction(idx);
                revoke_obj.properties.insert(
                    "__host_module".into(),
                    Value::String(Arc::from("ecma:proxy")),
                );
                revoke_obj
                    .properties
                    .insert("__host_name".into(), Value::String(Arc::from("apply")));
                revoke_obj
                    .properties
                    .insert("__host_idx".into(), Value::F64(idx as f64));
                revoke_obj.properties.insert(
                    "__proto__".into(),
                    crate::function::shared_function_prototype(),
                );
                revoke_obj
                    .properties
                    .insert("name".into(), Value::String(Arc::from("revoke")));
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
