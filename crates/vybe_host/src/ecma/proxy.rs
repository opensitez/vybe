use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

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
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_proxy(v: &Value) -> Option<Arc<Mutex<Object>>> {
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        if o.properties.contains_key(PROXY_TAG) {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

fn proxy_is_revoked(proxy: &Arc<Mutex<Object>>) -> bool {
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
            let o = obj.lock().unwrap();
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        }
        _ => Value::Undefined,
    }
}

fn target_set(target: &Value, key: &str, val: Value) {
    if let Value::Object(obj) = target {
        obj.lock().unwrap().properties.insert(key.to_string(), val);
    }
}

fn trap_return_value(trap: &Value) -> Option<Value> {
    if let Value::Object(t) = trap {
        let t = t.lock().unwrap();
        t.properties.get("__trap_return").cloned()
    } else {
        None
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => {
                    return target_get(
                        args.first().unwrap_or(&Value::Undefined),
                        args.get(1)
                            .and_then(|v| {
                                if let Value::String(s) = v {
                                    Some(s.as_ref())
                                } else {
                                    None
                                }
                            })
                            .unwrap_or(""),
                    );
                }
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Undefined;
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key = match args.get(1) {
                Some(Value::String(s)) => s.as_ref().to_string(),
                _ => return Value::Undefined,
            };
            if let Some(trap) = get_trap(&handler, "get") {
                if let Some(ret) = trap_return_value(&trap) {
                    return ret;
                }
                return target_get(&target, &key);
            }
            target_get(&target, &key)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "set",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Bool(false),
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Bool(false);
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key = match args.get(1) {
                Some(Value::String(s)) => s.as_ref().to_string(),
                _ => return Value::Bool(false),
            };
            let val = args.get(2).cloned().unwrap_or(Value::Undefined);
            if let Some(_trap) = get_trap(&handler, "set") {
                return Value::Bool(true);
            }
            target_set(&target, &key, val);
            Value::Bool(true)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "has",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Bool(false),
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Bool(false);
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key = match args.get(1) {
                Some(Value::String(s)) => s.as_ref().to_string(),
                _ => return Value::Bool(false),
            };
            if let Some(_trap) = get_trap(&handler, "has") {
                return Value::Bool(true);
            }
            if let Value::Object(t) = target {
                return Value::Bool(t.lock().unwrap().properties.contains_key(&key));
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "deleteProperty",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Bool(false),
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Bool(false);
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            let key = match args.get(1) {
                Some(Value::String(s)) => s.as_ref().to_string(),
                _ => return Value::Bool(false),
            };
            if let Some(_trap) = get_trap(&handler, "deleteProperty") {
                return Value::Bool(true);
            }
            if let Value::Object(t) = target {
                t.lock().unwrap().properties.remove(&key);
            }
            Value::Bool(true)
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "ownKeys",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Undefined,
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Undefined;
            }
            let handler = get_handler(&proxy_obj);
            let target = get_target(&proxy_obj);
            if let Some(_trap) = get_trap(&handler, "ownKeys") {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))));
            }
            if let Value::Object(t) = target {
                let keys: Vec<Value> = t
                    .lock()
                    .unwrap()
                    .properties
                    .keys()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "getPrototypeOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Null,
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Null;
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "setPrototypeOf",
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Undefined,
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Undefined;
            }
            let handler = get_handler(&proxy_obj);
            if let Some(trap) = get_trap(&handler, "apply") {
                if let Some(ret) = trap_return_value(&trap) {
                    return ret;
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:proxy",
        "construct",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proxy_obj = match args.first().and_then(is_proxy) {
                Some(p) => p,
                None => return Value::Undefined,
            };
            if proxy_is_revoked(&proxy_obj) {
                return Value::Undefined;
            }
            let target = get_target(&proxy_obj);
            match target {
                Value::Object(_) => {
                    let obj = Object::new();
                    Value::Object(Arc::new(Mutex::new(obj)))
                }
                _ => Value::Undefined,
            }
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
            let revoke = Value::Object(Arc::new(Mutex::new(revoke_obj)));
            let mut result = Object::new();
            result.properties.insert("proxy".into(), proxy);
            result.properties.insert("revoke".into(), revoke);
            Value::Object(Arc::new(Mutex::new(result)))
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
