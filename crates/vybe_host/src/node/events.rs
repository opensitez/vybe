//! `node:events` — Node.js EventEmitter module.
//!
//! Reference: <https://nodejs.org/api/events.html>.
//!
//! Listeners are stored as properties on the emitter Object so that
//! shared Arc<Mutex<Object>> clones see each other's mutations.
//! Storage keys:
//!   `__ev_<name>`  → Array of regular listeners
//!   `__evo_<name>` → Array of once listeners

use std::sync::{Arc, Mutex};
use vybe_bytecode::VM;
use vybe_bytecode::value::{Object, ObjectKind, Value};

fn empty_array() -> Value {
    arr_val(vec![])
}

fn arr_val(elems: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object {
        kind: ObjectKind::Array(elems),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: Vec::new(),
    })))
}

fn get_array(emitter: &Object, key: &str) -> Vec<Value> {
    if let Some(Value::Object(arr)) = emitter.properties.get(key) {
        let arr = arr.lock().unwrap();
        if let ObjectKind::Array(elems) = &arr.kind {
            return elems.clone();
        }
    }
    vec![]
}

fn set_array(emitter: &mut Object, key: &str, elems: Vec<Value>) {
    emitter.properties.insert(key.to_string(), arr_val(elems));
}

fn ev_key(event: &str) -> String {
    format!("__ev_{event}")
}
fn evo_key(event: &str) -> String {
    format!("__evo_{event}")
}

fn get_emitter_mut(v: &Value) -> Option<std::sync::MutexGuard<Object>> {
    if let Value::Object(o) = v {
        Some(o.lock().unwrap())
    } else {
        None
    }
}

fn listener_count_for(emitter: &Object, event: &str) -> usize {
    get_array(emitter, &ev_key(event)).len() + get_array(emitter, &evo_key(event)).len()
}

fn make_emitter() -> Value {
    let mut o = Object::new();
    o.properties.insert("__maxlisteners".into(), Value::I32(10));
    // Stub method names so has_method checks pass
    for m in [
        "on",
        "once",
        "off",
        "emit",
        "addListener",
        "removeListener",
        "removeAllListeners",
        "listenerCount",
        "listeners",
        "rawListeners",
        "eventNames",
        "setMaxListeners",
        "getMaxListeners",
        "prependListener",
        "prependOnceListener",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(Arc::new(Mutex::new(o)))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:events",
        "EventEmitter",
        Box::new(|_ctx, _args| make_emitter()),
    );

    // on / addListener
    for name in ["on", "addListener"] {
        vm.register_host_fn(
            "node:events",
            name,
            Box::new(|_ctx, args| {
                let event = match args.get(1) {
                    Some(Value::String(s)) => s.to_string(),
                    _ => return Value::Undefined,
                };
                let listener = args.get(2).cloned().unwrap_or(Value::Null);
                if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                    let k = ev_key(&event);
                    let mut arr = get_array(&em, &k);
                    arr.push(listener);
                    set_array(&mut em, &k, arr);
                }
                Value::Undefined
            }),
        );
    }

    // prependListener
    vm.register_host_fn(
        "node:events",
        "prependListener",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return Value::Undefined,
            };
            let listener = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let k = ev_key(&event);
                let mut arr = get_array(&em, &k);
                arr.insert(0, listener);
                set_array(&mut em, &k, arr);
            }
            Value::Undefined
        }),
    );

    // once
    vm.register_host_fn(
        "node:events",
        "once",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return Value::Undefined,
            };
            let listener = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let k = evo_key(&event);
                let mut arr = get_array(&em, &k);
                arr.push(listener);
                set_array(&mut em, &k, arr);
            }
            Value::Undefined
        }),
    );

    // prependOnceListener
    vm.register_host_fn(
        "node:events",
        "prependOnceListener",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return Value::Undefined,
            };
            let listener = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let k = evo_key(&event);
                let mut arr = get_array(&em, &k);
                arr.insert(0, listener);
                set_array(&mut em, &k, arr);
            }
            Value::Undefined
        }),
    );

    // off / removeListener — removes first matching listener from regular or once
    for name in ["off", "removeListener"] {
        vm.register_host_fn(
            "node:events",
            name,
            Box::new(|_ctx, args| {
                let event = match args.get(1) {
                    Some(Value::String(s)) => s.to_string(),
                    _ => return Value::Undefined,
                };
                let listener = args.get(2).cloned().unwrap_or(Value::Null);
                if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                    let k = ev_key(&event);
                    let mut arr = get_array(&em, &k);
                    if let Some(pos) = arr.iter().position(|v| v == &listener) {
                        arr.remove(pos);
                        set_array(&mut em, &k, arr);
                    } else {
                        let ko = evo_key(&event);
                        let mut oarr = get_array(&em, &ko);
                        if let Some(pos) = oarr.iter().position(|v| v == &listener) {
                            oarr.remove(pos);
                            set_array(&mut em, &ko, oarr);
                        }
                    }
                }
                Value::Undefined
            }),
        );
    }

    // removeAllListeners([event])
    vm.register_host_fn(
        "node:events",
        "removeAllListeners",
        Box::new(|_ctx, args| {
            if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                if let Some(Value::String(event)) = args.get(1) {
                    let ev = event.to_string();
                    em.properties.remove(&ev_key(&ev));
                    em.properties.remove(&evo_key(&ev));
                } else {
                    let keys: Vec<String> = em
                        .properties
                        .keys()
                        .filter(|k| k.starts_with("__ev_") || k.starts_with("__evo_"))
                        .cloned()
                        .collect();
                    for k in keys {
                        em.properties.remove(&k);
                    }
                }
            }
            Value::Undefined
        }),
    );

    // emit(ee, event) → Bool
    vm.register_host_fn(
        "node:events",
        "emit",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return Value::Bool(false),
            };
            if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let regular = get_array(&em, &ev_key(&event));
                let once = get_array(&em, &evo_key(&event));
                let had_listeners = !regular.is_empty() || !once.is_empty();
                if !once.is_empty() {
                    set_array(&mut em, &evo_key(&event), vec![]);
                }
                return Value::Bool(had_listeners);
            }
            Value::Bool(false)
        }),
    );

    // listenerCount(ee, event) → I32
    vm.register_host_fn(
        "node:events",
        "listenerCount",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return Value::I32(0),
            };
            if let Some(em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                return Value::I32(listener_count_for(&em, &event) as i32);
            }
            Value::I32(0)
        }),
    );

    // listeners(ee, event) → Array
    vm.register_host_fn(
        "node:events",
        "listeners",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return empty_array(),
            };
            if let Some(em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let mut all = get_array(&em, &ev_key(&event));
                all.extend(get_array(&em, &evo_key(&event)));
                return arr_val(all);
            }
            empty_array()
        }),
    );

    // rawListeners(ee, event) → Array (same as listeners for our purposes)
    vm.register_host_fn(
        "node:events",
        "rawListeners",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return empty_array(),
            };
            if let Some(em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let mut all = get_array(&em, &ev_key(&event));
                all.extend(get_array(&em, &evo_key(&event)));
                return arr_val(all);
            }
            empty_array()
        }),
    );

    // eventNames(ee) → Array of strings
    vm.register_host_fn(
        "node:events",
        "eventNames",
        Box::new(|_ctx, args| {
            if let Some(em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let mut seen = std::collections::HashSet::new();
                for k in em.properties.keys() {
                    if let Some(ev) = k.strip_prefix("__ev_").or_else(|| k.strip_prefix("__evo_")) {
                        seen.insert(ev.to_string());
                    }
                }
                // Only include events that still have listeners
                let names: Vec<Value> = seen
                    .into_iter()
                    .filter(|ev| listener_count_for(&em, ev) > 0)
                    .map(|ev| Value::String(Arc::from(ev.as_str())))
                    .collect();
                return arr_val(names);
            }
            empty_array()
        }),
    );

    // getMaxListeners(ee) → I32
    vm.register_host_fn(
        "node:events",
        "getMaxListeners",
        Box::new(|_ctx, args| {
            if let Some(em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                return em
                    .properties
                    .get("__maxlisteners")
                    .cloned()
                    .unwrap_or(Value::I32(10));
            }
            Value::I32(10)
        }),
    );

    // setMaxListeners(ee, n)
    vm.register_host_fn(
        "node:events",
        "setMaxListeners",
        Box::new(|_ctx, args| {
            let n = match args.get(1) {
                Some(Value::I32(n)) => Value::I32(*n),
                Some(Value::F64(f)) => Value::I32(*f as i32),
                _ => Value::I32(10),
            };
            if let Some(mut em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                em.properties.insert("__maxlisteners".into(), n);
            }
            Value::Undefined
        }),
    );

    // defaultMaxListeners → I32(10)
    vm.register_host_fn(
        "node:events",
        "defaultMaxListeners",
        Box::new(|_ctx, _args| Value::I32(10)),
    );

    // getEventListeners(ee, event) → Array
    vm.register_host_fn(
        "node:events",
        "getEventListeners",
        Box::new(|_ctx, args| {
            let event = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => return empty_array(),
            };
            if let Some(em) = get_emitter_mut(args.first().unwrap_or(&Value::Undefined)) {
                let mut all = get_array(&em, &ev_key(&event));
                all.extend(get_array(&em, &evo_key(&event)));
                return arr_val(all);
            }
            empty_array()
        }),
    );

    // errorMonitor — Symbol-like marker
    vm.register_host_fn(
        "node:events",
        "errorMonitor",
        Box::new(|_ctx, _args| Value::String(Arc::from("Symbol(events.errorMonitor)"))),
    );
}
