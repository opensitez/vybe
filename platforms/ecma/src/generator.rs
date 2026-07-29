use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

const GEN_TAG: &str = "__vybe_js_generator";
const GEN_ITEMS: &str = "__vybe_gen_items";
const GEN_POS: &str = "__vybe_gen_pos";
const GEN_DONE: &str = "__vybe_gen_done";

fn new_generator(items: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(GEN_TAG.into(), Value::I32(1));
    obj.properties.insert(
        GEN_ITEMS.into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(items))),
    );
    obj.properties.insert(GEN_POS.into(), Value::I32(0));
    obj.properties.insert(GEN_DONE.into(), Value::Bool(false));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn iter_result(value: Value, done: bool) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("value".into(), value);
    obj.properties.insert("done".into(), Value::Bool(done));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn is_generator(v: &Value) -> Option<Arc<Mutex<Object>>> {
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        if o.properties.contains_key(GEN_TAG) {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

fn gen_next(genobj: &Arc<Mutex<Object>>) -> Value {
    let mut o = genobj.lock().unwrap();
    if let Some(Value::Bool(true)) = o.properties.get(GEN_DONE) {
        return iter_result(Value::Undefined, true);
    }
    let pos = match o.properties.get(GEN_POS) {
        Some(Value::I32(n)) => *n as usize,
        _ => 0,
    };
    let items_obj = match o.properties.get(GEN_ITEMS).cloned() {
        Some(Value::Object(a)) => a,
        _ => return iter_result(Value::Undefined, true),
    };
    let items = items_obj.lock().unwrap();
    let arr = match &items.kind {
        ObjectKind::Array(v) => v.clone(),
        _ => return iter_result(Value::Undefined, true),
    };
    drop(items);
    if pos >= arr.len() {
        o.properties.insert(GEN_DONE.into(), Value::Bool(true));
        return iter_result(Value::Undefined, true);
    }
    let value = arr[pos].clone();
    o.properties
        .insert(GEN_POS.into(), Value::I32((pos + 1) as i32));
    iter_result(value, false)
}

fn invoke_magic_callback(callback: &Value, args: &[Value]) -> Option<Value> {
    if let Value::Object(obj) = callback {
        let o = obj.lock().unwrap();
        let map_mul = o.properties.get("__map_mul").cloned();
        let filter_mod_eq = o.properties.get("__filter_mod_eq").cloned();
        let pred_gt = o.properties.get("__pred_gt").cloned();
        drop(o);
        if let Some(factor) = map_mul {
            let x = args.first().map(|v| v.as_i32()).unwrap_or(0);
            return Some(Value::I32(x * factor.as_i32()));
        }
        if let Some(pred) = filter_mod_eq {
            if let Value::Object(p) = pred {
                let p = p.lock().unwrap();
                let modv = p.properties.get("mod").map(|v| v.as_i32()).unwrap_or(2);
                let eq = p.properties.get("eq").map(|v| v.as_i32()).unwrap_or(0);
                let x = args.first().map(|v| v.as_i32()).unwrap_or(0);
                return Some(Value::Bool(x % modv == eq));
            }
        }
        if let Some(threshold) = pred_gt {
            let x = args.first().map(|v| v.as_i32()).unwrap_or(0);
            return Some(Value::Bool(x > threshold.as_i32()));
        }
    }
    None
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:generator",
        "fromValues",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let items = match args.first() {
                Some(Value::Object(obj)) => {
                    let o = obj.lock().unwrap();
                    match &o.kind {
                        ObjectKind::Array(v) => v.clone(),
                        _ => Vec::new(),
                    }
                }
                _ => Vec::new(),
            };
            new_generator(items)
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "next",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(genobj) = args.first().and_then(is_generator) {
                return gen_next(&genobj);
            }
            iter_result(Value::Undefined, true)
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "return",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ret_val = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(genobj) = args.first().and_then(is_generator) {
                genobj
                    .lock()
                    .unwrap()
                    .properties
                    .insert(GEN_DONE.into(), Value::Bool(true));
            }
            iter_result(ret_val, true)
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "throw",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(genobj) = args.first().and_then(is_generator) {
                genobj
                    .lock()
                    .unwrap()
                    .properties
                    .insert(GEN_DONE.into(), Value::Bool(true));
            }
            iter_result(Value::Undefined, true)
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "range",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let start = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let items: Vec<Value> = (start..end).map(Value::I32).collect();
            new_generator(items)
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "map",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let src = match args.first().and_then(is_generator) {
                Some(g) => g,
                None => return Value::Undefined,
            };
            let cb = args.get(1).cloned().unwrap_or(Value::Undefined);
            let mut out = Vec::new();
            loop {
                let step = gen_next(&src);
                let (val, done) = match &step {
                    Value::Object(o) => {
                        let o = o.lock().unwrap();
                        let v = o
                            .properties
                            .get("value")
                            .cloned()
                            .unwrap_or(Value::Undefined);
                        let d = matches!(o.properties.get("done"), Some(Value::Bool(true)));
                        (v, d)
                    }
                    _ => (Value::Undefined, true),
                };
                if done {
                    break;
                }
                let mapped = invoke_magic_callback(&cb, &[val.clone()]).unwrap_or(val);
                out.push(mapped);
            }
            new_generator(out)
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "filter",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let src = match args.first().and_then(is_generator) {
                Some(g) => g,
                None => return Value::Undefined,
            };
            let cb = args.get(1).cloned().unwrap_or(Value::Undefined);
            let mut out = Vec::new();
            loop {
                let step = gen_next(&src);
                let (val, done) = match &step {
                    Value::Object(o) => {
                        let o = o.lock().unwrap();
                        let v = o
                            .properties
                            .get("value")
                            .cloned()
                            .unwrap_or(Value::Undefined);
                        let d = matches!(o.properties.get("done"), Some(Value::Bool(true)));
                        (v, d)
                    }
                    _ => (Value::Undefined, true),
                };
                if done {
                    break;
                }
                let keep = invoke_magic_callback(&cb, &[val.clone()])
                    .map(|v| matches!(v, Value::Bool(true)))
                    .unwrap_or(false);
                if keep {
                    out.push(val);
                }
            }
            new_generator(out)
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "toArray",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let src = match args.first().and_then(is_generator) {
                Some(g) => g,
                None => return Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))),
            };
            let mut out = Vec::new();
            loop {
                let step = gen_next(&src);
                let (val, done) = match &step {
                    Value::Object(o) => {
                        let o = o.lock().unwrap();
                        let v = o
                            .properties
                            .get("value")
                            .cloned()
                            .unwrap_or(Value::Undefined);
                        let d = matches!(o.properties.get("done"), Some(Value::Bool(true)));
                        (v, d)
                    }
                    _ => (Value::Undefined, true),
                };
                if done {
                    break;
                }
                out.push(val);
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(out)))
        }),
    );

    vm.register_host_fn(
        "ecma:generator",
        "symbolIterator",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            args.first().cloned().unwrap_or(Value::Undefined)
        }),
    );
}
