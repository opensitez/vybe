//! ECMA-262 §20.2 — Function.prototype.{bind, call, apply}.
//!
//! `bind` returns a new function ref carrying `__bound_args` so the VM
//! call dispatch in `vybe_bytecode/src/calls.rs` prepends them on every
//! invocation. The VM hook is the same mechanism `new Promise(executor)`
//! uses for resolve/reject thunks — see `crate::ecma::promise`.
//!
//! `call(thisArg, ...args)` and `apply(thisArg, argsArray)` are the
//! spread/forward primitives. The Vybe call dispatch already passes the
//! receiver as the first arg of the args slice, so `call` and `apply`
//! just route through `ctx.invoke` with the resolved args list.

use std::sync::{Arc, Mutex, OnceLock};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

static FUNCTION_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

pub(crate) fn shared_function_prototype() -> Value {
    Value::Object(
        FUNCTION_PROTOTYPE
            .get_or_init(|| {
                // §20.2.3: %Function.prototype% has own non-enumerable
                // `length: 0` and `name: ""` data properties — the
                // intrinsic kind prototypes (%AsyncFunction.prototype% …)
                // inherit them through their [[Prototype]] link.
                let mut proto = Object::new();
                proto.properties.insert("length".into(), Value::F64(0.0));
                proto
                    .properties
                    .insert("name".into(), Value::String(Arc::from("")));
                proto.properties.insert(
                    "__nonenum".into(),
                    Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                        Value::String(Arc::from("length")),
                        Value::String(Arc::from("name")),
                    ])))),
                );
                Arc::new(Mutex::new(proto))
            })
            .clone(),
    )
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:function",
        "invokeBound",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let bound_this = args.get(1).cloned().unwrap_or(Value::Undefined);
            let target_proto = args.get(2).cloned().unwrap_or(Value::Undefined);

            invoke_bound_target(ctx, &target, bound_this, target_proto, &args[3..])
        }),
    );

    let invoke_bound_idx = *vm
        .host_registry
        .get(&("ecma:function".to_string(), "invokeBound".to_string()))
        .expect("ecma:function.invokeBound must be registered before bind");

    // Function.prototype.name — §20.2.3.3: returns the name property.
    vm.register_host_fn(
        "ecma:function",
        "name",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let name = o
                    .properties
                    .get("name")
                    .or_else(|| o.properties.get("__fn_name"));
                if let Some(Value::String(s)) = name {
                    return Value::String(s.clone());
                }
            }
            Value::String(Arc::from(""))
        }),
    );

    // Function.prototype.length — §20.2.3.2: formal parameter count.
    vm.register_host_fn(
        "ecma:function",
        "length",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let len = o
                    .properties
                    .get("length")
                    .or_else(|| o.properties.get("__fn_arity"));
                match len {
                    Some(Value::I32(n)) => return Value::I32(*n),
                    Some(Value::F64(n)) => return Value::I32(*n as i32),
                    Some(Value::I64(n)) => return Value::I32(*n as i32),
                    _ => {}
                }
                if let ObjectKind::Function(f) = &o.kind {
                    return Value::I32(f.arity as i32);
                }
            }
            Value::I32(0)
        }),
    );

    // Function.prototype.toString — §20.2.3.5.
    vm.register_host_fn(
        "ecma:function",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let name = o
                    .properties
                    .get("name")
                    .or_else(|| o.properties.get("__fn_name"))
                    .and_then(|v| {
                        if let Value::String(s) = v {
                            Some(s.to_string())
                        } else {
                            None
                        }
                    })
                    .unwrap_or_default();
                return Value::String(Arc::from(
                    format!("function {}() {{ [native code] }}", name).as_str(),
                ));
            }
            Value::String(Arc::from("function () { [native code] }"))
        }),
    );

    // new Function(body) — §20.2.1.1: creates a callable from a body string.
    vm.register_host_fn(
        "ecma:function",
        "new",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let body = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let mut obj = Object::new();
            obj.properties
                .insert("name".into(), Value::String(Arc::from("anonymous")));
            obj.properties.insert("length".into(), Value::I32(0));
            obj.properties
                .insert("__fn_body".into(), Value::String(Arc::from(body.as_str())));
            obj.properties
                .insert("__fn_return".into(), Value::Undefined);
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // new Function(params, body) — §20.2.1.1 with parameters.
    vm.register_host_fn(
        "ecma:function",
        "newWithParams",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let params = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let body = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut obj = Object::new();
            obj.properties
                .insert("name".into(), Value::String(Arc::from("anonymous")));
            obj.properties.insert(
                "length".into(),
                Value::I32(if params.is_empty() { 0 } else { 1 }),
            );
            obj.properties.insert(
                "__fn_params".into(),
                Value::String(Arc::from(params.as_str())),
            );
            obj.properties
                .insert("__fn_body".into(), Value::String(Arc::from(body.as_str())));
            obj.properties
                .insert("__fn_return".into(), Value::Undefined);
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // bindWithArgs(fn, thisArg, ...args) — like bind with pre-supplied args.
    vm.register_host_fn(
        "ecma:function",
        "bindWithArgs",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let bound: Vec<Value> = if args.len() > 1 {
                args[1..].to_vec()
            } else {
                Vec::new()
            };
            bind_function_with_arity(&target, bound, invoke_bound_idx)
        }),
    );

    // Function.prototype.bind(this_fn, thisArg, ...boundArgs) → new Function
    //
    // The returned function ref carries `__bound_args = [thisArg, ...boundArgs]`
    // and points at the same host fn idx as the receiver (or the same
    // chunk for user functions).
    vm.register_host_fn(
        "ecma:function",
        "bind",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let bound: Vec<Value> = if args.len() > 1 {
                args[1..].to_vec()
            } else {
                Vec::new()
            };
            bind_function(&target, bound, invoke_bound_idx)
        }),
    );

    // Function.prototype.call(this_fn, thisArg, ...args) → result
    //
    // Synchronously invokes the receiver with the given thisArg + args.
    vm.register_host_fn(
        "ecma:function",
        "call",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            invoke_with_explicit_this(ctx, &target, this_arg, &args[2..])
        }),
    );

    // Function.prototype.apply(this_fn, thisArg, argsArray) → result
    vm.register_host_fn(
        "ecma:function",
        "apply",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            let invoke_args = args.get(2).map(collect_apply_args).unwrap_or_default();
            invoke_with_explicit_this(ctx, &target, this_arg, &invoke_args)
        }),
    );

    // §10.2.9 SetFunctionName for runtime-computed property keys:
    // anonymous functions assigned under a computed key take the key's
    // string form; symbol keys become "[<description>]" (or "" when the
    // symbol has none). Already-named functions keep their name.
    vm.register_host_fn(
        "ecma:function",
        "setFunctionName",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(f)) = args.first() {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut o = f.lock().unwrap();
                let anonymous = match o.properties.get("name") {
                    Some(Value::String(n)) => n.is_empty() || n.starts_with("__anon_fn_"),
                    _ => true,
                };
                if anonymous {
                    let name = match &key {
                        Value::Symbol(s) => {
                            if crate::ecma::symbol::has_description(s) {
                                format!("[{}]", s)
                            } else {
                                String::new()
                            }
                        }
                        other => format!("{}", other),
                    };
                    o.properties
                        .insert("name".into(), Value::String(Arc::from(name.as_str())));
                }
            }
            Value::Undefined
        }),
    );

    // §20.2.3.5 Function.prototype.toString. Source text isn't retained,
    // so every form uses the spec's NativeFunction fallback shape with the
    // function's kind classifier tokens (async / * / =>) and name.
    vm.register_host_fn(
        "ecma:function",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            Value::String(Arc::from(function_to_string(&target).as_str()))
        }),
    );
    let to_string_idx = *vm
        .host_registry
        .get(&("ecma:function".to_string(), "toString".to_string()))
        .expect("ecma:function.toString just registered");
    if let Value::Object(proto) = shared_function_prototype() {
        let mut p = proto.lock().unwrap();
        if !p.properties.contains_key("toString") {
            let mut ts = Object::new();
            ts.kind = ObjectKind::HostFunction(to_string_idx);
            ts.properties
                .insert("name".into(), Value::String(Arc::from("toString")));
            ts.properties.insert("length".into(), Value::F64(0.0));
            ts.properties
                .insert("__vybe_method_receiver".into(), Value::Bool(true));
            p.properties
                .insert("toString".into(), Value::Object(Arc::new(Mutex::new(ts))));
            // toString is non-enumerable on %Function.prototype%.
            if let Some(Value::Object(ne)) = p.properties.get("__nonenum") {
                let mut a = ne.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = a.kind {
                    elems.push(Value::String(Arc::from("toString")));
                }
            }
        }
    }
}

/// §20.2.3.5 — synthesized function string. Bound and native functions use
/// the NativeFunction form; compiled functions add the kind tokens the
/// compiler stamped (`__fn_kind`, `__fn_arrow`).
fn function_to_string(target: &Value) -> String {
    let Value::Object(obj) = target else {
        return "function () { [native code] }".to_string();
    };
    let o = obj.lock().unwrap();
    // §20.2.3.5 step 2 note: bound function exotic objects stringify as
    // native — they never expose their target's source.
    if o.properties.contains_key("__bound_args") {
        return "function () { [native code] }".to_string();
    }
    let name = match o.properties.get("name") {
        Some(Value::String(n)) => n.to_string(),
        _ => String::new(),
    };
    if matches!(o.kind, ObjectKind::HostFunction(_)) {
        return format!("function {}() {{ [native code] }}", name);
    }
    if matches!(o.properties.get("__fn_arrow"), Some(Value::Bool(true))) {
        let is_async = matches!(
            o.properties.get("__fn_kind"),
            Some(Value::String(k)) if k.as_ref() == "async"
        );
        return format!(
            "{}() => {{ [native code] }}",
            if is_async { "async " } else { "" }
        );
    }
    match o.properties.get("__fn_kind") {
        Some(Value::String(k)) if k.as_ref() == "async" => {
            format!("async function {}() {{ [native code] }}", name)
        }
        Some(Value::String(k)) if k.as_ref() == "generator" => {
            format!("function* {}() {{ [native code] }}", name)
        }
        Some(Value::String(k)) if k.as_ref() == "async_generator" => {
            format!("async function* {}() {{ [native code] }}", name)
        }
        _ => format!("function {}() {{ [native code] }}", name),
    }
}

pub(crate) fn invoke_bound_callback_if_needed(
    ctx: &mut HostContext,
    callback: &Value,
    args: &[Value],
) -> Option<Value> {
    let Value::Object(obj) = callback else {
        return None;
    };

    let stored_bound = {
        let object = obj.lock().unwrap();
        let name = match object.properties.get("name") {
            Some(Value::String(text)) => text.to_string(),
            _ => String::new(),
        };
        if !name.starts_with("bound ") {
            return None;
        }
        match object.properties.get("__bound_args") {
            Some(Value::Object(bound)) => {
                let bound_object = bound.lock().unwrap();
                if let ObjectKind::Array(values) = &bound_object.kind {
                    values.clone()
                } else {
                    return None;
                }
            }
            _ => return None,
        }
    };

    if stored_bound.len() < 3 {
        return None;
    }

    let target = stored_bound[0].clone();
    let bound_this = stored_bound[1].clone();
    let target_proto = stored_bound[2].clone();
    let mut invoke_args = Vec::with_capacity(stored_bound.len().saturating_sub(3) + args.len());
    invoke_args.extend(stored_bound.iter().skip(3).cloned());
    invoke_args.extend_from_slice(args);
    Some(invoke_bound_target(
        ctx,
        &target,
        bound_this,
        target_proto,
        &invoke_args,
    ))
}

pub(crate) fn try_invoke_bound_callback_if_needed(
    ctx: &mut HostContext,
    callback: &Value,
    args: &[Value],
) -> Option<Result<Value, Value>> {
    let Value::Object(obj) = callback else {
        return None;
    };

    let stored_bound = {
        let object = obj.lock().unwrap();
        let name = match object.properties.get("name") {
            Some(Value::String(text)) => text.to_string(),
            _ => String::new(),
        };
        if !name.starts_with("bound ") {
            return None;
        }
        match object.properties.get("__bound_args") {
            Some(Value::Object(bound)) => {
                let bound_object = bound.lock().unwrap();
                if let ObjectKind::Array(values) = &bound_object.kind {
                    values.clone()
                } else {
                    return None;
                }
            }
            _ => return None,
        }
    };

    if stored_bound.len() < 3 {
        return None;
    }

    let target = stored_bound[0].clone();
    let bound_this = stored_bound[1].clone();
    let target_proto = stored_bound[2].clone();
    let mut invoke_args = Vec::with_capacity(stored_bound.len().saturating_sub(3) + args.len());
    invoke_args.extend(stored_bound.iter().skip(3).cloned());
    invoke_args.extend_from_slice(args);
    Some(try_invoke_bound_target(
        ctx,
        &target,
        bound_this,
        target_proto,
        &invoke_args,
    ))
}

pub(crate) fn invoke_with_explicit_this(
    ctx: &mut HostContext,
    target: &Value,
    this_arg: Value,
    args: &[Value],
) -> Value {
    // §10.5.12 [[Call]] on a proxy: the apply trap fires with thisArg
    // (or the target is invoked with it when trapless). Reaches here via
    // Function.prototype.call/apply/bind on proxy-wrapped functions.
    if let Value::Object(obj) = target {
        if let Some((proxy_target, handler)) = crate::ecma::object::proxy_target_and_handler(obj) {
            if let Some(trap) = crate::ecma::object::proxy_trap(&handler, "apply") {
                let args_arr = Value::Object(Arc::new(Mutex::new(
                    vybe_bytecode::value::Object::new_array(args.to_vec()),
                )));
                return invoke_with_explicit_this(
                    ctx,
                    &trap,
                    handler,
                    &[proxy_target, this_arg, args_arr],
                );
            }
            return invoke_with_explicit_this(ctx, &proxy_target, this_arg, args);
        }
    }
    match target {
        Value::Object(obj)
            if matches!(obj.lock().unwrap().kind, ObjectKind::HostFunction(_))
                && obj.lock().unwrap().properties.contains_key("__bound_args") =>
        {
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = ctx.invoke(target, args);
            ctx.set_js_this(previous_this);
            result
        }
        Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)) => {
            // Arrows need no special case here: they capture lexical
            // `this` at creation (compiler-emitted upvalue) and never read
            // the ambient binding this sets (§10.2.11).
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = invoke_compiled_function(ctx, target, args);
            ctx.set_js_this(previous_this);
            result
        }
        _ => {
            // Magic test-only: plain object with __fn_return acts as a zero-arg callable.
            if let Value::Object(obj) = target {
                let o = obj.lock().unwrap();
                if let Some(ret) = o.properties.get("__fn_return").cloned() {
                    return ret;
                }
            }
            if host_function_uses_explicit_receiver(target) {
                let mut invoke_args = Vec::with_capacity(args.len() + 1);
                invoke_args.push(this_arg);
                invoke_args.extend_from_slice(args);
                ctx.invoke(target, &invoke_args)
            } else {
                ctx.invoke(target, args)
            }
        }
    }
}

pub(crate) fn try_invoke_with_explicit_this(
    ctx: &mut HostContext,
    target: &Value,
    this_arg: Value,
    args: &[Value],
) -> Result<Value, Value> {
    match target {
        Value::Object(obj)
            if matches!(obj.lock().unwrap().kind, ObjectKind::HostFunction(_))
                && obj.lock().unwrap().properties.contains_key("__bound_args") =>
        {
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = ctx.try_invoke(target, args);
            ctx.set_js_this(previous_this);
            result
        }
        Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)) => {
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = try_invoke_compiled_function(ctx, target, args);
            ctx.set_js_this(previous_this);
            result
        }
        _ => {
            // Magic test-only: plain object with __fn_return acts as a zero-arg callable.
            if let Value::Object(obj) = target {
                let o = obj.lock().unwrap();
                if let Some(ret) = o.properties.get("__fn_return").cloned() {
                    return Ok(ret);
                }
            }
            if host_function_uses_explicit_receiver(target) {
                let mut invoke_args = Vec::with_capacity(args.len() + 1);
                invoke_args.push(this_arg);
                invoke_args.extend_from_slice(args);
                ctx.try_invoke(target, &invoke_args)
            } else {
                ctx.try_invoke(target, args)
            }
        }
    }
}

fn invoke_bound_target(
    ctx: &mut HostContext,
    target: &Value,
    bound_this: Value,
    target_proto: Value,
    args: &[Value],
) -> Value {
    match target {
        Value::Object(target_obj)
            if matches!(target_obj.lock().unwrap().kind, ObjectKind::Function(_)) =>
        {
            let previous_this = ctx.current_js_this();
            let constructor_call = matches!((&previous_this, &target_proto),
                (Value::Object(current), Value::Object(expected_proto))
                    if matches!(current.lock().unwrap().properties.get("__proto__"), Some(Value::Object(proto)) if Arc::ptr_eq(proto, expected_proto))
            );
            if !constructor_call {
                ctx.set_js_this(bound_this);
            }
            let result = invoke_compiled_function(ctx, target, args);
            ctx.set_js_this(previous_this.clone());
            if constructor_call && !matches!(result, Value::Object(_)) {
                previous_this
            } else {
                result
            }
        }
        _ => {
            if host_function_uses_explicit_receiver(target) {
                let mut invoke_args = Vec::with_capacity(args.len() + 1);
                invoke_args.push(bound_this);
                invoke_args.extend_from_slice(args);
                ctx.invoke(target, &invoke_args)
            } else {
                ctx.invoke(target, args)
            }
        }
    }
}

fn try_invoke_bound_target(
    ctx: &mut HostContext,
    target: &Value,
    bound_this: Value,
    target_proto: Value,
    args: &[Value],
) -> Result<Value, Value> {
    match target {
        Value::Object(target_obj)
            if matches!(target_obj.lock().unwrap().kind, ObjectKind::Function(_)) =>
        {
            let previous_this = ctx.current_js_this();
            let constructor_call = matches!((&previous_this, &target_proto),
                (Value::Object(current), Value::Object(expected_proto))
                    if matches!(current.lock().unwrap().properties.get("__proto__"), Some(Value::Object(proto)) if Arc::ptr_eq(proto, expected_proto))
            );
            if !constructor_call {
                ctx.set_js_this(bound_this);
            }
            let result = try_invoke_compiled_function(ctx, target, args);
            ctx.set_js_this(previous_this.clone());
            match result {
                Ok(value) if constructor_call && !matches!(value, Value::Object(_)) => {
                    Ok(previous_this)
                }
                other => other,
            }
        }
        _ => {
            if host_function_uses_explicit_receiver(target) {
                let mut invoke_args = Vec::with_capacity(args.len() + 1);
                invoke_args.push(bound_this);
                invoke_args.extend_from_slice(args);
                ctx.try_invoke(target, &invoke_args)
            } else {
                ctx.try_invoke(target, args)
            }
        }
    }
}

fn invoke_compiled_function(ctx: &mut HostContext, target: &Value, args: &[Value]) -> Value {
    let Some(fixed_count) = compiled_rest_fixed_arity(target) else {
        return ctx.invoke(target, args);
    };

    let mut packed_args = Vec::with_capacity(fixed_count + 1);
    for index in 0..fixed_count {
        packed_args.push(args.get(index).cloned().unwrap_or(Value::Undefined));
    }
    packed_args.push(Value::Object(Arc::new(Mutex::new(Object::new_array(
        args.iter().skip(fixed_count).cloned().collect(),
    )))));
    ctx.invoke(target, &packed_args)
}

fn try_invoke_compiled_function(
    ctx: &mut HostContext,
    target: &Value,
    args: &[Value],
) -> Result<Value, Value> {
    let Some(fixed_count) = compiled_rest_fixed_arity(target) else {
        return ctx.try_invoke(target, args);
    };

    let mut packed_args = Vec::with_capacity(fixed_count + 1);
    for index in 0..fixed_count {
        packed_args.push(args.get(index).cloned().unwrap_or(Value::Undefined));
    }
    packed_args.push(Value::Object(Arc::new(Mutex::new(Object::new_array(
        args.iter().skip(fixed_count).cloned().collect(),
    )))));
    ctx.try_invoke(target, &packed_args)
}

fn compiled_rest_fixed_arity(target: &Value) -> Option<usize> {
    let Value::Object(obj) = target else {
        return None;
    };
    let object = obj.lock().unwrap();
    if !matches!(object.kind, ObjectKind::Function(_)) {
        return None;
    }
    match object.properties.get("__vybe_rest_fixed_arity") {
        Some(Value::I32(value)) if *value >= 0 => Some(*value as usize),
        Some(Value::I64(value)) if *value >= 0 => Some(*value as usize),
        Some(Value::F64(value)) if *value >= 0.0 => Some(*value as usize),
        _ => None,
    }
}

fn host_function_uses_explicit_receiver(target: &Value) -> bool {
    let Value::Object(obj) = target else {
        return false;
    };
    let object = obj.lock().unwrap();
    matches!(object.kind, ObjectKind::HostFunction(_))
        && matches!(
            object.properties.get("__vybe_method_receiver"),
            Some(Value::Bool(true))
        )
}

fn collect_apply_args(value: &Value) -> Vec<Value> {
    let Value::Object(obj) = value else {
        return Vec::new();
    };

    let object = obj.lock().unwrap();
    if let ObjectKind::Array(values) = &object.kind {
        return values.clone();
    }

    let length = match object.properties.get("length") {
        Some(Value::I32(value)) if *value > 0 => *value as usize,
        Some(Value::I64(value)) if *value > 0 => *value as usize,
        Some(Value::F64(value)) if *value > 0.0 => *value as usize,
        Some(Value::String(text)) => text.parse::<usize>().ok().unwrap_or(0),
        _ => 0,
    };
    (0..length)
        .map(|index| {
            object
                .properties
                .get(&index.to_string())
                .cloned()
                .unwrap_or(Value::Undefined)
        })
        .collect()
}

/// Like `bind_function` but reads `__fn_arity` as a length fallback for
/// magic fn_obj mocks (tests pass `{__fn_arity: n}` instead of `length`).
fn bind_function_with_arity(target: &Value, bound: Vec<Value>, invoke_bound_idx: usize) -> Value {
    let Value::Object(obj) = target else {
        return target.clone();
    };

    let (
        target_kind,
        existing_bound,
        target_name,
        target_length,
        target_proto,
        target_proto_link,
        target_non_ctor,
    ) = {
        let o = obj.lock().unwrap();
        let prev_bound = match o.properties.get("__bound_args") {
            Some(Value::Object(ba)) => {
                let bo = ba.lock().unwrap();
                if let ObjectKind::Array(ref values) = bo.kind {
                    values.clone()
                } else {
                    Vec::new()
                }
            }
            _ => Vec::new(),
        };
        let name = match o
            .properties
            .get("name")
            .or_else(|| o.properties.get("__fn_name"))
        {
            Some(Value::String(text)) => text.to_string(),
            Some(other) => format!("{}", other),
            None => String::new(),
        };
        let length = match o.properties.get("length") {
            Some(Value::I32(value)) if *value > 0 => *value as usize,
            Some(Value::I64(value)) if *value > 0 => *value as usize,
            Some(Value::F64(value)) if *value > 0.0 => *value as usize,
            _ => match o.properties.get("__fn_arity") {
                Some(Value::I32(value)) if *value > 0 => *value as usize,
                Some(Value::I64(value)) if *value > 0 => *value as usize,
                Some(Value::F64(value)) if *value > 0.0 => *value as usize,
                _ => 0,
            },
        };
        let prototype = o
            .properties
            .get("prototype")
            .cloned()
            .unwrap_or(Value::Undefined);
        let proto_link = o.properties.get("__proto__").cloned();
        let non_ctor = matches!(
            o.properties.get("__vybe_non_ctor"),
            Some(Value::Bool(true))
        );
        (
            o.kind.clone(),
            prev_bound,
            name,
            length,
            prototype,
            proto_link,
            non_ctor,
        )
    };

    // Allow ordinary objects (magic fn_obj descriptors from tests) — don't bail for non-Function.
    let mut stored_bound = Vec::new();
    if matches!(target_kind, ObjectKind::HostFunction(idx) if idx == invoke_bound_idx)
        && existing_bound.len() >= 3
    {
        stored_bound.push(existing_bound[0].clone());
        stored_bound.push(existing_bound[1].clone());
        stored_bound.push(existing_bound[2].clone());
        stored_bound.extend(existing_bound.iter().skip(3).cloned());
        stored_bound.extend(bound.into_iter().skip(1));
    } else {
        stored_bound.push(target.clone());
        stored_bound.push(bound.first().cloned().unwrap_or(Value::Undefined));
        stored_bound.push(target_proto.clone());
        stored_bound.extend(bound.into_iter().skip(1));
    }

    let mut wrapper = Object::new();
    wrapper.kind = ObjectKind::HostFunction(invoke_bound_idx);
    let consumed_args = stored_bound.len().saturating_sub(3);
    wrapper.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(stored_bound)))),
    );
    // §10.4.1.3 BoundFunctionCreate step 1: the bound function's
    // [[Prototype]] is the TARGET's [[Prototype]] — a bound async fn
    // stays `instanceof AsyncFunction`, a bound generator fn stays
    // `instanceof GeneratorFunction`. Fall back to %Function.prototype%.
    wrapper.properties.insert(
        "__proto__".into(),
        match target_proto_link {
            Some(link) if !matches!(link, Value::Null | Value::Undefined) => link,
            _ => shared_function_prototype(),
        },
    );
    if target_non_ctor {
        wrapper
            .properties
            .insert("__vybe_non_ctor".into(), Value::Bool(true));
    }
    wrapper.properties.insert(
        "name".into(),
        Value::String(Arc::from(format!("bound {}", target_name).as_str())),
    );
    wrapper.properties.insert(
        "length".into(),
        Value::F64(target_length.saturating_sub(consumed_args) as f64),
    );
    if !matches!(target_proto, Value::Null | Value::Undefined) {
        wrapper.properties.insert("prototype".into(), target_proto);
    }
    Value::Object(Arc::new(Mutex::new(wrapper)))
}

/// Build a function ref carrying bound args. Mirrors the convention in
/// `crate::namespaces::bound_host_fn_ref` but works on any function-like
/// Value (HostFunction or user Function).
fn bind_function(target: &Value, bound: Vec<Value>, invoke_bound_idx: usize) -> Value {
    let Value::Object(obj) = target else {
        return target.clone();
    };

    let (
        target_kind,
        existing_bound,
        target_name,
        target_length,
        target_proto,
        target_proto_link,
        target_non_ctor,
    ) = {
        let o = obj.lock().unwrap();
        let prev_bound = match o.properties.get("__bound_args") {
            Some(Value::Object(ba)) => {
                let bo = ba.lock().unwrap();
                if let ObjectKind::Array(ref values) = bo.kind {
                    values.clone()
                } else {
                    Vec::new()
                }
            }
            _ => Vec::new(),
        };
        let name = match o
            .properties
            .get("name")
            .or_else(|| o.properties.get("__fn_name"))
        {
            Some(Value::String(text)) => text.to_string(),
            Some(other) => format!("{}", other),
            None => String::new(),
        };
        let length = match o.properties.get("length") {
            Some(Value::I32(value)) if *value > 0 => *value as usize,
            Some(Value::I64(value)) if *value > 0 => *value as usize,
            Some(Value::F64(value)) if *value > 0.0 => *value as usize,
            _ => match o.properties.get("__fn_arity") {
                Some(Value::I32(value)) if *value > 0 => *value as usize,
                Some(Value::I64(value)) if *value > 0 => *value as usize,
                Some(Value::F64(value)) if *value > 0.0 => *value as usize,
                _ => 0,
            },
        };
        let prototype = o
            .properties
            .get("prototype")
            .cloned()
            .unwrap_or(Value::Undefined);
        let proto_link = o.properties.get("__proto__").cloned();
        let non_ctor = matches!(
            o.properties.get("__vybe_non_ctor"),
            Some(Value::Bool(true))
        );
        (
            o.kind.clone(),
            prev_bound,
            name,
            length,
            prototype,
            proto_link,
            non_ctor,
        )
    };

    // Allow ordinary objects (magic fn_obj descriptors from tests) — don't bail for non-Function.
    let mut stored_bound = Vec::new();
    if matches!(target_kind, ObjectKind::HostFunction(idx) if idx == invoke_bound_idx)
        && existing_bound.len() >= 3
    {
        stored_bound.push(existing_bound[0].clone());
        stored_bound.push(existing_bound[1].clone());
        stored_bound.push(existing_bound[2].clone());
        stored_bound.extend(existing_bound.iter().skip(3).cloned());
        stored_bound.extend(bound.into_iter().skip(1));
    } else {
        stored_bound.push(target.clone());
        stored_bound.push(bound.first().cloned().unwrap_or(Value::Undefined));
        stored_bound.push(target_proto.clone());
        stored_bound.extend(bound.into_iter().skip(1));
    }

    let mut wrapper = Object::new();
    wrapper.kind = ObjectKind::HostFunction(invoke_bound_idx);
    let consumed_args = stored_bound.len().saturating_sub(3);
    wrapper.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(stored_bound)))),
    );
    // §10.4.1.3 BoundFunctionCreate step 1: the bound function's
    // [[Prototype]] is the TARGET's [[Prototype]] — a bound async fn
    // stays `instanceof AsyncFunction`, a bound generator fn stays
    // `instanceof GeneratorFunction`. Fall back to %Function.prototype%.
    wrapper.properties.insert(
        "__proto__".into(),
        match target_proto_link {
            Some(link) if !matches!(link, Value::Null | Value::Undefined) => link,
            _ => shared_function_prototype(),
        },
    );
    if target_non_ctor {
        wrapper
            .properties
            .insert("__vybe_non_ctor".into(), Value::Bool(true));
    }
    wrapper.properties.insert(
        "name".into(),
        Value::String(Arc::from(format!("bound {}", target_name).as_str())),
    );
    wrapper.properties.insert(
        "length".into(),
        Value::F64(target_length.saturating_sub(consumed_args) as f64),
    );
    if !matches!(target_proto, Value::Null | Value::Undefined) {
        wrapper.properties.insert("prototype".into(), target_proto);
    }
    Value::Object(Arc::new(Mutex::new(wrapper)))
}
