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

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // Function.prototype.bind(this_fn, thisArg, ...boundArgs) → new Function
    //
    // The returned function ref carries `__bound_args = [thisArg, ...boundArgs]`
    // and points at the same host fn idx as the receiver (or the same
    // chunk for user functions).
    vm.register_host_fn("ecma:function", "bind", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let target = args.first().cloned().unwrap_or(Value::Undefined);
        let bound: Vec<Value> = if args.len() > 1 { args[1..].to_vec() } else { Vec::new() };
        bind_function(&target, bound)
    }));

    // Function.prototype.call(this_fn, thisArg, ...args) → result
    //
    // Synchronously invokes the receiver with the given thisArg + args.
    vm.register_host_fn("ecma:function", "call", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let target = args.first().cloned().unwrap_or(Value::Undefined);
        // ECMA semantics pass thisArg explicitly; Vybe's calling convention
        // passes `this` as args[0] of the host fn callee, so we forward
        // the rest of the args (which already starts with thisArg) verbatim.
        let invoke_args: &[Value] = if args.len() > 1 { &args[1..] } else { &[] };
        ctx.invoke(&target, invoke_args)
    }));

    // Function.prototype.apply(this_fn, thisArg, argsArray) → result
    vm.register_host_fn("ecma:function", "apply", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let target = args.first().cloned().unwrap_or(Value::Undefined);
        let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
        let mut invoke_args: Vec<Value> = vec![this_arg];
        if let Some(Value::Object(arr)) = args.get(2) {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                invoke_args.extend(v.iter().cloned());
            }
        }
        ctx.invoke(&target, &invoke_args)
    }));
}

/// Build a function ref carrying bound args. Mirrors the convention in
/// `crate::namespaces::bound_host_fn_ref` but works on any function-like
/// Value (HostFunction or user Function).
fn bind_function(target: &Value, bound: Vec<Value>) -> Value {
    if let Value::Object(obj) = target {
        let (existing_kind, existing_bound) = {
            let o = obj.lock().unwrap();
            let prev_bound = match o.properties.get("__bound_args") {
                Some(Value::Object(ba)) => {
                    let bo = ba.lock().unwrap();
                    if let ObjectKind::Array(ref v) = bo.kind { v.clone() } else { Vec::new() }
                }
                _ => Vec::new(),
            };
            (o.kind.clone(), prev_bound)
        };
        let mut combined = existing_bound;
        combined.extend(bound);
        let mut new_obj = Object::new();
        new_obj.kind = existing_kind;
        new_obj.properties.insert("__bound_args".into(), Value::Object(
            Arc::new(Mutex::new(Object::new_array(combined)))
        ));
        return Value::Object(Arc::new(Mutex::new(new_obj)));
    }
    target.clone()
}
