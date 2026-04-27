//! Function-call protocol, method dispatch, and property resolution.
//!
//! - `call_value` / `call_function` — the function-call entry points
//!   used by `CALL`, `CALL_REF`, `CALL_IMPORT` opcode handlers.
//! - `try_dunder_binary` — Python-style `__add__` / `__eq__` fallback
//!   during dynamic binary ops.
//! - `resolve_property` / `method_to_value` — WASM GC-style method
//!   resolution: getter → instance → vtable → Object universals.

use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use crate::error::VMError;
use crate::value::{Function, Object, ObjectKind, Value};
use crate::vm::{
    VM, CallFrame, HostFn, MAX_FRAMES,
};

impl VM {
    pub(crate) fn try_dunder_binary(&mut self, obj: &Arc<Mutex<crate::value::Object>>, arg: &Value, dunder: &str) -> Option<Value> {
        let method = {
            let o = obj.lock().unwrap();
            o.properties.get(dunder).cloned()
        };
        if let Some(func_val) = method {
            // Call dunder(self, arg) — push func, self, arg, call(2)
            let self_val = Value::Object(obj.clone());
            self.push(func_val).ok()?;
            self.push(self_val).ok()?;
            self.push(arg.clone()).ok()?;
            self.call_value(2).ok()?;
            // Execute until the function returns
            self.execute_until(self.frames.len()).ok()?;
            Some(self.pop())
        } else {
            None
        }
    }

    pub(crate) fn call_value(&mut self, argc: usize) -> Result<(), VMError> {
        self.call_value_inner(argc, false)
    }

    /// Like `call_value` but bypasses the generator intercept — used
    /// from RESUME / GEN_NEXT when we genuinely want the generator
    /// body to execute.
    pub(crate) fn call_value_direct(&mut self, argc: usize) -> Result<(), VMError> {
        self.call_value_inner(argc, true)
    }

    fn call_value_inner(&mut self, argc: usize, bypass_generator: bool) -> Result<(), VMError> {
        let callee_idx = self.stack.len() - 1 - argc;
        let callee = self.stack[callee_idx].clone();

        match &callee {
            Value::Object(obj) => {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Function(func) => {
                        let func = func.clone();
                        drop(o);
                        // Remove callee from stack (WASM convention: only args, no callee)
                        self.stack.remove(callee_idx);
                        if bypass_generator {
                            self.call_function_direct(&func, argc)?;
                        } else {
                            self.call_function(&func, argc)?;
                        }
                    }
                    ObjectKind::HostFunction(idx) => {
                        let idx = *idx;
                        // Function.prototype.bind support: if the function-ref
                        // Object carries `__bound_args` (an Array of values)
                        // they are prepended to the runtime args before
                        // invocation. This is the standard ECMA-262 §20.2.3.2
                        // semantics applied to host fns — callers build bound
                        // refs with `bound_host_fn_ref` in vybe_host.
                        let bound: Vec<Value> = match o.properties.get("__bound_args") {
                            Some(Value::Object(arr)) => {
                                let a = arr.lock().unwrap();
                                if let ObjectKind::Array(ref elems) = a.kind {
                                    elems.clone()
                                } else { Vec::new() }
                            }
                            _ => Vec::new(),
                        };
                        drop(o);
                        let mut args: Vec<Value> = Vec::with_capacity(bound.len() + argc);
                        args.extend(bound);
                        args.extend(self.stack[self.stack.len() - argc..].iter().cloned());
                        for _ in 0..argc { self.stack.pop(); }
                        self.stack.pop();
                        let placeholder: HostFn = Arc::new(|_, _| Value::Null);
                        let host_fn = std::mem::replace(&mut self.host_fns[idx], placeholder);
                        let result = {
                            let mut ctx = self.make_host_context();
                            host_fn(&mut ctx, &args)
                        };
                        self.host_fns[idx] = host_fn;
                        self.push(result)?;
                    }
                    _other => {
                        // Check for __call__ dunder (Python callable objects)
                        let call_fn = o.properties.get("__call__").cloned();
                        let kind_name = format!("{:?}", std::mem::discriminant(&o.kind));
                        drop(o);
                        if let Some(func) = call_fn {
                            self.stack[callee_idx] = func;
                            return self.call_value(argc);
                        }
                        let chunk_name = if !self.frames.is_empty() {
                            self.chunks[self.frame().chunk_index].name.clone()
                        } else { "?".into() };
                        return Err(VMError::new(format!("Not a function in chunk '{}' (kind: {})",
                            chunk_name, kind_name)));
                    }
                }
            }
            _ => {
                let stack = self.capture_call_stack();
                return Err(VMError::new(format!("{} is not callable (type: {})", callee.type_tag(), callee)).with_stack(stack));
            }
        }
        Ok(())
    }

    pub(crate) fn call_function(&mut self, func: &Function, argc: usize) -> Result<(), VMError> {
        self.call_function_inner(func, argc, false)
    }

    /// Direct entry-body call that bypasses the `is_generator`
    /// intercept — used from `RESUME` / `GEN_NEXT` when we want the
    /// generator's body to execute (rather than re-wrap as a nested
    /// continuation).
    pub(crate) fn call_function_direct(&mut self, func: &Function, argc: usize) -> Result<(), VMError> {
        self.call_function_inner(func, argc, true)
    }

    fn call_function_inner(&mut self, func: &Function, argc: usize, bypass_generator: bool) -> Result<(), VMError> {
        if self.frames.len() >= MAX_FRAMES {
            return Err(VMError::new("Stack overflow"));
        }

        let chunk_index = func.chunk_index;
        // Generator intercept: if the target chunk is flagged as a
        // generator, calling it doesn't enter the body — we build a
        // `Continuation` bound to a reified Function value and the
        // passed-through args, push it on the stack, and return. The
        // caller drives the generator by RESUMEing the continuation.
        if !bypass_generator && self.chunks[chunk_index].is_generator {
            use crate::value::{ContinuationState, ContinuationPhase};
            // Collect args — they'll be bound into the continuation.
            let mut args: Vec<Value> = Vec::with_capacity(argc);
            for _ in 0..argc { args.push(self.pop()); }
            args.reverse();
            // Re-wrap the Function value so entry can re-call it later.
            let fn_obj = Object {
                properties: HashMap::new(),
                kind: ObjectKind::Function(func.clone()),
                type_id: 0,
                fields: Vec::new(),
            };
            let entry = Value::Object(Arc::new(Mutex::new(fn_obj)));
            let state = ContinuationState {
                entry,
                saved: std::sync::Mutex::new(None),
                state: std::sync::Mutex::new(ContinuationPhase::Ready),
            };
            let mut cont = Object {
                properties: HashMap::new(),
                kind: ObjectKind::Continuation(state),
                type_id: 0,
                fields: Vec::new(),
            };
            if !args.is_empty() {
                // Stash bound args so the first RESUME can re-push them.
                let bound = Object {
                    properties: HashMap::new(),
                    kind: ObjectKind::Array(args),
                    type_id: 0,
                    fields: Vec::new(),
                };
                cont.properties.insert(
                    "__bound_args".into(),
                    Value::Object(Arc::new(Mutex::new(bound))),
                );
            }
            self.push(Value::Object(Arc::new(Mutex::new(cont))))?;
            return Ok(());
        }
        let arity = func.arity as usize;
        // WASM-compliant: slot 0 = first arg (not callee).
        // The caller must remove the callee from the stack before this call.
        let base = self.stack.len() - argc;

        // Arity validation: pad missing args, truncate extras (dynamic language semantics)
        if argc > arity && arity > 0 {
            for _ in 0..(argc - arity) {
                self.pop();
            }
        }
        for _ in argc..arity {
            self.push(Value::Null)?;
        }

        let local_count = self.chunks[chunk_index].local_count as usize;
        let total = local_count.max(arity);
        let have = self.stack.len() - base;
        for _ in have..total {
            self.push(Value::Null)?;
        }

        let upvalues = func.upvalues.clone();
        self.frames.push(CallFrame { chunk_index, ip: 0, base, upvalues });
        Ok(())
    }

    pub fn resolve_property(&self, obj: &Value, name: &str) -> Result<Value, VMError> {
        match obj {
            Value::Object(o) => {
                let ob = o.lock().unwrap();
                // 1. Instance property (getters handled in struct_get opcode directly)
                let val = ob.get(name);
                if !matches!(val, Value::Null) {
                    return Ok(val);
                }
                // 1b. Case-insensitive fallback for case-sensitive
                // languages (C#, Dart) reading PascalCase fields
                // (`btn.Location`) off an object whose setter-backed
                // write stored them lowercased (`location`). The
                // lowercase key is the canonical .NET wrapper storage;
                // falling back finds it without forcing every host
                // write-path to duplicate the value under two keys.
                let name_lc = name.to_lowercase();
                if name_lc != name {
                    let val = ob.get(&name_lc);
                    if !matches!(val, Value::Null) {
                        return Ok(val);
                    }
                }

                // 3. TypeRegistry vtable
                let type_id = ob.type_id;
                drop(ob); // release borrow before accessing self

                if type_id > 0 {
                    if let Some(method) = self.type_registry.resolve_method(type_id, name) {
                        return Ok(self.method_to_value(method));
                    }
                }

                // Also try inferring type from ObjectKind or __type property
                let ob = o.lock().unwrap();
                let inferred_type = ob.properties.get("__type")
                    .map(|v| format!("{}", v).to_lowercase())
                    .unwrap_or_else(|| match &ob.kind {
                        ObjectKind::Array(_) => "list".into(),
                        _ => String::new(),
                    });
                drop(ob);

                if !inferred_type.is_empty() {
                    if let Some(tid) = self.type_registry.get_id(&inferred_type) {
                        if let Some(method) = self.type_registry.resolve_method(tid, name) {
                            return Ok(self.method_to_value(method));
                        }
                    }
                }

                // 3. Universal Object methods (type 0)
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }

                Ok(Value::Null)
            }
            Value::String(s) => {
                if name == "length" {
                    return Ok(Value::F64(s.len() as f64));
                }
                if let Some(tid) = self.type_registry.get_id("string") {
                    if let Some(method) = self.type_registry.resolve_method(tid, name) {
                        return Ok(self.method_to_value(method));
                    }
                }
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }
                Ok(Value::Null)
            }
            _ => {
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }
                Ok(Value::Null)
            }
        }
    }

    /// Convert a Method (from TypeRegistry) to a callable Value.
    /// Uses the function table for zero-allocation dispatch.
    pub(crate) fn method_to_value(&self, method: &crate::typedef::Method) -> Value {
        match method {
            crate::typedef::Method::HostFn(idx) => {
                // Return existing entry from function table — no allocation
                if *idx < self.func_table.len() {
                    self.func_table[*idx].clone()
                } else {
                    // Fallback: create new (shouldn't happen if registered properly)
                    let mut obj = Object::new();
                    obj.kind = ObjectKind::HostFunction(*idx);
                    Value::Object(Arc::new(Mutex::new(obj)))
                }
            }
            crate::typedef::Method::ChunkFn(idx) => {
                let chunk = &self.chunks[*idx];
                let func = Function {
                    name: Some(chunk.name.clone()),
                    arity: chunk.arity,
                    chunk_index: *idx,
                    upvalues: Vec::new(),
                };
                let obj = Object { properties: HashMap::new(), kind: ObjectKind::Function(func), type_id: 0, fields: Vec::new() };
                Value::Object(Arc::new(Mutex::new(obj)))
            }
        }
    }
}
