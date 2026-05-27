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
use crate::value::{Function, Object, ObjectKind, TypedArrayState, TypedElemKind, Value};
use crate::vm::{
    VM, CallFrame, MAX_FRAMES,
};

pub(crate) fn attach_continuation_protocols(
    properties: &mut HashMap<String, Value>,
    globals: &HashMap<String, Value>,
) {
    if let Some(next) = globals.get("__vybe_generator_next").cloned() {
        properties.insert("next".into(), next);
    }
    if let Some(iter) = globals.get("__vybe_generator_self").cloned() {
        properties.insert("iterator".into(), iter.clone());
        properties.insert("asyncIterator".into(), iter);
    }
}

fn typed_array_live_length(ta: &TypedArrayState) -> usize {
    let buf = ta.buffer.lock().unwrap();
    let bpe = ta.elem.bytes_per_element();
    if ta.byte_offset >= buf.len() {
        return 0;
    }
    let available_bytes = buf.len() - ta.byte_offset;
    let available_elems = available_bytes / bpe;
    ta.length.min(available_elems)
}

fn read_typed_array_element(ta: &TypedArrayState, index: usize) -> Value {
    let bpe = ta.elem.bytes_per_element();
    let buf = ta.buffer.lock().unwrap();
    let abs = ta.byte_offset + index * bpe;
    if abs + bpe > buf.len() {
        return match ta.elem {
            TypedElemKind::F32 | TypedElemKind::F64 => Value::F64(0.0),
            TypedElemKind::BigI64 | TypedElemKind::BigU64 => Value::I64(0),
            _ => Value::I32(0),
        };
    }
    match ta.elem {
        TypedElemKind::I8 => Value::I32(buf[abs] as i8 as i32),
        TypedElemKind::U8 | TypedElemKind::U8Clamped => Value::I32(buf[abs] as i32),
        TypedElemKind::I16 => Value::I32(i16::from_le_bytes([buf[abs], buf[abs + 1]]) as i32),
        TypedElemKind::U16 => Value::I32(u16::from_le_bytes([buf[abs], buf[abs + 1]]) as i32),
        TypedElemKind::I32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            Value::I32(i32::from_le_bytes(bytes))
        }
        TypedElemKind::U32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            Value::I32(u32::from_le_bytes(bytes) as i32)
        }
        TypedElemKind::F32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            Value::F64(f32::from_le_bytes(bytes) as f64)
        }
        TypedElemKind::F64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::F64(f64::from_le_bytes(bytes))
        }
        TypedElemKind::BigI64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::I64(i64::from_le_bytes(bytes))
        }
        TypedElemKind::BigU64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::I64(u64::from_le_bytes(bytes) as i64)
        }
    }
}

impl VM {
    pub(crate) fn raise_exception_value(&mut self, val: Value) -> Result<(), VMError> {
        let mut matched_idx = None;
        for i in (0..self.exception_handlers.len()).rev() {
            let handler = &self.exception_handlers[i];
            if handler.tag == 0 {
                matched_idx = Some(i);
                break;
            }
            let tag_idx = handler.tag as usize;
            let tag_name = self.chunks.get(0)
                .and_then(|c| c.exception_tags.get(tag_idx))
                .cloned()
                .unwrap_or_default();
            if !tag_name.is_empty() {
                let matches = self.test_type(&val, &tag_name.to_lowercase())
                    || self.exception_value_matches(&val, &tag_name);
                if matches {
                    matched_idx = Some(i);
                    break;
                }
            }
        }

        if let Some(idx) = matched_idx {
            let handler = self.exception_handlers[idx].clone();
            self.exception_handlers.truncate(idx);
            while self.frames.len() > handler.frame_depth {
                let base = self.frames.last().unwrap().base;
                self.close_upvalues(base);
                self.frames.pop();
            }
            self.stack.truncate(handler.stack_depth);
            self.push(val)?;
            let f = self.frame_mut();
            f.ip = handler.catch_ip;
            Ok(())
        } else {
            self.last_exception = Some(val.clone());
            let stack = self.capture_call_stack();
            Err(VMError::new(format!("{}", val)).with_stack(stack))
        }
    }

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
                        let host_fn = self.host_fns[idx].clone();
                        let result = {
                            let mut ctx = self.make_host_context();
                            host_fn(&mut ctx, &args)
                        };
                        if let Some(exc) = self.last_exception.take() {
                            self.raise_exception_value(exc)?;
                            return Ok(());
                        }
                        self.push(result)?;
                    }
                    ObjectKind::Array(elems) => {
                        let handlers = elems.clone();
                        drop(o);

                        let args: Vec<Value> = self.stack[self.stack.len() - argc..].to_vec();
                        for _ in 0..argc { self.stack.pop(); }
                        self.stack.pop();

                        let mut last = Value::Null;
                        for handler in handlers {
                            self.push(handler)?;
                            for arg in &args { self.push(arg.clone())?; }
                            let depth = self.frames.len();
                            self.call_value(args.len())?;
                            // Some handlers (host/callable shims) complete without
                            // pushing a new frame, while bytecode functions do push one.
                            // Only run execute_until when a nested frame exists.
                            if self.frames.len() > depth {
                                last = self.execute_until(depth)?;
                            } else {
                                last = self.pop();
                            }
                        }
                        self.push(last)?;
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
            attach_continuation_protocols(&mut cont.properties, &self.globals);
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

        // Arity validation: pad missing args, truncate extras (dynamic language semantics).
        //
        // Missing positional args land as `Undefined` per ECMA-262 §10.2.1.1
        // (matches V8 / QuickJS internals). Distinct from `Null` so callers
        // can tell `f()` from `f(null)` — required for spec-compliant default
        // parameters. WASM core dispatch is fixed-arity; padding with
        // Undefined is the standard JS-engine convention used by every
        // browser-grade JS-on-WASM implementation.
        if argc > arity && arity > 0 {
            for _ in 0..(argc - arity) {
                self.pop();
            }
        }
        for _ in argc..arity {
            self.push(Value::Undefined)?;
        }

        let local_count = self.chunks[chunk_index].local_count as usize;
        let total = local_count.max(arity);
        let have = self.stack.len() - base;
        // Local slots beyond the arity range are uninitialized variables,
        // not missing args — Null is the right default here.
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
                if let Some(v) = ob.properties.get(name) {
                    return Ok(v.clone());
                }
                if let ObjectKind::Array(ref elems) = ob.kind {
                    if let Ok(idx) = name.parse::<usize>() {
                        if idx < elems.len() {
                            return Ok(elems[idx].clone());
                        }
                    }
                }
                if let ObjectKind::TypedArray(ref ta) = ob.kind {
                    if let Ok(idx) = name.parse::<usize>() {
                        if idx < typed_array_live_length(ta) {
                            return Ok(read_typed_array_element(ta, idx));
                        }
                    }
                    match name {
                        "buffer" => return Ok(Value::Object(ta.buffer_obj.clone())),
                        "length" => return Ok(Value::I32(typed_array_live_length(ta) as i32)),
                        "byteOffset" => return Ok(Value::I32(ta.byte_offset as i32)),
                        "byteLength" => {
                            let byte_length = typed_array_live_length(ta) * ta.elem.bytes_per_element();
                            return Ok(Value::I32(byte_length as i32));
                        }
                        "BYTES_PER_ELEMENT" => return Ok(Value::I32(ta.elem.bytes_per_element() as i32)),
                        _ => {}
                    }
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
                    if let Some(v) = ob.properties.get(&name_lc) {
                        return Ok(v.clone());
                    }
                    if let ObjectKind::Array(ref elems) = ob.kind {
                        if let Ok(idx) = name_lc.parse::<usize>() {
                            if idx < elems.len() {
                                return Ok(elems[idx].clone());
                            }
                        }
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

                Ok(Value::Undefined)
            }
            Value::String(s) => {
                if name == "length" {
                    return Ok(Value::F64(s.encode_utf16().count() as f64));
                }
                if let Some(tid) = self.type_registry.get_id("string") {
                    if let Some(method) = self.type_registry.resolve_method(tid, name) {
                        return Ok(self.method_to_value(method));
                    }
                }
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }
                Ok(Value::Undefined)
            }
            _ => {
                if let Some(method) = self.type_registry.resolve_method(0, name) {
                    return Ok(self.method_to_value(method));
                }
                Ok(Value::Undefined)
            }
        }
    }

    /// Convert a Method (from TypeRegistry) to a callable Value.
    /// Uses the function table for zero-allocation dispatch.
    pub(crate) fn method_to_value(&self, method: &crate::typedef::Method) -> Value {
        const RECEIVER_MARKER: &str = "__vybe_method_receiver";

        match method {
            crate::typedef::Method::HostFn(idx) => {
                if *idx < self.func_table.len() {
                    if let Value::Object(func_obj) = &self.func_table[*idx] {
                        let mut wrapped = Object::new();
                        wrapped.kind = func_obj.lock().unwrap().kind.clone();
                        wrapped.properties.insert(RECEIVER_MARKER.into(), Value::Bool(true));
                        Value::Object(Arc::new(Mutex::new(wrapped)))
                    } else {
                        self.func_table[*idx].clone()
                    }
                } else {
                    // Fallback: create new (shouldn't happen if registered properly)
                    let mut obj = Object::new();
                    obj.kind = ObjectKind::HostFunction(*idx);
                    obj.properties.insert(RECEIVER_MARKER.into(), Value::Bool(true));
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
                let mut properties = HashMap::new();
                properties.insert(RECEIVER_MARKER.into(), Value::Bool(true));
                let obj = Object { properties, kind: ObjectKind::Function(func), type_id: 0, fields: Vec::new() };
                Value::Object(Arc::new(Mutex::new(obj)))
            }
        }
    }
}
