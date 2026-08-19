//! Function-call protocol, method dispatch, and property resolution.
//!
//! - `call_value` / `call_function` — the function-call entry points
//!   used by `CALL`, `CALL_REF`, `CALL_IMPORT` opcode handlers.
//! - `try_dunder_binary` — Python-style `__add__` / `__eq__` fallback
//!   during dynamic binary ops.
//! - `resolve_property` / `method_to_value` — WASM GC-style method
//!   resolution: getter → instance → vtable → Object universals.

use crate::error::VMError;
use crate::value::{Function, Object, ObjectKind, TypedArrayState, TypedElemKind, Value};
use crate::vm::{CallFrame, MAX_FRAMES, VM};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

/// True when a continuation's entry Function points at an async chunk —
/// selects the promise-wrapping `next` driver in the protocol attach.
pub(crate) fn continuation_entry_is_async(chunks: &[crate::chunk::Chunk], entry: &Value) -> bool {
    if let Value::Object(obj) = entry {
        if let ObjectKind::Function(f) = &obj.lock().unwrap().kind {
            return chunks
                .get(f.chunk_index)
                .map(|c| c.is_async)
                .unwrap_or(false);
        }
    }
    false
}

pub(crate) fn attach_continuation_protocols(
    // Insertion-ordered — see `Object::properties`.
    //
    // Globals are now a `Vec<Value>` indexed by globalidx (WASM's model), so
    // this takes the three values it needs already RESOLVED rather than a map
    // to look them up in. It only ever asked three name questions.
    properties: &mut indexmap::IndexMap<String, Value>,
    resolved: ContinuationGlobals,
    is_async: bool,
) {
    // §27.6.1.2: an ASYNC generator's `next()` returns a promise-wrapped
    // IteratorResult — wire the async driver when the entry chunk is
    // async (falls back to the sync driver if the async stdlib chunk
    // wasn't bundled).
    let next = if is_async && resolved.async_next.is_some() {
        // Stamp the object so the compiler's inline `.next()` fast path
        // can defer to this promise-returning driver instead.
        properties.insert("__vybe_async_gen".into(), Value::Bool(true));
        resolved.async_next
    } else {
        resolved.sync_next
    };
    if let Some(next) = next {
        properties.insert("next".into(), next);
    }
    if let Some(iter) = resolved.generator_self {
        properties.insert("iterator".into(), iter.clone());
        properties.insert("asyncIterator".into(), iter);
    }
}

/// The three generator-protocol globals `attach_continuation_protocols` needs,
/// resolved by name BEFORE the call — the only three name questions it asked.
pub(crate) struct ContinuationGlobals {
    pub async_next: Option<Value>,
    pub sync_next: Option<Value>,
    pub generator_self: Option<Value>,
}

impl crate::vm::VM {
    pub(crate) fn continuation_globals(&self) -> ContinuationGlobals {
        ContinuationGlobals {
            async_next: self.global("__vybe_async_generator_next").cloned(),
            sync_next: self.global("__vybe_generator_next").cloned(),
            generator_self: self.global("__vybe_generator_self").cloned(),
        }
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
            TypedElemKind::BigI64 | TypedElemKind::BigU64 => Value::bigint_i64(0),
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
        // §10.4.5: BigInt64/BigUint64 elements ARE BigInts — the elem
        // stamp picks the signed/unsigned reading of the same 64 bits
        // (js-types JS-API: ToBigInt64 / ToBigUint64).
        TypedElemKind::BigI64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::bigint_i64(i64::from_le_bytes(bytes))
        }
        TypedElemKind::BigU64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::bigint_u64(u64::from_le_bytes(bytes))
        }
    }
}

/// Build a `RangeError: Maximum call stack size exceeded` value with the
/// canonical exception shape (`name` / `__type` / `__exception_type` /
/// `message`) so it is catchable by JS `try/catch` and matches
/// `e instanceof RangeError` (ECMA-262 §6.2.3 — stack overflow is a
/// `RangeError`, not an uncatchable host trap).
fn make_stack_overflow_error() -> Value {
    let mut obj = Object::new();
    let name = Value::String(Arc::from("RangeError"));
    obj.properties.insert("name".into(), name.clone());
    obj.properties.insert("__type".into(), name.clone());
    obj.properties.insert("__exception_type".into(), name);
    obj.properties.insert(
        "message".into(),
        Value::String(Arc::from("Maximum call stack size exceeded")),
    );
    Value::Object(crate::heap::alloc(obj))
}

/// The exception object a TRAP surfaces as once it crosses into host code.
/// `WebAssembly.RuntimeError` is what the WebAssembly JS Interface names for
/// exactly this, and the shape mirrors `make_stack_overflow_error` above —
/// which already surfaces a VM-level condition as a catchable ECMA error.
fn make_runtime_error(message: &str) -> Value {
    let mut obj = Object::new();
    let name = Value::String(Arc::from("RuntimeError"));
    obj.properties.insert("name".into(), name.clone());
    obj.properties.insert("__type".into(), name.clone());
    obj.properties.insert("__exception_type".into(), name);
    obj.properties
        .insert("message".into(), Value::String(Arc::from(message)));
    Value::Object(crate::heap::alloc(obj))
}

impl VM {
    /// Legacy raise — every value-shaped throw (host `throw_value`, RETHROW,
    /// VM-internal errors) is a `throw` of the host `vybe:exception` tag
    /// (entity 0) with the value as its 1-ary payload.
    pub(crate) fn raise_exception_value(&mut self, val: Value) -> Result<(), VMError> {
        self.raise_exception(0, vec![val], 0)
    }

    pub(crate) fn raise_exception_value_skipping(
        &mut self,
        val: Value,
        skip_handlers: usize,
    ) -> Result<(), VMError> {
        self.raise_exception(0, vec![val], skip_handlers)
    }

    /// Spec EH throw: find the innermost matching catch clause by TAG
    /// IDENTITY (exception-handling proposal — "catch clauses use a tag to
    /// identify the thrown exception"; the payload is NEVER inspected),
    /// unwind to it, and deliver per the clause kind:
    ///   catch          → payload values
    ///   catch_ref      → payload values, exnref
    ///   catch_all      → nothing
    ///   catch_all_ref  → exnref
    /// No clause anywhere → the exception escapes as a runtime error.
    pub(crate) fn raise_exception(
        &mut self,
        tag_entity: usize,
        payload: Vec<Value>,
        skip_handlers: usize,
    ) -> Result<(), VMError> {
        self.raise_exception_inner(tag_entity, payload, skip_handlers, false)
    }

    /// As [`raise_exception`], but for a TRAP — which two specs together say
    /// only HOST-level code may catch:
    ///   * WASM 3.0 core: a trap is not catchable by `try_table`. Not by
    ///     `catch $tag`, and NOT by `catch_all` either.
    ///   * WebAssembly JS Interface: a trap reaching host code surfaces there
    ///     as a catchable `WebAssembly.RuntimeError`.
    /// Composed: a trap passes through every `try_table` clause and stops at
    /// the first host-level handler. Nothing distinguishes the two layers by
    /// opcode — both compile to `TRY_TABLE` — but the TAG already does: a
    /// host `try/catch` is `emit_try_start`'s single `vybe:exception`
    /// (entity 0) `catch` clause, while a wast `try_table` carries a
    /// `wast:tag:*` entity or a `catch_all` kind.
    pub(crate) fn raise_trap(&mut self, message: &str) -> Result<(), VMError> {
        let err = make_runtime_error(message);
        self.raise_exception_inner(0, vec![err], 0, true)
    }

    fn raise_exception_inner(
        &mut self,
        tag_entity: usize,
        payload: Vec<Value>,
        skip_handlers: usize,
        is_trap: bool,
    ) -> Result<(), VMError> {
        use crate::vm::{
            CATCH_KIND_CATCH, CATCH_KIND_CATCH_ALL, CATCH_KIND_CATCH_ALL_REF, CATCH_KIND_CATCH_REF,
        };
        let mut matched_idx = None;
        let search_len = self.exception_handlers.len().saturating_sub(skip_handlers);
        for i in (0..search_len).rev() {
            let handler = &self.exception_handlers[i];
            let matches = if is_trap {
                // HOST-level clauses only — see `raise_trap`. `catch_all` is a
                // `try_table` clause and must let the trap through.
                matches!(handler.kind, CATCH_KIND_CATCH | CATCH_KIND_CATCH_REF)
                    && handler.tag_entity == 0
            } else {
                match handler.kind {
                    CATCH_KIND_CATCH_ALL | CATCH_KIND_CATCH_ALL_REF => true,
                    CATCH_KIND_CATCH | CATCH_KIND_CATCH_REF => handler.tag_entity == tag_entity,
                    _ => false,
                }
            };
            if matches {
                matched_idx = Some(i);
                break;
            }
        }

        // The user-facing escape value: for the language exception tag the
        // payload IS the exception object; foreign tags surface as exnref.
        let escape_value = |payload: &[Value], entity: usize| -> Value {
            if entity == 0 {
                payload.first().cloned().unwrap_or(Value::Null)
            } else {
                Self::pack_exnref(entity, payload.to_vec())
            }
        };

        if let Some(idx) = matched_idx {
            let handler = self.exception_handlers[idx].clone();
            // The handler's frame sits BELOW the innermost dispatch loop's
            // floor: unwinding here would leave this nested loop executing
            // an OUTER loop's frames (fatal "no frame" when the outer loop
            // resumes on an empty stack). Defer instead — last_exception
            // propagates through the host-call chain and re-raises at the
            // outer loop's host-call site, which CAN unwind cleanly.
            if let Some(&floor) = self.exec_floors.last() {
                if handler.frame_depth < floor {
                    let val = escape_value(&payload, tag_entity);
                    self.last_exception = Some(val.clone());
                    let stack = self.capture_call_stack();
                    return Err(VMError::new(format!("{}", val)).with_stack(stack));
                }
            }
            // Remove the matched clause AND its sibling clauses (same
            // try_table group), plus everything nested above them.
            let group = handler.group;
            let mut group_start = idx;
            while group_start > 0 && self.exception_handlers[group_start - 1].group == group {
                group_start -= 1;
            }
            self.exception_handlers.truncate(group_start);
            while self.frames.len() > handler.frame_depth {
                let base = self.frames.last().unwrap().base;
                self.close_upvalues(base);
                self.frames.pop();
            }
            // Unwind the structured-control label stack to the try's level:
            // a throw skips the `end`s of any nested block/loop/if inside the
            // try body, so those entries must be dropped or later `br`s in the
            // handler's frame (e.g. a loop re-entering this try) mis-target.
            self.label_stack.truncate(handler.label_depth);
            self.stack.truncate(handler.stack_depth);
            match handler.kind {
                CATCH_KIND_CATCH => {
                    for v in payload {
                        self.push(v)?;
                    }
                }
                CATCH_KIND_CATCH_REF => {
                    let exn = Self::pack_exnref(tag_entity, payload.clone());
                    for v in payload {
                        self.push(v)?;
                    }
                    self.push(exn)?;
                }
                CATCH_KIND_CATCH_ALL => {} // spec: no values pushed
                CATCH_KIND_CATCH_ALL_REF => {
                    self.push(Self::pack_exnref(tag_entity, payload))?;
                }
                _ => {}
            }
            // Spec: a matching clause BRANCHES to its `labelidx`, carrying the
            // values just pushed as that label's results. Resolve the depth
            // through the SAME helper `br`/`br_if` use, so the two can never
            // drift — and so the handler target is block structure rather than
            // a byte offset that truncates past a 64KB try body.
            let depth = handler.catch_label as usize;
            match self.label_stack.iter().rev().nth(depth).copied() {
                Some(entry) => {
                    self.branch_to_label(depth, entry);
                    Ok(())
                }
                None => Err(VMError::new(format!(
                    "try_table catch label {depth} out of range: {} label(s) in scope",
                    self.label_stack.len()
                ))),
            }
        } else if let Some(ac) = self.active_continuations.pop() {
            // Stack-switching proposal: an exception not handled inside a
            // resumed continuation propagates to the PARENT at the `resume`
            // site — the continuation completes exceptionally. Mark it Done,
            // restore the caller fiber (saved by RESUME/GEN_NEXT), and
            // re-raise there so the caller's try/catch handlers fire.
            if let Value::Object(ref obj) = ac.cont {
                let o = obj.lock().unwrap();
                if let ObjectKind::Continuation(cs) = &o.kind {
                    *cs.state.lock().unwrap() = crate::value::ContinuationPhase::Done;
                }
            }
            self.resume_fiber_with(ac.caller_fiber, None)?;
            self.raise_exception(tag_entity, payload, 0)
        } else {
            let val = escape_value(&payload, tag_entity);
            self.last_exception = Some(val.clone());
            let stack = self.capture_call_stack();
            let msg = if tag_entity == 0 {
                format!("{}", val)
            } else {
                let tag_name = self
                    .tag_entities
                    .get(tag_entity)
                    .map(|t| t.debug_name.as_str())
                    .unwrap_or("?");
                format!("uncaught exception (tag {tag_entity} '{tag_name}')")
            };
            Err(VMError::new(msg).with_stack(stack))
        }
    }

    /// Internal exnref representation: an opaque object carrying the tag
    /// entity + payload so `throw_ref` can rethrow the EXACT exception.
    /// (Spec exnref is an opaque reference type; its internal shape is
    /// engine-private, like our fixed-width operand encoding.)
    pub(crate) fn pack_exnref(tag_entity: usize, payload: Vec<Value>) -> Value {
        let mut obj = crate::value::Object::new();
        obj.properties
            .insert("__exnref_tag".into(), Value::I32(tag_entity as i32));
        obj.properties.insert(
            "__exnref_payload".into(),
            Value::Object(std::sync::Arc::new(std::sync::Mutex::new(
                crate::value::Object::new_array(payload),
            ))),
        );
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)))
    }

    /// Reverse of `pack_exnref`. `None` when the value is not an exnref.
    pub(crate) fn unpack_exnref(val: &Value) -> Option<(usize, Vec<Value>)> {
        let Value::Object(obj) = val else { return None };
        let o = obj.lock().unwrap();
        let Some(Value::I32(tag)) = o.properties.get("__exnref_tag") else {
            return None;
        };
        let Some(Value::Object(arr)) = o.properties.get("__exnref_payload") else {
            return None;
        };
        let a = arr.lock().unwrap();
        let ObjectKind::Array(items) = &a.kind else {
            return None;
        };
        Some((*tag as usize, items.clone()))
    }

    #[allow(dead_code)]
    pub(crate) fn try_dunder_binary(
        &mut self,
        obj: &Arc<Mutex<crate::value::Object>>,
        arg: &Value,
        dunder: &str,
    ) -> Option<Value> {
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
                                } else {
                                    Vec::new()
                                }
                            }
                            _ => Vec::new(),
                        };
                        drop(o);
                        let mut args: Vec<Value> = Vec::with_capacity(bound.len() + argc);
                        args.extend(bound);
                        args.extend(self.stack[self.stack.len() - argc..].iter().cloned());
                        for _ in 0..argc {
                            self.stack.pop();
                        }
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
                        for _ in 0..argc {
                            self.stack.pop();
                        }
                        self.stack.pop();

                        let mut last = Value::Null;
                        for handler in handlers {
                            self.push(handler)?;
                            for arg in &args {
                                self.push(arg.clone())?;
                            }
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
                        } else {
                            "?".into()
                        };
                        return Err(VMError::new(format!(
                            "Not a function in chunk '{}' (kind: {})",
                            chunk_name, kind_name
                        )));
                    }
                }
            }
            _ => {
                let stack = self.capture_call_stack();
                return Err(VMError::new(format!(
                    "{} is not callable (type: {})",
                    callee.type_tag(),
                    callee
                ))
                .with_stack(stack));
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
    pub(crate) fn call_function_direct(
        &mut self,
        func: &Function,
        argc: usize,
    ) -> Result<(), VMError> {
        self.call_function_inner(func, argc, true)
    }

    fn call_function_inner(
        &mut self,
        func: &Function,
        argc: usize,
        bypass_generator: bool,
    ) -> Result<(), VMError> {
        if self.frames.len() >= MAX_FRAMES {
            // Catchable RangeError rather than an uncatchable host trap — JS
            // recursion guards (`try { recurse() } catch { … }`) depend on it.
            return self.raise_exception_value(make_stack_overflow_error());
        }

        let chunk_index = func.chunk_index;
        // JSPI promising boundary: calling an async function is delimited at
        // this call. The body runs inline until it returns (result Promise on
        // the stack) or suspends at an `await`, in which case only the async
        // frames are captured, the caller receives a pending Promise and
        // KEEPS RUNNING — resumption comes off the ready queue.
        if !bypass_generator
            && self.chunks[chunk_index].is_async
            && !self.chunks[chunk_index].is_generator
        {
            // Async GENERATORS fall through to the continuation branch —
            // calling one builds the generator object; the async surface
            // is its promise-returning `next()` (§27.6.1.2).
            let func = func.clone();
            return self.call_async(&func, argc);
        }
        // Generator intercept: if the target chunk is flagged as a
        // generator, calling it doesn't enter the body — we build a
        // `Continuation` bound to a reified Function value and the
        // passed-through args, push it on the stack, and return. The
        // caller drives the generator by RESUMEing the continuation.
        if !bypass_generator && self.chunks[chunk_index].is_generator {
            use crate::value::{ContinuationPhase, ContinuationState};
            // Collect args — they'll be bound into the continuation.
            let mut args: Vec<Value> = Vec::with_capacity(argc);
            for _ in 0..argc {
                args.push(self.pop());
            }
            args.reverse();
            // Re-wrap the Function value so entry can re-call it later.
            let fn_obj = Object {
                properties: indexmap::IndexMap::new(),
                kind: ObjectKind::Function(func.clone()),
                type_id: 0,
                fields: Vec::new(),
            };
            let entry = Value::Object(crate::heap::alloc(fn_obj));
            let state = ContinuationState {
                entry,
                saved: std::sync::Mutex::new(None),
                state: std::sync::Mutex::new(ContinuationPhase::Ready),
            };
            let mut cont = Object {
                properties: indexmap::IndexMap::new(),
                kind: ObjectKind::Continuation(state),
                type_id: 0,
                fields: Vec::new(),
            };
            let entry_is_async = self.chunks[chunk_index].is_async;
            let cg = self.continuation_globals();
            attach_continuation_protocols(&mut cont.properties, cg, entry_is_async);
            if !args.is_empty() {
                // Stash bound args so the first RESUME can re-push them.
                let bound = Object {
                    properties: indexmap::IndexMap::new(),
                    kind: ObjectKind::Array(args),
                    type_id: 0,
                    fields: Vec::new(),
                };
                cont.properties.insert(
                    "__bound_args".into(),
                    Value::Object(crate::heap::alloc(bound)),
                );
            }
            self.push(Value::Object(crate::heap::alloc(cont)))?;
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
        if argc > arity {
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
        for _ in have..total {
            self.push(Value::Null)?;
        }

        let capture_count = self.chunks[chunk_index].capture_count as usize;
        let capture_base = self.chunks[chunk_index].capture_base as usize;
        for (i, uv) in func.upvalues.iter().enumerate().take(capture_count) {
            let val = match &uv.lock().unwrap().location {
                // Lazy-locals convention: LOCAL_SET grows the stack on
                // demand, so a captured slot that was never written may
                // lie beyond the current stack — it reads as Null, same
                // as a direct LOCAL_GET of an untouched local.
                crate::value::UpvalueLocation::Open(si) => {
                    self.stack.get(*si).cloned().unwrap_or(Value::Null)
                }
                crate::value::UpvalueLocation::Closed(v) => v.clone(),
            };
            let slot = base + capture_base + i;
            if slot < self.stack.len() {
                self.stack[slot] = val;
            }
        }

        let upvalues = func.upvalues.clone();
        self.frames.push(CallFrame {
            chunk_index,
            ip: 0,
            base,
            label_base: self.label_stack.len(),
            upvalues,
        });
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

                // 1b. Typed fields from TypeDef (for typed objects like Error with field layout)
                let type_id = ob.type_id;
                if type_id > 0 {
                    if let Some(td) = self.type_registry.get(type_id) {
                        if let Some(field_idx) = td.field_index(name) {
                            if let Some(v) = ob.fields.get(field_idx) {
                                return Ok(v.clone());
                            }
                        }
                    }
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
                            let byte_length =
                                typed_array_live_length(ta) * ta.elem.bytes_per_element();
                            return Ok(Value::I32(byte_length as i32));
                        }
                        "BYTES_PER_ELEMENT" => {
                            return Ok(Value::I32(ta.elem.bytes_per_element() as i32));
                        }
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
                let inferred_type = ob
                    .properties
                    .get("__type")
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

                // 3. Universal Object methods (type 0).
                // §20.1.2.2: a bare object (`Object.create(null)`, marked
                // by an explicit `__proto__: Null`) has NO reachable
                // %Object.prototype% methods — skip the universal vtable.
                let bare = {
                    let ob = o.lock().unwrap();
                    matches!(ob.properties.get("__proto__"), Some(Value::Null))
                };
                if !bare {
                    if let Some(method) = self.type_registry.resolve_method(0, name) {
                        return Ok(self.method_to_value(method));
                    }
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
                        wrapped
                            .properties
                            .insert(RECEIVER_MARKER.into(), Value::Bool(true));
                        Value::Object(crate::heap::alloc(wrapped))
                    } else {
                        self.func_table[*idx].clone()
                    }
                } else {
                    // Fallback: create new (shouldn't happen if registered properly)
                    let mut obj = Object::new();
                    obj.kind = ObjectKind::HostFunction(*idx);
                    obj.properties
                        .insert(RECEIVER_MARKER.into(), Value::Bool(true));
                    Value::Object(crate::heap::alloc(obj))
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
                let mut properties = indexmap::IndexMap::new();
                properties.insert(RECEIVER_MARKER.into(), Value::Bool(true));
                let obj = Object {
                    properties,
                    kind: ObjectKind::Function(func),
                    type_id: 0,
                    fields: Vec::new(),
                };
                Value::Object(crate::heap::alloc(obj))
            }
        }
    }
}
