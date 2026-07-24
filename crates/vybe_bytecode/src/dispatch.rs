//! The opcode dispatch loop — `execute_until`.
//!
//! This is the interpreter core: the giant `match` over every opcode.
//! Weighs ~3000 lines because it covers 300+ opcodes across every WASM
//! proposal (core, GC, SIMD, threads, exception-handling, stack-switching,
//! memory64, component-model) plus Vybe-internal semantic ops.
//!
//! Pulled into its own file so other concerns (function-call protocol,
//! fiber save/restore, SIMD helpers) stay readable without scrolling
//! past every opcode arm.

use crate::error::VMError;
use crate::opcode::{Op, read_leb_u32};
use crate::value::{Function, Object, ObjectKind, TypedArrayState, TypedElemKind, Upvalue, Value};
use crate::vm::{
    ActiveContinuation, BlockTargets, ExceptionHandler, ImportTarget, LabelEntry, ResumeMode, VM,
};
use std::collections::HashMap;

impl VM {
    /// Whether `o` is a WASM GC array (spec trap-on-out-of-bounds) rather than a
    /// dynamic-language array (lenient subscript).
    ///
    /// The distinction is the object's rtt: `array.new $t` stamps the instance
    /// with the registry id of an `(array …)` defined type, so we read that
    /// type's composite kind back — exactly the mechanism `ref.test`/`ref.cast`
    /// use. Dynamic arrays carry `type_id == 0` (`Object`, a `Struct` kind), so
    /// they never match and keep lenient semantics.
    #[inline]
    pub(crate) fn is_gc_array_obj(&self, o: &std::sync::Arc<std::sync::Mutex<Object>>) -> bool {
        let ob = o.lock().unwrap();
        if !matches!(ob.kind, ObjectKind::Array(_)) {
            return false;
        }
        self.type_registry
            .get(ob.type_id)
            .is_some_and(|td| td.is_array())
    }

    /// Resolve an `array.new` type immediate to the instance rtt (registry id).
    ///
    /// The immediate is a 1-based index into the running module's own type
    /// table (`module_type_names`, in `chunk.types` order); `0` means a
    /// dynamic-language array with no GC type (rtt `0` = `Object`). Any named
    /// type resolves through the registry *by name*, so the host's builtin
    /// types — registered ahead of the module's — don't skew the mapping the
    /// way a raw compile-time table position would.
    #[inline]
    /// Materialize the canonical capture-free funcref for function chunk
    /// `func_idx` — the same object `REF_FUNC` produces (interned via
    /// `funcref_cache`), so identity is stable. Used to populate passive element
    /// segments at instantiation.
    pub(crate) fn make_funcref(&mut self, func_idx: usize) -> Value {
        if let Some(cached) = self.funcref_cache.get(&func_idx) {
            return cached.clone();
        }
        let chunk = &self.chunks[func_idx];
        let arity = chunk.arity;
        let name = if chunk.name == "<script>" {
            None
        } else {
            Some(chunk.name.clone())
        };
        let func = crate::value::Function {
            name,
            arity,
            chunk_index: func_idx,
            upvalues: Vec::new(),
        };
        let mut obj = Object {
            properties: std::collections::HashMap::new(),
            kind: ObjectKind::Function(func),
            type_id: 0,
            fields: Vec::new(),
        };
        let table_idx = self.func_table.len();
        obj.properties
            .insert("__table_idx".into(), Value::F64(table_idx as f64));
        let func_val = Value::Object(Arc::new(Mutex::new(obj)));
        self.func_table.push(func_val.clone());
        self.funcref_cache.insert(func_idx, func_val.clone());
        func_val
    }

    pub(crate) fn resolve_gc_array_rtt(&self, type_imm: usize) -> usize {
        if type_imm == 0 {
            return 0;
        }
        self.module_type_names
            .get(type_imm - 1)
            .and_then(|name| self.type_registry.get_id(name))
            .unwrap_or(0)
    }
}
use std::sync::{Arc, Mutex};

fn make_operation_cancelled_error() -> Value {
    let mut obj = Object::new();
    let name = Value::String(Arc::from("OperationCanceledException"));
    obj.properties.insert("name".into(), name.clone());
    obj.properties.insert("__type".into(), name.clone());
    obj.properties.insert("__exception_type".into(), name);
    obj.properties.insert(
        "message".into(),
        Value::String(Arc::from("The operation was canceled.")),
    );
    Value::Object(Arc::new(Mutex::new(obj)))
}

// ── Block table ──────────────────────────────────────────────────────────────

/// Scan `code` once and build a map from every BLOCK/LOOP/IF/ELSE opcode
/// position (first byte) to its jump targets.
///
/// Format (WASM-compliant): every BLOCK/LOOP/IF carries exactly 1 blocktype
/// byte. ELSE and END carry no operands.
pub(crate) fn build_block_table(code: &[u8]) -> HashMap<usize, BlockTargets> {
    let mut table: HashMap<usize, BlockTargets> = HashMap::new();
    // Stack of opcode_starts for open BLOCK/LOOP/IF scopes.
    let mut stack: Vec<usize> = Vec::new();
    // Maps IF opcode_start → ELSE opcode_start (populated when ELSE is seen).
    let mut else_of: HashMap<usize, usize> = HashMap::new();

    let mut ip = 0usize;
    while ip + 3 < code.len() {
        let opcode_start = ip;
        let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
        let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
        let op = match Op::decode(group as u16, sub as u16) {
            Some(op) => op,
            None => {
                ip += 4;
                continue;
            }
        };
        ip += 4;

        if op == Op::BLOCK || op == Op::LOOP || op == Op::IF {
            ip += 1; // skip result_count byte
            stack.push(opcode_start);
        } else if op == Op::TRY_TABLE {
            // Spec `try_table` IS a block instruction: it opens a
            // handler-scoped block closed by a matching `end`. Skip its
            // variable immediate, then treat it as a nesting level so the
            // block table pairs it with its `end` (whose `is_try` label pop
            // removes the exception handler — replacing the old TRY_END).
            ip += op.operand_format().size_in(code, ip);
            stack.push(opcode_start);
        } else if op == Op::ELSE {
            // Associate ELSE with the top-of-stack IF entry.
            if let Some(&if_start) = stack.last() {
                table
                    .entry(if_start)
                    .or_insert(BlockTargets {
                        else_ip: None,
                        end_ip: 0,
                    })
                    .else_ip = Some(opcode_start);
                else_of.insert(if_start, opcode_start);
            }
            // No stack push — ELSE does not add a nesting level.
        } else if op == Op::END {
            if let Some(entry_start) = stack.pop() {
                // end_ip = ip PAST the END opcode (2-byte internal encoding).
                // BR always jumps here, bypassing END. END only fires via
                // sequential execution, ensuring the label is popped exactly once.
                let end_ip = opcode_start + 4;
                table
                    .entry(entry_start)
                    .or_insert(BlockTargets {
                        else_ip: None,
                        end_ip: 0,
                    })
                    .end_ip = end_ip;
                // ELSE also needs the same end_ip so it can jump past END.
                if let Some(&else_start) = else_of.get(&entry_start) {
                    table
                        .entry(else_start)
                        .or_insert(BlockTargets {
                            else_ip: None,
                            end_ip: 0,
                        })
                        .end_ip = end_ip;
                }
            }
        } else {
            ip += op.operand_format().size_in(code, ip);
        }
    }
    table
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

#[inline]
fn wasm_bool(value: bool) -> Value {
    Value::I32(if value { 1 } else { 0 })
}

/// The stringref "WTF-8 position treatment": a byte offset past the end clamps
/// to the length; an offset that lands inside a multi-byte codepoint is advanced
/// forward to the next codepoint boundary (or the end). Used by the
/// `stringview_wtf8.*` cursor ops.
fn wtf8_treat(s: &str, pos: usize) -> usize {
    if pos >= s.len() {
        return s.len();
    }
    let mut p = pos;
    while p < s.len() && !s.is_char_boundary(p) {
        p += 1;
    }
    p
}

fn typed_array_read(ta: &TypedArrayState, idx: usize) -> Option<Value> {
    if idx >= typed_array_live_length(ta) {
        return None;
    }
    let bpe = ta.elem.bytes_per_element();
    let buf = ta.buffer.lock().unwrap();
    let abs = ta.byte_offset + idx * bpe;
    Some(match ta.elem {
        TypedElemKind::I8 => Value::I32(buf[abs] as i8 as i32),
        TypedElemKind::U8 | TypedElemKind::U8Clamped => Value::I32(buf[abs] as i32),
        TypedElemKind::I16 => {
            let bytes = [buf[abs], buf[abs + 1]];
            Value::I32(i16::from_le_bytes(bytes) as i32)
        }
        TypedElemKind::U16 => {
            let bytes = [buf[abs], buf[abs + 1]];
            Value::I32(u16::from_le_bytes(bytes) as i32)
        }
        TypedElemKind::I32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            Value::I32(i32::from_le_bytes(bytes))
        }
        TypedElemKind::U32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            // Uint32 spans [0, 2^32) — beyond i32 — so surface as an F64
            // JS number (e.g. 4294967295), not a wrapped i32.
            Value::F64(u32::from_le_bytes(bytes) as f64)
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
        // stamp selects the signed/unsigned reading of the 64 bits
        // (js-types JS-API ToBigInt64 / ToBigUint64).
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
    })
}

fn read_le<const N: usize>(bytes: &[u8]) -> [u8; N] {
    bytes.try_into().unwrap_or([0; N])
}

fn read_leb_u64(code: &[u8], ip: &mut usize) -> u64 {
    let mut result = 0u64;
    let mut shift = 0;
    while *ip < code.len() {
        let byte = code[*ip];
        *ip += 1;
        result |= ((byte & 0x7f) as u64) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    result
}

/// The element storage byte-width and decode `kind` of a GC array element
/// storage type name. `kind` feeds `decode_le_numeric`: 0=i32, 1=i64, 2=f32,
/// 3=f64, 4=i8 (packed), 5=i16 (packed). Ref element types return None.
fn array_elem_storage_kind(name: &str) -> Option<(usize, u8)> {
    Some(match name {
        "i8" => (1, 4),
        "i16" => (2, 5),
        "i32" => (4, 0),
        "i64" => (8, 1),
        "f32" => (4, 2),
        "f64" => (8, 3),
        _ => return None,
    })
}

/// Decode `bytes` (little-endian, exactly the element width) into the numeric
/// Value for an untyped GC array element. `kind`: 0=i32, 1=i64, 2=f32, 3=f64,
/// 4=i8, 5=i16 (packed lanes are read as their raw unsigned storage; a later
/// `array.get_s`/`get_u` performs the sign/zero extension).
fn decode_le_numeric(kind: u8, bytes: &[u8]) -> Value {
    match kind {
        1 => Value::I64(i64::from_le_bytes(bytes.try_into().unwrap_or([0; 8]))),
        2 => Value::F32(f32::from_le_bytes(bytes.try_into().unwrap_or([0; 4]))),
        3 => Value::F64(f64::from_le_bytes(bytes.try_into().unwrap_or([0; 8]))),
        4 => Value::I32(bytes.first().map(|b| *b as i32).unwrap_or(0)),
        5 => Value::I32(u16::from_le_bytes(bytes.try_into().unwrap_or([0; 2])) as i32),
        _ => Value::I32(i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4]))),
    }
}

/// Decode `bytes` (little-endian, exactly `kind.bytes_per_element()`) into the
/// numeric Value a `TypedArray` of the given element kind stores.
fn decode_typed_le(kind: crate::value::TypedElemKind, bytes: &[u8]) -> Value {
    use crate::value::TypedElemKind::*;
    match kind {
        I8 => Value::I32(bytes.first().map(|b| (*b as i8) as i32).unwrap_or(0)),
        U8 | U8Clamped => Value::I32(bytes.first().map(|b| *b as i32).unwrap_or(0)),
        I16 => Value::I32(i16::from_le_bytes(bytes.try_into().unwrap_or([0; 2])) as i32),
        U16 => Value::I32(u16::from_le_bytes(bytes.try_into().unwrap_or([0; 2])) as i32),
        I32 | U32 => Value::I32(i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4]))),
        F32 => Value::F32(f32::from_le_bytes(bytes.try_into().unwrap_or([0; 4]))),
        F64 => Value::F64(f64::from_le_bytes(bytes.try_into().unwrap_or([0; 8]))),
        BigI64 | BigU64 => Value::I64(i64::from_le_bytes(bytes.try_into().unwrap_or([0; 8]))),
    }
}

fn typed_array_write(ta: &TypedArrayState, idx: usize, value: &Value) -> bool {
    if idx >= typed_array_live_length(ta) {
        return false;
    }
    let bpe = ta.elem.bytes_per_element();
    let mut buf = ta.buffer.lock().unwrap();
    let abs = ta.byte_offset + idx * bpe;
    match ta.elem {
        TypedElemKind::I8 => {
            buf[abs] = (value.as_i32() as i8) as u8;
        }
        TypedElemKind::U8 => {
            buf[abs] = (value.as_i32() & 0xFF) as u8;
        }
        TypedElemKind::U8Clamped => {
            let n = value.as_f64();
            let clamped = if n.is_nan() {
                0
            } else {
                n.clamp(0.0, 255.0).round() as i32
            };
            buf[abs] = clamped as u8;
        }
        TypedElemKind::I16 => {
            let bytes = (value.as_i32() as i16).to_le_bytes();
            buf[abs..abs + 2].copy_from_slice(&bytes);
        }
        TypedElemKind::U16 => {
            let bytes = ((value.as_i32() & 0xFFFF) as u16).to_le_bytes();
            buf[abs..abs + 2].copy_from_slice(&bytes);
        }
        TypedElemKind::I32 => {
            let bytes = value.as_i32().to_le_bytes();
            buf[abs..abs + 4].copy_from_slice(&bytes);
        }
        TypedElemKind::U32 => {
            // §7.1.7 ToUint32: truncate toward zero, reduce mod 2^32. Using
            // `as_i32()` here saturated large numbers (4294967295 → i32::MAX);
            // ToUint32 must wrap instead.
            let n = value.as_f64();
            let u = if n.is_finite() {
                n.trunc().rem_euclid(4294967296.0) as u32
            } else {
                0
            };
            buf[abs..abs + 4].copy_from_slice(&u.to_le_bytes());
        }
        TypedElemKind::F32 => {
            let bytes = (value.as_f64() as f32).to_le_bytes();
            buf[abs..abs + 4].copy_from_slice(&bytes);
        }
        TypedElemKind::F64 => {
            let bytes = value.as_f64().to_le_bytes();
            buf[abs..abs + 8].copy_from_slice(&bytes);
        }
        TypedElemKind::BigI64 => {
            // §10.4.5 SetValueInBuffer: ToBigInt64 wrap of the value.
            let bits = match value {
                Value::BigInt(n) => n.to_i64_wrapping(),
                Value::I64(n) => *n,
                other => other.as_i32() as i64,
            };
            buf[abs..abs + 8].copy_from_slice(&bits.to_le_bytes());
        }
        TypedElemKind::BigU64 => {
            // ToBigUint64 wrap.
            let bits = match value {
                Value::BigInt(n) => n.to_u64_wrapping(),
                Value::I64(n) => *n as u64,
                other => other.as_i32() as u64,
            };
            buf[abs..abs + 8].copy_from_slice(&bits.to_le_bytes());
        }
    }
    true
}

impl VM {
    /// Ensure the block table for chunk `ci` is computed and cached.
    pub(crate) fn ensure_block_table(&mut self, ci: usize) {
        if !self.block_tables.contains_key(&ci) {
            let table = build_block_table(&self.chunks[ci].code);
            self.block_tables.insert(ci, table);
        }
    }

    fn suspend_for_pending_promise(&mut self, promise_id: u64) -> VMError {
        let fiber = self.save_fiber();
        self.event_loop
            .borrow_mut()
            .suspend_fiber(promise_id, fiber);
        VMError::new(format!("__jspi__:{}", promise_id))
    }

    /// Top-level settled/plain-value await (no promising boundary, not inside
    /// a driven continuation): ECMA-262 §6.2.3.1 still requires one job tick.
    /// Save the whole fiber exactly like a pending top-level await and wake it
    /// immediately off the microtask queue with the value (or rejection).
    fn tick_top_level_await(&mut self, value: Value, is_exception: bool) -> VMError {
        let id = self.event_loop.borrow_mut().next_promise_id();
        let err = self.suspend_for_pending_promise(id);
        let mut el = self.event_loop.borrow_mut();
        let woken = if is_exception {
            el.reject_promise(id, value)
        } else {
            el.resolve_promise(id, value)
        };
        if let Some(fiber) = woken {
            el.microtasks
                .push_back(crate::event_loop::Task::ResumeFiber(fiber));
        }
        err
    }

    /// True when an `await` here should take the one-tick event-queue path
    /// via whole-fiber save: no async-call boundary on THIS fiber. Applies at
    /// top level AND inside RESUME-driven continuations — the driver's state
    /// travels in the fiber's active-continuation chain, and `call_async`
    /// ignores suspensions arriving on foreign fibers, so the whole-save
    /// composes (§6.2.3.1: await always yields one job tick).
    fn top_level_await_ticks(&self) -> bool {
        self.async_floors.is_empty()
    }

    /// `await val` — the JSPI suspend behaviour, reached via a `call` to the
    /// `jspi.await` suspending import (WebAssembly.Suspending; see `emit_await`).
    ///
    /// ECMA-262 §27.5.1.3.2 Await semantics, modelled on stack switching:
    ///   - fulfilled promise → unwrap and push the value (resume the fiber)
    ///   - rejected promise  → throw the rejection reason (walk try/catch)
    ///   - pending promise   → suspend the fiber until the promise settles
    ///   - non-promise value → pass through (Await(v) = Await(Promise.resolve(v)))
    ///
    /// On Ok the dispatch loop continues at the (possibly relocated, for a
    /// caught rejection) ip with the value on the stack. On Err it propagates a
    /// JSPI suspension signal or an unhandled rejection up the stack.
    ///
    /// `await` flattens recursively (ECMA-262 §27.2.1.3.2 Await resolves a
    /// thenable, and the resolution may itself be a thenable): a fulfilled
    /// promise whose value is *another* promise is awaited again. This loop
    /// keeps unwrapping fulfilled-promise chains; it suspends the moment it
    /// reaches a pending link and throws the moment it reaches a rejected one.
    fn do_await(&mut self, val: Value) -> Result<(), VMError> {
        let mut val = val;
        loop {
            // Clone the Arc so `val` is free to be reassigned for the next
            // flatten iteration without a borrow conflict.
            let arc = match &val {
                Value::Object(o) => o.clone(),
                // Primitive: ECMA-262 §6.2.3.1 Await performs
                // PromiseResolve(v) and ALWAYS resumes as a job — one
                // microtask tick even for plain values. Inside an async
                // boundary, suspend and schedule the immediate resume; at
                // top level (no boundary) keep the direct return.
                _ => {
                    if !self.async_floors.is_empty() {
                        let id = self.event_loop.borrow_mut().next_promise_id();
                        self.pending_settled_await = Some((id, val, false));
                        return Err(VMError::new(format!("__jspi__:{}", id)));
                    }
                    if self.top_level_await_ticks() {
                        return Err(self.tick_top_level_await(val, false));
                    }
                    self.push(val)?;
                    return Ok(());
                }
            };
            let o = arc.lock().unwrap();
            let ty = o
                .properties
                .get("__type")
                .map(|v| format!("{}", v))
                .unwrap_or_default();
            if ty == "Task" {
                drop(o);
                return self.await_task_object(arc);
            }
            if ty != "Promise" {
                // Raw thenable (callable `then`): §27.2.1.3.2 PromiseResolve
                // adopts its eventual state. The host promise engine already
                // implements assimilation — wrap through ecma:promise.resolve
                // and re-enter the loop on the resulting Promise.
                let then_callable = matches!(
                    o.properties.get("then"),
                    Some(Value::Object(f)) if matches!(
                        f.lock().unwrap().kind,
                        crate::value::ObjectKind::Function(_)
                            | crate::value::ObjectKind::HostFunction(_)
                    )
                );
                drop(o);
                if then_callable {
                    let idx = self
                        .host_registry
                        .get(&("ecma:promise".to_string(), "resolve".to_string()))
                        .copied();
                    if let Some(idx) = idx {
                        let host_fn = self.host_fns[idx].clone();
                        let arg = [val.clone()];
                        let resolved = {
                            let mut ctx = self.make_host_context();
                            host_fn(&mut ctx, &arg)
                        };
                        val = resolved;
                        continue;
                    }
                    // No host engine registered — fall through to the
                    // non-thenable handling below (passthrough/tick).
                }
                // Non-thenable object: same one-tick rule as primitives.
                if !self.async_floors.is_empty() {
                    let id = self.event_loop.borrow_mut().next_promise_id();
                    self.pending_settled_await = Some((id, val, false));
                    return Err(VMError::new(format!("__jspi__:{}", id)));
                }
                if self.top_level_await_ticks() {
                    return Err(self.tick_top_level_await(val, false));
                }
                self.push(val)?;
                return Ok(());
            }
            let state = o
                .properties
                .get("__state")
                .map(|v| format!("{}", v))
                .unwrap_or_default();
            if state == "pending" {
                let promise_id = o
                    .properties
                    .get("__id")
                    .map(|v| v.as_f64() as u64)
                    .unwrap_or(0);
                drop(o);
                if !self.async_floors.is_empty() {
                    // Inside a JSPI promising boundary: the innermost
                    // `call_async` performs the delimited capture — do NOT
                    // save the whole program here.
                    return Err(VMError::new(format!("__jspi__:{}", promise_id)));
                }
                return Err(self.suspend_for_pending_promise(promise_id));
            }
            let value = o.properties.get("__value").cloned().unwrap_or(Value::Null);
            if state == "rejected" {
                // JSPI: even a settled promise resumes "by the event queue
                // task runner" — inside an async boundary, suspend (bounded)
                // and schedule the rejection to be thrown into the resumed
                // fiber as a microtask (its captured try/catch handlers fire
                // there). No synchronous shortcut.
                if !self.async_floors.is_empty() {
                    drop(o);
                    let id = self.event_loop.borrow_mut().next_promise_id();
                    self.pending_settled_await = Some((id, value, true));
                    return Err(VMError::new(format!("__jspi__:{}", id)));
                }
                // await on a rejected promise throws the rejection reason —
                // exactly like THROW. At the genuine top level this still
                // takes the one-tick path (rejection thrown into the resumed
                // fiber). Inside driven continuations, raise directly:
                // raise_exception_value handles handler matching, frame/stack
                // unwinding AND label-stack truncation plus propagation
                // across continuation boundaries.
                drop(o);
                if self.top_level_await_ticks() {
                    return Err(self.tick_top_level_await(value, true));
                }
                return self.raise_exception_value(value);
            }
            // fulfilled — flatten if the resolved value is itself a promise,
            // otherwise unwrap and continue.
            drop(o);
            if matches!(&value, Value::Object(inner)
                if inner.lock().unwrap().properties.get("__type")
                    .map(|v| format!("{}", v)).as_deref() == Some("Promise"))
            {
                val = value;
                continue;
            }
            // JSPI: a resolved promise still resumes via the event queue task
            // runner. Inside an async boundary, suspend (bounded) and schedule
            // an immediate microtask resume with the fulfilled value — the
            // spec "await always yields one tick" ordering, no sync shortcut.
            if !self.async_floors.is_empty() {
                let id = self.event_loop.borrow_mut().next_promise_id();
                self.pending_settled_await = Some((id, value, false));
                return Err(VMError::new(format!("__jspi__:{}", id)));
            }
            if self.top_level_await_ticks() {
                return Err(self.tick_top_level_await(value, false));
            }
            self.push(value)?;
            return Ok(());
        }
    }

    fn await_task_object(&mut self, task_obj: Arc<Mutex<Object>>) -> Result<(), VMError> {
        self.join_task_object_if_needed(&task_obj);

        let delay_token_cancelled = self.task_delay_token_cancelled(&task_obj);
        let (status, result, exception) = {
            let task = task_obj.lock().unwrap();
            let status = task
                .properties
                .get("status")
                .map(|v| format!("{}", v))
                .unwrap_or_default();
            let result = task.properties.get("result").cloned().unwrap_or(Value::Null);
            let exception = task.properties.get("exception").cloned();
            (status, result, exception)
        };

        let faulted = delay_token_cancelled
            || status.eq_ignore_ascii_case("Faulted")
            || status.eq_ignore_ascii_case("Canceled")
            || exception
                .as_ref()
                .is_some_and(|v| !matches!(v, Value::Null | Value::Undefined));

        if faulted {
            let reason = exception.unwrap_or_else(|| {
                if delay_token_cancelled || status.eq_ignore_ascii_case("Canceled") {
                    make_operation_cancelled_error()
                } else {
                    Value::String(Arc::from("Task faulted"))
                }
            });
            if !self.async_floors.is_empty() {
                let id = self.event_loop.borrow_mut().next_promise_id();
                self.pending_settled_await = Some((id, reason, true));
                return Err(VMError::new(format!("__jspi__:{}", id)));
            }
            if self.top_level_await_ticks() {
                return Err(self.tick_top_level_await(reason, true));
            }
            return self.raise_exception_value(reason);
        }

        if !self.async_floors.is_empty() {
            let id = self.event_loop.borrow_mut().next_promise_id();
            self.pending_settled_await = Some((id, result, false));
            return Err(VMError::new(format!("__jspi__:{}", id)));
        }
        if self.top_level_await_ticks() {
            return Err(self.tick_top_level_await(result, false));
        }
        self.push(result)?;
        Ok(())
    }

    fn task_delay_token_cancelled(&self, task_obj: &Arc<Mutex<Object>>) -> bool {
        let token = {
            let task = task_obj.lock().unwrap();
            task.properties.get("__dotnet_delay_token").cloned()
        };
        let Some(Value::Object(token_obj)) = token else {
            return false;
        };
        let token = token_obj.lock().unwrap();
        ["__dotnet_cancelled", "IsCancellationRequested"]
            .iter()
            .any(|key| {
                token
                    .properties
                    .get(*key)
                    .is_some_and(|value| matches!(value, Value::Bool(true)))
            })
    }

    fn join_task_object_if_needed(&mut self, task_obj: &Arc<Mutex<Object>>) {
        let tid = {
            let task = task_obj.lock().unwrap();
            task.properties
                .get("__thread_id")
                .map(|v| v.as_f64() as i32)
                .unwrap_or(-1)
        };
        if let Some(handle) = self.thread_handles.remove(&tid) {
            let success = match handle.join() {
                Ok(result) => result.first().copied().unwrap_or(1) == 0,
                Err(_) => false,
            };
            let mut task = task_obj.lock().unwrap();
            task.properties
                .insert("iscompleted".into(), Value::Bool(true));
            task.properties.insert("isalive".into(), Value::Bool(false));
            task.properties.insert("hasexited".into(), Value::Bool(true));
            task.properties.insert(
                "exitcode".into(),
                Value::I32(if success { 0 } else { -1 }),
            );
            task.properties.insert(
                "status".into(),
                Value::String(Arc::from(if success {
                    "RanToCompletion"
                } else {
                    "Faulted"
                })),
            );
        }
    }

    fn next_bytes_decode_opcode(&self) -> bool {
        let f = self.frame();
        let code = &self.chunks[f.chunk_index].code;
        f.ip + 3 < code.len()
            && Op::decode(
                ((code[f.ip] as u16) << 8) | code[f.ip + 1] as u16,
                ((code[f.ip + 2] as u16) << 8) | code[f.ip + 3] as u16,
            )
            .is_some()
    }

    pub(crate) fn read_optional_memidx_immediate(&mut self) -> usize {
        let chunk_idx = self.frame().chunk_index;
        let code = &self.chunks[chunk_idx].code;
        let ip = self.frame().ip;
        // Multi-memory selector. VM instructions are always 4 bytes, so the
        // memidx selector is a fixed 4-byte block — `0xEE 0x00 <memidx u16 BE>`
        // — keeping the following instruction 4-aligned. Only emitted for a
        // non-default memory; absent means memidx 0.
        if code.get(ip) == Some(&0xEE) && code.get(ip + 1) == Some(&0x00) {
            let memidx = ((code[ip + 2] as usize) << 8) | (code[ip + 3] as usize);
            self.frame_mut().ip = ip + 4;
            return memidx;
        }
        0
    }

    /// Pop a stringref operand, trapping on a null reference (WASM stringref
    /// spec: every stringref-consuming op except `string.eq` traps on null).
    fn pop_stringref(&mut self) -> Result<Arc<str>, VMError> {
        match self.pop() {
            Value::String(s) => Ok(s),
            Value::Null | Value::TypedNull(_) | Value::Undefined => {
                Err(VMError::new("trap: null string reference"))
            }
            other => Err(VMError::new(format!(
                "trap: expected stringref, got {}",
                other.type_tag()
            ))),
        }
    }

    /// Read a codepoint-iterator view (`string.as_iter` result): its backing
    /// string and current codepoint index, stashed in the object's properties.
    fn read_string_iter(&mut self, view: &Value) -> Result<(Arc<str>, usize), VMError> {
        let obj = match view {
            Value::Object(o) => o,
            _ => return Err(VMError::new("trap: null stringview_iter reference")),
        };
        let guard = obj.lock().unwrap();
        let s = match guard.properties.get("__iter_str") {
            Some(Value::String(s)) => s.clone(),
            _ => return Err(VMError::new("trap: not a stringview_iter")),
        };
        let pos = match guard.properties.get("__iter_pos") {
            Some(Value::I32(p)) => *p as usize,
            _ => 0,
        };
        Ok((s, pos))
    }

    /// Update the codepoint index of an iterator view in place.
    fn write_string_iter_pos(&mut self, view: &Value, pos: usize) -> Result<(), VMError> {
        if let Value::Object(o) = view {
            o.lock()
                .unwrap()
                .properties
                .insert("__iter_pos".to_string(), Value::I32(pos as i32));
        }
        Ok(())
    }

    /// Read `[start, end)` bytes from a GC array of i8/i16 elements (used by
    /// the `string.new_*_array` ops). Each element is one code unit.
    fn read_array_code_units(
        &mut self,
        arr: &Value,
        start: usize,
        end: usize,
    ) -> Result<Vec<u32>, VMError> {
        let obj = match arr {
            Value::Object(o) => o,
            _ => return Err(VMError::new("trap: null array reference")),
        };
        let guard = obj.lock().unwrap();
        let elems = match &guard.kind {
            ObjectKind::Array(v) => v,
            _ => return Err(VMError::new("trap: expected array reference")),
        };
        if start > end || end > elems.len() {
            return Err(VMError::new("trap: array access out of bounds"));
        }
        Ok(elems[start..end]
            .iter()
            .map(|v| v.as_i32() as u32)
            .collect())
    }

    pub(crate) fn read_optional_memarg(&mut self) -> (usize, usize) {
        // Explicit multi-memory selector: the compiler folds any static offset
        // into the address, so a non-default memory is carried by the same
        // fixed 4-byte `0xEE 0x00 <memidx u16 BE>` sentinel `memory.size`/`grow`
        // use — keeping the stream 4-aligned. Offset is therefore 0 here.
        // Checked before the opcode-lookahead so the sentinel is never mistaken
        // for the next instruction.
        {
            let chunk_idx = self.frame().chunk_index;
            let code = &self.chunks[chunk_idx].code;
            let ip = self.frame().ip;
            if code.get(ip) == Some(&0xEE) && code.get(ip + 1) == Some(&0x00) {
                let memidx = ((code[ip + 2] as usize) << 8) | (code[ip + 3] as usize);
                self.frame_mut().ip = ip + 4;
                return (0, memidx);
            }
        }
        if self.next_bytes_decode_opcode() {
            return (0, 0);
        }
        let chunk_idx = self.frame().chunk_index;
        let code = &self.chunks[chunk_idx].code;
        let mut ip = self.frame().ip;
        let align = read_leb_u32(code, &mut ip);
        // Offset is read as u64 so a 64-bit memory's memarg (memory64
        // proposal) decodes correctly. A 32-bit u32 offset decodes to the
        // same value, so this is backward-compatible.
        let offset = read_leb_u64(code, &mut ip) as usize;
        let memidx = if align & 0x40 != 0 {
            read_leb_u32(code, &mut ip) as usize
        } else {
            0
        };
        self.frame_mut().ip = ip;
        (offset, memidx)
    }

    /// Pop a load/store address operand, widening per the target memory's
    /// index type: a 64-bit (memory64) memory pops an i64 address, a 32-bit
    /// memory pops an unsigned i32. Memory64 adds no opcodes — the width comes
    /// from the memory's declared type, exactly as the spec requires.
    fn effective_addr(&mut self, memidx: usize, offset: usize) -> usize {
        if self.mem_is_64(memidx) {
            (self.pop().as_i64() as u64 as usize).saturating_add(offset)
        } else {
            (self.pop().as_i32() as u32 as usize).saturating_add(offset)
        }
    }

    /// Whether memory `memidx` has a 64-bit index type (memory64 proposal).
    fn mem_is_64(&self, memidx: usize) -> bool {
        self.memory_is_64.get(memidx).copied().unwrap_or(false)
    }

    /// Whether table `tidx` has a 64-bit index type (table64 proposal).
    fn tbl_is_64(&self, tidx: usize) -> bool {
        self.table_is_64.get(tidx).copied().unwrap_or(false)
    }

    /// Pop a table index/count operand: i64 (trapping on negative) for a
    /// 64-bit table, else `max(0)` i32. Used by table.grow/fill/copy/init.
    fn pop_table_count(&mut self, is64: bool) -> Result<usize, VMError> {
        if is64 {
            Self::table64_index(self.pop(), "table")
        } else {
            Ok(self.pop().as_i32().max(0) as usize)
        }
    }

    /// Pop a memory-op count/index operand, widening per the memory's index
    /// type: i64 for a 64-bit memory, unsigned i32 otherwise. Used by
    /// `memory.size/grow/copy/fill` — all standard opcodes; memory64 adds none.
    fn pop_mem_index(&mut self, is64: bool) -> usize {
        if is64 {
            self.pop().as_i64().max(0) as usize
        } else {
            self.pop().as_i32().max(0) as usize
        }
    }

    fn memory64_effective_addr(&self, base: i64, offset: u64) -> Result<usize, VMError> {
        let addr = (base as u64)
            .checked_add(offset)
            .ok_or_else(|| VMError::new("trap: memory64 address overflow"))?;
        usize::try_from(addr).map_err(|_| VMError::new("trap: memory64 address out of range"))
    }

    fn read_optional_simd_memarg(&mut self) -> (u64, usize, bool) {
        let chunk_idx = self.frame().chunk_index;
        let code = &self.chunks[chunk_idx].code;
        let mut ip = self.frame().ip;
        let align = read_leb_u32(code, &mut ip);
        if align & 0x80 == 0 {
            return (0, 0, false);
        }
        let memory64 = align & 0x100 != 0;
        let offset = if memory64 {
            read_leb_u64(code, &mut ip)
        } else {
            read_leb_u32(code, &mut ip) as u64
        };
        let memidx = if align & 0x40 != 0 {
            read_leb_u32(code, &mut ip) as usize
        } else {
            0
        };
        self.frame_mut().ip = ip;
        (offset, memidx, memory64)
    }

    fn pop_simd_addr(&mut self) -> Result<(usize, usize), VMError> {
        let (offset, memidx, memory64) = self.read_optional_simd_memarg();
        let base = self.pop();
        let addr = self.simd_effective_addr(base, offset, memory64)?;
        Ok((memidx, addr))
    }

    fn simd_effective_addr(
        &self,
        base: Value,
        offset: u64,
        memory64: bool,
    ) -> Result<usize, VMError> {
        if memory64 {
            self.memory64_effective_addr(base.as_i64(), offset)
        } else {
            let base = base.as_i32() as u32 as usize;
            base.checked_add(offset as usize)
                .ok_or_else(|| VMError::new("trap: simd memory address overflow"))
        }
    }

    fn table64_index(value: Value, context: &str) -> Result<usize, VMError> {
        let idx = value.as_i64();
        if idx < 0 {
            return Err(VMError::new(format!(
                "trap: {context} negative table index"
            )));
        }
        usize::try_from(idx).map_err(|_| VMError::new(format!("trap: {context} index too large")))
    }

    pub(crate) fn execute_until(&mut self, min_depth: usize) -> Result<Value, VMError> {
        // Track this loop's floor so exception unwinding defers instead of
        // crossing it (see `raise_exception_value`). Pop on every exit.
        self.exec_floors.push(min_depth);
        let result = self.execute_until_inner(min_depth);
        self.exec_floors.pop();
        result
    }

    fn execute_until_inner(&mut self, min_depth: usize) -> Result<Value, VMError> {
        // The fiber this loop was entered on. Stack-switching swaps whole frame
        // stacks mid-loop; `min_depth` is only meaningful for THIS fiber, so
        // the return/end boundaries below are gated on still running it.
        let entry_fiber_id = self.cur_fiber_id;
        let mut dbg_last_op: Option<Op> = None; // TEMP diagnostics (VYBE_DEBUG_AC)
        loop {
            if self.frames.is_empty() && std::env::var("VYBE_DEBUG_AC").is_ok() {
                eprintln!(
                    "[AC-DEBUG] loop-top EMPTY frames: last_op={:?} last_import={:?} stack_len={} ac_len={} fiber={} floors={:?}",
                    dbg_last_op,
                    self.dbg_last_import,
                    self.stack.len(),
                    self.active_continuations.len(),
                    self.cur_fiber_id,
                    self.async_floors
                );
            }
            // Floor check on EVERY iteration, not only at RETURN/END sites:
            // an exception raised inside this (nested) loop can unwind to a
            // handler in a frame BELOW our floor (raise_exception_value jumps
            // ip/frames directly). Continuing here would execute the OUTER
            // loop's frames inside this one — running them to completion and
            // leaving the outer dispatch loop facing an empty frame stack
            // (fatal "no frame"). Yield control back to the Rust caller; the
            // outer loop resumes at the handler ip the unwind established.
            if self.frames.len() < min_depth && self.cur_fiber_id == entry_fiber_id {
                return Ok(Value::Null);
            }
            let f = self.frame();
            let chunk = &self.chunks[f.chunk_index];

            if f.ip >= chunk.code.len() {
                if self.frames.len() <= 1.max(min_depth + 1) && self.cur_fiber_id == entry_fiber_id
                {
                    return Ok(self.stack.pop().unwrap_or(Value::Null));
                }
                let base = self.frame().base;
                self.frames.pop();
                self.stack.truncate(base);
                self.push(Value::Null)?;
                continue;
            }

            // Uniform 2-byte opcode decode: [prefix, sub].
            // Standard WASM ops use prefix 0x00 with the WASM byte as sub.
            // Extended proposals use 0xFC/0xFD/0xFE prefix.
            // Vybe custom ops use 0xFF prefix.
            let opcode_start = self.frame().ip;
            let group = self.read_u16();
            let sub = self.read_u16();
            let op = match Op::decode(group as u16, sub as u16) {
                Some(op) => op,
                None => {
                    return Err(VMError::new(format!(
                        "Invalid opcode: 0x{:04X} 0x{:04X}",
                        group, sub
                    )));
                }
            };
            dbg_last_op = Some(op);

            // ── Instrumentation (step debugger + execution trace) ────────
            // Single hot-path gate: false in normal runs. `opcode_start` is the
            // instruction's own offset (ip has already advanced past the opcode).
            if self.instrumented {
                // Step debugger: taken out of `self` for the call so it can
                // borrow the VM freely, then put back. May BLOCK (pause) or
                // return `__debug_quit__` to terminate.
                if let Some(mut dbg) = self.debugger.take() {
                    let r = dbg.on_instruction(self, opcode_start, op);
                    self.debugger = Some(dbg);
                    r?;
                }
                if self.trace {
                    let f = self.frame();
                    let chunk_name = &self.chunks[f.chunk_index].name;
                    let should_trace = self
                        .trace_chunk_filter
                        .as_ref()
                        .map(|filter| filter == chunk_name)
                        .unwrap_or(true);
                    if should_trace {
                        let ip = f.ip;
                        let stack_top = if self.stack.is_empty() {
                            "[]".to_string()
                        } else {
                            let top = &self.stack[self.stack.len() - 1];
                            let depth = self.stack.len();
                            format!("[{}] (depth={})", top, depth)
                        };
                        eprintln!(
                            "  TRACE {:>12} @{:04} {:?}  stack: {}",
                            chunk_name,
                            ip.saturating_sub(1),
                            op,
                            stack_top
                        );
                    }
                }
            }

            match op {
                _ if op == Op::HALT => {
                    if self.frames.len() <= 1 {
                        // Top-level halt: terminate execution
                        self.close_upvalues(0);
                        return Ok(if self.stack.is_empty() {
                            Value::Null
                        } else {
                            self.pop()
                        });
                    } else {
                        // Nested halt (e.g. script chunk called via bootstrap):
                        // act like return — pop frame and return null
                        let base = self.frame().base;
                        self.close_upvalues(base);
                        self.frames.pop();
                        self.stack.truncate(base);
                        self.push(Value::Null)?;
                    }
                }
                _ if op == Op::UNREACHABLE => {
                    return Err(VMError::new("trap: unreachable executed"));
                }
                _ if op == Op::NOP => { /* no-op */ }

                _ if op == Op::CONST => {
                    let idx = self.read_u16();
                    let val = self.get_constant(idx);
                    self.push(val)?;
                }
                _ if op == Op::DROP => {
                    if self.stack.len() > self.stack_floor() {
                        self.pop();
                    }
                }

                // -- Variables --
                _ if op == Op::LOCAL_GET => {
                    let slot = self.read_u16() as usize;
                    let base = self.frame().base;
                    let idx = base + slot;
                    let val = self
                        .stack
                        .get(idx)
                        .cloned()
                        .ok_or_else(|| VMError::new("trap: local index out of bounds"))?;
                    if let Some(rec) = self.type_recorder.as_mut() {
                        let chunk_idx = self.frames.last().unwrap().chunk_index;
                        rec.record(chunk_idx, slot, &val);
                    }
                    self.push(val)?;
                }
                _ if op == Op::LOCAL_SET => {
                    let slot = self.read_u16() as usize;
                    let val = self.pop();
                    if let Some(rec) = self.type_recorder.as_mut() {
                        let chunk_idx = self.frames.last().unwrap().chunk_index;
                        rec.record(chunk_idx, slot, &val);
                    }
                    let base = self.frame().base;
                    let idx = base + slot;
                    if idx < self.stack.len() {
                        self.stack[idx] = val;
                    } else {
                        let ci = self.frame().chunk_index;
                        let need = base + self.chunks[ci].local_count as usize;
                        if self.stack.len() < need {
                            self.stack.resize(need, Value::Null);
                        }
                        if idx < self.stack.len() {
                            self.stack[idx] = val;
                        }
                    }
                }
                _ if op == Op::LOCAL_TEE => {
                    let slot = self.read_u16() as usize;
                    let val = self.peek(0).clone();
                    if let Some(rec) = self.type_recorder.as_mut() {
                        let chunk_idx = self.frames.last().unwrap().chunk_index;
                        rec.record(chunk_idx, slot, &val);
                    }
                    let base = self.frame().base;
                    let idx = base + slot;
                    // Locals exist (zero-initialized) from function entry
                    // (spec §4.4.9) — grow to the declared local frame the
                    // same way LOCAL_SET does, so a tee before any set
                    // doesn't trap.
                    let ci = self.frame().chunk_index;
                    let need = base + self.chunks[ci].local_count as usize;
                    if self.stack.len() < need && idx < need {
                        self.stack.resize(need, Value::Null);
                    }
                    let dst = self
                        .stack
                        .get_mut(idx)
                        .ok_or_else(|| VMError::new("trap: local index out of bounds"))?;
                    *dst = val;
                }
                _ if op == Op::GLOBAL_GET => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    // In strict isolation mode, prefix globals with module name
                    // to prevent cross-module access
                    let key = if self.strict_isolation {
                        if let Some(ref prefix) = self.module_prefix {
                            let prefixed = format!("{}::{}", prefix, name);
                            // Try prefixed first, then unprefixed (for exports)
                            if self.globals.contains_key(&prefixed) {
                                prefixed
                            } else {
                                name
                            }
                        } else {
                            name
                        }
                    } else {
                        name
                    };
                    let val = self.globals.get(&key).cloned().unwrap_or(Value::Undefined);
                    self.push(val)?;
                }
                _ if op == Op::GLOBAL_SET => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let key = if self.strict_isolation {
                        if let Some(ref prefix) = self.module_prefix {
                            format!("{}::{}", prefix, name)
                        } else {
                            name
                        }
                    } else {
                        name
                    };
                    let val = self.pop();
                    self.globals.insert(key, val);
                }

                // -- Properties --
                _ if op == Op::STRUCT_GET => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let obj = self.pop();
                    // WASM GC `struct.get` traps on a null ref. Only a TYPED null
                    // (a GC reference) traps; a plain null — a dynamic-language
                    // `obj.field` on null — stays lenient (handled below).
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: struct.get on null reference"));
                    }
                    // Auto-join thread when accessing .result on a Task/Thread object
                    if let Value::Object(ref o) = obj {
                        let needs_join = {
                            let o_ref = o.lock().unwrap();
                            (name == "result" || name == "exitcode")
                                && o_ref.properties.contains_key("__thread_id")
                                && !o_ref
                                    .properties
                                    .get("iscompleted")
                                    .map(|v| v.as_bool())
                                    .unwrap_or(true)
                        };
                        if needs_join {
                            let tid = o
                                .lock()
                                .unwrap()
                                .properties
                                .get("__thread_id")
                                .map(|v| v.as_f64() as i32)
                                .unwrap_or(-1);
                            if let Some(handle) = self.thread_handles.remove(&tid) {
                                let _ = handle.join();
                                // Task object was updated by child thread
                            }
                        }
                        // Check for getter
                        let getter_key = format!("__get_{}", name);
                        let getter = o.lock().unwrap().properties.get(&getter_key).cloned();
                        if let Some(getter_fn) = getter {
                            self.push(getter_fn)?;
                            self.push(obj)?;
                            self.call_value(1)?;
                            continue;
                        }
                    }
                    self.push(self.resolve_property(&obj, &name)?)?;
                }
                _ if op == Op::STRUCT_SET => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let val = self.pop();
                    let obj = self.pop();
                    // WASM GC `struct.set` traps on a typed null (GC ref); a plain
                    // null (dynamic-language write) stays lenient.
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: struct.set on null reference"));
                    }
                    if let Value::Object(o) = &obj {
                        // Check for setter: __set_{name}. Property setters
                        // installed by the .NET class wrappers use a
                        // lowercased key (`__set_location`), so fall back
                        // to a case-insensitive lookup for case-sensitive
                        // languages (C#, Dart) whose AST preserves
                        // PascalCase field names.
                        let setter_key = format!("__set_{}", name);
                        let setter_key_lc = format!("__set_{}", name.to_lowercase());
                        let setter = {
                            let props = &o.lock().unwrap().properties;
                            props
                                .get(&setter_key)
                                .cloned()
                                .or_else(|| props.get(&setter_key_lc).cloned())
                        };
                        if let Some(setter_fn) = setter {
                            // Call the setter synchronously. Save stack depth
                            // and restore after — invoke_callback leaks the
                            // return value and intermediate locals on the stack.
                            let stack_save = self.stack.len();
                            let _result =
                                self.invoke_callback(&setter_fn, &[obj.clone(), val.clone()]);
                            self.stack.truncate(stack_save);
                            self.push(val)?;
                        } else {
                            // Set property in properties HashMap
                            o.lock().unwrap().set(name.clone(), val.clone());
                            // For typed objects, also update the fields Vec if this property is a field
                            let type_id = o.lock().unwrap().type_id;
                            if type_id > 0 {
                                if let Some(td) = self.type_registry.get(type_id) {
                                    if let Some(field_idx) = td.field_index(&name) {
                                        let mut ob = o.lock().unwrap();
                                        if field_idx < ob.fields.len() {
                                            ob.fields[field_idx] = val.clone();
                                        }
                                    }
                                }
                            }
                            self.push(val)?;
                        }
                    } else {
                        self.push(val)?;
                    }
                }
                _ if op == Op::ARRAY_GET => {
                    let key = self.pop();
                    let obj = self.pop();
                    // WASM GC `array.get` traps on a typed null (GC array ref);
                    // a plain null (dynamic subscript) stays lenient.
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: array.get on null reference"));
                    }
                    match &obj {
                        Value::Object(o) => {
                            // WASM GC `array.get`: spec (trap on out-of-bounds),
                            // distinct from the lenient dynamic subscript below.
                            if self.is_gc_array_obj(o) {
                                let ob = o.lock().unwrap();
                                if let ObjectKind::Array(a) = &ob.kind {
                                    let idx = match &key {
                                        Value::I32(n) if *n >= 0 => Some(*n as usize),
                                        Value::I64(n) if *n >= 0 => Some(*n as usize),
                                        Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => {
                                            Some(*n as usize)
                                        }
                                        _ => None,
                                    };
                                    match idx {
                                        Some(i) if i < a.len() => {
                                            let v = a[i].clone();
                                            drop(ob);
                                            self.push(v)?;
                                            continue;
                                        }
                                        _ => {
                                            return Err(VMError::new(
                                                "trap: array.get out of bounds",
                                            ));
                                        }
                                    }
                                }
                            }
                            {
                                let ob = o.lock().unwrap();
                                if let ObjectKind::TypedArray(ref ta) = ob.kind {
                                    let numeric_idx = match &key {
                                        Value::I32(n) if *n >= 0 => Some(*n as usize),
                                        Value::I64(n) if *n >= 0 => Some(*n as usize),
                                        Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => {
                                            Some(*n as usize)
                                        }
                                        Value::String(s) => s.parse::<usize>().ok(),
                                        _ => None,
                                    };
                                    if let Some(idx) = numeric_idx {
                                        let val =
                                            typed_array_read(ta, idx).unwrap_or(Value::Undefined);
                                        drop(ob);
                                        self.push(val)?;
                                        continue;
                                    }
                                }
                            }
                            // Map (associative collection — Python dict, PHP keyed
                            // array, Ruby hash, JS Map): IndexMap lookup by Value
                            // key. Distinct from member access on a Map (`m.foo`
                            // → struct_get → Object::get → property bag) which
                            // matches JS semantics (`m.foo !== m.get("foo")`).
                            // Mirrors `ecma:map.get` (ECMA-262 §24.1.3.4).
                            {
                                let ob = o.lock().unwrap();
                                if let ObjectKind::Map(ref m) = ob.kind {
                                    let lookup_key = match &key {
                                        Value::String(_)
                                        | Value::I32(_)
                                        | Value::I64(_)
                                        | Value::F64(_) => key.clone(),
                                        other => {
                                            Value::String(Arc::from(format!("{}", other).as_str()))
                                        }
                                    };
                                    if let Some(v) = m.get(&lookup_key) {
                                        let v = v.clone();
                                        drop(ob);
                                        self.push(v)?;
                                        continue;
                                    }
                                    // Numeric/string key coercion (PHP idiom): "0" ↔ 0.
                                    if let Value::String(s) = &key {
                                        if let Ok(n) = s.parse::<i32>() {
                                            if let Some(v) = m.get(&Value::I32(n)) {
                                                let v = v.clone();
                                                drop(ob);
                                                self.push(v)?;
                                                continue;
                                            }
                                        }
                                    } else if let Value::I32(n) = &key {
                                        if let Some(v) =
                                            m.get(&Value::String(Arc::from(n.to_string().as_str())))
                                        {
                                            let v = v.clone();
                                            drop(ob);
                                            self.push(v)?;
                                            continue;
                                        }
                                    }
                                    drop(ob);
                                    self.push(Value::Undefined)?;
                                    continue;
                                }
                            }
                            // Array / property-bag dispatch — spec-clean indexed
                            // access (no negative-index wrap). ECMA §10.4.2.1
                            // returns undefined for `arr[-1]` and any missing
                            // property; languages whose syntax wraps (Python,
                            // Ruby) normalize the index ahead of this opcode
                            // via the compiler's `emit_negative_index_wrap`
                            // adapter. Missing-key returns Undefined (not Null)
                            // so JS predicates like `cache[k] !== undefined`
                            // work correctly.
                            let k = format!("{}", key);
                            let mut val = o.lock().unwrap().get(&k);
                            if matches!(val, Value::Null) {
                                let ob = o.lock().unwrap();
                                let exists = ob.properties.contains_key(&k)
                                    || matches!(&ob.kind, ObjectKind::Array(a)
                                        if k.parse::<usize>().map(|i| i < a.len()).unwrap_or(false));
                                if !exists {
                                    val = Value::Undefined;
                                }
                            }
                            // If not found and object has __getitem__, call it
                            if matches!(val, Value::Null) {
                                let getitem =
                                    o.lock().unwrap().properties.get("__getitem__").cloned();
                                if let Some(func) = getitem {
                                    self.push(func)?;
                                    self.push(obj.clone())?; // self
                                    self.push(key)?; // key arg
                                    self.call_value(2)?;
                                    continue;
                                }
                            }
                            self.push(val)?;
                        }
                        Value::String(s) => {
                            let i = key.as_f64() as usize;
                            if let Some(ch) = s.chars().nth(i) {
                                self.push(Value::String(Arc::from(ch.to_string().as_str())))?;
                            } else {
                                self.push(Value::Null)?;
                            }
                        }
                        _ => self.push(Value::Null)?,
                    }
                }
                _ if op == Op::ARRAY_SET => {
                    let val = self.pop();
                    let key = self.pop();
                    let obj = self.pop();
                    // WASM GC `array.set` traps on a typed null (GC array ref);
                    // a plain null (dynamic subscript) stays lenient.
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: array.set on null reference"));
                    }
                    if let Value::Object(o) = &obj {
                        // WASM GC `array.set`: spec (trap on out-of-bounds).
                        if self.is_gc_array_obj(o) {
                            let idx = match &key {
                                Value::I32(n) if *n >= 0 => Some(*n as usize),
                                Value::I64(n) if *n >= 0 => Some(*n as usize),
                                Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
                                _ => None,
                            };
                            let mut ob = o.lock().unwrap();
                            if let ObjectKind::Array(a) = &mut ob.kind {
                                match idx {
                                    Some(i) if i < a.len() => {
                                        a[i] = val.clone();
                                        drop(ob);
                                        self.push(val)?;
                                        continue;
                                    }
                                    _ => {
                                        return Err(VMError::new("trap: array.set out of bounds"));
                                    }
                                }
                            }
                        }
                        {
                            let ob = o.lock().unwrap();
                            if let ObjectKind::TypedArray(ref ta) = ob.kind {
                                let numeric_idx = match &key {
                                    Value::I32(n) if *n >= 0 => Some(*n as usize),
                                    Value::I64(n) if *n >= 0 => Some(*n as usize),
                                    Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => {
                                        Some(*n as usize)
                                    }
                                    Value::String(s) => s.parse::<usize>().ok(),
                                    _ => None,
                                };
                                if let Some(idx) = numeric_idx {
                                    typed_array_write(ta, idx, &val);
                                    drop(ob);
                                    self.push(val)?;
                                    continue;
                                }
                            }
                        }
                        // Check for __setitem__ dunder
                        let setitem = o.lock().unwrap().properties.get("__setitem__").cloned();
                        if let Some(func) = setitem {
                            self.push(func)?;
                            self.push(obj.clone())?; // self
                            self.push(key)?; // key
                            self.push(val.clone())?; // value
                            self.call_value(3)?;
                            self.pop(); // discard __setitem__ return
                            self.push(val)?;
                            continue;
                        }
                        // Map (Python dict, PHP keyed array, Ruby hash, JS Map):
                        // insert into the IndexMap by Value-keyed entry, not the
                        // property bag. Mirrors `ecma:map.set` (ECMA-262 §24.1.3.9).
                        {
                            let mut ob = o.lock().unwrap();
                            if let ObjectKind::Map(ref mut m) = ob.kind {
                                let map_key = match &key {
                                    Value::String(_)
                                    | Value::I32(_)
                                    | Value::I64(_)
                                    | Value::F64(_) => key.clone(),
                                    other => {
                                        Value::String(Arc::from(format!("{}", other).as_str()))
                                    }
                                };
                                m.insert(map_key, val.clone());
                                drop(ob);
                                self.push(val)?;
                                continue;
                            }
                        }
                        let k = format!("{}", key);
                        o.lock().unwrap().set(k, val.clone());
                    }
                    self.push(val)?;
                }

                // -- F32 arithmetic (f32 precision, stored as F64) --
                _ if op == Op::F32_ADD => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a + b))?;
                }
                _ if op == Op::F32_SUB => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a - b))?;
                }
                _ if op == Op::F32_MUL => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a * b))?;
                }
                _ if op == Op::F32_DIV => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a / b))?;
                }
                // -- Float arithmetic --
                _ if op == Op::F64_ADD => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a + b))?;
                }
                _ if op == Op::F64_SUB => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a - b))?;
                }
                _ if op == Op::F64_MUL => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a * b))?;
                }
                _ if op == Op::F64_DIV => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a / b))?;
                }
                // f64_mod: removed (non-WASM, use __stdlib_fmod)
                _ if op == Op::F64_NEG => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(-a))?;
                }

                // -- Integer arithmetic --
                _ if op == Op::I32_ADD => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_add(b)))?;
                }
                _ if op == Op::I32_SUB => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_sub(b)))?;
                }
                _ if op == Op::I32_MUL => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_mul(b)))?;
                }
                _ if op == Op::I32_DIV_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    if a == i32::MIN && b == -1 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(a / b))?;
                }
                _ if op == Op::I32_DIV_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I32((a / b) as i32))?;
                }
                _ if op == Op::I32_REM_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I32(a.wrapping_rem(b)))?;
                }
                _ if op == Op::I32_REM_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I32((a % b) as i32))?;
                }

                // -- i64 arithmetic --
                _ if op == Op::I64_ADD => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.wrapping_add(b)))?;
                }
                _ if op == Op::I64_SUB => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.wrapping_sub(b)))?;
                }
                _ if op == Op::I64_MUL => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.wrapping_mul(b)))?;
                }
                _ if op == Op::I64_DIV_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    if a == i64::MIN && b == -1 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(a / b))?;
                }
                _ if op == Op::I64_DIV_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I64((a / b) as i64))?;
                }
                _ if op == Op::I64_REM_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I64(a.wrapping_rem(b)))?;
                }
                _ if op == Op::I64_REM_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I64((a % b) as i64))?;
                }
                _ if op == Op::I64_AND => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a & b))?;
                }
                _ if op == Op::I64_OR => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a | b))?;
                }
                _ if op == Op::I64_XOR => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a ^ b))?;
                }
                _ if op == Op::I64_SHL => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a << (b & 0x3f)))?;
                }
                _ if op == Op::I64_SHR_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a >> (b & 0x3f)))?;
                }
                _ if op == Op::I64_SHR_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::I64((a >> (b & 0x3f)) as i64))?;
                }
                _ if op == Op::I64_ROTL => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::I64(a.rotate_left((b & 0x3f) as u32) as i64))?;
                }
                _ if op == Op::I64_ROTR => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::I64(a.rotate_right((b & 0x3f) as u32) as i64))?;
                }
                _ if op == Op::I64_CLZ => {
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.leading_zeros() as i64))?;
                }
                _ if op == Op::I64_CTZ => {
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.trailing_zeros() as i64))?;
                }
                _ if op == Op::I64_POPCNT => {
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.count_ones() as i64))?;
                }

                // -- f64 math --
                _ if op == Op::F64_ABS => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.abs()))?;
                }
                _ if op == Op::F64_CEIL => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.ceil()))?;
                }
                _ if op == Op::F64_FLOOR => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.floor()))?;
                }
                _ if op == Op::F64_TRUNC => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.trunc()))?;
                }
                _ if op == Op::F64_NEAREST => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.round_ties_even()))?;
                }
                _ if op == Op::F64_SQRT => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.sqrt()))?;
                }
                _ if op == Op::F64_MIN => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(if a.is_nan() || b.is_nan() {
                        f64::NAN
                    } else {
                        a.min(b)
                    }))?;
                }
                _ if op == Op::F64_MAX => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(if a.is_nan() || b.is_nan() {
                        f64::NAN
                    } else {
                        a.max(b)
                    }))?;
                }
                _ if op == Op::F64_COPYSIGN => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.copysign(b)))?;
                }

                // -- f32 (promoted to f64) --
                _ if op == Op::F32_ABS => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.abs()))?;
                }
                _ if op == Op::F32_NEG => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(-a))?;
                }
                _ if op == Op::F32_CEIL => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.ceil()))?;
                }
                _ if op == Op::F32_FLOOR => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.floor()))?;
                }
                _ if op == Op::F32_TRUNC => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.trunc()))?;
                }
                _ if op == Op::F32_NEAREST => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.round_ties_even()))?;
                }
                _ if op == Op::F32_SQRT => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.sqrt()))?;
                }
                _ if op == Op::F32_MIN => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(if a.is_nan() || b.is_nan() {
                        f32::NAN
                    } else {
                        a.min(b)
                    }))?;
                }
                _ if op == Op::F32_MAX => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(if a.is_nan() || b.is_nan() {
                        f32::NAN
                    } else {
                        a.max(b)
                    }))?;
                }
                _ if op == Op::F32_COPYSIGN => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.copysign(b)))?;
                }

                // -- WASM select --
                _ if op == Op::SELECT => {
                    let cond = self.pop().as_i32();
                    let val2 = self.pop();
                    let val1 = self.pop();
                    self.push(if cond != 0 { val1 } else { val2 })?;
                }
                // Typed select (`select t`): same runtime semantics as
                // untyped select; the result-type vec is a validation-time
                // hint. The emitter writes `0x1C <count> <valtype>*`; VM
                // side just pops and picks.
                _ if op == Op::SELECT_T => {
                    let cond = self.pop().as_i32();
                    let val2 = self.pop();
                    let val1 = self.pop();
                    self.push(if cond != 0 { val1 } else { val2 })?;
                }

                // Reference-types `table.get tbl` — pop i32 index, push
                // the table slot as a value. Table 0 is `func_table`
                // (the function-reference table). Tables 1+ live in
                // `wasm_tables`, indexed directly by tableidx.
                // `table.get tbl` — the index is i64 for a 64-bit (table64)
                // table, else i32. table64 adds no new opcodes.
                _ if op == Op::TABLE_GET => {
                    let table_idx = self.read_byte() as usize;
                    let idx = if self.tbl_is_64(table_idx) {
                        Self::table64_index(self.pop(), "table.get")?
                    } else {
                        self.pop().as_i32() as usize
                    };
                    let table = self
                        .table_ref(table_idx)
                        .ok_or_else(|| VMError::new("trap: table.get unknown table"))?;
                    let val = table
                        .get(idx)
                        .cloned()
                        .ok_or_else(|| VMError::new("trap: table.get out of bounds"))?;
                    self.push(val)?;
                }
                // `table.set tbl` — pop value + index, write into table.
                // Trap on out-of-bounds index per spec.
                _ if op == Op::TABLE_SET => {
                    let table_idx = self.read_byte() as usize;
                    let val = self.pop();
                    let idx = if self.tbl_is_64(table_idx) {
                        Self::table64_index(self.pop(), "table.set")?
                    } else {
                        self.pop().as_i32() as usize
                    };
                    let table = self
                        .table_mut(table_idx)
                        .ok_or_else(|| VMError::new("trap: table.set unknown table"))?;
                    if idx >= table.len() {
                        return Err(VMError::new("trap: table.set out of bounds"));
                    }
                    table[idx] = val;
                }

                // -- i32 rotation and bit counting --
                _ if op == Op::I32_ROTL => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.rotate_left(b & 0x1f) as i32))?;
                }
                _ if op == Op::I32_ROTR => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.rotate_right(b & 0x1f) as i32))?;
                }
                _ if op == Op::I32_CLZ => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.leading_zeros() as i32))?;
                }
                _ if op == Op::I32_CTZ => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.trailing_zeros() as i32))?;
                }
                _ if op == Op::I32_POPCNT => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.count_ones() as i32))?;
                }

                // -- eqz --
                _ if op == Op::I32_EQZ => {
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a == 0))?;
                }
                _ if op == Op::I64_EQZ => {
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a == 0))?;
                }

                // -- String --

                // -- Bitwise --
                _ if op == Op::I32_AND => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a & b))?;
                }
                _ if op == Op::I32_OR => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a | b))?;
                }
                _ if op == Op::I32_XOR => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a ^ b))?;
                }
                // i32_not: removed (non-WASM, use i32.const -1 + i32.xor)
                _ if op == Op::I32_SHL => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a.wrapping_shl((b as u32) & 0x1f)))?;
                }
                _ if op == Op::I32_SHR_S => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a >> (b & 0x1f)))?;
                }
                _ if op == Op::I32_SHR_U => {
                    let b = self.pop().to_ecma_int32() as u32;
                    let a = self.pop().to_ecma_int32() as u32;
                    self.push(Value::I32((a >> (b & 0x1f)) as i32))?;
                }

                // -- Comparison --
                // i32 comparisons (WASM MVP 0x46–0x4F)
                _ if op == Op::I32_EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(wasm_bool(a.eq(&b)))?;
                }
                _ if op == Op::I32_NE => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(wasm_bool(!a.eq(&b)))?;
                }
                _ if op == Op::I32_LT_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a < b))?;
                }
                _ if op == Op::I32_LT_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a < b))?;
                }
                _ if op == Op::I32_GT_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a > b))?;
                }
                _ if op == Op::I32_GT_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a > b))?;
                }
                _ if op == Op::I32_LE_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a <= b))?;
                }
                _ if op == Op::I32_LE_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a <= b))?;
                }
                _ if op == Op::I32_GE_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a >= b))?;
                }
                _ if op == Op::I32_GE_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a >= b))?;
                }
                // i64 comparisons (WASM MVP 0x51–0x5A)
                _ if op == Op::I64_EQ => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a == b))?;
                }
                _ if op == Op::I64_NE => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a != b))?;
                }
                _ if op == Op::I64_LT_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a < b))?;
                }
                _ if op == Op::I64_LT_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a < b))?;
                }
                _ if op == Op::I64_GT_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a > b))?;
                }
                _ if op == Op::I64_GT_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a > b))?;
                }
                _ if op == Op::I64_LE_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a <= b))?;
                }
                _ if op == Op::I64_LE_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a <= b))?;
                }
                _ if op == Op::I64_GE_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a >= b))?;
                }
                _ if op == Op::I64_GE_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a >= b))?;
                }
                // f32 comparisons (WASM MVP 0x5B–0x60) — operate on f32 precision
                _ if op == Op::F32_EQ => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a == b))?;
                }
                _ if op == Op::F32_NE => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a != b))?;
                }
                _ if op == Op::F32_LT => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a < b))?;
                }
                _ if op == Op::F32_GT => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a > b))?;
                }
                _ if op == Op::F32_LE => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a <= b))?;
                }
                _ if op == Op::F32_GE => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a >= b))?;
                }
                // f64 comparisons (WASM MVP 0x61–0x66)
                _ if op == Op::F64_EQ => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a == b))?;
                }
                _ if op == Op::F64_NE => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a != b))?;
                }
                _ if op == Op::F64_LT => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a < b))?;
                }
                _ if op == Op::F64_GT => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a > b))?;
                }
                _ if op == Op::F64_LE => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a <= b))?;
                }
                _ if op == Op::F64_GE => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a >= b))?;
                }
                // str_lt, str_gt: removed (non-WASM, were unused)

                // -- Logical --
                // bool_not: removed (non-WASM, use dyn_to_bool + i32_eqz)

                // -- Control flow --
                _ if op == Op::BR => {
                    let ci = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let depth = read_leb_u32(&self.chunks[ci].code, &mut ip) as usize;
                    self.frame_mut().ip = ip;
                    if let Some(entry) = self.label_stack.iter().rev().nth(depth).copied() {
                        self.branch_to_label(depth, entry);
                    }
                }
                _ if op == Op::BR_IF => {
                    let ci = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let depth = read_leb_u32(&self.chunks[ci].code, &mut ip) as usize;
                    self.frame_mut().ip = ip;
                    let cond = match self.pop() {
                        crate::value::Value::I32(n) => n,
                        crate::value::Value::Bool(b) => b as i32,
                        other => {
                            return Err(VMError::new(format!(
                                "type mismatch: br_if expected i32 condition, got {}",
                                other.tag().name()
                            )));
                        }
                    };
                    if cond != 0 {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth).copied() {
                            self.branch_to_label(depth, entry);
                        }
                    }
                }

                // -- Functions --
                _ if op == Op::CALL => {
                    let argc = self.read_byte() as usize;
                    self.call_value(argc)?;
                }
                _ if op == Op::CALL_REF => {
                    // Direct call through a function reference — same as call
                    // but the func ref is already on the stack (no table lookup).
                    let argc = self.read_byte() as usize;
                    self.call_value(argc)?;
                }
                // Multi-value RETURN: the current chunk declares how many
                // results it produces via `result_arity` (defaults to 1 —
                // the single-value baseline). We pop N values from the
                // top of the stack, unwind the frame, then push them back
                // onto the caller's stack in original order. When the
                // outermost frame unwinds we surface the last value as
                // the "final" return (callers that want every value can
                // read them off the stack before the frame is popped;
                // `Ok(...)` is a scalar channel).
                _ if op == Op::RETURN => {
                    let frame_chunk = self.frame().chunk_index;
                    let frame_label_base = self.frame().label_base;
                    let n = (self.chunks[frame_chunk].result_arity as usize).max(1);
                    let split = self.stack.len().saturating_sub(n);
                    let mut results: Vec<Value> = self.stack.split_off(split);
                    let base = self.frame().base;
                    self.close_upvalues(base);
                    self.label_stack.truncate(frame_label_base);
                    self.frames.pop();
                    // A returning frame's exception handlers die with the frame:
                    // a JS async-wrapper `try_table`, or a `return` from inside a
                    // `try`, leaves handlers that nothing else pops (a `return`
                    // is a frame exit, not a structural `br`). Drop them here,
                    // frame-scoped — matching how `save_fiber` partitions
                    // handlers by frame depth. This replaces the async-wrapper
                    // `TRY_END` the frontend used to emit before RETURN. The
                    // empty-frames continuation-completion case below is owned by
                    // the fiber/cont machinery, so leave its handlers alone.
                    if !self.frames.is_empty() {
                        let live = self.frames.len();
                        self.exception_handlers.retain(|h| h.frame_depth <= live);
                    }
                    // Continuation completion: per the WASM stack-switching
                    // proposal a continuation owns its stack; `save_fiber` drains
                    // the caller's frames when the continuation is resumed, so
                    // during its execution `frames` holds the continuation's
                    // frames alone and a genuine body completion empties them.
                    // Mark the cont Done and transfer control back to the
                    // caller of RESUME/GEN_NEXT rather than exiting the VM.
                    if self.frames.is_empty() {
                        if std::env::var("VYBE_DEBUG_AC").is_ok() {
                            eprintln!(
                                "[AC-DEBUG] empty-frames RETURN: chunk={} ip={} min_depth={} ac_len={} fiber={} stack_len={}",
                                self.chunks[frame_chunk].name,
                                opcode_start,
                                min_depth,
                                self.active_continuations.len(),
                                self.cur_fiber_id,
                                self.stack.len()
                            );
                        }
                        if let Some(ac) = self.active_continuations.pop() {
                            if let Value::Object(ref obj) = ac.cont {
                                let o = obj.lock().unwrap();
                                if let ObjectKind::Continuation(cs) = &o.kind {
                                    *cs.state.lock().unwrap() =
                                        crate::value::ContinuationPhase::Done;
                                }
                            }
                            let ret_val = results.pop().unwrap_or(Value::Null);
                            self.resume_fiber_with(ac.caller_fiber, Some(ret_val))?;
                            if ac.mode == ResumeMode::Iterator {
                                // has_more = 0 — generator is exhausted
                                self.push(Value::I32(0))?;
                            }
                            continue;
                        }
                        let last = results.pop().unwrap_or(Value::Null);
                        return Ok(last);
                    }
                    // `execute_until` boundary: a nested loop (e.g. a host
                    // `invoke_callback`) returns when frames fall below its
                    // floor — but ONLY while running on the same fiber it was
                    // entered on. A continuation resumed from inside a callback
                    // runs on a different fiber whose nested returns must NOT
                    // trip the callback's (now-stale, pre-`save_fiber`) floor;
                    // those take the normal-return path below.
                    if self.frames.len() < min_depth && self.cur_fiber_id == entry_fiber_id {
                        let last = results.pop().unwrap_or(Value::Null);
                        return Ok(last);
                    }
                    self.stack.truncate(base);
                    for r in results {
                        self.push(r)?;
                    }
                }
                _ if op == Op::REF_FUNC => {
                    let func_idx = self.read_u16() as usize;
                    // The uv_count byte's high bit (0x80) is a "do not intern"
                    // flag: a bound method stamps a per-receiver property on its
                    // funcref, so it must be a FRESH object per binding — never
                    // the shared interned canonical one. The low 7 bits are the
                    // real upvalue count (methods never have ≥128 captures).
                    let uv_raw = self.read_byte();
                    let no_intern = uv_raw & 0x80 != 0;
                    let uv_count = (uv_raw & 0x7f) as usize;

                    // Capture-free funcref: return the interned canonical object
                    // so two `ref.func $f` tear-offs are reference-identical. This
                    // is what lets `ref.eq` be a pure `Arc::ptr_eq` — identity is
                    // established at CREATION, not faked at comparison time. (No
                    // upvalue bytes follow when uv_count == 0, so the instruction
                    // stream is already fully consumed.)
                    if uv_count == 0 && !no_intern {
                        if let Some(cached) = self.funcref_cache.get(&func_idx) {
                            let v = cached.clone();
                            self.push(v)?;
                            continue;
                        }
                    }

                    let chunk = &self.chunks[func_idx];
                    let arity = chunk.arity;
                    let name = if chunk.name == "<script>" {
                        None
                    } else {
                        Some(chunk.name.clone())
                    };

                    let mut upvalues: Vec<Arc<Mutex<Upvalue>>> = Vec::with_capacity(uv_count);
                    for _ in 0..uv_count {
                        let is_local = self.read_byte() != 0;
                        // u16 like every other slot operand — a parent
                        // frame can have far more than 255 locals.
                        let index = self.read_u16() as usize;
                        if is_local {
                            let base = self.frame().base;
                            let uv = self.capture_upvalue(base + index);
                            upvalues.push(uv);
                        } else {
                            let uv = self.frame().upvalues[index].clone();
                            upvalues.push(uv);
                        }
                    }

                    let func = Function {
                        name,
                        arity,
                        chunk_index: func_idx,
                        upvalues,
                    };
                    let mut obj = Object {
                        properties: HashMap::new(),
                        kind: ObjectKind::Function(func),
                        type_id: 0,
                        fields: Vec::new(),
                    };
                    // Add to function table for call_indirect
                    let table_idx = self.func_table.len();
                    obj.properties
                        .insert("__table_idx".into(), Value::F64(table_idx as f64));
                    let func_val = Value::Object(Arc::new(Mutex::new(obj)));
                    self.func_table.push(func_val.clone());
                    // Intern the canonical capture-free funcref for reuse.
                    if uv_count == 0 {
                        self.funcref_cache.insert(func_idx, func_val.clone());
                    }
                    self.push(func_val)?;
                }

                // -- Host functions --
                _ if op == Op::CALL_IMPORT => {
                    let import_idx = self.read_u16() as usize;
                    let argc = self.read_byte() as usize;
                    let chunk_index = self.frame().chunk_index;

                    let target = match self.resolve_chunk_import(chunk_index, import_idx)? {
                        Some(target) => target,
                        None => {
                            if import_idx >= self.import_table.len() {
                                return Err(VMError::new(format!(
                                    "Unresolved import index: {}",
                                    import_idx
                                )));
                            }
                            self.import_table[import_idx].clone()
                        }
                    };

                    match target {
                        ImportTarget::Host(host_idx) => {
                            let base = self.stack.len() - argc;
                            let args: Vec<Value> = self.stack[base..].to_vec();
                            self.stack.truncate(base);

                            if std::env::var("VYBE_DEBUG_AC").is_ok() {
                                self.dbg_last_import = self
                                    .host_registry
                                    .iter()
                                    .find(|(_, v)| **v == host_idx)
                                    .map(|((m, n), _)| format!("{}:{}", m, n));
                            }
                            let host_fn = self.host_fns[host_idx].clone();
                            let result = {
                                let mut ctx = self.make_host_context();
                                host_fn(&mut ctx, &args)
                            };
                            if let Some(exc) = self.last_exception.take() {
                                self.raise_exception_value(exc)?;
                                continue;
                            }

                            // A host fn (wasi:cli/exit) requested clean run
                            // termination: unwind every frame and hand control
                            // back to the embedder — like Op::HALT's top-frame
                            // case, but without a process exit.
                            if self.pending_exit {
                                self.pending_exit = false;
                                self.close_upvalues(0);
                                return Ok(Value::Null);
                            }

                            // A host fn returning a pending promise is just a
                            // VALUE (`new Promise(...)`, `fetch()`, a deferred
                            // promise). Per ECMA-262 / JSPI, suspension happens
                            // ONLY at an `await` (the explicit `PROMISE_SUSPEND`
                            // the compiler emits), never implicitly at the call
                            // that produced the promise — auto-suspending here
                            // wrongly suspended non-awaited promises (deferred
                            // resolvers, `.then` chains, `Promise.race`).
                            self.push(result)?;
                        }
                        ImportTarget::ChunkFn { chunk_index, arity } => {
                            // Component-exported function: build Function value, push below args, call.
                            let func = crate::value::Function {
                                name: None,
                                arity,
                                chunk_index,
                                upvalues: Vec::new(),
                            };
                            let mut obj = crate::value::Object::new();
                            obj.kind = crate::value::ObjectKind::Function(func);
                            let func_val = Value::Object(Arc::new(Mutex::new(obj)));
                            let args_start = self.stack.len() - argc;
                            self.stack.insert(args_start, func_val);
                            self.call_value(argc)?;
                        }
                        ImportTarget::StdlibRedirect(ref global_name) => {
                            if let Some(func_val) = self.globals.get(global_name).cloned() {
                                let args_start = self.stack.len() - argc;
                                self.stack.insert(args_start, func_val);
                                self.call_value(argc)?;
                            } else {
                                return Err(VMError::new(format!(
                                    "Stdlib redirect not found: {}",
                                    global_name
                                )));
                            }
                        }
                        ImportTarget::JspiSuspend => {
                            let val = if argc == 0 {
                                Value::Undefined
                            } else {
                                self.pop()
                            };
                            for _ in 1..argc {
                                self.pop();
                            }
                            self.do_await(val)?;
                        }
                        ImportTarget::StringConst(ref s) => {
                            for _ in 0..argc {
                                self.pop();
                            }
                            self.push(Value::String(s.clone()))?;
                        }
                    }
                }

                // -- Object/Array --
                _ if op == Op::STRUCT_NEW => {
                    let count = self.read_u16() as usize;
                    let mut obj = Object::new();
                    let needed = count * 2;
                    let available = self.stack.len();
                    let start = if needed <= available {
                        available - needed
                    } else {
                        0
                    };
                    for i in 0..count {
                        let key = format!("{}", self.stack[start + i * 2]);
                        let val = self.stack[start + i * 2 + 1].clone();
                        obj.set(key, val);
                    }
                    self.stack.truncate(start);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                // `array.new_fixed $t N` — pops N values off the stack
                // and allocates an N-element array initialised from them.
                _ if op == Op::ARRAY_NEW_FIXED => {
                    let count = self.read_u16() as usize;
                    let count = count.min(self.stack.len());
                    let start = self.stack.len() - count;
                    let elems: Vec<Value> = self.stack[start..].to_vec();
                    self.stack.truncate(start);
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(
                        elems,
                    )))))?;
                }
                // `array.new $t` — [value, length] -> [array of length,
                // every lane = value].
                _ if op == Op::ARRAY_NEW => {
                    // Immediate is a 1-based index into the script chunk's type
                    // table naming an `(array …)` defined type (0 = a
                    // dynamic-language array, no GC type). Resolved to the
                    // instance's rtt (registry id) via the type name — the
                    // compile-time table position can't be the registry id
                    // directly because the VM pre-registers builtin types ahead
                    // of the module's own. Stamping the rtt makes `array.get`/
                    // `set`/`copy` trap per spec for GC arrays and stay lenient
                    // for dynamic arrays (id 0 = `Object`, a `Struct` kind).
                    let typeidx = self.read_u16() as usize;
                    let len = self.pop().as_i32().max(0) as usize;
                    let value = self.pop();
                    let elems = vec![value; len];
                    let mut obj = Object::new_array(elems);
                    obj.type_id = self.resolve_gc_array_rtt(typeidx);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                // `array.new_default $t` — [length] -> [array of length,
                // zero-initialised]. We use `Value::Null` as the default
                // for externref lanes (the only lane type we actually
                // support) per the "null is the default for refs" rule.
                _ if op == Op::ARRAY_NEW_DEFAULT => {
                    // Immediate is a 1-based script-chunk type-table index (0 =
                    // dynamic); resolved to the instance rtt so a defaulted GC
                    // array traps per spec, matching `array.new`.
                    let typeidx = self.read_u16() as usize;
                    let len = self.pop().as_i32().max(0) as usize;
                    let elems = vec![Value::Null; len];
                    let mut obj = Object::new_array(elems);
                    obj.type_id = self.resolve_gc_array_rtt(typeidx);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                // `array.new_data $t $d` / `array.new_elem $t $e` — allocate
                // a new array initialised from a data or element segment.
                // Our VM doesn't (yet) model data/element segments, so we
                // produce an empty array rather than silently returning
                // garbage. Emitted WASM still carries the spec-correct
                // opcode bytes, so engines with real segment support
                // execute these correctly.
                _ if op == Op::ARRAY_NEW_DATA => {
                    let _typeidx = self.read_u16();
                    let dataidx = self.read_u16() as u32;
                    if self.dropped_data.contains(&dataidx) {
                        return Err(VMError::new("array.new_data: data segment dropped"));
                    }
                    let size = self.pop().as_i32().max(0) as usize;
                    let offset = self.pop().as_i32().max(0) as usize;
                    let data = self
                        .data_segments
                        .get(dataidx as usize)
                        .ok_or_else(|| VMError::new("array.new_data: missing data segment"))?;
                    let end = offset.saturating_add(size);
                    if end > data.len() {
                        return Err(VMError::new("array.new_data: out of bounds"));
                    }
                    let elems = data[offset..end]
                        .iter()
                        .map(|b| Value::I32(*b as i32))
                        .collect();
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(
                        elems,
                    )))))?;
                }
                _ if op == Op::ARRAY_NEW_ELEM => {
                    let _typeidx = self.read_u16();
                    let elemidx = self.read_u16() as u32;
                    if self.dropped_elems.contains(&elemidx) {
                        return Err(VMError::new("array.new_elem: element segment dropped"));
                    }
                    let size = self.pop().as_i32().max(0) as usize;
                    let offset = self.pop().as_i32().max(0) as usize;
                    let elems = self
                        .elem_segments
                        .get(elemidx as usize)
                        .ok_or_else(|| VMError::new("array.new_elem: missing element segment"))?;
                    let end = offset.saturating_add(size);
                    if end > elems.len() {
                        return Err(VMError::new("array.new_elem: out of bounds"));
                    }
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(
                        elems[offset..end].to_vec(),
                    )))))?;
                }
                // `array.get_s $t` / `array.get_u $t` — only applicable to
                // arrays of packed element types (i8/i16). Our array model
                // is externref-only, so no packing conversion is needed:
                // both behave identically to `array.get`.
                // `array.get_s $t` / `array.get_u $t` — packed-array read.
                // Spec applies these only to arrays whose field type is
                // packed (i8 / i16); the byte value is sign-extended (S)
                // or zero-extended (U) to i32. We honour that on typed
                // storage (TypedArray / ArrayBuffer) and fall back to a
                // plain read for Value arrays.
                _ if op == Op::ARRAY_GET_S || op == Op::ARRAY_GET_U => {
                    let _typeidx = self.read_u16();
                    let is_signed = op == Op::ARRAY_GET_S;
                    let idx = self.pop().as_i32().max(0) as usize;
                    let arr = self.pop();
                    let val = if let Value::Object(obj) = arr {
                        let o = obj.lock().unwrap();
                        match &o.kind {
                            ObjectKind::TypedArray(ta) => {
                                let buf = ta.buffer.lock().unwrap();
                                let bpe = ta.elem.bytes_per_element();
                                let base = ta.byte_offset + idx * bpe;
                                match bpe {
                                    1 => {
                                        let b = buf.get(base).copied().unwrap_or(0);
                                        let v = if is_signed {
                                            (b as i8) as i32
                                        } else {
                                            b as i32
                                        };
                                        Value::I32(v)
                                    }
                                    2 => {
                                        let lo = buf.get(base).copied().unwrap_or(0) as u16;
                                        let hi = buf.get(base + 1).copied().unwrap_or(0) as u16;
                                        let raw = lo | (hi << 8);
                                        let v = if is_signed {
                                            (raw as i16) as i32
                                        } else {
                                            raw as i32
                                        };
                                        Value::I32(v)
                                    }
                                    _ => Value::Null,
                                }
                            }
                            ObjectKind::ArrayBuffer(ab) => {
                                let buf = ab.bytes.lock().unwrap();
                                let b = buf.get(idx).copied().unwrap_or(0);
                                let v = if is_signed {
                                    (b as i8) as i32
                                } else {
                                    b as i32
                                };
                                Value::I32(v)
                            }
                            ObjectKind::Array(elems) => {
                                elems.get(idx).cloned().unwrap_or(Value::Null)
                            }
                            _ => Value::Null,
                        }
                    } else {
                        Value::Null
                    };
                    self.push(val)?;
                }
                // `array.init_data $t $d` / `array.init_elem $t $e` — copy
                // elements into an existing array. Stub to a no-op (same
                // rationale as new_data / new_elem above).
                // `array.init_data $t $d` — copy `count` ELEMENTS from data
                // segment `$d` into an array (WASM GC). Stack: [array, dst_elem,
                // src_byte_offset, count]. Each element occupies `elemsize` bytes
                // in the segment, read little-endian; `src` is a BYTE offset so
                // the source span is `[src, src + count·elemsize)`.
                _ if op == Op::ARRAY_INIT_DATA => {
                    let _typeidx = self.read_u16();
                    let dataidx = self.read_u16() as u32;
                    if self.dropped_data.contains(&dataidx) {
                        return Err(VMError::new("array.init_data: data segment dropped"));
                    }
                    let size = self.pop().as_i32().max(0) as usize;
                    let src_offset = self.pop().as_i32().max(0) as usize;
                    let dst_offset = self.pop().as_i32().max(0) as usize;
                    let array = self.pop();
                    let data = self
                        .data_segments
                        .get(dataidx as usize)
                        .ok_or_else(|| VMError::new("array.init_data: missing data segment"))?
                        .clone();
                    let check_src = |elem_size: usize| -> Result<(), VMError> {
                        let end = src_offset.saturating_add(size.saturating_mul(elem_size));
                        if end > data.len() {
                            return Err(VMError::new("array.init_data: source out of bounds"));
                        }
                        Ok(())
                    };
                    if let Value::Object(obj) = array {
                        // The value model stores i32/f32/f64 all as f64, so the
                        // element byte width cannot be read from the runtime
                        // value — recover it from the array's rtt (its element
                        // storage type, kept as the type's single "field").
                        let (elem_size, kind) = {
                            let tid = obj.lock().unwrap().type_id;
                            self.type_registry
                                .get(tid)
                                .and_then(|td| td.field_defs.first())
                                .and_then(|f| array_elem_storage_kind(&f.name))
                                .unwrap_or((4, 0))
                        };
                        let mut o = obj.lock().unwrap();
                        match &mut o.kind {
                            ObjectKind::Array(elems) => {
                                check_src(elem_size)?;
                                let dst_end = dst_offset.saturating_add(size);
                                if dst_end > elems.len() {
                                    return Err(VMError::new(
                                        "array.init_data: destination out of bounds",
                                    ));
                                }
                                for i in 0..size {
                                    let base = src_offset + i * elem_size;
                                    elems[dst_offset + i] =
                                        decode_le_numeric(kind, &data[base..base + elem_size]);
                                }
                            }
                            ObjectKind::TypedArray(ta) => {
                                let elem_size = ta.elem.bytes_per_element();
                                check_src(elem_size)?;
                                let dst_end = dst_offset.saturating_add(size);
                                if dst_end > typed_array_live_length(ta) {
                                    return Err(VMError::new(
                                        "array.init_data: destination out of bounds",
                                    ));
                                }
                                for i in 0..size {
                                    let base = src_offset + i * elem_size;
                                    let v = decode_typed_le(ta.elem, &data[base..base + elem_size]);
                                    typed_array_write(ta, dst_offset + i, &v);
                                }
                            }
                            _ => return Err(VMError::new("array.init_data: not an array")),
                        }
                    } else {
                        return Err(VMError::new("array.init_data: not an array"));
                    }
                }
                _ if op == Op::ARRAY_INIT_ELEM => {
                    let _typeidx = self.read_u16();
                    let elemidx = self.read_u16() as u32;
                    if self.dropped_elems.contains(&elemidx) {
                        return Err(VMError::new("array.init_elem: element segment dropped"));
                    }
                    let size = self.pop().as_i32().max(0) as usize;
                    let src_offset = self.pop().as_i32().max(0) as usize;
                    let dst_offset = self.pop().as_i32().max(0) as usize;
                    let array = self.pop();
                    let source = self
                        .elem_segments
                        .get(elemidx as usize)
                        .ok_or_else(|| VMError::new("array.init_elem: missing element segment"))?;
                    let src_end = src_offset.saturating_add(size);
                    if src_end > source.len() {
                        return Err(VMError::new("array.init_elem: source out of bounds"));
                    }
                    if let Value::Object(obj) = array {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(elems) = &mut o.kind {
                            let dst_end = dst_offset.saturating_add(size);
                            if dst_end > elems.len() {
                                return Err(VMError::new(
                                    "array.init_elem: destination out of bounds",
                                ));
                            }
                            elems[dst_offset..dst_end]
                                .clone_from_slice(&source[src_offset..src_end]);
                        } else {
                            return Err(VMError::new("array.init_elem: not an array"));
                        }
                    } else {
                        return Err(VMError::new("array.init_elem: not an array"));
                    }
                }
                // `struct.new_default $t` — no per-field values on stack;
                // produce an all-null struct. Matches our externref model.
                _ if op == Op::STRUCT_NEW_DEFAULT => {
                    let _typeidx = self.read_u16();
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new()))))?;
                }
                // ── Custom Descriptors proposal ───────────────────────────
                // See `proposals/custom-descriptors/`.
                //
                // `struct.new_desc $t`   — [field_0 .. field_{N-1}, descriptor]
                //                          → [ref to $t with descriptor attached]
                // `struct.new_default_desc $t` — [descriptor]
                //                                → [default-initialised ref with descriptor]
                // `ref.get_desc $t`      — [ref] → [descriptor]
                //
                // Vybe's VM doesn't use descriptors for method dispatch at
                // runtime (we go through TypeRegistry / __type properties),
                // but we honour the opcodes so emitted .wasm that leans on
                // descriptor semantics stays correct on engines that do.
                // Descriptors are stashed in the object's `__descriptor`
                // property slot — reading them back via REF_GET_DESC just
                // returns whatever was stamped at construction.
                _ if op == Op::STRUCT_NEW_DESC => {
                    let _typeidx = self.read_u16();
                    let descriptor = self.pop();
                    // Pop N field values — spec takes exactly the type's
                    // declared field count, but our VM treats all fields
                    // as an untyped blob, so we snapshot whatever's on the
                    // stack that came in with STRUCT_NEW-style pairing is
                    // not used here (descriptor variant is lightweight).
                    // Users driving this opcode directly push a single
                    // descriptor + zero fields; languages that need fields
                    // prefer STRUCT_NEW.
                    let mut obj = Object::new();
                    obj.properties.insert("__descriptor".into(), descriptor);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                _ if op == Op::STRUCT_NEW_DEFAULT_DESC => {
                    let _typeidx = self.read_u16();
                    let descriptor = self.pop();
                    let mut obj = Object::new();
                    obj.properties.insert("__descriptor".into(), descriptor);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                _ if op == Op::REF_GET_DESC => {
                    let _typeidx = self.read_u16();
                    let val = self.pop();
                    let desc = if let Value::Object(o) = &val {
                        o.lock()
                            .unwrap()
                            .properties
                            .get("__descriptor")
                            .cloned()
                            .unwrap_or(Value::Null)
                    } else {
                        Value::Null
                    };
                    self.push(desc)?;
                }
                // `struct.get_s $t i` / `struct.get_u $t i` — packed field
                // variants. Our structs have externref fields only, so
                // there's no sign extension to do — both behave like
                // `struct.get`.
                _ if op == Op::STRUCT_GET_S || op == Op::STRUCT_GET_U => {
                    let field_idx = self.read_u16();
                    let obj = self.pop();
                    let val = if let Value::Object(o) = obj {
                        let o = o.lock().unwrap();
                        o.fields
                            .get(field_idx as usize)
                            .cloned()
                            .unwrap_or(Value::Null)
                    } else {
                        Value::Null
                    };
                    self.push(val)?;
                }
                // `ref.test_null ht` / `ref.cast_null ht` — same as their
                // non-null variants but succeed when the operand is null.
                // Our VM already treats null as assignable to externref,
                // so these short-circuit to true / pass-through on null.
                _ if op == Op::REF_TEST_NULL => {
                    let typeidx = self.read_u16();
                    let val = self.pop();
                    let result = if val.is_null_ref() {
                        true
                    } else {
                        let target_name = self.constant_str(typeidx);
                        self.test_type(&val, &target_name)
                    };
                    self.push(Value::I32(if result { 1 } else { 0 }))?;
                }
                _ if op == Op::REF_CAST_NULL => {
                    let typeidx = self.read_u16();
                    let val = self.peek(0).clone();
                    if !val.is_null_ref() {
                        let target_name = self.constant_str(typeidx);
                        if !self.test_type(&val, &target_name) {
                            return Err(VMError::new(&format!(
                                "ref.cast_null failed: value is not {}",
                                target_name
                            )));
                        }
                    }
                }
                // `any.convert_extern` / `extern.convert_any` — identity at
                // runtime for us: our value ABI is a universal externref.
                // Spec says composing the two yields the original value, so
                // emitting them as nops is semantically correct.
                _ if op == Op::ANY_CONVERT_EXTERN || op == Op::EXTERN_CONVERT_ANY => {}
                // `ref.as_non_null` — trap if the operand is null, otherwise
                // pass the value through unchanged.
                _ if op == Op::REF_AS_NON_NULL => {
                    if self.stack.last().map_or(false, |v| v.is_null_ref()) {
                        return Err(VMError::new("trap: ref.as_non_null on null reference"));
                    }
                }
                // `br_on_null $l` — if TOS is null, pop it and branch;
                // otherwise leave the value on the stack and fall through.
                // `br_on_non_null $l` — if TOS is non-null, branch with
                // the value; otherwise pop and fall through.
                _ if op == Op::BR_ON_NULL => {
                    let offset = self.read_i16();
                    let is_null = self.stack.last().map_or(false, |v| v.is_null_ref());
                    if is_null {
                        self.pop();
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                _ if op == Op::BR_ON_NON_NULL => {
                    let offset = self.read_i16();
                    let is_null = self.stack.last().map_or(false, |v| v.is_null_ref());
                    if !is_null {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    } else {
                        self.pop();
                    }
                }

                // -- Immediates --
                _ if op == Op::NULL => self.push(Value::Null)?,
                // `ref.null none` — a WASM GC typed null (traps on GC accessors).
                _ if op == Op::NULL_NONE => self.push(Value::TypedNull(0))?,

                _ if op == Op::I32_CONST => {
                    let v = self.read_leb_i32();
                    self.push(Value::I32(v))?;
                }
                _ if op == Op::I64_CONST => {
                    let v = self.read_leb_i64();
                    self.push(Value::I64(v))?;
                }
                _ if op == Op::F32_CONST => {
                    let v = self.read_f32();
                    self.push(Value::F32(v))?;
                }
                _ if op == Op::F64_CONST => {
                    let v = self.read_f64();
                    self.push(Value::F64(v))?;
                }

                // ref.eq (GC proposal) — reference identity equality.
                // Two references are equal iff they point at the same
                // underlying object. Null-null is also true. Used by JS
                // `===` for object identity.
                _ if op == Op::REF_EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    let eq = match (&a, &b) {
                        // All nulls (typed or plain) are ref.eq.
                        _ if a.is_null_ref() && b.is_null_ref() => true,
                        (Value::Undefined, Value::Undefined) => true,
                        // Pure WASM `ref.eq`: reference identity only. Two
                        // `ref.func $f` tear-offs of the same capture-free
                        // function are identical because `REF_FUNC` INTERNS them
                        // (one canonical object per function) — identity is
                        // established at creation, not faked here. Closures with
                        // captures stay distinct, as do bound methods (which
                        // capture `self`), so `C().f is C().f` is correctly false.
                        (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
                        (Value::Symbol(a), Value::Symbol(b)) => Arc::ptr_eq(a, b),
                        (Value::String(a), Value::String(b)) => Arc::ptr_eq(a, b),
                        _ => false,
                    };
                    self.push(wasm_bool(eq))?;
                }

                // -- Type checks --
                // ref_test: TypeOf...Is using TypeRegistry hierarchy.
                // Delegates to test_type() which handles: type_id lookup,
                // __type string, __types array (JS inheritance), __control_type.
                _ if op == Op::REF_TEST => {
                    let type_name_idx = self.read_u16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.pop();
                    let result = self.test_type(&val, &target_name);
                    self.push(wasm_bool(result))?;
                }
                _ if op == Op::REF_CAST => {
                    let type_name_idx = self.read_u16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    let is_type = self.test_type(&val, &target_name);
                    if !is_type {
                        return Err(VMError::new(&format!(
                            "ref.cast failed: value is not {}",
                            target_name
                        )));
                    }
                    // Value stays on stack (cast is a no-op if it passes)
                }
                // `br_on_cast l ht` / `br_on_cast_fail l ht` — structured
                // branch keyed off a runtime type test. Operand is
                // (u16 type-name-idx, u8 label-depth), matching core `br`'s
                // label-stack discipline so the VM can honour the branch
                // without a parallel byte-offset table.
                _ if op == Op::BR_ON_CAST => {
                    let type_name_idx = self.read_u16();
                    let depth = self.read_byte() as usize;
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    if self.test_type(&val, &target_name) {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth).copied() {
                            self.branch_to_label(depth, entry);
                        }
                    }
                }
                _ if op == Op::BR_ON_CAST_FAIL => {
                    let type_name_idx = self.read_u16();
                    let depth = self.read_byte() as usize;
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    if !self.test_type(&val, &target_name) {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth).copied() {
                            self.branch_to_label(depth, entry);
                        }
                    }
                }

                // -- i31ref (tagged small integers) --
                _ if op == Op::I31_NEW => {
                    // Box i32 as i31ref. In our VM, I32 is already unboxed,
                    // so this is a no-op identity. The optimization is that
                    // the VM can use I32 directly without heap allocation.
                    let v = self.pop().as_i32();
                    self.push(Value::I32(v & 0x7FFF_FFFF))?; // mask to 31 bits
                }
                _ if op == Op::I31_GET_S => {
                    let v = self.pop().as_i32();
                    // Sign extend from 31 bits
                    let extended = if v & 0x4000_0000 != 0 {
                        v | !0x7FFF_FFFF_u32 as i32
                    } else {
                        v
                    };
                    self.push(Value::I32(extended))?;
                }
                _ if op == Op::I31_GET_U => {
                    let v = self.pop().as_i32();
                    self.push(Value::I32(v & 0x7FFF_FFFF))?;
                }

                // ── Stringref proposal ────────────────────────────────────
                // Strings are `Value::String`. The `$mem` immediate defaults to
                // memory 0 (no immediate bytes read — matches operand_format).
                _ if op == Op::STRING_NEW_UTF8 || op == Op::STRING_NEW_WTF8 => {
                    let len = self.pop().as_i32() as u32 as usize;
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let bytes = self.read_memory_bytes(0, ptr, len)?;
                    // string.new_utf8 traps on invalid UTF-8; new_wtf8 accepts
                    // valid UTF-8 identically (WTF-8 surrogate forms unsupported).
                    let s = String::from_utf8(bytes)
                        .map_err(|_| VMError::new("trap: invalid UTF-8"))?;
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::STRING_NEW_LOSSY_UTF8 => {
                    let len = self.pop().as_i32() as u32 as usize;
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let bytes = self.read_memory_bytes(0, ptr, len)?;
                    let s = String::from_utf8_lossy(&bytes).into_owned();
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::STRING_NEW_UTF8_ARRAY || op == Op::STRING_NEW_WTF16_ARRAY => {
                    let end = self.pop().as_i32() as u32 as usize;
                    let start = self.pop().as_i32() as u32 as usize;
                    let arr = self.pop();
                    let units = self.read_array_code_units(&arr, start, end)?;
                    let s: String = if op == Op::STRING_NEW_WTF16_ARRAY {
                        let u16s: Vec<u16> = units.iter().map(|&u| u as u16).collect();
                        String::from_utf16_lossy(&u16s)
                    } else {
                        let bytes: Vec<u8> = units.iter().map(|&u| u as u8).collect();
                        String::from_utf8(bytes).map_err(|_| VMError::new("trap: invalid UTF-8"))?
                    };
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::STRING_MEASURE_UTF8 || op == Op::STRING_MEASURE_WTF8 => {
                    let s = self.pop_stringref()?;
                    self.push(Value::I32(s.len() as i32))?;
                }
                _ if op == Op::STRING_MEASURE_WTF16 => {
                    let s = self.pop_stringref()?;
                    self.push(Value::I32(s.encode_utf16().count() as i32))?;
                }
                _ if op == Op::STRING_ENCODE_UTF8 => {
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let bytes = s.as_bytes().to_vec();
                    self.write_memory_bytes(0, ptr, &bytes)?;
                    self.push(Value::I32(bytes.len() as i32))?;
                }
                _ if op == Op::STRING_ENCODE_WTF16 => {
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let units: Vec<u16> = s.encode_utf16().collect();
                    let mut bytes = Vec::with_capacity(units.len() * 2);
                    for u in &units {
                        bytes.extend_from_slice(&u.to_le_bytes());
                    }
                    self.write_memory_bytes(0, ptr, &bytes)?;
                    self.push(Value::I32(units.len() as i32))?;
                }
                _ if op == Op::STRING_ENCODE_UTF8_ARRAY || op == Op::STRING_ENCODE_WTF16_ARRAY => {
                    let start = self.pop().as_i32() as u32 as usize;
                    let arr = self.pop();
                    let s = self.pop_stringref()?;
                    let obj = match &arr {
                        Value::Object(o) => o.clone(),
                        _ => return Err(VMError::new("trap: null array reference")),
                    };
                    let wtf16 = op == Op::STRING_ENCODE_WTF16_ARRAY;
                    let units: Vec<u32> = if wtf16 {
                        s.encode_utf16().map(|u| u as u32).collect()
                    } else {
                        s.as_bytes().iter().map(|&b| b as u32).collect()
                    };
                    let mut guard = obj.lock().unwrap();
                    let elems = match &mut guard.kind {
                        ObjectKind::Array(v) => v,
                        _ => return Err(VMError::new("trap: expected array reference")),
                    };
                    if start.saturating_add(units.len()) > elems.len() {
                        return Err(VMError::new("trap: array access out of bounds"));
                    }
                    for (i, u) in units.iter().enumerate() {
                        elems[start + i] = Value::I32(*u as i32);
                    }
                    self.push(Value::I32(units.len() as i32))?;
                }
                _ if op == Op::STRING_CONCAT => {
                    let b = self.pop_stringref()?;
                    let a = self.pop_stringref()?;
                    let mut s = String::with_capacity(a.len() + b.len());
                    s.push_str(&a);
                    s.push_str(&b);
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::STRING_EQ => {
                    // Does NOT trap on null: null == null → 1, null vs string → 0.
                    let b = self.pop();
                    let a = self.pop();
                    let eq = match (&a, &b) {
                        _ if a.is_null_ref() && b.is_null_ref() => true,
                        (Value::String(x), Value::String(y)) => x == y,
                        _ => false,
                    };
                    self.push(Value::I32(i32::from(eq)))?;
                }
                _ if op == Op::STRING_AS_WTF8 || op == Op::STRING_AS_WTF16 => {
                    // Views over the same content: return the string unchanged
                    // (trap on null). The stringview cursor ops below operate on
                    // this string directly (position is an explicit operand).
                    let s = self.pop_stringref()?;
                    self.push(Value::String(s))?;
                }

                // ── Stringref: additional encodings (WTF-16 / WTF-8 / lossy) ──
                // Native strings are always valid UTF-8 (= valid USV sequences,
                // no lone surrogates), so WTF-8 and lossy-UTF-8 encode/decode
                // identically to UTF-8; WTF-16 mirrors the existing UTF-16 paths.
                _ if op == Op::STRING_NEW_WTF16 => {
                    // (ptr, codeunits): read codeunits × 2 bytes as little-endian u16.
                    let units = self.pop().as_i32() as u32 as usize;
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let bytes = self.read_memory_bytes(0, ptr, units * 2)?;
                    let u16s: Vec<u16> = bytes
                        .chunks_exact(2)
                        .map(|c| u16::from_le_bytes([c[0], c[1]]))
                        .collect();
                    let s = String::from_utf16_lossy(&u16s);
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::STRING_ENCODE_WTF8 || op == Op::STRING_ENCODE_LOSSY_UTF8 => {
                    // (str, ptr): write the UTF-8 bytes, return the byte count.
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let bytes = s.as_bytes().to_vec();
                    self.write_memory_bytes(0, ptr, &bytes)?;
                    self.push(Value::I32(bytes.len() as i32))?;
                }
                _ if op == Op::STRING_NEW_WTF8_ARRAY || op == Op::STRING_NEW_LOSSY_UTF8_ARRAY => {
                    // (array, start, end): decode the byte range into a string.
                    let end = self.pop().as_i32() as u32 as usize;
                    let start = self.pop().as_i32() as u32 as usize;
                    let arr = self.pop();
                    let units = self.read_array_code_units(&arr, start, end)?;
                    let bytes: Vec<u8> = units.iter().map(|&u| u as u8).collect();
                    // WTF-8 is strict like UTF-8 here (we can't hold surrogates);
                    // lossy replaces malformed sequences with U+FFFD.
                    let s = if op == Op::STRING_NEW_LOSSY_UTF8_ARRAY {
                        String::from_utf8_lossy(&bytes).into_owned()
                    } else {
                        String::from_utf8(bytes).map_err(|_| VMError::new("trap: invalid UTF-8"))?
                    };
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::STRING_ENCODE_WTF8_ARRAY
                    || op == Op::STRING_ENCODE_LOSSY_UTF8_ARRAY =>
                {
                    // (str, array, start): write the UTF-8 bytes into the array.
                    let start = self.pop().as_i32() as u32 as usize;
                    let arr = self.pop();
                    let s = self.pop_stringref()?;
                    let obj = match &arr {
                        Value::Object(o) => o.clone(),
                        _ => return Err(VMError::new("trap: null array reference")),
                    };
                    let units: Vec<u32> = s.as_bytes().iter().map(|&b| b as u32).collect();
                    let mut guard = obj.lock().unwrap();
                    let elems = match &mut guard.kind {
                        ObjectKind::Array(v) => v,
                        _ => return Err(VMError::new("trap: expected array reference")),
                    };
                    if start.saturating_add(units.len()) > elems.len() {
                        return Err(VMError::new("trap: array access out of bounds"));
                    }
                    for (i, u) in units.iter().enumerate() {
                        elems[start + i] = Value::I32(*u as i32);
                    }
                    self.push(Value::I32(units.len() as i32))?;
                }
                _ if op == Op::STRING_IS_USV_SEQUENCE => {
                    // A native Rust string is always a valid Unicode scalar-value
                    // sequence (no lone surrogates possible) → always 1 for a
                    // non-null string.
                    let _s = self.pop_stringref()?;
                    self.push(Value::I32(1))?;
                }

                // ── Stringview cursor ops (positions are explicit operands) ──
                // The view IS the string (string.as_wtf8/wtf16 return it). WTF-8
                // positions are byte offsets with the spec's "position treatment"
                // (snap forward to the next codepoint boundary, clamp to length);
                // WTF-16 positions are code-unit offsets, clamped to length.
                _ if op == Op::STRINGVIEW_WTF16_LENGTH => {
                    let s = self.pop_stringref()?;
                    self.push(Value::I32(s.encode_utf16().count() as i32))?;
                }
                _ if op == Op::STRINGVIEW_WTF16_GET_CODEUNIT => {
                    let pos = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let units: Vec<u16> = s.encode_utf16().collect();
                    let u = *units.get(pos).ok_or_else(|| {
                        VMError::new("trap: stringview_wtf16 index out of bounds")
                    })?;
                    self.push(Value::I32(u as i32))?;
                }
                _ if op == Op::STRINGVIEW_WTF16_SLICE => {
                    let end = self.pop().as_i32() as u32 as usize;
                    let start = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let units: Vec<u16> = s.encode_utf16().collect();
                    let a = start.min(units.len());
                    let b = end.min(units.len());
                    let out = if a < b {
                        String::from_utf16_lossy(&units[a..b])
                    } else {
                        String::new()
                    };
                    self.push(Value::String(Arc::from(out.as_str())))?;
                }
                _ if op == Op::STRINGVIEW_WTF16_ENCODE => {
                    // (view, ptr, pos, len) → code units written.
                    let len = self.pop().as_i32() as u32 as usize;
                    let pos = self.pop().as_i32() as u32 as usize;
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let units: Vec<u16> = s.encode_utf16().collect();
                    let start = pos.min(units.len());
                    let count = len.min(units.len() - start);
                    let mut bytes = Vec::with_capacity(count * 2);
                    for u in &units[start..start + count] {
                        bytes.extend_from_slice(&u.to_le_bytes());
                    }
                    self.write_memory_bytes(0, ptr, &bytes)?;
                    self.push(Value::I32(count as i32))?;
                }
                _ if op == Op::STRINGVIEW_WTF8_ADVANCE => {
                    // (view, pos, bytes) → next byte offset (highest codepoint
                    // boundary ≤ treated_pos + bytes).
                    let bytes = self.pop().as_i32() as u32 as usize;
                    let pos = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let start = wtf8_treat(&s, pos);
                    let mut target = start.saturating_add(bytes).min(s.len());
                    while target > start && !s.is_char_boundary(target) {
                        target -= 1;
                    }
                    self.push(Value::I32(target as i32))?;
                }
                _ if op == Op::STRINGVIEW_WTF8_SLICE => {
                    // (view, start, end) → substring over the treated byte range.
                    let end = self.pop().as_i32() as u32 as usize;
                    let start = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let a = wtf8_treat(&s, start);
                    let b = wtf8_treat(&s, end);
                    let out = if a < b {
                        s[a..b].to_string()
                    } else {
                        String::new()
                    };
                    self.push(Value::String(Arc::from(out.as_str())))?;
                }
                _ if op == Op::STRINGVIEW_WTF8_ENCODE_UTF8 => {
                    // (view, ptr, pos, bytes) → (next_pos, bytes_written). Writes
                    // whole codepoints only, never splitting one across the limit.
                    let max = self.pop().as_i32() as u32 as usize;
                    let pos = self.pop().as_i32() as u32 as usize;
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let start = wtf8_treat(&s, pos);
                    let mut end = start.saturating_add(max).min(s.len());
                    while end > start && !s.is_char_boundary(end) {
                        end -= 1;
                    }
                    let written = &s.as_bytes()[start..end];
                    self.write_memory_bytes(0, ptr, written)?;
                    // Two results: next position, then bytes written (spec order).
                    self.push(Value::I32(end as i32))?;
                    self.push(Value::I32((end - start) as i32))?;
                }

                // ── Codepoint iterators ──────────────────────────────────────
                // `string.as_iter` yields a cursor object: an ordinary object
                // carrying the string plus a codepoint index in its properties
                // (`__iter_str` / `__iter_pos`). The iter ops read/advance it.
                _ if op == Op::STRING_AS_ITER => {
                    let s = self.pop_stringref()?;
                    let mut obj = crate::value::Object::new();
                    obj.properties
                        .insert("__iter_str".to_string(), Value::String(s));
                    obj.properties
                        .insert("__iter_pos".to_string(), Value::I32(0));
                    self.push(Value::Object(Arc::new(std::sync::Mutex::new(obj))))?;
                }
                _ if op == Op::STRINGVIEW_ITER_NEXT => {
                    let view = self.pop();
                    let (s, pos) = self.read_string_iter(&view)?;
                    let total = s.chars().count();
                    if pos >= total {
                        self.push(Value::I32(-1))?;
                    } else {
                        let cp = s.chars().nth(pos).unwrap() as i32;
                        self.write_string_iter_pos(&view, pos + 1)?;
                        self.push(Value::I32(cp))?;
                    }
                }
                _ if op == Op::STRINGVIEW_ITER_ADVANCE => {
                    let n = self.pop().as_i32() as u32 as usize;
                    let view = self.pop();
                    let (s, pos) = self.read_string_iter(&view)?;
                    let total = s.chars().count();
                    let new_pos = pos.saturating_add(n).min(total);
                    self.write_string_iter_pos(&view, new_pos)?;
                    self.push(Value::I32((new_pos - pos) as i32))?;
                }
                _ if op == Op::STRINGVIEW_ITER_REWIND => {
                    let n = self.pop().as_i32() as u32 as usize;
                    let view = self.pop();
                    let (_s, pos) = self.read_string_iter(&view)?;
                    let new_pos = pos.saturating_sub(n);
                    self.write_string_iter_pos(&view, new_pos)?;
                    self.push(Value::I32((pos - new_pos) as i32))?;
                }
                _ if op == Op::STRINGVIEW_ITER_SLICE => {
                    // Substring of up to `n` codepoints from the cursor; does NOT
                    // advance the iterator.
                    let n = self.pop().as_i32() as u32 as usize;
                    let view = self.pop();
                    let (s, pos) = self.read_string_iter(&view)?;
                    let out: String = s.chars().skip(pos).take(n).collect();
                    self.push(Value::String(Arc::from(out.as_str())))?;
                }

                _ if op == Op::REF_IS_NULL => {
                    let v = self.pop();
                    self.push(Value::I32(
                        if v.is_null_ref() || matches!(v, Value::Undefined) {
                            1
                        } else {
                            0
                        },
                    ))?;
                }
                // -- Conversions --
                _ if op == Op::F64_FROM_I32 => {
                    let v = self.pop();
                    self.push(Value::F64(v.as_f64()))?;
                }
                _ if op == Op::F64_CONVERT_I32_U => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::F64(a as f64))?;
                }
                _ if op == Op::F64_CONVERT_I64_S => {
                    let a = self.pop().as_i64();
                    self.push(Value::F64(a as f64))?;
                }
                _ if op == Op::F64_CONVERT_I64_U => {
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::F64(a as f64))?;
                }
                _ if op == Op::F32_CONVERT_I32_S => {
                    let a = self.pop().as_i32();
                    self.push(Value::F32(a as f32))?;
                }
                _ if op == Op::F32_CONVERT_I32_U => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::F32(a as f32))?;
                }
                _ if op == Op::F32_CONVERT_I64_S => {
                    let a = self.pop().as_i64();
                    self.push(Value::F32(a as f32))?;
                }
                _ if op == Op::F32_CONVERT_I64_U => {
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::F32(a as f32))?;
                }
                _ if op == Op::I32_FROM_F64 => {
                    let v = self.pop().as_f64();
                    if v.is_nan() || v >= 2147483648.0 || v < -2147483648.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as i32))?;
                }
                _ if op == Op::I32_TRUNC_F64_U => {
                    let v = self.pop().as_f64();
                    if v.is_nan() || v < 0.0 || v >= 4294967296.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as u32 as i32))?;
                }
                _ if op == Op::I32_TRUNC_F32_S => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() || v >= 2147483648.0f32 || v < -2147483648.0f32 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as i32))?;
                }
                _ if op == Op::I32_TRUNC_F32_U => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() || v < 0.0f32 || v >= 4294967296.0f32 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as u32 as i32))?;
                }
                _ if op == Op::I64_TRUNC_F32_S => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() || v >= 9223372036854775808.0f32 || v < -9223372036854775808.0f32
                    {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(v as i64))?;
                }
                _ if op == Op::I64_TRUNC_F32_U => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() || v < 0.0f32 || v >= 18446744073709551616.0f32 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(v as u64 as i64))?;
                }

                // nontrapping-float-to-int-conversions proposal (0xFC 0x00–0x07).
                // Rust `as` casts saturate since 1.45: NaN → 0, overflow → min/max.
                _ if op == Op::I32_TRUNC_SAT_F32_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32(v as i32))?;
                }
                _ if op == Op::I32_TRUNC_SAT_F32_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32((v as u32) as i32))?;
                }
                _ if op == Op::I32_TRUNC_SAT_F64_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32(v as i32))?;
                }
                _ if op == Op::I32_TRUNC_SAT_F64_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32((v as u32) as i32))?;
                }
                _ if op == Op::I64_TRUNC_SAT_F32_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64(v as i64))?;
                }
                _ if op == Op::I64_TRUNC_SAT_F32_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64((v as u64) as i64))?;
                }
                _ if op == Op::I64_TRUNC_SAT_F64_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64(v as i64))?;
                }
                _ if op == Op::I64_TRUNC_SAT_F64_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64((v as u64) as i64))?;
                }

                // -- Async (await) --
                // r#await: removed (duplicate of promise_suspend, use JSPI proposal name)

                // -- Exceptions (WASM exception-handling proposal, final) --
                // Normal exit from a try block is handled by the structural
                // `end` (Op::END, `is_try` label) — see the END dispatch. The
                // custom TRY_END opcode has been retired.
                _ if op == Op::THROW => {
                    // Spec `throw <tagidx>`: the tag index immediate selects
                    // the tag entity; the payload (per the tag's signature
                    // arity) is popped off the stack.
                    let tag_idx = self.read_u16();
                    let chunk_index = self.frame().chunk_index;
                    let entity = self.resolve_chunk_tag(chunk_index, tag_idx)?;
                    let arity = self.tag_entities[entity].arity as usize;
                    let mut payload = Vec::with_capacity(arity);
                    for _ in 0..arity {
                        payload.push(self.pop());
                    }
                    payload.reverse();
                    self.raise_exception(entity, payload, 0)?;
                }
                _ if op == Op::THROW_REF => {
                    // Spec `throw_ref`: rethrow the exception an exnref
                    // refers to — same tag identity, same payload.
                    let val = self.pop();
                    let (entity, payload) = Self::unpack_exnref(&val)
                        .ok_or_else(|| VMError::new("throw_ref: operand is not an exnref"))?;
                    self.raise_exception(entity, payload, 0)?;
                }
                _ if op == Op::RETHROW => {
                    // Legacy EH rethrow — carries the exception object as a
                    // value; re-raises through the vybe:exception tag.
                    let chunk_idx = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let _depth = read_leb_u32(&self.chunks[chunk_idx].code, &mut ip);
                    self.frame_mut().ip = ip;
                    let val = self.pop();
                    self.raise_exception_value(val)?;
                }
                _ if op == Op::DELEGATE => {
                    let chunk_idx = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let depth = read_leb_u32(&self.chunks[chunk_idx].code, &mut ip);
                    self.frame_mut().ip = ip;
                    let val = self.pop();
                    self.raise_exception_value_skipping(val, depth as usize)?;
                }
                _ if op == Op::TRY_TABLE => {
                    // Spec try_table. Internal fixed-width encoding:
                    //   [try_table, u8 clause_count, per clause:
                    //    u8 kind (0=catch 1=catch_ref 2=catch_all 3=catch_all_ref),
                    //    u16 tag_idx (ignored for catch_all kinds),
                    //    u16 offset (forward from the end of this clause)]
                    // Matching is TAG IDENTITY only — clauses are tried in
                    // order (pushed reversed so the first clause is on top).
                    let clause_count = self.read_byte() as usize;
                    let chunk_index = self.frame().chunk_index;
                    self.try_group_counter += 1;
                    let group = self.try_group_counter;
                    let mut handlers = Vec::with_capacity(clause_count);
                    for _ in 0..clause_count {
                        let kind = self.read_byte();
                        let tag_idx = self.read_u16();
                        let offset = self.read_u16();
                        let ip = self.frame().ip + offset as usize;
                        let tag_entity = if kind == crate::vm::CATCH_KIND_CATCH
                            || kind == crate::vm::CATCH_KIND_CATCH_REF
                        {
                            self.resolve_chunk_tag(chunk_index, tag_idx)?
                        } else {
                            0 // unused for catch_all kinds
                        };
                        handlers.push(ExceptionHandler {
                            catch_ip: ip,
                            stack_depth: self.stack.len(),
                            frame_depth: self.frames.len(),
                            label_depth: self.label_stack.len(),
                            _chunk_index: chunk_index,
                            kind,
                            tag_entity,
                            group,
                        });
                    }
                    // Push in reverse so the FIRST clause is on top (spec:
                    // "catch clauses are tried in the order they appear").
                    for h in handlers.into_iter().rev() {
                        self.exception_handlers.push(h);
                    }
                    // try_table is a block: push its structural label so the
                    // matching `end` closes the handler scope. `is_try` tells
                    // END / branch_to_label to also pop the handler group. The
                    // handlers above recorded `label_depth` BEFORE this push,
                    // so it names the OUTER level — a thrown exception unwinds
                    // to there (catch code lives past this block's `end`).
                    let ci = self.frame().chunk_index;
                    self.ensure_block_table(ci);
                    let end_ip = self.block_tables[&ci]
                        .get(&opcode_start)
                        .map(|t| t.end_ip)
                        .unwrap_or(self.frame().ip);
                    self.label_stack.push(LabelEntry {
                        target: end_ip,
                        is_loop: false,
                        is_try: true,
                        result_arity: 0,
                        stack_height: self.stack.len(),
                    });
                }

                // -- Tail call --
                _ if op == Op::RETURN_CALL => {
                    let argc = self.read_byte() as usize;
                    // Reuse current frame: move callee + args down to base-1.
                    // Stack: [..., callee, arg0, ..., argN-1]
                    // After: [..., callee, arg0, ..., argN-1] starting at base-1
                    // call_value will pop the callee and set base = stack.len() - argc
                    let old_base = self.frame().base;
                    // Place callee + args starting at old_base - 0 (function frame had callee ABOVE old_base,
                    // but in new convention there's no callee in the frame). Just truncate down to old_base
                    // and keep callee + args on top.
                    let callee_idx = self.stack.len() - argc - 1;
                    // Copy callee + args down to start at old_base
                    for i in 0..=argc {
                        self.stack[old_base + i] = self.stack[callee_idx + i].clone();
                    }
                    self.stack.truncate(old_base + 1 + argc);
                    self.frames.pop();
                    self.call_value(argc)?;
                }
                _ if op == Op::RETURN_CALL_INDIRECT => {
                    // Tail-call form of `call_indirect`: same immediate shape
                    // (argc, tableidx, expected results), same `wasm_tables`
                    // lookup and runtime type-shape check — but reuses the
                    // current frame (spec tail call) so unbounded indirect tail
                    // recursion runs in O(1) stack.
                    let argc = self.read_byte() as usize;
                    let tableidx = self.read_byte() as usize;
                    let expected_results = self.read_byte() as usize;
                    // Spec layout: the i32 table index is on TOP, above the args.
                    let raw_idx = self.pop().as_f64();
                    let funcref = {
                        let table = self.table_ref(tableidx).ok_or_else(|| {
                            VMError::new("trap: return_call_indirect unknown table")
                        })?;
                        if raw_idx < 0.0 || raw_idx.is_nan() || raw_idx >= table.len() as f64 {
                            return Err(VMError::new(format!(
                                "trap: return_call_indirect: invalid table index {}",
                                raw_idx
                            )));
                        }
                        table[raw_idx as usize].clone()
                    };
                    if let Value::Object(o) = &funcref {
                        let ob = o.lock().unwrap();
                        if let crate::value::ObjectKind::Function(f) = &ob.kind {
                            let ch = &self.chunks[f.chunk_index];
                            if ch.param_count as usize != argc
                                || ch.result_arity as usize != expected_results
                            {
                                return Err(VMError::new(format!(
                                    "trap: return_call_indirect: signature mismatch \
                                     (callee {}→{}, expected {}→{})",
                                    ch.param_count, ch.result_arity, argc, expected_results
                                )));
                            }
                        }
                    }
                    // Splice the funcref in below the args, then reuse the frame.
                    let callee_idx = self.stack.len() - argc;
                    self.stack.insert(callee_idx, funcref);
                    let old_base = self.frame().base;
                    for i in 0..=argc {
                        self.stack[old_base + i] = self.stack[callee_idx + i].clone();
                    }
                    self.stack.truncate(old_base + 1 + argc);
                    self.frames.pop();
                    self.call_value(argc)?;
                }
                _ if op == Op::RETURN_CALL_REF => {
                    let argc = self.read_byte() as usize;
                    let old_base = self.frame().base;
                    let callee_idx = self.stack.len() - argc - 1;
                    for i in 0..=argc {
                        self.stack[old_base + i] = self.stack[callee_idx + i].clone();
                    }
                    self.stack.truncate(old_base + 1 + argc);
                    self.frames.pop();
                    self.call_value(argc)?;
                }

                // -- Linear memory --
                _ if op == Op::MEMORY_SIZE => {
                    let memidx = self.read_optional_memidx_immediate();
                    let pages = self.mem_len(memidx) / 65536;
                    // memory64: page count is i64; 32-bit memory: i32.
                    if self.mem_is_64(memidx) {
                        self.push(Value::I64(pages as i64))?;
                    } else {
                        self.push(Value::I32(pages as i32))?;
                    }
                }
                _ if op == Op::MEMORY_GROW => {
                    let memidx = self.read_optional_memidx_immediate();
                    let is64 = self.mem_is_64(memidx);
                    let pages = self.pop_mem_index(is64);
                    let old_pages = self.mem_grow(memidx, pages);
                    let failed = old_pages == usize::MAX;
                    if is64 {
                        self.push(Value::I64(if failed { -1 } else { old_pages as i64 }))?;
                    } else {
                        self.push(Value::I32(if failed { -1 } else { old_pages as i32 }))?;
                    }
                }
                _ if op == Op::I32_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    self.push(Value::I32(i32::from_le_bytes(read_le(&bytes))))?;
                }
                _ if op == Op::I32_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i32();
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                _ if op == Op::I64_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    self.push(Value::I64(i64::from_le_bytes(read_le(&bytes))))?;
                }
                _ if op == Op::I64_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64();
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                _ if op == Op::F64_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    self.push(Value::F64(f64::from_le_bytes(read_le(&bytes))))?;
                }
                _ if op == Op::F64_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_f64();
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                _ if op == Op::I32_LOAD8_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I32(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i32,
                    ))?;
                }
                _ if op == Op::I32_STORE8 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i32() as u8;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &[val])?;
                }
                _ if op == Op::F32_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    let val = f32::from_le_bytes(read_le(&bytes));
                    self.push(Value::F32(val))?;
                }
                _ if op == Op::F32_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_f64() as f32;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                _ if op == Op::I32_LOAD8_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I32(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i8 as i32,
                    ))?;
                }
                _ if op == Op::I32_LOAD16_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = i16::from_le_bytes(read_le(&bytes)) as i32;
                    self.push(Value::I32(val))?;
                }
                _ if op == Op::I32_LOAD16_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = u16::from_le_bytes(read_le(&bytes)) as i32;
                    self.push(Value::I32(val))?;
                }
                _ if op == Op::I32_STORE16 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i32() as i16;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                _ if op == Op::I64_LOAD8_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I64(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i8 as i64,
                    ))?;
                }
                _ if op == Op::I64_LOAD8_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I64(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i64,
                    ))?;
                }
                _ if op == Op::I64_LOAD16_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = i16::from_le_bytes(read_le(&bytes)) as i64;
                    self.push(Value::I64(val))?;
                }
                _ if op == Op::I64_LOAD16_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = u16::from_le_bytes(read_le(&bytes)) as i64;
                    self.push(Value::I64(val))?;
                }
                _ if op == Op::I64_LOAD32_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    self.push(Value::I64(i32::from_le_bytes(read_le(&bytes)) as i64))?;
                }
                _ if op == Op::I64_LOAD32_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    self.push(Value::I64(i32::from_le_bytes(read_le(&bytes)) as u32 as i64))?;
                }
                _ if op == Op::I64_STORE8 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64() as u8;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &[val])?;
                }
                _ if op == Op::I64_STORE16 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64() as i16;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                _ if op == Op::I64_STORE32 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64() as i32;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }

                // -- Conversions --
                _ if op == Op::I32_WRAP_I64 => {
                    let a = self.pop().as_i64();
                    self.push(Value::I32(a as i32))?;
                }
                _ if op == Op::I64_EXTEND_I32_S => {
                    let a = self.pop().as_i32();
                    self.push(Value::I64(a as i64))?;
                }
                _ if op == Op::I64_EXTEND_I32_U => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I64(a as i64))?;
                }
                _ if op == Op::I64_TRUNC_F64_S => {
                    let a = self.pop().as_f64();
                    if a.is_nan() || a >= 9223372036854775808.0 || a < -9223372036854775808.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(a as i64))?;
                }
                _ if op == Op::I64_TRUNC_F64_U => {
                    let a = self.pop().as_f64();
                    if a.is_nan() || a < 0.0 || a >= 18446744073709551616.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(a as u64 as i64))?;
                }
                _ if op == Op::F64_PROMOTE_F32 => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a))?;
                }
                _ if op == Op::F32_DEMOTE_F64 => {
                    let a = self.pop().as_f64();
                    self.push(Value::F32(a as f32))?;
                }
                _ if op == Op::I32_REINTERPRET_F32 => {
                    let a = self.pop().as_f64() as f32;
                    self.push(Value::I32(a.to_bits() as i32))?;
                }
                _ if op == Op::I64_REINTERPRET_F64 => {
                    let a = self.pop().as_f64();
                    self.push(Value::I64(a.to_bits() as i64))?;
                }
                _ if op == Op::F32_REINTERPRET_I32 => {
                    let a = self.pop().as_i32();
                    self.push(Value::F32(f32::from_bits(a as u32)))?;
                }
                _ if op == Op::F64_REINTERPRET_I64 => {
                    let a = self.pop().as_i64();
                    self.push(Value::F64(f64::from_bits(a as u64)))?;
                }

                // -- Sign extension --
                _ if op == Op::I32_EXTEND8_S => {
                    let a = self.pop().as_i32() as i8;
                    self.push(Value::I32(a as i32))?;
                }
                _ if op == Op::I32_EXTEND16_S => {
                    let a = self.pop().as_i32() as i16;
                    self.push(Value::I32(a as i32))?;
                }
                _ if op == Op::I64_EXTEND8_S => {
                    let a = self.pop().as_i64() as i8;
                    self.push(Value::I64(a as i64))?;
                }
                _ if op == Op::I64_EXTEND16_S => {
                    let a = self.pop().as_i64() as i16;
                    self.push(Value::I64(a as i64))?;
                }
                _ if op == Op::I64_EXTEND32_S => {
                    let a = self.pop().as_i64() as i32;
                    self.push(Value::I64(a as i64))?;
                }

                // -- Multi-value --
                // pack, unpack: removed (non-WASM, were unused by compilers)

                // -- Block/loop/if structured control (WASM-compliant) --
                _ if op == Op::BLOCK => {
                    let result_arity = self.read_byte(); // 0=void, 1=single, 2+=multi-value
                    let ci = self.frame().chunk_index;
                    self.ensure_block_table(ci);
                    let end_ip = self.block_tables[&ci]
                        .get(&opcode_start)
                        .map(|t| t.end_ip)
                        .unwrap_or(self.frame().ip);
                    self.label_stack.push(LabelEntry {
                        target: end_ip,
                        is_loop: false,
                        is_try: false,
                        result_arity,
                        stack_height: self.stack.len(),
                    });
                }
                _ if op == Op::LOOP => {
                    let result_arity = self.read_byte();
                    // Loop target is the ip right after the blocktype byte —
                    // that is where `br 0` restarts (the loop body start).
                    let loop_body_start = self.frame().ip;
                    self.label_stack.push(LabelEntry {
                        target: loop_body_start,
                        is_loop: true,
                        is_try: false,
                        result_arity,
                        stack_height: self.stack.len(),
                    });
                }
                _ if op == Op::IF => {
                    let result_arity = self.read_byte();
                    let ci = self.frame().chunk_index;
                    self.ensure_block_table(ci);
                    let targets = self.block_tables[&ci]
                        .get(&opcode_start)
                        .copied()
                        .unwrap_or(BlockTargets {
                            else_ip: None,
                            end_ip: self.frame().ip,
                        });
                    // WASM `if` consumes an i32 condition (spec §4.4.1).
                    // The compiler must lower dynamic truthiness to i32 before this opcode.
                    // We also accept Bool (from VM-internal ops like REF_IS_*, emit_dyn_eq etc.)
                    // to be ECMA-runtime compliant without requiring a separate coercion step.
                    let cond = match self.pop() {
                        crate::value::Value::I32(n) => n,
                        crate::value::Value::Bool(b) => b as i32,
                        other => {
                            return Err(VMError::new(format!(
                                "type mismatch: if expected i32 condition, got {}",
                                other.tag().name()
                            )));
                        }
                    };
                    if cond != 0 {
                        // Condition true — push label and fall through to then-body.
                        // END pops the label after sequential body execution.
                        self.label_stack.push(LabelEntry {
                            target: targets.end_ip,
                            is_loop: false,
                            is_try: false,
                            result_arity,
                            stack_height: self.stack.len(),
                        });
                    } else if let Some(else_ip) = targets.else_ip {
                        // Condition false, ELSE exists — push label and jump into else-body.
                        // The else-body ends at END which pops the label.
                        // (We jump past the ELSE opcode itself to reach the else-body start.)
                        self.label_stack.push(LabelEntry {
                            target: targets.end_ip,
                            is_loop: false,
                            is_try: false,
                            result_arity,
                            stack_height: self.stack.len(),
                        });
                        self.frame_mut().ip = else_ip + 4; // +4 skips the ELSE opcode bytes
                    } else {
                        // Condition false, no ELSE — skip the block entirely.
                        // No label push needed; jump past END directly.
                        self.frame_mut().ip = targets.end_ip;
                    }
                }
                _ if op == Op::ELSE => {
                    // Then-body completed normally. ELSE behaves like `br 0` for the IF
                    // block: pop the IF label and jump past the else-body (to past END).
                    // The else-body is thus never executed on the then-path.
                    self.label_stack.pop();
                    let ci = self.frame().chunk_index;
                    self.ensure_block_table(ci);
                    let end_ip = self.block_tables[&ci]
                        .get(&opcode_start)
                        .map(|t| t.end_ip)
                        .unwrap_or(self.frame().ip);
                    self.frame_mut().ip = end_ip;
                }
                _ if op == Op::END => {
                    // Closing a block. If it is a try_table block (`is_try`),
                    // also remove its exception-handler group — the structural
                    // end of the protected region, replacing the retired
                    // TRY_END opcode. Reached ONLY on normal completion: a
                    // caught exception jumps past this `end` to the catch code
                    // (raise_exception already truncated the group), and a `br`
                    // out is handled by branch_to_label — so the group is
                    // popped exactly once, on whichever path fires.
                    if let Some(label) = self.label_stack.pop() {
                        if label.is_try {
                            if let Some(top) = self.exception_handlers.last() {
                                let group = top.group;
                                while self
                                    .exception_handlers
                                    .last()
                                    .is_some_and(|h| h.group == group)
                                {
                                    self.exception_handlers.pop();
                                }
                            }
                        }
                    }
                }
                _ if op == Op::BR_TABLE => {
                    let ci = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let count = read_leb_u32(&self.chunks[ci].code, &mut ip) as usize;
                    let mut labels = Vec::with_capacity(count);
                    for _ in 0..count {
                        labels.push(read_leb_u32(&self.chunks[ci].code, &mut ip) as usize);
                    }
                    let default_depth = read_leb_u32(&self.chunks[ci].code, &mut ip) as usize;
                    self.frame_mut().ip = ip;
                    let idx = self.pop().as_i32() as usize;
                    let depth = if idx < count {
                        labels[idx]
                    } else {
                        default_depth
                    };
                    if let Some(entry) = self.label_stack.iter().rev().nth(depth).copied() {
                        self.branch_to_label(depth, entry);
                    }
                }

                // -- call_indirect --
                _ if op == Op::CALL_INDIRECT => {
                    let argc = self.read_byte() as usize;
                    let tableidx = self.read_byte() as usize;
                    let expected_results = self.read_byte() as usize;
                    // Spec `call_indirect`: `[t* i32] → [t'*]` — the i32 table
                    // index is on TOP of the stack, above the `argc` call
                    // arguments. Pop it, resolve the funcref, then splice the
                    // funcref in below the args so `call_value` sees
                    // `[funcref, args…]`.
                    let raw_idx = self.pop().as_f64();
                    let funcref = {
                        let table = self
                            .table_ref(tableidx)
                            .ok_or_else(|| VMError::new("trap: call_indirect unknown table"))?;
                        if raw_idx < 0.0 || raw_idx.is_nan() || raw_idx >= table.len() as f64 {
                            return Err(VMError::new(format!(
                                "trap: call_indirect: invalid table index {}",
                                raw_idx
                            )));
                        }
                        table[raw_idx as usize].clone()
                    };
                    // Spec runtime type check: the funcref's declared type shape
                    // (params → results) must match the call's static `(type
                    // $sig)`. The VM is untyped, so equality is over the
                    // param/result COUNTS carried on the callee's chunk.
                    if let Value::Object(o) = &funcref {
                        let ob = o.lock().unwrap();
                        if let crate::value::ObjectKind::Function(f) = &ob.kind {
                            let ch = &self.chunks[f.chunk_index];
                            if ch.param_count as usize != argc
                                || ch.result_arity as usize != expected_results
                            {
                                return Err(VMError::new(format!(
                                    "trap: call_indirect: signature mismatch \
                                     (callee {}→{}, expected {}→{})",
                                    ch.param_count, ch.result_arity, argc, expected_results
                                )));
                            }
                        }
                    }
                    let insert_pos = self.stack.len() - argc;
                    self.stack.insert(insert_pos, funcref);
                    self.call_value(argc)?;
                }

                // -- Component Model --
                _ if op == Op::CANON_LIFT => {
                    let type_idx = self.read_u16() as usize;
                    // Lift: convert core value to component interface type.
                    // For now: if value is an object, stamp its type_id.
                    // In full CM, this would validate/convert the value shape.
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let mut o = obj.lock().unwrap();
                        if o.type_id == 0 && type_idx < self.type_registry.types.len() {
                            o.type_id = type_idx;
                        }
                    }
                    self.push(val)?;
                }
                _ if op == Op::CANON_LOWER => {
                    let type_idx = self.read_u16() as usize;
                    // Lower: convert component interface type to core value.
                    // For now: validate type_id matches, strip interface metadata.
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        if type_idx < self.type_registry.types.len() && o.type_id != type_idx {
                            // Type mismatch — could trap, for now allow
                        }
                    }
                    self.push(val)?;
                }

                // -- Shared-Everything Threads (shared GC objects) --

                // -- Weak References & Finalizers --

                // -- Multi-Memory --
                _ if op == Op::MEMORY_INIT => {
                    let data_idx = self.read_byte() as u32;
                    let memidx = self.read_optional_memidx_immediate() as usize;
                    if self.dropped_data.contains(&data_idx) {
                        return Err(VMError::new("memory.init: data segment dropped"));
                    }
                    let count = self.pop().as_i32().max(0) as usize;
                    let src = self.pop().as_i32().max(0) as usize;
                    let dst = self.pop().as_i32().max(0) as usize;
                    if count == 0 {
                        continue;
                    }
                    let data = self
                        .data_segments
                        .get(data_idx as usize)
                        .ok_or_else(|| VMError::new("memory.init: missing data segment"))?;
                    if src.saturating_add(count) > data.len() {
                        return Err(VMError::new("trap: memory.init source out of bounds"));
                    }
                    let bytes = data[src..src + count].to_vec();
                    self.write_memory_bytes(memidx, dst, &bytes)?;
                }
                // ── reference-types: table operations ─────────────────
                // Each op reads a `u8 table_idx` operand per spec. Tables
                // route through `table_ref`/`table_mut` so the multi-table
                // proposal works: index 0 maps to `func_table`, indexes
                // indexed directly in `wasm_tables`.
                _ if op == Op::TABLE_SIZE => {
                    let tidx = self.read_byte() as usize;
                    let size = self
                        .table_ref(tidx)
                        .ok_or_else(|| VMError::new("trap: table.size unknown table"))?
                        .len();
                    // table64: size is i64; 32-bit table: i32.
                    if self.tbl_is_64(tidx) {
                        self.push(Value::I64(size as i64))?;
                    } else {
                        self.push(Value::I32(size as i32))?;
                    }
                }
                _ if op == Op::TABLE_GROW => {
                    let tidx = self.read_byte() as usize;
                    let is64 = self.tbl_is_64(tidx);
                    let delta = self.pop_table_count(is64)?;
                    let init = self.pop();
                    // WASM spec: growing past the declared max fails, returning -1
                    // (as the index type) without resizing.
                    let max = self.wasm_table_maxes.get(tidx).copied().flatten();
                    let table = self
                        .table_mut(tidx)
                        .ok_or_else(|| VMError::new("trap: table.grow unknown table"))?;
                    let old_size = table.len();
                    let new_size = old_size.saturating_add(delta);
                    let exceeds_max = max.is_some_and(|m| new_size > m);
                    if exceeds_max {
                        if is64 {
                            self.push(Value::I64(-1))?;
                        } else {
                            self.push(Value::I32(-1))?;
                        }
                    } else {
                        table.resize(new_size, init);
                        if is64 {
                            self.push(Value::I64(old_size as i64))?;
                        } else {
                            self.push(Value::I32(old_size as i32))?;
                        }
                    }
                }
                _ if op == Op::TABLE_FILL => {
                    let tidx = self.read_byte() as usize;
                    let is64 = self.tbl_is_64(tidx);
                    let count = self.pop_table_count(is64)?;
                    let value = self.pop();
                    let dst = self.pop_table_count(is64)?;
                    let table = self
                        .table_mut(tidx)
                        .ok_or_else(|| VMError::new("trap: table.fill unknown table"))?;
                    let end = dst.saturating_add(count);
                    if end > table.len() {
                        return Err(VMError::new("trap: table.fill out of bounds"));
                    }
                    for i in dst..end {
                        table[i] = value.clone();
                    }
                }
                _ if op == Op::TABLE_COPY => {
                    let dst_table_idx = self.read_byte() as usize;
                    let src_table_idx = self.read_byte() as usize;
                    // table64: operands are i64 if either table is 64-bit.
                    let is64 = self.tbl_is_64(dst_table_idx) || self.tbl_is_64(src_table_idx);
                    let count = self.pop_table_count(is64)?;
                    let src = self.pop_table_count(is64)?;
                    let dst = self.pop_table_count(is64)?;
                    let source = self
                        .table_ref(src_table_idx)
                        .ok_or_else(|| VMError::new("trap: table.copy unknown table"))?;
                    if src.saturating_add(count) > source.len() {
                        return Err(VMError::new("trap: table.copy out of bounds".to_string()));
                    }
                    let values: Vec<Value> = source[src..src + count].to_vec();
                    let destination = self
                        .table_mut(dst_table_idx)
                        .ok_or_else(|| VMError::new("trap: table.copy unknown table"))?;
                    if dst.saturating_add(count) > destination.len() {
                        return Err(VMError::new("trap: table.copy out of bounds".to_string()));
                    }
                    destination[dst..dst + count].clone_from_slice(&values);
                }
                _ if op == Op::TABLE_INIT => {
                    let elem_idx = self.read_byte() as u32;
                    let table_idx = self.read_byte() as usize;
                    if self.dropped_elems.contains(&elem_idx) {
                        return Err(VMError::new("table.init: element segment dropped"));
                    }
                    let is64 = self.tbl_is_64(table_idx);
                    let count = self.pop_table_count(is64)?;
                    let src = self.pop_table_count(is64)?;
                    let dst = self.pop_table_count(is64)?;
                    let elems = self
                        .elem_segments
                        .get(elem_idx as usize)
                        .ok_or_else(|| VMError::new("table.init: missing element segment"))?;
                    if src.saturating_add(count) > elems.len() {
                        return Err(VMError::new("trap: table.init source out of bounds"));
                    }
                    let values: Vec<Value> = elems[src..src + count].to_vec();
                    let table = self
                        .table_mut(table_idx)
                        .ok_or_else(|| VMError::new("trap: table.init unknown table"))?;
                    if dst.saturating_add(count) > table.len() {
                        return Err(VMError::new("trap: table.init destination out of bounds"));
                    }
                    table[dst..dst + count].clone_from_slice(&values);
                }
                _ if op == Op::ELEM_DROP => {
                    let elem_idx = self.read_byte() as u32;
                    self.dropped_elems.insert(elem_idx);
                }
                _ if op == Op::DATA_DROP => {
                    let data_idx = self.read_byte() as u32;
                    self.dropped_data.insert(data_idx);
                }
                _ if op == Op::MEMORY_COPY => {
                    let dst_mem = self.read_optional_memidx_immediate();
                    let src_mem = self.read_optional_memidx_immediate();
                    // memory64: operands are i64 if either memory is 64-bit.
                    let is64 = self.mem_is_64(dst_mem) || self.mem_is_64(src_mem);
                    let count = self.pop_mem_index(is64);
                    let src = self.pop_mem_index(is64);
                    let dst = self.pop_mem_index(is64);
                    let buf = self.read_memory_bytes(src_mem, src, count)?;
                    self.write_memory_bytes(dst_mem, dst, &buf)?;
                }
                _ if op == Op::MEMORY_FILL => {
                    let memidx = self.read_optional_memidx_immediate();
                    let is64 = self.mem_is_64(memidx);
                    let count = self.pop_mem_index(is64);
                    let val = self.pop().as_i32() as u8;
                    let dst = self.pop_mem_index(is64);
                    let buf = vec![val; count];
                    self.write_memory_bytes(memidx, dst, &buf)?;
                }

                // Type discrimination opcodes

                // -- Array builtins --
                _ if op == Op::ARRAY_LENGTH => {
                    let arr = self.pop();
                    let len = if let Value::Object(obj) = &arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            a.len() as i32
                        } else {
                            0
                        }
                    } else if let Value::String(s) = &arr {
                        s.chars().count() as i32
                    } else {
                        0
                    };
                    self.push(Value::I32(len))?;
                }
                // REMOVED (Phase E): the 9 non-spec `0xFF` ARRAY_* dispatch
                // arms — ARRAY_PUSH / POP / SLICE / JOIN / REVERSE /
                // CONTAINS / INDEX_OF (here) and ARRAY_CONCAT / SHIFT
                // (below). All callers migrated to `ecma:array.*`
                // imports (Vybe handlers in vybe_host, native on v8 via
                // JS glue, polyfill on wasmtime). Any bytecode still
                // carrying these opcode bytes will hit the "unknown op"
                // path at `Op::decode`'s None branch and trap cleanly.

                // WASM GC array ops
                _ if op == Op::ARRAY_NEW_DEFAULT => {
                    let len = self.pop().as_i32().max(0) as usize;
                    let elems = vec![Value::Null; len];
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(
                        elems,
                    )))))?;
                }
                _ if op == Op::ARRAY_FILL => {
                    // Spec `array.fill $t`: stack `[arrayref, index, value, count]`,
                    // so popping off the top yields count, value, index, arrayref.
                    let count = self.pop().as_i32().max(0) as usize;
                    let val = self.pop();
                    let start = self.pop().as_i32().max(0) as usize;
                    let arr = self.pop();
                    if matches!(arr, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: array.fill on null reference"));
                    }
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind {
                            let end = (start + count).min(a.len());
                            for i in start..end {
                                a[i] = val.clone();
                            }
                        }
                    }
                }
                _ if op == Op::ARRAY_COPY => {
                    let len = self.pop().as_i32().max(0) as usize;
                    let src_off = self.pop().as_i32().max(0) as usize;
                    let src = self.pop();
                    let dst_off = self.pop().as_i32().max(0) as usize;
                    let dst = self.pop();
                    if matches!(src, Value::TypedNull(_)) || matches!(dst, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: array.copy on null reference"));
                    }
                    // WASM GC `array.copy` traps when the copy region is out of
                    // bounds of a stamped GC array; dynamic-language arrays stay
                    // lenient (clamped). A null src/dst carries no rtt, so its
                    // trap is guarded compiler-side.
                    let src_is_gc = matches!(&src, Value::Object(o) if self.is_gc_array_obj(o));
                    let dst_is_gc = matches!(&dst, Value::Object(o) if self.is_gc_array_obj(o));
                    if src_is_gc || dst_is_gc {
                        let arr_len = |v: &Value| -> usize {
                            match v {
                                Value::Object(o) => match &o.lock().unwrap().kind {
                                    ObjectKind::Array(a) => a.len(),
                                    _ => 0,
                                },
                                _ => 0,
                            }
                        };
                        if src_off + len > arr_len(&src) || dst_off + len > arr_len(&dst) {
                            return Err(VMError::new("trap: array.copy out of bounds"));
                        }
                    }
                    // Read source slice
                    let src_vals: Vec<Value> = if let Value::Object(obj) = &src {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            let end = (src_off + len).min(a.len());
                            a[src_off.min(a.len())..end].to_vec()
                        } else {
                            vec![]
                        }
                    } else {
                        vec![]
                    };
                    // Write to destination
                    if let Value::Object(obj) = &dst {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind {
                            for (i, v) in src_vals.into_iter().enumerate() {
                                let idx = dst_off + i;
                                if idx < a.len() {
                                    a[idx] = v;
                                }
                            }
                        }
                    }
                }
                // ARRAY_CONCAT and ARRAY_SHIFT removed with the cluster
                // above — both were the non-spec `0xFF` variants. Use
                // `ecma:array.concat` / `ecma:array.shift` instead.

                // ── Stack-switching proposal — real coroutine semantics ──
                // Each continuation is an `ObjectKind::Continuation` carrying
                // its entry function plus an optional captured `Fiber` from
                // the last suspend. The active-continuation stack
                // (`self.active_continuations`) records which cont owns the
                // current execution — suspend reads the topmost entry to
                // decide where to stash the fresh fiber.
                _ if op == Op::CONT_NEW => {
                    let func_val = self.pop();
                    if func_val.is_null_ref() {
                        return Err(VMError::new("cont.new: null function reference"));
                    }
                    let state = crate::value::ContinuationState {
                        entry: func_val,
                        saved: std::sync::Mutex::new(None),
                        state: std::sync::Mutex::new(crate::value::ContinuationPhase::Ready),
                    };
                    let obj = Object {
                        properties: HashMap::new(),
                        kind: ObjectKind::Continuation(state),
                        type_id: 0,
                        fields: Vec::new(),
                    };
                    let mut obj = obj;
                    let entry_async = match &obj.kind {
                        ObjectKind::Continuation(cs) => {
                            crate::calls::continuation_entry_is_async(&self.chunks, &cs.entry)
                        }
                        _ => false,
                    };
                    crate::calls::attach_continuation_protocols(
                        &mut obj.properties,
                        &self.globals,
                        entry_async,
                    );
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                _ if op == Op::SUSPEND => {
                    let tag = self.read_u16();
                    let val = self.pop();
                    // Yield a value from the innermost active continuation.
                    // We save the current VM state as a `Fiber`, stash it
                    // into the continuation's saved slot, restore the
                    // caller's pre-RESUME state, then push the yielded
                    // value onto the caller's stack.
                    match self.active_continuations.pop() {
                        Some(ActiveContinuation {
                            cont,
                            caller_fiber,
                            mode,
                            handlers,
                        }) => {
                            let fiber = self.save_fiber();
                            if let Value::Object(ref obj) = cont {
                                let o = obj.lock().unwrap();
                                if let ObjectKind::Continuation(cs) = &o.kind {
                                    *cs.saved.lock().unwrap() = Some(fiber);
                                    *cs.state.lock().unwrap() =
                                        crate::value::ContinuationPhase::Suspended;
                                }
                            }
                            // Restore caller with the yielded value. For
                            // iterator-mode callers, append `has_more=1`
                            // so a GEN_NEXT-driven loop can check without
                            // a second API call.
                            self.resume_fiber_with(caller_fiber, Some(val))?;
                            let handled = handlers
                                .iter()
                                .find(|h| h.kind == 0 && h.tag_index == tag as u32)
                                .map(|h| h.label_index as usize);
                            if let Some(handler_ip) = handled {
                                self.frame_mut().ip = handler_ip;
                            } else if mode == ResumeMode::Iterator {
                                self.push(Value::I32(1))?;
                            }
                        }
                        None => {
                            // No active cont — legacy behaviour: return the
                            // yielded value from the current frame.
                            return Ok(val);
                        }
                    }
                }
                _ if op == Op::RESUME => {
                    let _tag = self.read_u16();
                    let resume_handlers = self.chunks[self.frame().chunk_index]
                        .stack_switch_handlers
                        .get(&opcode_start)
                        .cloned()
                        .unwrap_or_default();
                    let resume_val = self.pop();
                    let cont = self.pop();
                    if let Value::Object(ref obj) = cont {
                        let (phase, entry) = {
                            let o = obj.lock().unwrap();
                            if let ObjectKind::Continuation(cs) = &o.kind {
                                let phase = *cs.state.lock().unwrap();
                                let entry = cs.entry.clone();
                                (phase, entry)
                            } else {
                                return Err(VMError::new("resume: not a continuation"));
                            }
                        };
                        // Capture the caller's state so SUSPEND can restore
                        // us here when the coroutine yields. Important:
                        // push the active-continuation entry AFTER body
                        // state is in place — `resume_fiber_with`
                        // overwrites `active_continuations` with the
                        // saved fiber's copy, and we want our new entry
                        // on top of that.
                        let caller_fiber = self.save_fiber();
                        match phase {
                            crate::value::ContinuationPhase::Ready => {
                                let bound: Vec<Value> = {
                                    let o = obj.lock().unwrap();
                                    match o.properties.get("__bound_args") {
                                        Some(Value::Object(arr)) => {
                                            let a = arr.lock().unwrap();
                                            if let ObjectKind::Array(v) = &a.kind {
                                                v.clone()
                                            } else {
                                                Vec::new()
                                            }
                                        }
                                        _ => Vec::new(),
                                    }
                                };
                                let argc = bound.len() + 1;
                                self.push(entry)?;
                                for b in bound {
                                    self.push(b)?;
                                }
                                self.push(resume_val)?;
                                self.call_value_direct(argc)?;
                                // Fresh continuation runs on its own fiber (see
                                // GEN_NEXT) so internal returns don't trip an
                                // enclosing callback's stale execute_until floor.
                                self.cur_fiber_id = self.next_fiber_id;
                                self.next_fiber_id += 1;
                            }
                            crate::value::ContinuationPhase::Suspended => {
                                let saved = {
                                    let o = obj.lock().unwrap();
                                    if let ObjectKind::Continuation(cs) = &o.kind {
                                        cs.saved.lock().unwrap().take()
                                    } else {
                                        None
                                    }
                                };
                                if let Some(fiber) = saved {
                                    self.resume_fiber_with(fiber, Some(resume_val))?;
                                }
                            }
                            crate::value::ContinuationPhase::Done => {
                                // Resuming a completed coroutine traps per spec.
                                return Err(VMError::new("trap: resume on completed continuation"));
                            }
                        }
                        // Body state is now live; push the AC on top so
                        // SUSPEND finds it.
                        self.active_continuations.push(ActiveContinuation {
                            cont: cont.clone(),
                            caller_fiber,
                            mode: ResumeMode::Raw,
                            handlers: resume_handlers,
                        });
                    } else {
                        return Err(VMError::new("resume: not a continuation"));
                    }
                }
                _ if op == Op::SWITCH => {
                    let tag = self.read_u16();
                    // Symmetric swap: suspend the current cont (top of the
                    // active stack) and resume the target cont in one step.
                    let val = self.pop();
                    let target = self.pop();
                    let Some(current) = self.active_continuations.pop() else {
                        return Err(VMError::new("switch: no active continuation handler"));
                    };
                    if !current
                        .handlers
                        .iter()
                        .any(|h| h.kind == 1 && h.tag_index == tag as u32)
                    {
                        self.active_continuations.push(current);
                        return Err(VMError::new("switch: no matching continuation handler"));
                    }
                    let (phase, entry) = if let Value::Object(ref obj) = target {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Continuation(cs) = &o.kind {
                            (*cs.state.lock().unwrap(), cs.entry.clone())
                        } else {
                            self.active_continuations.push(current);
                            return Err(VMError::new("switch: not a continuation"));
                        }
                    } else {
                        self.active_continuations.push(current);
                        return Err(VMError::new("switch: not a continuation"));
                    };
                    if matches!(phase, crate::value::ContinuationPhase::Done) {
                        self.active_continuations.push(current);
                        return Err(VMError::new("trap: switch to completed continuation"));
                    }
                    let fiber = self.save_fiber();
                    if let Value::Object(ref obj) = current.cont {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Continuation(cs) = &o.kind {
                            *cs.saved.lock().unwrap() = Some(fiber);
                            *cs.state.lock().unwrap() = crate::value::ContinuationPhase::Suspended;
                        }
                    }
                    match phase {
                        crate::value::ContinuationPhase::Ready => {
                            let bound: Vec<Value> = {
                                let o = match &target {
                                    Value::Object(obj) => obj.lock().unwrap(),
                                    _ => unreachable!(),
                                };
                                match o.properties.get("__bound_args") {
                                    Some(Value::Object(arr)) => {
                                        let a = arr.lock().unwrap();
                                        if let ObjectKind::Array(v) = &a.kind {
                                            v.clone()
                                        } else {
                                            Vec::new()
                                        }
                                    }
                                    _ => Vec::new(),
                                }
                            };
                            let argc = bound.len() + 1;
                            self.push(entry)?;
                            for b in bound {
                                self.push(b)?;
                            }
                            self.push(val)?;
                            self.call_value_direct(argc)?;
                        }
                        crate::value::ContinuationPhase::Suspended => {
                            let saved = {
                                let o = match &target {
                                    Value::Object(obj) => obj.lock().unwrap(),
                                    _ => unreachable!(),
                                };
                                if let ObjectKind::Continuation(cs) = &o.kind {
                                    cs.saved.lock().unwrap().take()
                                } else {
                                    None
                                }
                            };
                            if let Some(fiber) = saved {
                                self.resume_fiber_with(fiber, Some(val))?;
                            }
                        }
                        crate::value::ContinuationPhase::Done => unreachable!(),
                    }
                    self.active_continuations.push(ActiveContinuation {
                        cont: target.clone(),
                        caller_fiber: current.caller_fiber,
                        mode: current.mode,
                        handlers: current.handlers,
                    });
                }
                // `cont.bind argc` — partially apply `argc` args to a
                // continuation. Stack: [cont, arg0, ..., arg(argc-1)] →
                // [cont'] where cont' is a fresh continuation that will
                // receive those args on first resume.
                _ if op == Op::CONT_BIND => {
                    let argc = self.read_byte() as usize;
                    let mut args: Vec<Value> = Vec::with_capacity(argc);
                    for _ in 0..argc {
                        args.push(self.pop());
                    }
                    args.reverse();
                    let cont_val = self.pop();
                    let new_cont = if let Value::Object(ref obj) = cont_val {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Continuation(cs) = &o.kind {
                            if matches!(
                                *cs.state.lock().unwrap(),
                                crate::value::ContinuationPhase::Done
                            ) {
                                return Err(VMError::new(
                                    "cont.bind: continuation already consumed",
                                ));
                            }
                            let entry = cs.entry.clone();
                            *cs.saved.lock().unwrap() = None;
                            *cs.state.lock().unwrap() = crate::value::ContinuationPhase::Done;
                            // Build a shim entry function: when the new
                            // cont is resumed, it calls the original entry
                            // with the bound args prefixed. Since our
                            // Value model can't carry a closure tuple
                            // directly, we stash the bound args in the
                            // continuation's properties and apply them on
                            // first resume.
                            let mut new_obj = Object {
                                properties: HashMap::new(),
                                kind: ObjectKind::Continuation(crate::value::ContinuationState {
                                    entry,
                                    saved: std::sync::Mutex::new(None),
                                    state: std::sync::Mutex::new(
                                        crate::value::ContinuationPhase::Ready,
                                    ),
                                }),
                                type_id: 0,
                                fields: Vec::new(),
                            };
                            let entry_async = match &new_obj.kind {
                                ObjectKind::Continuation(cs) => {
                                    crate::calls::continuation_entry_is_async(
                                        &self.chunks,
                                        &cs.entry,
                                    )
                                }
                                _ => false,
                            };
                            crate::calls::attach_continuation_protocols(
                                &mut new_obj.properties,
                                &self.globals,
                                entry_async,
                            );
                            // Store the bound args as an array property
                            // keyed `__bound_args`; RESUME sees this on
                            // first fire.
                            let bound = Object {
                                properties: HashMap::new(),
                                kind: ObjectKind::Array(args),
                                type_id: 0,
                                fields: Vec::new(),
                            };
                            new_obj.properties.insert(
                                "__bound_args".into(),
                                Value::Object(Arc::new(Mutex::new(bound))),
                            );
                            Value::Object(Arc::new(Mutex::new(new_obj)))
                        } else {
                            return Err(VMError::new("cont.bind: not a continuation"));
                        }
                    } else if cont_val.is_null_ref() {
                        return Err(VMError::new("cont.bind: null continuation"));
                    } else {
                        return Err(VMError::new("cont.bind: not a continuation"));
                    };
                    self.push(new_cont)?;
                }
                // `resume_throw $ct $tag handlers` — resume a continuation
                // by throwing an exception into it. Stack:
                // [cont, exn_value] → control transfers into the cont's
                // nearest try_table matching the throw tag.
                _ if op == Op::RESUME_THROW => {
                    let _tag_idx = self.read_u16();
                    let resume_handlers = self.chunks[self.frame().chunk_index]
                        .stack_switch_handlers
                        .get(&opcode_start)
                        .cloned()
                        .unwrap_or_default();
                    let exn = self.pop();
                    let cont = self.pop();
                    if let Value::Object(ref obj) = cont {
                        let (phase, entry) = {
                            let o = obj.lock().unwrap();
                            if let ObjectKind::Continuation(cs) = &o.kind {
                                (*cs.state.lock().unwrap(), cs.entry.clone())
                            } else {
                                return Err(VMError::new("resume_throw: not a continuation"));
                            }
                        };
                        if matches!(phase, crate::value::ContinuationPhase::Done) {
                            return Err(VMError::new(
                                "trap: resume_throw on completed continuation",
                            ));
                        }
                        let caller_fiber = self.save_fiber();
                        // If suspended, restore fiber then immediately
                        // throw the exception. If fresh (ready), we
                        // first call entry with the exn as its arg so
                        // user-level code can choose to forward.
                        match phase {
                            crate::value::ContinuationPhase::Suspended => {
                                let saved = {
                                    let o = obj.lock().unwrap();
                                    if let ObjectKind::Continuation(cs) = &o.kind {
                                        cs.saved.lock().unwrap().take()
                                    } else {
                                        None
                                    }
                                };
                                if let Some(fiber) = saved {
                                    self.resume_fiber_with(fiber, None)?;
                                }
                                self.active_continuations.push(ActiveContinuation {
                                    cont: cont.clone(),
                                    caller_fiber,
                                    mode: ResumeMode::Raw,
                                    handlers: resume_handlers,
                                });
                                if self.raise_exception_value(exn).is_err() {
                                    let thrown = self.last_exception.take().unwrap_or(Value::Null);
                                    if let Some(ac) = self.active_continuations.pop() {
                                        if let Value::Object(ref obj) = ac.cont {
                                            let o = obj.lock().unwrap();
                                            if let ObjectKind::Continuation(cs) = &o.kind {
                                                *cs.state.lock().unwrap() =
                                                    crate::value::ContinuationPhase::Done;
                                            }
                                        }
                                        self.resume_fiber_with(ac.caller_fiber, None)?;
                                        self.raise_exception_value(thrown)?;
                                    } else {
                                        return Err(VMError::new(format!("{}", thrown)));
                                    }
                                }
                            }
                            crate::value::ContinuationPhase::Ready => {
                                self.active_continuations.push(ActiveContinuation {
                                    cont: cont.clone(),
                                    caller_fiber,
                                    mode: ResumeMode::Raw,
                                    handlers: resume_handlers,
                                });
                                self.push(entry)?;
                                self.push(exn)?;
                                self.call_value(1)?;
                                // Fresh continuation runs on its own fiber so
                                // its internal returns don't trip an enclosing
                                // callback's stale execute_until floor.
                                self.cur_fiber_id = self.next_fiber_id;
                                self.next_fiber_id += 1;
                            }
                            _ => unreachable!(),
                        }
                    } else {
                        return Err(VMError::new("resume_throw: operand is not a continuation"));
                    }
                }

                _ if op == Op::RESUME_THROW_REF => {
                    // `resume_throw_ref $ct (handler)*` — like resume_throw
                    // but the exception is the exnref already on the stack
                    // (no tag immediate). Stack: [cont, exnref].
                    let resume_handlers = self.chunks[self.frame().chunk_index]
                        .stack_switch_handlers
                        .get(&opcode_start)
                        .cloned()
                        .unwrap_or_default();
                    let exn = self.pop();
                    if exn.is_null_ref() {
                        return Err(VMError::new("resume_throw_ref: null exception reference"));
                    }
                    let cont = self.pop();
                    if let Value::Object(ref obj) = cont {
                        let (phase, entry) = {
                            let o = obj.lock().unwrap();
                            if let ObjectKind::Continuation(cs) = &o.kind {
                                (*cs.state.lock().unwrap(), cs.entry.clone())
                            } else {
                                return Err(VMError::new("resume_throw_ref: not a continuation"));
                            }
                        };
                        if matches!(phase, crate::value::ContinuationPhase::Done) {
                            return Err(VMError::new(
                                "trap: resume_throw_ref on completed continuation",
                            ));
                        }
                        let caller_fiber = self.save_fiber();
                        match phase {
                            crate::value::ContinuationPhase::Suspended => {
                                let saved = {
                                    let o = obj.lock().unwrap();
                                    if let ObjectKind::Continuation(cs) = &o.kind {
                                        cs.saved.lock().unwrap().take()
                                    } else {
                                        None
                                    }
                                };
                                if let Some(fiber) = saved {
                                    self.resume_fiber_with(fiber, None)?;
                                }
                                self.active_continuations.push(ActiveContinuation {
                                    cont: cont.clone(),
                                    caller_fiber,
                                    mode: ResumeMode::Raw,
                                    handlers: resume_handlers,
                                });
                                if self.raise_exception_value(exn).is_err() {
                                    let thrown = self.last_exception.take().unwrap_or(Value::Null);
                                    if let Some(ac) = self.active_continuations.pop() {
                                        if let Value::Object(ref obj) = ac.cont {
                                            let o = obj.lock().unwrap();
                                            if let ObjectKind::Continuation(cs) = &o.kind {
                                                *cs.state.lock().unwrap() =
                                                    crate::value::ContinuationPhase::Done;
                                            }
                                        }
                                        self.resume_fiber_with(ac.caller_fiber, None)?;
                                        self.raise_exception_value(thrown)?;
                                    } else {
                                        return Err(VMError::new(format!("{}", thrown)));
                                    }
                                }
                            }
                            crate::value::ContinuationPhase::Ready => {
                                self.active_continuations.push(ActiveContinuation {
                                    cont: cont.clone(),
                                    caller_fiber,
                                    mode: ResumeMode::Raw,
                                    handlers: resume_handlers,
                                });
                                self.push(entry)?;
                                self.push(exn)?;
                                self.call_value(1)?;
                                // Fresh continuation runs on its own fiber so
                                // its internal returns don't trip an enclosing
                                // callback's stale execute_until floor.
                                self.cur_fiber_id = self.next_fiber_id;
                                self.next_fiber_id += 1;
                            }
                            _ => unreachable!(),
                        }
                    } else {
                        return Err(VMError::new(
                            "resume_throw_ref: operand is not a continuation",
                        ));
                    }
                }

                // -- wasi-threads: real OS thread spawning --
                _ if op == Op::THREAD_SPAWN => {
                    // [start_arg, func_ref] → [task_object]
                    //
                    // Matches the wasi-threads `thread.spawn(start_arg) -> i32`
                    // signature: pops a single i32-shaped start argument
                    // alongside the function ref and forwards it to the
                    // spawned function as its slot-0 parameter. This is how
                    // closure-free closure capture works in wasi-threads —
                    // a `Task.Delay(ms)` lowering can push `[ms, worker_fn,
                    // THREAD_SPAWN]` and the worker reads ms from slot 0
                    // without ever needing parent-stack upvalues.
                    //
                    // Value is now Arc-based (Send+Sync), so chunks and host_fns
                    // can be shared directly — no serialization needed.
                    let func_val = self.pop();
                    let start_arg = self.pop();

                    let function = match &func_val {
                        Value::Object(obj) => {
                            let o = obj.lock().unwrap();
                            match &o.kind {
                                ObjectKind::Function(f) => Some(f.clone()),
                                _ => None,
                            }
                        }
                        _ => None,
                    };

                    if let Some(func) = function {
                        let tid = self.next_thread_id;
                        self.next_thread_id += 1;

                        // Create task object FIRST so child can write result to it
                        let mut obj = Object::new();
                        obj.properties
                            .insert("__type".into(), Value::String(Arc::from("Task")));
                        obj.properties.insert("__thread_id".into(), Value::I32(tid));
                        obj.properties
                            .insert("iscompleted".into(), Value::Bool(false));
                        obj.properties.insert("isalive".into(), Value::Bool(true));
                        obj.properties.insert("result".into(), Value::Null);
                        obj.properties
                            .insert("status".into(), Value::String(Arc::from("Running")));
                        let task_obj = Arc::new(Mutex::new(obj));
                        let task_for_child = task_obj.clone();

                        // Share directly — Value is Send+Sync now
                        let child_chunks = self.chunks.clone();
                        let child_memory = self.memory.clone();
                        let child_host_fns = self.host_fns.clone();
                        let child_host_registry = self.host_registry.clone();
                        let child_import_table = self.import_table.clone();
                        let child_globals = self.globals.clone();
                        let child_type_registry = self.type_registry.clone();
                        let child_func_table = self.func_table.clone();
                        let child_wasm_tables = self.wasm_tables.clone();
                        let child_case_aliases = self.case_aliases.clone();
                        let child_strict_isolation = self.strict_isolation;
                        let child_module_prefix = self.module_prefix.clone();

                        let handle = std::thread::spawn(move || {
                            let mut child_vm = VM::new();
                            child_vm.chunks = child_chunks;
                            child_vm.memory = child_memory;
                            child_vm.host_fns = child_host_fns;
                            child_vm.host_registry = child_host_registry;
                            child_vm.import_table = child_import_table;
                            child_vm.globals = child_globals;
                            child_vm.type_registry = child_type_registry;
                            child_vm.func_table = child_func_table;
                            child_vm.wasm_tables = child_wasm_tables;
                            child_vm.case_aliases = child_case_aliases;
                            child_vm.strict_isolation = child_strict_isolation;
                            child_vm.module_prefix = child_module_prefix;

                            // Push the start_arg onto the child VM's stack so
                            // call_function lays it out at slot 0 of the spawned
                            // function's frame (per wasi-threads spec). For
                            // arity-0 worker fns the value sits in an unread
                            // slot and is harmless; for arity-1 workers (e.g.
                            // the Task.Delay sleep worker) it's the start_arg.
                            // Direct push — child VM stack is fresh, can't
                            // overflow.
                            child_vm.stack.push(start_arg);
                            let result = match child_vm
                                .call_function(&func, 1)
                                .and_then(|_| child_vm.execute())
                            {
                                Ok(val) => {
                                    // Store return value in the shared task object
                                    let mut t = task_for_child.lock().unwrap();
                                    t.properties.insert("result".into(), val.clone());
                                    t.properties.insert("iscompleted".into(), Value::Bool(true));
                                    t.properties.insert("isalive".into(), Value::Bool(false));
                                    t.properties.insert("hasexited".into(), Value::Bool(true));
                                    t.properties.insert("exitcode".into(), Value::I32(0));
                                    t.properties.insert(
                                        "status".into(),
                                        Value::String(Arc::from("RanToCompletion")),
                                    );
                                    vec![0u8]
                                }
                                Err(e) => {
                                    let thrown = child_vm.last_exception.take().unwrap_or_else(|| {
                                        Value::String(Arc::from(e.message.as_str()))
                                    });
                                    let mut t = task_for_child.lock().unwrap();
                                    t.properties.insert("exception".into(), thrown);
                                    t.properties.insert("iscompleted".into(), Value::Bool(true));
                                    t.properties.insert("isalive".into(), Value::Bool(false));
                                    t.properties.insert("hasexited".into(), Value::Bool(true));
                                    t.properties.insert("exitcode".into(), Value::I32(-1));
                                    t.properties.insert(
                                        "status".into(),
                                        Value::String(Arc::from("Faulted")),
                                    );
                                    eprintln!("[thread {}] error: {}", tid, e.message);
                                    vec![1u8]
                                }
                            };
                            result
                        });

                        self.thread_handles.insert(tid, handle);
                        self.push(Value::Object(task_obj))?;
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                _ if op == Op::THREAD_JOIN => {
                    // [task_object] → [status: i32]
                    // Wait for a thread to complete. Accepts either a task object
                    // (with __thread_id) or a raw i32 thread ID.
                    let task_val = self.pop();
                    let tid = match &task_val {
                        Value::Object(obj) => {
                            let o = obj.lock().unwrap();
                            o.properties
                                .get("__thread_id")
                                .map(|v| v.as_f64() as i32)
                                .unwrap_or(-1)
                        }
                        Value::I32(n) => *n,
                        _ => task_val.as_f64() as i32,
                    };

                    if let Some(handle) = self.thread_handles.remove(&tid) {
                        let success = match handle.join() {
                            Ok(result) => result.first().copied().unwrap_or(1) == 0,
                            Err(_) => false,
                        };
                        // Update the task object properties
                        if let Value::Object(obj) = &task_val {
                            let mut o = obj.lock().unwrap();
                            o.properties.insert("iscompleted".into(), Value::Bool(true));
                            o.properties.insert("isalive".into(), Value::Bool(false));
                            o.properties.insert("hasexited".into(), Value::Bool(true));
                            o.properties.insert(
                                "exitcode".into(),
                                Value::I32(if success { 0 } else { -1 }),
                            );
                            o.properties.insert(
                                "status".into(),
                                Value::String(Arc::from(if success {
                                    "RanToCompletion"
                                } else {
                                    "Faulted"
                                })),
                            );
                        }
                        self.push(Value::I32(if success { 0 } else { -1 }))?;
                    } else {
                        self.push(Value::I32(-1))?;
                    }
                }

                // -- Extended Const Expressions --

                // -- Typed Continuations --

                // ── SIMD (128-bit vectors) ────────────────────────────────────
                // Memory
                _ if op == Op::V128_LOAD => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let mut b = [0u8; 16];
                    b.copy_from_slice(&self.read_memory_bytes(memidx, addr, 16)?);
                    self.push(Value::V128(b))?;
                }
                _ if op == Op::V128_LOAD8X8_S => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..8 {
                        let b = bytes[i];
                        let v = b as i8 as i16;
                        out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD8X8_U => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..8 {
                        let b = bytes[i];
                        let v = b as u16;
                        out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD16X4_S => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        let lo = bytes[i * 2];
                        let hi = bytes[i * 2 + 1];
                        let v = u16::from_le_bytes([lo, hi]) as i16 as i32;
                        out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD16X4_U => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        let lo = bytes[i * 2];
                        let hi = bytes[i * 2 + 1];
                        let v = u16::from_le_bytes([lo, hi]) as u32;
                        out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD32X2_S => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..2 {
                        let v = i32::from_le_bytes(read_le(&bytes[i * 4..i * 4 + 4])) as i64;
                        out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD32X2_U => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..2 {
                        let v = u32::from_le_bytes(read_le(&bytes[i * 4..i * 4 + 4])) as u64;
                        out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD8_SPLAT => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let b = self.read_memory_bytes(memidx, addr, 1)?[0];
                    self.push(Value::V128([b; 16]))?;
                }
                _ if op == Op::V128_LOAD16_SPLAT => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let lo = bytes[0];
                    let hi = bytes[1];
                    let mut out = [0u8; 16];
                    for i in 0..8 {
                        out[i * 2] = lo;
                        out[i * 2 + 1] = hi;
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD32_SPLAT => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    let v = i32::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD64_SPLAT => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let v = i64::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    out[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_STORE => {
                    let val = self.pop();
                    let (memidx, addr) = self.pop_simd_addr()?;
                    if let Value::V128(b) = val {
                        self.write_memory_bytes(memidx, addr, &b)?;
                    }
                }
                _ if op == Op::V128_CONST => {
                    let mut b = [0u8; 16];
                    for i in 0..16 {
                        b[i] = self.read_byte();
                    }
                    self.push(Value::V128(b))?;
                }
                _ if op == Op::V128_LOAD8_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 15;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(mut v) = val {
                        v[lane] = self.read_memory_bytes(memidx, addr, 1)?[0];
                        self.push(Value::V128(v))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_LOAD16_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 7;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(mut v) = val {
                        let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                        v[lane * 2] = bytes[0];
                        v[lane * 2 + 1] = bytes[1];
                        self.push(Value::V128(v))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_LOAD32_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 3;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(mut v) = val {
                        let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                        v[lane * 4..lane * 4 + 4].copy_from_slice(&bytes);
                        self.push(Value::V128(v))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_LOAD64_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 1;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(mut v) = val {
                        let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                        v[lane * 8..lane * 8 + 8].copy_from_slice(&bytes);
                        self.push(Value::V128(v))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_STORE8_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 15;
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = self.pop() {
                        self.write_memory_bytes(memidx, addr, &[v[lane]])?;
                    }
                }
                _ if op == Op::V128_STORE16_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 7;
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = self.pop() {
                        self.write_memory_bytes(memidx, addr, &v[lane * 2..lane * 2 + 2])?;
                    }
                }
                _ if op == Op::V128_STORE32_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 3;
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = self.pop() {
                        self.write_memory_bytes(memidx, addr, &v[lane * 4..lane * 4 + 4])?;
                    }
                }
                _ if op == Op::V128_STORE64_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 1;
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = self.pop() {
                        self.write_memory_bytes(memidx, addr, &v[lane * 8..lane * 8 + 8])?;
                    }
                }
                _ if op == Op::V128_LOAD32_ZERO => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    let v = i32::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    out[0..4].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::V128_LOAD64_ZERO => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let v = i64::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                // Shuffle / swizzle
                _ if op == Op::I8X16_SHUFFLE => {
                    let mut idx = [0u8; 16];
                    for i in 0..16 {
                        idx[i] = self.read_byte();
                    }
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let combined: Vec<u8> = va.iter().chain(vb.iter()).copied().collect();
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = combined.get(idx[i] as usize).copied().unwrap_or(0);
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I8X16_SWIZZLE => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            let n = vb[i] as usize;
                            out[i] = if n < 16 { va[n] } else { 0 };
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // Splat
                _ if op == Op::I8X16_SPLAT => {
                    let v = self.pop().as_i32() as u8;
                    self.push(Value::V128([v; 16]))?;
                }
                _ if op == Op::I16X8_SPLAT => {
                    let v = self.pop().as_i32() as i16;
                    let b = v.to_le_bytes();
                    let mut out = [0u8; 16];
                    for i in 0..8 {
                        out[i * 2..i * 2 + 2].copy_from_slice(&b);
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::I32X4_SPLAT => {
                    let v = self.pop().as_i32();
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::I64X2_SPLAT => {
                    let v = self.pop().as_i64();
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    out[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::F32X4_SPLAT => {
                    let v = self.pop().as_f64() as f32;
                    let b = v.to_le_bytes();
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        out[i * 4..i * 4 + 4].copy_from_slice(&b);
                    }
                    self.push(Value::V128(out))?;
                }
                _ if op == Op::F64X2_SPLAT => {
                    let v = self.pop().as_f64();
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    out[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                // Extract / replace lane
                _ if op == Op::I8X16_EXTRACT_LANE_S => {
                    let l = self.read_byte() as usize & 15;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(a[l] as i8 as i32))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I8X16_EXTRACT_LANE_U => {
                    let l = self.read_byte() as usize & 15;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(a[l] as i32))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I8X16_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 15;
                    let v = self.pop().as_i32() as u8;
                    if let Value::V128(mut a) = self.pop() {
                        a[l] = v;
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTRACT_LANE_S => {
                    let l = self.read_byte() as usize & 7;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(
                            i16::from_le_bytes([a[l * 2], a[l * 2 + 1]]) as i32
                        ))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I16X8_EXTRACT_LANE_U => {
                    let l = self.read_byte() as usize & 7;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(
                            u16::from_le_bytes([a[l * 2], a[l * 2 + 1]]) as i32
                        ))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I16X8_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 7;
                    let v = self.pop().as_i32() as i16;
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 2..l * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTRACT_LANE => {
                    let l = self.read_byte() as usize & 3;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(i32::from_le_bytes(
                            a[l * 4..l * 4 + 4].try_into().unwrap(),
                        )))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I32X4_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 3;
                    let v = self.pop().as_i32();
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 4..l * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_EXTRACT_LANE => {
                    let l = self.read_byte() as usize & 1;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I64(i64::from_le_bytes(
                            a[l * 8..l * 8 + 8].try_into().unwrap(),
                        )))?;
                    } else {
                        self.push(Value::I64(0))?;
                    }
                }
                _ if op == Op::I64X2_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 1;
                    let v = self.pop().as_i64();
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 8..l * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F32X4_EXTRACT_LANE => {
                    let l = self.read_byte() as usize & 3;
                    if let Value::V128(a) = self.pop() {
                        // Result type is `f32` (spec): push a Value::F32 so it
                        // displays as WAT float text and feeds later f32 ops.
                        self.push(Value::F32(f32::from_le_bytes(
                            a[l * 4..l * 4 + 4].try_into().unwrap(),
                        )))?;
                    } else {
                        self.push(Value::F32(0.0))?;
                    }
                }
                _ if op == Op::F32X4_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 3;
                    let v = self.pop().as_f64() as f32;
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 4..l * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_EXTRACT_LANE => {
                    let l = self.read_byte() as usize & 1;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::F64(f64::from_le_bytes(
                            a[l * 8..l * 8 + 8].try_into().unwrap(),
                        )))?;
                    } else {
                        self.push(Value::F64(0.0))?;
                    }
                }
                _ if op == Op::F64X2_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 1;
                    let v = self.pop().as_f64();
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 8..l * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // i8x16 comparisons
                _ if op == Op::I8X16_EQ => {
                    self.simd_i8x16_binop(|a, b| if a == b { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_NE => {
                    self.simd_i8x16_binop(|a, b| if a != b { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_LT_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) < (b as i8) { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_LT_U => {
                    self.simd_i8x16_binop(|a, b| if a < b { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_GT_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) > (b as i8) { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_GT_U => {
                    self.simd_i8x16_binop(|a, b| if a > b { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_LE_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) <= (b as i8) { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_LE_U => {
                    self.simd_i8x16_binop(|a, b| if a <= b { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_GE_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) >= (b as i8) { 0xFF } else { 0 })?;
                }
                _ if op == Op::I8X16_GE_U => {
                    self.simd_i8x16_binop(|a, b| if a >= b { 0xFF } else { 0 })?;
                }
                // i16x8 comparisons
                _ if op == Op::I16X8_EQ => {
                    self.simd_i16x8_binop(|a, b| if a == b { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_NE => {
                    self.simd_i16x8_binop(|a, b| if a != b { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_LT_S => {
                    self.simd_i16x8_binop(|a, b| if a < b { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_LT_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) < (b as u16) { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_GT_S => {
                    self.simd_i16x8_binop(|a, b| if a > b { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_GT_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) > (b as u16) { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_LE_S => {
                    self.simd_i16x8_binop(|a, b| if a <= b { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_LE_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) <= (b as u16) { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_GE_S => {
                    self.simd_i16x8_binop(|a, b| if a >= b { -1 } else { 0 })?;
                }
                _ if op == Op::I16X8_GE_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) >= (b as u16) { -1 } else { 0 })?;
                }
                // i32x4 comparisons
                _ if op == Op::I32X4_EQ => {
                    self.simd_i32x4_binop(|a, b| if a == b { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_NE => {
                    self.simd_i32x4_binop(|a, b| if a != b { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_LT_S => {
                    self.simd_i32x4_binop(|a, b| if a < b { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_LT_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) < (b as u32) { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_GT_S => {
                    self.simd_i32x4_binop(|a, b| if a > b { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_GT_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) > (b as u32) { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_LE_S => {
                    self.simd_i32x4_binop(|a, b| if a <= b { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_LE_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) <= (b as u32) { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_GE_S => {
                    self.simd_i32x4_binop(|a, b| if a >= b { -1 } else { 0 })?;
                }
                _ if op == Op::I32X4_GE_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) >= (b as u32) { -1 } else { 0 })?;
                }
                // f32x4 comparisons
                _ if op == Op::F32X4_EQ => {
                    self.simd_f32x4_cmp(|a, b| a == b)?;
                }
                _ if op == Op::F32X4_NE => {
                    self.simd_f32x4_cmp(|a, b| a != b)?;
                }
                _ if op == Op::F32X4_LT => {
                    self.simd_f32x4_cmp(|a, b| a < b)?;
                }
                _ if op == Op::F32X4_GT => {
                    self.simd_f32x4_cmp(|a, b| a > b)?;
                }
                _ if op == Op::F32X4_LE => {
                    self.simd_f32x4_cmp(|a, b| a <= b)?;
                }
                _ if op == Op::F32X4_GE => {
                    self.simd_f32x4_cmp(|a, b| a >= b)?;
                }
                // f64x2 comparisons
                _ if op == Op::F64X2_EQ => {
                    self.simd_f64x2_cmp(|a, b| a == b)?;
                }
                _ if op == Op::F64X2_NE => {
                    self.simd_f64x2_cmp(|a, b| a != b)?;
                }
                _ if op == Op::F64X2_LT => {
                    self.simd_f64x2_cmp(|a, b| a < b)?;
                }
                _ if op == Op::F64X2_GT => {
                    self.simd_f64x2_cmp(|a, b| a > b)?;
                }
                _ if op == Op::F64X2_LE => {
                    self.simd_f64x2_cmp(|a, b| a <= b)?;
                }
                _ if op == Op::F64X2_GE => {
                    self.simd_f64x2_cmp(|a, b| a >= b)?;
                }
                // v128 bitwise
                _ if op == Op::V128_NOT => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = !a[i];
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_AND => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(a), Value::V128(b)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = a[i] & b[i];
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_ANDNOT => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(a), Value::V128(b)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = a[i] & !b[i];
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_OR => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(a), Value::V128(b)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = a[i] | b[i];
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_XOR => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(a), Value::V128(b)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = a[i] ^ b[i];
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_BITSELECT => {
                    let m = self.pop();
                    let v2 = self.pop();
                    let v1 = self.pop();
                    if let (Value::V128(a), Value::V128(b), Value::V128(m)) = (v1, v2, m) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = (a[i] & m[i]) | (b[i] & !m[i]);
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::V128_ANY_TRUE => {
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(if a.iter().any(|&b| b != 0) { 1 } else { 0 }))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                // Promote / demote
                _ if op == Op::F32X4_DEMOTE_F64X2_ZERO => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f =
                                f64::from_le_bytes(a[i * 8..i * 8 + 8].try_into().unwrap()) as f32;
                            out[i * 4..i * 4 + 4].copy_from_slice(&f.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_PROMOTE_LOW_F32X4 => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f =
                                f32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) as f64;
                            out[i * 8..i * 8 + 8].copy_from_slice(&f.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // i8x16 unary
                _ if op == Op::I8X16_ABS => {
                    self.simd_i8x16_unop(|a| (a as i8).unsigned_abs())?;
                }
                _ if op == Op::I8X16_NEG => {
                    self.simd_i8x16_unop(|a| (a as i8).wrapping_neg() as u8)?;
                }
                _ if op == Op::I8X16_POPCNT => {
                    self.simd_i8x16_unop(|a| a.count_ones() as u8)?;
                }
                _ if op == Op::I8X16_ALL_TRUE => {
                    self.simd_i8x16_testop(|a| a != 0)?;
                }
                _ if op == Op::I8X16_BITMASK => {
                    if let Value::V128(a) = self.pop() {
                        let mut mask = 0i32;
                        for i in 0..16 {
                            if (a[i] as i8) < 0 {
                                mask |= 1 << i;
                            }
                        }
                        self.push(Value::I32(mask))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I8X16_NARROW_I16X8_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = i16::from_le_bytes([va[i * 2], va[i * 2 + 1]]).clamp(-128, 127)
                                as i8 as u8;
                            out[i] = v;
                            let v2 = i16::from_le_bytes([vb[i * 2], vb[i * 2 + 1]]).clamp(-128, 127)
                                as i8 as u8;
                            out[8 + i] = v2;
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I8X16_NARROW_I16X8_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v =
                                i16::from_le_bytes([va[i * 2], va[i * 2 + 1]]).clamp(0, 255) as u8;
                            out[i] = v;
                            let v2 =
                                i16::from_le_bytes([vb[i * 2], vb[i * 2 + 1]]).clamp(0, 255) as u8;
                            out[8 + i] = v2;
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // f32x4 unary
                _ if op == Op::F32X4_CEIL => {
                    self.simd_f32x4_unop(|a| a.ceil())?;
                }
                _ if op == Op::F32X4_FLOOR => {
                    self.simd_f32x4_unop(|a| a.floor())?;
                }
                _ if op == Op::F32X4_TRUNC => {
                    self.simd_f32x4_unop(|a| a.trunc())?;
                }
                _ if op == Op::F32X4_NEAREST => {
                    self.simd_f32x4_unop(|a| a.round_ties_even())?;
                }
                // i8x16 shifts
                _ if op == Op::I8X16_SHL => {
                    let sh = self.pop().as_i32() as u32 & 7;
                    self.simd_i8x16_unop(|a| a.wrapping_shl(sh))?;
                }
                _ if op == Op::I8X16_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 7;
                    self.simd_i8x16_unop(|a| ((a as i8).wrapping_shr(sh)) as u8)?;
                }
                _ if op == Op::I8X16_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 7;
                    self.simd_i8x16_unop(|a| a.wrapping_shr(sh))?;
                }
                // i8x16 arithmetic
                _ if op == Op::I8X16_ADD => {
                    self.simd_i8x16_binop(|a, b| a.wrapping_add(b))?;
                }
                _ if op == Op::I8X16_ADD_SAT_S => {
                    self.simd_i8x16_binop(|a, b| ((a as i8).saturating_add(b as i8)) as u8)?;
                }
                _ if op == Op::I8X16_ADD_SAT_U => {
                    self.simd_i8x16_binop(|a, b| a.saturating_add(b))?;
                }
                _ if op == Op::I8X16_SUB => {
                    self.simd_i8x16_binop(|a, b| a.wrapping_sub(b))?;
                }
                _ if op == Op::I8X16_SUB_SAT_S => {
                    self.simd_i8x16_binop(|a, b| ((a as i8).saturating_sub(b as i8)) as u8)?;
                }
                _ if op == Op::I8X16_SUB_SAT_U => {
                    self.simd_i8x16_binop(|a, b| a.saturating_sub(b))?;
                }
                _ if op == Op::I8X16_MIN_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) < (b as i8) { a } else { b })?;
                }
                _ if op == Op::I8X16_MIN_U => {
                    self.simd_i8x16_binop(|a, b| a.min(b))?;
                }
                _ if op == Op::I8X16_MAX_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) > (b as i8) { a } else { b })?;
                }
                _ if op == Op::I8X16_MAX_U => {
                    self.simd_i8x16_binop(|a, b| a.max(b))?;
                }
                _ if op == Op::I8X16_AVGR_U => {
                    self.simd_i8x16_binop(|a, b| ((a as u16 + b as u16 + 1) / 2) as u8)?;
                }
                // f64x2 unary
                _ if op == Op::F64X2_CEIL => {
                    self.simd_f64x2_unop(|a| a.ceil())?;
                }
                _ if op == Op::F64X2_FLOOR => {
                    self.simd_f64x2_unop(|a| a.floor())?;
                }
                _ if op == Op::F64X2_TRUNC => {
                    self.simd_f64x2_unop(|a| a.trunc())?;
                }
                _ if op == Op::F64X2_NEAREST => {
                    self.simd_f64x2_unop(|a| a.round_ties_even())?;
                }
                // extadd pairwise
                _ if op == Op::I16X8_EXTADD_PAIRWISE_I8X16_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = (a[i * 2] as i8 as i16).wrapping_add(a[i * 2 + 1] as i8 as i16);
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTADD_PAIRWISE_I8X16_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = (a[i * 2] as i16).wrapping_add(a[i * 2 + 1] as i16);
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTADD_PAIRWISE_I16X8_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let la = i16::from_le_bytes([a[i * 4], a[i * 4 + 1]]) as i32;
                            let lb = i16::from_le_bytes([a[i * 4 + 2], a[i * 4 + 3]]) as i32;
                            out[i * 4..i * 4 + 4]
                                .copy_from_slice(&la.wrapping_add(lb).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTADD_PAIRWISE_I16X8_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let la = u16::from_le_bytes([a[i * 4], a[i * 4 + 1]]) as i32;
                            let lb = u16::from_le_bytes([a[i * 4 + 2], a[i * 4 + 3]]) as i32;
                            out[i * 4..i * 4 + 4]
                                .copy_from_slice(&la.wrapping_add(lb).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // i16x8 unary
                _ if op == Op::I16X8_ABS => {
                    self.simd_i16x8_unop(|a| a.unsigned_abs() as i16)?;
                }
                _ if op == Op::I16X8_NEG => {
                    self.simd_i16x8_unop(|a| a.wrapping_neg())?;
                }
                _ if op == Op::I16X8_Q15MULR_SAT_S => {
                    self.simd_i16x8_binop(|a, b| {
                        let r = (a as i32 * b as i32 + 0x4000) >> 15;
                        r.clamp(i16::MIN as i32, i16::MAX as i32) as i16
                    })?;
                }
                _ if op == Op::I16X8_ALL_TRUE => {
                    self.simd_i16x8_testop(|a| a != 0)?;
                }
                _ if op == Op::I16X8_BITMASK => {
                    if let Value::V128(a) = self.pop() {
                        let mut mask = 0i32;
                        for i in 0..8 {
                            if i16::from_le_bytes([a[i * 2], a[i * 2 + 1]]) < 0 {
                                mask |= 1 << i;
                            }
                        }
                        self.push(Value::I32(mask))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I16X8_NARROW_I32X4_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap())
                                .clamp(-32768, 32767) as i16;
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                            let v2 = i32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap())
                                .clamp(-32768, 32767) as i16;
                            out[(4 + i) * 2..(4 + i) * 2 + 2].copy_from_slice(&v2.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_NARROW_I32X4_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap())
                                .clamp(0, 65535) as u16;
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                            let v2 = i32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap())
                                .clamp(0, 65535) as u16;
                            out[(4 + i) * 2..(4 + i) * 2 + 2].copy_from_slice(&v2.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTEND_LOW_I8X16_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = a[i] as i8 as i16;
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTEND_HIGH_I8X16_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = a[8 + i] as i8 as i16;
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTEND_LOW_I8X16_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = a[i] as i16;
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTEND_HIGH_I8X16_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = a[8 + i] as i16;
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_SHL => {
                    let sh = self.pop().as_i32() as u32 & 15;
                    self.simd_i16x8_unop(|a| a.wrapping_shl(sh))?;
                }
                _ if op == Op::I16X8_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 15;
                    self.simd_i16x8_unop(|a| a.wrapping_shr(sh))?;
                }
                _ if op == Op::I16X8_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 15;
                    self.simd_i16x8_unop(|a| (a as u16).wrapping_shr(sh) as i16)?;
                }
                _ if op == Op::I16X8_ADD => {
                    self.simd_i16x8_binop(|a, b| a.wrapping_add(b))?;
                }
                _ if op == Op::I16X8_ADD_SAT_S => {
                    self.simd_i16x8_binop(|a, b| a.saturating_add(b))?;
                }
                _ if op == Op::I16X8_ADD_SAT_U => {
                    self.simd_i16x8_binop(|a, b| ((a as u16).saturating_add(b as u16)) as i16)?;
                }
                _ if op == Op::I16X8_SUB => {
                    self.simd_i16x8_binop(|a, b| a.wrapping_sub(b))?;
                }
                _ if op == Op::I16X8_SUB_SAT_S => {
                    self.simd_i16x8_binop(|a, b| a.saturating_sub(b))?;
                }
                _ if op == Op::I16X8_SUB_SAT_U => {
                    self.simd_i16x8_binop(|a, b| ((a as u16).saturating_sub(b as u16)) as i16)?;
                }
                _ if op == Op::I16X8_MUL => {
                    self.simd_i16x8_binop(|a, b| a.wrapping_mul(b))?;
                }
                _ if op == Op::I16X8_MIN_S => {
                    self.simd_i16x8_binop(|a, b| a.min(b))?;
                }
                _ if op == Op::I16X8_MIN_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) < (b as u16) { a } else { b })?;
                }
                _ if op == Op::I16X8_MAX_S => {
                    self.simd_i16x8_binop(|a, b| a.max(b))?;
                }
                _ if op == Op::I16X8_MAX_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) > (b as u16) { a } else { b })?;
                }
                _ if op == Op::I16X8_AVGR_U => {
                    self.simd_i16x8_binop(|a, b| {
                        (((a as u16 as u32) + (b as u16 as u32) + 1) / 2) as i16
                    })?;
                }
                _ if op == Op::I16X8_EXTMUL_LOW_I8X16_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = (va[i] as i8 as i16).wrapping_mul(vb[i] as i8 as i16);
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTMUL_HIGH_I8X16_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = (va[8 + i] as i8 as i16).wrapping_mul(vb[8 + i] as i8 as i16);
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTMUL_LOW_I8X16_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = (va[i] as i16).wrapping_mul(vb[i] as i16);
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_EXTMUL_HIGH_I8X16_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let v = (va[8 + i] as i16).wrapping_mul(vb[8 + i] as i16);
                            out[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // i32x4 unary
                _ if op == Op::I32X4_ABS => {
                    self.simd_i32x4_unop(|a| a.unsigned_abs() as i32)?;
                }
                _ if op == Op::I32X4_NEG => {
                    self.simd_i32x4_unop(|a| a.wrapping_neg())?;
                }
                _ if op == Op::I32X4_ALL_TRUE => {
                    self.simd_i32x4_testop(|a| a != 0)?;
                }
                _ if op == Op::I32X4_BITMASK => {
                    if let Value::V128(a) = self.pop() {
                        let mut mask = 0i32;
                        for i in 0..4 {
                            if i32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) < 0 {
                                mask |= 1 << i;
                            }
                        }
                        self.push(Value::I32(mask))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I32X4_EXTEND_LOW_I16X8_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i16::from_le_bytes([a[i * 2], a[i * 2 + 1]]) as i32;
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTEND_HIGH_I16X8_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i16::from_le_bytes([a[(4 + i) * 2], a[(4 + i) * 2 + 1]]) as i32;
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTEND_LOW_I16X8_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = u16::from_le_bytes([a[i * 2], a[i * 2 + 1]]) as i32;
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTEND_HIGH_I16X8_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = u16::from_le_bytes([a[(4 + i) * 2], a[(4 + i) * 2 + 1]]) as i32;
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_SHL => {
                    let sh = self.pop().as_i32() as u32 & 31;
                    self.simd_i32x4_unop(|a| a.wrapping_shl(sh))?;
                }
                _ if op == Op::I32X4_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 31;
                    self.simd_i32x4_unop(|a| a.wrapping_shr(sh))?;
                }
                _ if op == Op::I32X4_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 31;
                    self.simd_i32x4_unop(|a| (a as u32).wrapping_shr(sh) as i32)?;
                }
                _ if op == Op::I32X4_ADD => {
                    self.simd_i32x4_binop(|a, b| a.wrapping_add(b))?;
                }
                _ if op == Op::I32X4_SUB => {
                    self.simd_i32x4_binop(|a, b| a.wrapping_sub(b))?;
                }
                _ if op == Op::I32X4_MUL => {
                    self.simd_i32x4_binop(|a, b| a.wrapping_mul(b))?;
                }
                _ if op == Op::I32X4_MIN_S => {
                    self.simd_i32x4_binop(|a, b| a.min(b))?;
                }
                _ if op == Op::I32X4_MIN_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) < (b as u32) { a } else { b })?;
                }
                _ if op == Op::I32X4_MAX_S => {
                    self.simd_i32x4_binop(|a, b| a.max(b))?;
                }
                _ if op == Op::I32X4_MAX_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) > (b as u32) { a } else { b })?;
                }
                _ if op == Op::I32X4_DOT_I16X8_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let a0 = i16::from_le_bytes([va[i * 4], va[i * 4 + 1]]) as i32;
                            let b0 = i16::from_le_bytes([vb[i * 4], vb[i * 4 + 1]]) as i32;
                            let a1 = i16::from_le_bytes([va[i * 4 + 2], va[i * 4 + 3]]) as i32;
                            let b1 = i16::from_le_bytes([vb[i * 4 + 2], vb[i * 4 + 3]]) as i32;
                            out[i * 4..i * 4 + 4]
                                .copy_from_slice(&(a0 * b0 + a1 * b1).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTMUL_LOW_I16X8_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v =
                                (i16::from_le_bytes([va[i * 2], va[i * 2 + 1]]) as i32)
                                    .wrapping_mul(
                                        i16::from_le_bytes([vb[i * 2], vb[i * 2 + 1]]) as i32
                                    );
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTMUL_HIGH_I16X8_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = (i16::from_le_bytes([va[(4 + i) * 2], va[(4 + i) * 2 + 1]])
                                as i32)
                                .wrapping_mul(i16::from_le_bytes([
                                    vb[(4 + i) * 2],
                                    vb[(4 + i) * 2 + 1],
                                ]) as i32);
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTMUL_LOW_I16X8_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v =
                                (u16::from_le_bytes([va[i * 2], va[i * 2 + 1]]) as i32)
                                    .wrapping_mul(
                                        u16::from_le_bytes([vb[i * 2], vb[i * 2 + 1]]) as i32
                                    );
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_EXTMUL_HIGH_I16X8_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = (u16::from_le_bytes([va[(4 + i) * 2], va[(4 + i) * 2 + 1]])
                                as i32)
                                .wrapping_mul(u16::from_le_bytes([
                                    vb[(4 + i) * 2],
                                    vb[(4 + i) * 2 + 1],
                                ]) as i32);
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // i64x2
                _ if op == Op::I64X2_ABS => {
                    self.simd_i64x2_unop(|a| a.unsigned_abs() as i64)?;
                }
                _ if op == Op::I64X2_NEG => {
                    self.simd_i64x2_unop(|a| a.wrapping_neg())?;
                }
                _ if op == Op::I64X2_ALL_TRUE => {
                    self.simd_i64x2_testop(|a| a != 0)?;
                }
                _ if op == Op::I64X2_BITMASK => {
                    if let Value::V128(a) = self.pop() {
                        let mut mask = 0i32;
                        for i in 0..2 {
                            if i64::from_le_bytes(a[i * 8..i * 8 + 8].try_into().unwrap()) < 0 {
                                mask |= 1 << i;
                            }
                        }
                        self.push(Value::I32(mask))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                _ if op == Op::I64X2_EXTEND_LOW_I32X4_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v =
                                i32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) as i64;
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_EXTEND_HIGH_I32X4_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = i32::from_le_bytes(
                                a[(2 + i) * 4..(2 + i) * 4 + 4].try_into().unwrap(),
                            ) as i64;
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_EXTEND_LOW_I32X4_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v =
                                u32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) as i64;
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_EXTEND_HIGH_I32X4_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = u32::from_le_bytes(
                                a[(2 + i) * 4..(2 + i) * 4 + 4].try_into().unwrap(),
                            ) as i64;
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_SHL => {
                    let sh = self.pop().as_i32() as u32 & 63;
                    self.simd_i64x2_unop(|a| a.wrapping_shl(sh))?;
                }
                _ if op == Op::I64X2_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 63;
                    self.simd_i64x2_unop(|a| a.wrapping_shr(sh))?;
                }
                _ if op == Op::I64X2_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 63;
                    self.simd_i64x2_unop(|a| (a as u64).wrapping_shr(sh) as i64)?;
                }
                _ if op == Op::I64X2_ADD => {
                    self.simd_i64x2_binop(|a, b| a.wrapping_add(b))?;
                }
                _ if op == Op::I64X2_SUB => {
                    self.simd_i64x2_binop(|a, b| a.wrapping_sub(b))?;
                }
                _ if op == Op::I64X2_MUL => {
                    self.simd_i64x2_binop(|a, b| a.wrapping_mul(b))?;
                }
                _ if op == Op::I64X2_EQ => {
                    self.simd_i64x2_cmp(|a, b| a == b)?;
                }
                _ if op == Op::I64X2_NE => {
                    self.simd_i64x2_cmp(|a, b| a != b)?;
                }
                _ if op == Op::I64X2_LT_S => {
                    self.simd_i64x2_cmp(|a, b| a < b)?;
                }
                _ if op == Op::I64X2_GT_S => {
                    self.simd_i64x2_cmp(|a, b| a > b)?;
                }
                _ if op == Op::I64X2_LE_S => {
                    self.simd_i64x2_cmp(|a, b| a <= b)?;
                }
                _ if op == Op::I64X2_GE_S => {
                    self.simd_i64x2_cmp(|a, b| a >= b)?;
                }
                _ if op == Op::I64X2_EXTMUL_LOW_I32X4_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = (i32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap())
                                as i64)
                                .wrapping_mul(i32::from_le_bytes(
                                    vb[i * 4..i * 4 + 4].try_into().unwrap(),
                                ) as i64);
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_EXTMUL_HIGH_I32X4_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = (i32::from_le_bytes(
                                va[(2 + i) * 4..(2 + i) * 4 + 4].try_into().unwrap(),
                            ) as i64)
                                .wrapping_mul(i32::from_le_bytes(
                                    vb[(2 + i) * 4..(2 + i) * 4 + 4].try_into().unwrap(),
                                ) as i64);
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_EXTMUL_LOW_I32X4_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = (u32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap())
                                as i64)
                                .wrapping_mul(u32::from_le_bytes(
                                    vb[i * 4..i * 4 + 4].try_into().unwrap(),
                                ) as i64);
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_EXTMUL_HIGH_I32X4_U => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = (u32::from_le_bytes(
                                va[(2 + i) * 4..(2 + i) * 4 + 4].try_into().unwrap(),
                            ) as i64)
                                .wrapping_mul(u32::from_le_bytes(
                                    vb[(2 + i) * 4..(2 + i) * 4 + 4].try_into().unwrap(),
                                ) as i64);
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // f32x4
                _ if op == Op::F32X4_ABS => {
                    self.simd_f32x4_unop(|a| a.abs())?;
                }
                _ if op == Op::F32X4_NEG => {
                    self.simd_f32x4_unop(|a| -a)?;
                }
                _ if op == Op::F32X4_SQRT => {
                    self.simd_f32x4_unop(|a| a.sqrt())?;
                }
                _ if op == Op::F32X4_ADD => {
                    self.simd_f32x4_binop(|a, b| a + b)?;
                }
                _ if op == Op::F32X4_SUB => {
                    self.simd_f32x4_binop(|a, b| a - b)?;
                }
                _ if op == Op::F32X4_MUL => {
                    self.simd_f32x4_binop(|a, b| a * b)?;
                }
                _ if op == Op::F32X4_DIV => {
                    self.simd_f32x4_binop(|a, b| a / b)?;
                }
                _ if op == Op::F32X4_MIN => {
                    self.simd_f32x4_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f32::NAN
                        } else {
                            a.min(b)
                        }
                    })?;
                }
                _ if op == Op::F32X4_MAX => {
                    self.simd_f32x4_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f32::NAN
                        } else {
                            a.max(b)
                        }
                    })?;
                }
                _ if op == Op::F32X4_PMIN => {
                    self.simd_f32x4_binop(|a, b| if b < a { b } else { a })?;
                }
                _ if op == Op::F32X4_PMAX => {
                    self.simd_f32x4_binop(|a, b| if a < b { b } else { a })?;
                }
                // f64x2
                _ if op == Op::F64X2_ABS => {
                    self.simd_f64x2_unop(|a| a.abs())?;
                }
                _ if op == Op::F64X2_NEG => {
                    self.simd_f64x2_unop(|a| -a)?;
                }
                _ if op == Op::F64X2_SQRT => {
                    self.simd_f64x2_unop(|a| a.sqrt())?;
                }
                _ if op == Op::F64X2_ADD => {
                    self.simd_f64x2_binop(|a, b| a + b)?;
                }
                _ if op == Op::F64X2_SUB => {
                    self.simd_f64x2_binop(|a, b| a - b)?;
                }
                _ if op == Op::F64X2_MUL => {
                    self.simd_f64x2_binop(|a, b| a * b)?;
                }
                _ if op == Op::F64X2_DIV => {
                    self.simd_f64x2_binop(|a, b| a / b)?;
                }
                _ if op == Op::F64X2_MIN => {
                    self.simd_f64x2_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f64::NAN
                        } else {
                            a.min(b)
                        }
                    })?;
                }
                _ if op == Op::F64X2_MAX => {
                    self.simd_f64x2_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f64::NAN
                        } else {
                            a.max(b)
                        }
                    })?;
                }
                _ if op == Op::F64X2_PMIN => {
                    self.simd_f64x2_binop(|a, b| if b < a { b } else { a })?;
                }
                _ if op == Op::F64X2_PMAX => {
                    self.simd_f64x2_binop(|a, b| if a < b { b } else { a })?;
                }
                // Conversions
                _ if op == Op::I32X4_TRUNC_SAT_F32X4_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let f = f32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap());
                            let v = if f.is_nan() {
                                0
                            } else {
                                f.clamp(i32::MIN as f32, i32::MAX as f32) as i32
                            };
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_TRUNC_SAT_F32X4_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let f = f32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap());
                            let v: u32 = if f.is_nan() || f < 0.0 {
                                0
                            } else {
                                f.clamp(0.0, u32::MAX as f32) as u32
                            };
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F32X4_CONVERT_I32X4_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v =
                                i32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) as f32;
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F32X4_CONVERT_I32X4_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v =
                                u32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) as f32;
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_TRUNC_SAT_F64X2_S_ZERO => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f = f64::from_le_bytes(a[i * 8..i * 8 + 8].try_into().unwrap());
                            let v = if f.is_nan() {
                                0
                            } else {
                                f.clamp(i32::MIN as f64, i32::MAX as f64) as i32
                            };
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_TRUNC_SAT_F64X2_U_ZERO => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f = f64::from_le_bytes(a[i * 8..i * 8 + 8].try_into().unwrap());
                            let v: u32 = if f.is_nan() || f < 0.0 {
                                0
                            } else {
                                f.clamp(0.0, u32::MAX as f64) as u32
                            };
                            out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_CONVERT_LOW_I32X4_S => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v =
                                i32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) as f64;
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_CONVERT_LOW_I32X4_U => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v =
                                u32::from_le_bytes(a[i * 4..i * 4 + 4].try_into().unwrap()) as f64;
                            out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }

                _ if op.group() == 0xFE && self.execute_threads_op(op)? => {}

                // -- Memory64 --

                // -- Relaxed-SIMD proposal (prefix 0xDD internal, 0xFD 0x100+ in WASM) --
                //
                // All 20 ops implemented deterministically. The "relaxed"
                // semantics give the implementation freedom on edge cases
                // (NaN sign, out-of-range truncation, lane-select bit
                // policy) — we pick one policy per op and stick with it
                // so results are reproducible across platforms.
                _ if op == Op::I8X16_RELAXED_SWIZZLE => {
                    // Same as i8x16.swizzle; relaxed allows host to mask
                    // indices >= 16 to 0 or return unspecified bytes. We
                    // pick the safe mask-to-zero variant.
                    let idx = self.pop();
                    let src = self.pop();
                    if let (Value::V128(src), Value::V128(idx)) = (src, idx) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            let n = idx[i];
                            out[i] = if n < 16 { src[n as usize] } else { 0 };
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F32X4_S => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let f = f32::from_le_bytes(bytes[i * 4..i * 4 + 4].try_into().unwrap());
                            let r = if f.is_nan() { 0 } else { f as i32 };
                            out[i * 4..i * 4 + 4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F32X4_U => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let f = f32::from_le_bytes(bytes[i * 4..i * 4 + 4].try_into().unwrap());
                            let r: u32 = if f.is_nan() || f < 0.0 { 0 } else { f as u32 };
                            out[i * 4..i * 4 + 4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F64X2_S_ZERO => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f = f64::from_le_bytes(bytes[i * 8..i * 8 + 8].try_into().unwrap());
                            let r = if f.is_nan() { 0 } else { f as i32 };
                            out[i * 4..i * 4 + 4].copy_from_slice(&r.to_le_bytes());
                        }
                        // Upper two lanes stay zero.
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F64X2_U_ZERO => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f = f64::from_le_bytes(bytes[i * 8..i * 8 + 8].try_into().unwrap());
                            let r: u32 = if f.is_nan() || f < 0.0 { 0 } else { f as u32 };
                            out[i * 4..i * 4 + 4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F32X4_RELAXED_MADD => {
                    let c = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap());
                            let fc = f32::from_le_bytes(vc[i * 4..i * 4 + 4].try_into().unwrap());
                            out[i * 4..i * 4 + 4]
                                .copy_from_slice(&fa.mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F32X4_RELAXED_NMADD => {
                    let c = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap());
                            let fc = f32::from_le_bytes(vc[i * 4..i * 4 + 4].try_into().unwrap());
                            out[i * 4..i * 4 + 4]
                                .copy_from_slice(&(-fa).mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_RELAXED_MADD => {
                    let c = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                            let fc = f64::from_le_bytes(vc[i * 8..i * 8 + 8].try_into().unwrap());
                            out[i * 8..i * 8 + 8]
                                .copy_from_slice(&fa.mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_RELAXED_NMADD => {
                    let c = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                            let fc = f64::from_le_bytes(vc[i * 8..i * 8 + 8].try_into().unwrap());
                            out[i * 8..i * 8 + 8]
                                .copy_from_slice(&(-fa).mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // -- Relaxed laneselect --
                // mask bit policy: we use the full bit (all 8 / 16 / 32
                // bits of the mask lane compared to 0) — picking the
                // "any non-zero bit" interpretation consistently.
                _ if op == Op::I8X16_RELAXED_LANESELECT => {
                    let mask = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = if vm[i] & 0x80 != 0 { va[i] } else { vb[i] };
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I16X8_RELAXED_LANESELECT => {
                    let mask = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let m = u16::from_le_bytes([vm[i * 2], vm[i * 2 + 1]]);
                            let pick_a = m & 0x8000 != 0;
                            out[i * 2..i * 2 + 2].copy_from_slice(if pick_a {
                                &va[i * 2..i * 2 + 2]
                            } else {
                                &vb[i * 2..i * 2 + 2]
                            });
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_RELAXED_LANESELECT => {
                    let mask = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let m = u32::from_le_bytes(vm[i * 4..i * 4 + 4].try_into().unwrap());
                            let pick_a = m & 0x8000_0000 != 0;
                            out[i * 4..i * 4 + 4].copy_from_slice(if pick_a {
                                &va[i * 4..i * 4 + 4]
                            } else {
                                &vb[i * 4..i * 4 + 4]
                            });
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I64X2_RELAXED_LANESELECT => {
                    let mask = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let m = u64::from_le_bytes(vm[i * 8..i * 8 + 8].try_into().unwrap());
                            let pick_a = m & 0x8000_0000_0000_0000 != 0;
                            out[i * 8..i * 8 + 8].copy_from_slice(if pick_a {
                                &va[i * 8..i * 8 + 8]
                            } else {
                                &vb[i * 8..i * 8 + 8]
                            });
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // -- Relaxed min / max --
                // NaN handling: relaxed variants are allowed to return
                // either operand on NaN input (vs MVP which must return
                // NaN). We pick `a` when `a` is NaN, `b` otherwise — the
                // x86 `minps/maxps` behavior.
                _ if op == Op::F32X4_RELAXED_MIN => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap());
                            let r = if fa < fb { fa } else { fb };
                            out[i * 4..i * 4 + 4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F32X4_RELAXED_MAX => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i * 4..i * 4 + 4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i * 4..i * 4 + 4].try_into().unwrap());
                            let r = if fa > fb { fa } else { fb };
                            out[i * 4..i * 4 + 4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_RELAXED_MIN => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                            let r = if fa < fb { fa } else { fb };
                            out[i * 8..i * 8 + 8].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::F64X2_RELAXED_MAX => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i * 8..i * 8 + 8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i * 8..i * 8 + 8].try_into().unwrap());
                            let r = if fa > fb { fa } else { fb };
                            out[i * 8..i * 8 + 8].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // -- q15 multiply-round-saturate (same semantics as MVP) --
                _ if op == Op::I16X8_RELAXED_Q15MULR_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let av =
                                i16::from_le_bytes(va[i * 2..i * 2 + 2].try_into().unwrap()) as i32;
                            let bv =
                                i16::from_le_bytes(vb[i * 2..i * 2 + 2].try_into().unwrap()) as i32;
                            let r = ((av * bv) + (1 << 14)) >> 15;
                            let r = r.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
                            out[i * 2..i * 2 + 2].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                // -- Relaxed integer dot products --
                // The i8x16 x i7x16 ops assume the second operand's high
                // bit is zero (7-bit). The "relaxed" part is that the
                // implementation may saturate or wrap — we wrap via i32.
                _ if op == Op::I16X8_RELAXED_DOT_I8X16_I7X16_S => {
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let av0 = va[i * 2] as i8 as i16;
                            let av1 = va[i * 2 + 1] as i8 as i16;
                            let bv0 = (vb[i * 2] & 0x7F) as i16;
                            let bv1 = (vb[i * 2 + 1] & 0x7F) as i16;
                            let sum = av0.wrapping_mul(bv0).wrapping_add(av1.wrapping_mul(bv1));
                            out[i * 2..i * 2 + 2].copy_from_slice(&sum.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                _ if op == Op::I32X4_RELAXED_DOT_I8X16_I7X16_ADD_S => {
                    let c = self.pop();
                    let b = self.pop();
                    let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let mut sum: i32 =
                                i32::from_le_bytes(vc[i * 4..i * 4 + 4].try_into().unwrap());
                            for j in 0..4 {
                                let av = va[i * 4 + j] as i8 as i32;
                                let bv = (vb[i * 4 + j] & 0x7F) as i32;
                                sum = sum.wrapping_add(av.wrapping_mul(bv));
                            }
                            out[i * 4..i * 4 + 4].copy_from_slice(&sum.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }

                // -- CM3 / WASI 0.3 async (Track B) --
                _ if op == Op::STREAM_READ => {
                    use crate::value::ObjectKind;
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Stream { id } = o.kind {
                            let stream_id = id;
                            drop(o);
                            let has_item = self.event_loop.borrow().stream_has_item(stream_id);
                            let is_eof = self.event_loop.borrow().stream_is_eof(stream_id);
                            if has_item {
                                let item = self
                                    .event_loop
                                    .borrow_mut()
                                    .stream_pop(stream_id)
                                    .unwrap_or(Value::Null);
                                self.push(item)?;
                            } else if is_eof {
                                self.push(Value::Null)?;
                            } else {
                                let fiber = self.save_fiber();
                                self.event_loop
                                    .borrow_mut()
                                    .suspend_stream_reader(stream_id, fiber);
                                return Err(VMError::new(format!("__stream_read__:{}", stream_id)));
                            }
                        } else {
                            drop(o);
                            self.push(Value::Null)?;
                        }
                    } else {
                        self.push(Value::Null)?;
                    }
                }

                _ if op == Op::STREAM_WRITE => {
                    use crate::value::ObjectKind;
                    let item = self.pop();
                    let val = self.pop();
                    // The stream is either the high-level Stream value or a
                    // CM3 writable-end i32 handle (canon stream.new pushes
                    // i32 handles per CanonicalABI §HandleTable).
                    let stream_id = match val {
                        Value::Object(ref obj) => {
                            let o = obj.lock().unwrap();
                            if let ObjectKind::Stream { id } = o.kind {
                                Some(id)
                            } else {
                                None
                            }
                        }
                        Value::I32(handle) => match self.handle_table.get(handle as u32) {
                            Some(crate::handle_table::HandleEntry::WritableStreamEnd(id)) => {
                                Some(*id)
                            }
                            _ => None,
                        },
                        _ => None,
                    };
                    if let Some(stream_id) = stream_id {
                        let mut el = self.event_loop.borrow_mut();
                        if let Some(fiber) = el.stream_push(stream_id, item) {
                            el.microtasks
                                .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                        }
                    }
                }

                _ if op == Op::STREAM_CANCEL_READ => {
                    use crate::value::ObjectKind;
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Stream { id } = o.kind {
                            let stream_id = id;
                            drop(o);
                            let mut el = self.event_loop.borrow_mut();
                            if let Some(fiber) = el.stream_close(stream_id) {
                                el.microtasks
                                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                            }
                        }
                    }
                }

                // ── CM3 Canonical ABI — Track A ─────────────────────────────────
                _ if op == Op::TASK_RETURN => {
                    // canon task.return — pop result, mark active task as Returned.
                    // A second task.return on the same task is a trap per spec.
                    let result = self.pop();
                    if let Some(task) = self.cm_tasks.last_mut() {
                        if !task.mark_returned() {
                            return Err(VMError::new(
                                "task.return called twice on same task (trap)",
                            ));
                        }
                    }
                    // Push the result back — the function body may continue running.
                    self.push(result)?;
                }

                _ if op == Op::TASK_CANCEL => {
                    // canon task.cancel — cancel the current task.
                    if let Some(task) = self.cm_tasks.last_mut() {
                        task.phase = crate::cm_task::TaskPhase::Returned;
                    }
                }

                _ if op == Op::SUBTASK_CANCEL => {
                    // canon subtask.cancel — pops subtask handle (i32), cancels the subtask.
                    let handle = self.pop().as_i32() as u32;
                    let fid = if let Some(crate::handle_table::HandleEntry::Subtask {
                        future_id,
                        ..
                    }) = self.handle_table.get(handle)
                    {
                        Some(*future_id)
                    } else {
                        None
                    };
                    if let Some(fid) = fid {
                        let mut el = self.event_loop.borrow_mut();
                        if let Some(fiber) =
                            el.reject_future(fid, Value::String(Arc::from("cancelled")))
                        {
                            el.microtasks
                                .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                        }
                    }
                }

                _ if op == Op::SUBTASK_DROP => {
                    // canon subtask.drop — pops subtask handle (i32), removes from handle table.
                    let handle = self.pop().as_i32() as u32;
                    self.handle_table.remove(handle);
                }

                _ if op == Op::WAITABLE_SET_NEW => {
                    // canon waitable-set.new — create a new waitable set, push its handle (i32).
                    let set_id = self.waitable_sets.create();
                    self.push(Value::I32(set_id as i32))?;
                }

                _ if op == Op::WAITABLE_JOIN => {
                    // canon waitable.join — pops [waitable_handle_i32, set_handle_i32];
                    // looks up waitable in handle table, adds to set.
                    let set_handle = self.pop().as_i32() as u32;
                    let waitable_handle = self.pop().as_i32() as u32;
                    let waitable = match self.handle_table.get(waitable_handle) {
                        Some(crate::handle_table::HandleEntry::ReadableStreamEnd(sid)) => {
                            Some(crate::waitable::Waitable::Stream(*sid))
                        }
                        Some(crate::handle_table::HandleEntry::ReadableFutureEnd(fid)) => {
                            Some(crate::waitable::Waitable::Future(*fid))
                        }
                        Some(crate::handle_table::HandleEntry::Subtask { future_id, .. }) => {
                            Some(crate::waitable::Waitable::Subtask(*future_id))
                        }
                        _ => None,
                    };
                    if let Some(w) = waitable {
                        if let Some(set) = self.waitable_sets.get_mut(set_handle) {
                            set.join(w);
                        }
                    }
                }

                _ if op == Op::WAITABLE_SET_WAIT => {
                    // canon waitable-set.wait — pops [set_handle_i32, memory_ptr_i32];
                    // writes (event_code, handle_id, 0) to memory; pushes event_code (i32).
                    // If nothing is ready, returns NONE immediately (MVP — true blocking TBD).
                    let memory_ptr = self.pop().as_i32() as usize;
                    let set_handle = self.pop().as_i32() as u32;
                    let ready = {
                        let el = self.event_loop.borrow();
                        self.waitable_sets
                            .get(set_handle)
                            .and_then(|set| set.poll_ready(&el))
                    };
                    let (code, handle_id) = ready.unwrap_or((crate::waitable::EventCode::None, 0));
                    if memory_ptr + 12 <= self.memory.len() {
                        self.memory.store_i32(memory_ptr, code as i32)?;
                        self.memory.store_i32(memory_ptr + 4, handle_id as i32)?;
                        self.memory.store_i32(memory_ptr + 8, 0)?;
                    }
                    self.push(Value::I32(code as i32))?;
                }

                _ if op == Op::WAITABLE_SET_POLL => {
                    // canon waitable-set.poll — non-blocking version of WAITABLE_SET_WAIT.
                    // Pushes EventCode::None (0) immediately if nothing is ready.
                    let memory_ptr = self.pop().as_i32() as usize;
                    let set_handle = self.pop().as_i32() as u32;
                    let ready = {
                        let el = self.event_loop.borrow();
                        self.waitable_sets
                            .get(set_handle)
                            .and_then(|set| set.poll_ready(&el))
                    };
                    let (code, handle_id) = ready.unwrap_or((crate::waitable::EventCode::None, 0));
                    if memory_ptr + 12 <= self.memory.len() {
                        self.memory.store_i32(memory_ptr, code as i32)?;
                        self.memory.store_i32(memory_ptr + 4, handle_id as i32)?;
                        self.memory.store_i32(memory_ptr + 8, 0)?;
                    }
                    self.push(Value::I32(code as i32))?;
                }

                _ if op == Op::STREAM_NEW => {
                    // canon stream.new — create a stream; push readable_handle and writable_handle (i32).
                    let stream_id = self.event_loop.borrow_mut().create_stream();
                    let rd = self.handle_table.insert(
                        crate::handle_table::HandleEntry::ReadableStreamEnd(stream_id),
                    );
                    let wr = self.handle_table.insert(
                        crate::handle_table::HandleEntry::WritableStreamEnd(stream_id),
                    );
                    self.push(Value::I32(rd as i32))?;
                    self.push(Value::I32(wr as i32))?;
                }

                _ if op == Op::STREAM_DROP_RD => {
                    // canon stream.drop-readable — pops readable stream handle (i32).
                    let handle = self.pop().as_i32() as u32;
                    if let Some(crate::handle_table::HandleEntry::ReadableStreamEnd(sid)) =
                        self.handle_table.remove(handle)
                    {
                        // Close the stream so waiting writers don't block forever.
                        let mut el = self.event_loop.borrow_mut();
                        if let Some(fiber) = el.stream_close(sid) {
                            el.microtasks
                                .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                        }
                    }
                }

                _ if op == Op::STREAM_DROP_WR => {
                    // canon stream.drop-writable — pops writable stream handle (i32).
                    let handle = self.pop().as_i32() as u32;
                    if let Some(crate::handle_table::HandleEntry::WritableStreamEnd(sid)) =
                        self.handle_table.remove(handle)
                    {
                        // Closing the write end signals EOF to the reader.
                        let mut el = self.event_loop.borrow_mut();
                        if let Some(fiber) = el.stream_close(sid) {
                            el.microtasks
                                .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                        }
                    }
                }

                _ if op == Op::FUTURE_NEW => {
                    // canon future.new — create a future; push readable_handle and writable_handle (i32).
                    let future_id = self.event_loop.borrow_mut().create_future();
                    let rd = self.handle_table.insert(
                        crate::handle_table::HandleEntry::ReadableFutureEnd(future_id),
                    );
                    let wr = self.handle_table.insert(
                        crate::handle_table::HandleEntry::WritableFutureEnd(future_id),
                    );
                    self.push(Value::I32(rd as i32))?;
                    self.push(Value::I32(wr as i32))?;
                }

                _ if op == Op::FUTURE_DROP_RD => {
                    // canon future.drop-readable — pops readable future handle (i32).
                    let handle = self.pop().as_i32() as u32;
                    self.handle_table.remove(handle);
                }

                _ if op == Op::FUTURE_DROP_WR => {
                    // canon future.drop-writable — pops writable future handle (i32).
                    let handle = self.pop().as_i32() as u32;
                    if let Some(crate::handle_table::HandleEntry::WritableFutureEnd(fid)) =
                        self.handle_table.remove(handle)
                    {
                        // Dropping the write end without resolving rejects the future.
                        let mut el = self.event_loop.borrow_mut();
                        if let Some(fiber) =
                            el.reject_future(fid, Value::String(Arc::from("future dropped")))
                        {
                            el.microtasks
                                .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                        }
                    }
                }

                _ if op == Op::BACKPRESSURE_SET => {
                    // canon backpressure.set — pops enabled_i32, sets/clears backpressure on active task.
                    let enabled = self.pop().as_i32() != 0;
                    if let Some(task) = self.cm_tasks.last_mut() {
                        task.backpressure = enabled;
                    }
                }

                _ if op == Op::CONTEXT_GET => {
                    // canon context.get — pops index_i32, pushes context slot value.
                    let index = self.pop().as_i32() as usize;
                    let val = self
                        .context_slots
                        .get(index)
                        .cloned()
                        .unwrap_or(Value::Undefined);
                    self.push(val)?;
                }

                _ if op == Op::CONTEXT_SET => {
                    // canon context.set — pops [value, index_i32], sets context slot.
                    let index = self.pop().as_i32() as usize;
                    let val = self.pop();
                    if index >= self.context_slots.len() {
                        self.context_slots.resize(index + 1, Value::Undefined);
                    }
                    self.context_slots[index] = val;
                }

                // -- WASM GC Type System --
                _ if op == Op::SET_TYPE_ID => {
                    // Stack: [obj, type_id_i32] → [obj]
                    let type_id = self.pop().as_i32() as usize;
                    // Stamping a null yields a WASM GC TYPED null (`ref.null $t`):
                    // it behaves like a plain null everywhere except the GC
                    // accessors, which trap on it per spec.
                    if self.peek(0).is_null_ref() {
                        self.pop();
                        self.push(Value::TypedNull(type_id))?;
                        continue;
                    }
                    let obj = self.peek(0);
                    if let Value::Object(o) = obj {
                        let mut obj_mut = o.lock().unwrap();
                        obj_mut.type_id = type_id;

                        // Populate __nonenum with non-enumerable field names from TypeDef
                        if let Some(td) = self.type_registry.get(type_id) {
                            let nonenum_fields: Vec<Value> = td
                                .field_defs
                                .iter()
                                .filter(|f| !f.descriptor.enumerable)
                                .map(|f| Value::String(Arc::from(f.name.as_str())))
                                .collect();

                            if !nonenum_fields.is_empty() {
                                let nonenum_arr = Object::new_array(nonenum_fields);
                                obj_mut.properties.insert(
                                    "__nonenum".to_string(),
                                    Value::Object(Arc::new(Mutex::new(nonenum_arr))),
                                );
                            }
                        }
                    }
                }

                // -- Iteration protocol --
                // iter_get, iter_next: removed (non-WASM, were unused by compilers)
                // class_new, method_def, inherit: removed (non-WASM, were NOPs)
                _ => {
                    return Err(VMError::new(format!(
                        "Unhandled opcode: {:?} (0x{:04X})",
                        op, op.0
                    )));
                }
            }
        }
    }
}
