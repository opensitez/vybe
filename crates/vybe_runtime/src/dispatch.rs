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
use crate::opcode::heaptype::HeapType;
use crate::opcode::{Op, read_leb_u32};
use crate::value::{Function, Object, ObjectKind, TypedArrayState, TypedElemKind, Upvalue, Value};
use crate::vm::{
    ActiveContinuation, BlockTargets, ExceptionHandler, ImportTarget, LabelEntry, ResumeMode, VM,
};
use std::collections::HashMap;

// ── Task protocol ───────────────────────────────────────────────────────────
//
// The VM's language-neutral vocabulary for a task object. A plugin stamps these
// from its own adapter and translates them back into its surface names
// (`platforms/dotnet` maps `__state` onto `TaskStatus`). The VM never spells a
// language's status names itself.

/// Settled state of a task, in the same vocabulary the promise objects use.
pub(crate) const TASK_STATE: &str = "__state";
pub(crate) const TASK_STATE_FULFILLED: &str = "fulfilled";
pub(crate) const TASK_STATE_REJECTED: &str = "rejected";
/// Base of a spawned task's `wasi:threads` argument block. The STATUS WORD is
/// at `+ 4` — 1 done, 2 faulted, 0 still running — stamped and notified by the
/// child at thread exit. It is the single source of truth for a thread's
/// outcome; nothing mirrors it.
pub(crate) const TASK_FUTEX: &str = "__futex";
/// The cancellation token a task carries, and the flag the VM reads off it.
pub(crate) const TASK_CANCEL_TOKEN: &str = "__cancel_token";
pub(crate) const TASK_CANCELLED: &str = "__cancelled";

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

    /// Resolve a type immediate (`struct.new`, `array.new`, …) to the
    /// instance rtt.
    ///
    /// The immediate is a 1-based index into the running module's own type
    /// index space, exactly as the spec addresses types; `0` means "no GC
    /// type" (rtt `0` = `Object`), which is what dynamic-language allocations
    /// carry. The mapping to registry ids was bound once at load
    /// (`bind_module_type_ids`) — no name is looked up here, so two modules
    /// can define same-named types without colliding at the instruction.
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
            properties: indexmap::IndexMap::new(),
            kind: ObjectKind::Function(func),
            type_id: 0,
            fields: Vec::new(),
        };
        let table_idx = self.func_table.len();
        obj.properties
            .insert("__table_idx".into(), Value::F64(table_idx as f64));
        let func_val = Value::Object(crate::heap::alloc(obj));
        self.func_table.push(func_val.clone());
        self.funcref_cache.insert(func_idx, func_val.clone());
        func_val
    }

    /// `wasi:threads`.`thread-spawn(start_arg) -> tid` — the VM is the
    /// embedder-side implementation of the wasi-threads import (as wasmtime
    /// is). `start_arg` points at `{fn_table_index: i32, status_word: i32}`
    /// in shared linear memory (the wasi-libc `pthread_create` pattern).
    /// The child invokes the module's `__wasi_thread_start` chunk with
    /// `(tid, start_arg)`; that dispatcher runs `table[fn_table_index]`
    /// and stamps + notifies the status word. There is NO thread opcode —
    /// this import is the entire surface.
    pub(crate) fn wasi_thread_spawn(&mut self, start_arg: i32) -> Result<i32, VMError> {
        let fn_idx = self.memory.atomic_load_i32(start_arg as usize) as usize;
        let status_addr = start_arg as usize + 4;
        // The start function travels through funcref table 0 (spec object);
        // close Open upvalues against the SPAWNING stack before it crosses —
        // the child VM has a fresh stack (wasi-threads has no shared stack).
        let function = {
            let val = self
                .wasm_tables
                .first()
                .and_then(|t| t.get(fn_idx))
                .cloned()
                .ok_or_else(|| {
                    VMError::new("wasi:threads/thread-spawn: start_arg names no table entry")
                })?;
            match &val {
                Value::Object(obj) => {
                    let o = obj.lock().unwrap();
                    match &o.kind {
                        ObjectKind::Function(f) => {
                            for uv in &f.upvalues {
                                let mut u = uv.lock().unwrap();
                                if let crate::value::UpvalueLocation::Open(slot) = u.location {
                                    let v = self.stack.get(slot).cloned().unwrap_or(Value::Null);
                                    u.location = crate::value::UpvalueLocation::Closed(v);
                                }
                            }
                            f.clone()
                        }
                        _ => {
                            return Err(VMError::new(
                                "wasi:threads/thread-spawn: table entry is not a function",
                            ));
                        }
                    }
                }
                _ => {
                    return Err(VMError::new(
                        "wasi:threads/thread-spawn: table entry is not a function",
                    ));
                }
            }
        };

        let tid = self.next_thread_id;
        self.next_thread_id += 1;

        let child_chunks = self.chunks.clone();
        let child_memory = self.memory.clone();
        let child_host_fns = self.host_fns.clone();
        let child_host_registry = self.host_registry.clone();
        let child_import_table = self.import_table.clone();
        let child_globals = self.globals.clone();
        // `globals_assigned` is PART OF the global store, not a side table:
        // index i of one is index i of the other. Cloning `globals` without it
        // left the child on `VM::new`'s `vec![true]` (len 1), so every
        // `GLOBAL_SET` at idx ≥ 1 in a spawned thread panicked — and because
        // the worker died mid-flight, `pthread_join` waited forever and it
        // surfaced as a TIMEOUT rather than a crash.
        let child_globals_assigned = self.globals_assigned.clone();
        let child_type_registry = self.type_registry.clone();
        let child_func_table = self.func_table.clone();
        let child_wasm_tables = self.wasm_tables.clone();
        let child_case_aliases = self.case_aliases.clone();
        // Register BEFORE the thread runs so a waiter never observes `live`
        // short (all-asleep deadlock detection).
        child_memory.thread_started();

        let handle = std::thread::spawn(move || {
            let mut child_vm = VM::new();
            child_vm.chunks = child_chunks;
            child_vm.memory = child_memory;
            child_vm.host_fns = child_host_fns;
            child_vm.host_registry = child_host_registry;
            child_vm.import_table = child_import_table;
            child_vm.globals = child_globals;
            child_vm.globals_assigned = child_globals_assigned;
            child_vm.type_registry = child_type_registry;
            child_vm.func_table = child_func_table;
            child_vm.wasm_tables = child_wasm_tables;
            child_vm.case_aliases = child_case_aliases;
            // record[+8] = user_arg → the start function's slot-0 parameter
            // (arity-0 closures leave it in an unread slot, harmless — the
            // same contract the old opcode documented).
            let user_arg = child_vm.memory.atomic_load_i32(start_arg as usize + 8);
            child_vm.stack.push(Value::I32(user_arg));
            let result = child_vm
                .call_function(&function, 1)
                .and_then(|_| child_vm.execute());
            let ok = match result {
                Ok(_) => true,
                Err(e) => {
                    eprintln!("[thread {}] error: {}", tid, e.message);
                    false
                }
            };
            // Thread-exit contract (embedder side, like wasi-threads hosts):
            // stamp the status word — 1 done, 2 faulted — and wake joiners.
            child_vm
                .memory
                .atomic_store_i32(status_addr, if ok { 1 } else { 2 });
            child_vm.memory.notify(status_addr, i32::MAX);
            child_vm.memory.thread_exited();
            if ok { vec![0u8] } else { vec![1u8] }
        });
        self.thread_handles.insert(tid, handle);
        Ok(tid)
    }

    /// Pop an i32 operand that the spec reads as UNSIGNED.
    ///
    /// ⛔ CLAMPING A NEGATIVE TO ZERO SUPPRESSES THE TRAP IT SHOULD CAUSE.
    /// Every offset, index and count in the bulk and GC-array instructions is
    /// an unsigned i32, so `0x8000_0000` means 2147483648, not "negative, call
    /// it 0" — and calling it 0 turns the most out-of-bounds request there is
    /// into the most in-bounds one. `gc/array.wast`'s `new-overflow` asks for
    /// `0x8000_0000` elements at offset `0x8000_0000` and must trap; clamped,
    /// it allocated an empty array and returned.
    fn pop_u32_operand(&mut self) -> usize {
        self.pop().as_i32() as u32 as usize
    }

    pub(crate) fn resolve_gc_rtt(&self, type_imm: usize) -> usize {
        if type_imm == 0 {
            return 0;
        }
        // Relative to the module that emitted the instruction — the executing
        // chunk names it.
        let base = self
            .frames
            .last()
            .and_then(|frame| self.chunk_type_base.get(frame.chunk_index))
            .copied()
            .unwrap_or(0);
        self.module_type_ids
            .get(base + type_imm - 1)
            .copied()
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
    Value::Object(crate::heap::alloc(obj))
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
            // Skip the blocktype immediates format-driven (U8_U8: params +
            // results) so this walk can never disagree with the declaration.
            ip += op.operand_format().size_in(code, ip);
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

/// Where a custom descriptor lives on an instance.
///
/// A reserved property key, not a user-reachable name: the descriptor is an
/// engine-managed RTT reference in the proposal, so nothing in source may
/// read or overwrite it. Kept as a single constant with one writer
/// (`set_descriptor`) and one reader (`descriptor_of`) so the allocation and
/// cast paths can never disagree about the location.
///
/// ⚠ This is a string-keyed property, which is the name-keyed representation
/// `flexclassplan` §0-bis is driving to zero. It stays for now because moving
/// it to a dedicated `Object` field changes a shared `vybe_runtime` struct.
const DESCRIPTOR_SLOT: &str = "__descriptor";

/// Attach a custom descriptor to a freshly allocated instance.
fn set_descriptor(obj: &mut Object, descriptor: Value) {
    obj.properties.insert(DESCRIPTOR_SLOT.into(), descriptor);
}

/// The custom descriptor attached to a reference, or `Null` when it has none.
///
/// This is the single reader used by `ref.get_desc` and by the
/// descriptor-comparing casts.
fn descriptor_of(value: &Value) -> Value {
    match value {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(DESCRIPTOR_SLOT)
            .cloned()
            .unwrap_or(Value::Null),
        _ => Value::Null,
    }
}

/// `ref.eq` — reference identity. Shared by the `REF_EQ` opcode and by the
/// descriptor-comparing casts, which the proposal defines in terms of
/// descriptor *equality*, i.e. the same reference.
fn ref_eq(a: &Value, b: &Value) -> bool {
    match (a, b) {
        // All nulls (typed or plain) are ref.eq.
        _ if a.is_null_ref() && b.is_null_ref() => true,
        (Value::Undefined, Value::Undefined) => true,
        // Pure WASM `ref.eq`: reference identity only. Two `ref.func $f`
        // tear-offs of the same capture-free function are identical because
        // `REF_FUNC` INTERNS them (one canonical object per function) —
        // identity is established at creation, not faked here. Closures with
        // captures stay distinct, as do bound methods (which capture `self`),
        // so `C().f is C().f` is correctly false.
        (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
        (Value::Symbol(a), Value::Symbol(b)) => Arc::ptr_eq(a, b),
        (Value::String(a), Value::String(b)) => Arc::ptr_eq(a, b),
        // An `i31ref` is UNBOXED — `I31_NEW` masks to 31 bits and pushes a
        // plain `I32`, with no allocation to take the address of. Its value
        // therefore IS its identity: `ref.eq (ref.i31 7) (ref.i31 7)` is 1,
        // where the pointer-identity arms above would answer 0 because there
        // is no pointer. Validation restricts `ref.eq` to `eqref` operands, so
        // a plain integer never reaches this arm.
        (Value::I32(a), Value::I32(b)) => a == b,
        _ => false,
    }
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
            // Bit-preserving, like `f32.store` — no f64 round-trip, which would
            // quiet a signalling NaN.
            let bytes = value.as_f32().to_le_bytes();
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

    /// Pop the current frame the way a tail call must: a `return_call` REPLACES
    /// the frame, so that frame's structured state dies with it — its labels,
    /// and, spec-visibly, its `try_table` handlers.
    ///
    /// `frames.pop()` alone left the caller's catch clauses ARMED for the
    /// callee, so a tail call out of a protected body had its callee's throw
    /// caught by a handler that no longer existed (`try_table.wast`
    /// `return-call-in-try-catch`, which asserts the throw ESCAPES). This is the
    /// same cleanup `Op::RETURN` does; only the tail-call arms were missing it.
    fn pop_frame_for_tail_call(&mut self) {
        let frame_label_base = self.frame().label_base;
        self.label_stack.truncate(frame_label_base);
        self.frames.pop();
        if !self.frames.is_empty() {
            let live = self.frames.len();
            self.exception_handlers.retain(|h| h.frame_depth <= live);
        }
    }

    fn suspend_for_pending_promise(&mut self, promise_id: u64) -> VMError {
        let fiber = self.save_fiber();
        self.event_loop
            .borrow_mut()
            .suspend_fiber(promise_id, fiber);
        VMError::new(format!("__jspi__:{}", promise_id))
    }

    /// Park the current fiber inside a SYNCHRONOUS `canon stream.read` /
    /// `future.read` — `CanonicalABI.md` §`canon stream.{read,write}`, where
    /// only the `async` variant may hand `BLOCKED` back to the guest.
    ///
    /// The copy travels WITH the fiber (`Fiber::pending_copy`) and is redone by
    /// `resume_fiber` once a producer appears, because by now the guest's read
    /// instruction has retired: its operands are popped and its result slot is
    /// the one value a resume pushes. Nothing about the copy is decided here —
    /// only that it has not happened yet.
    ///
    /// The end stays in `COPYING` exactly as the BLOCKED path left it: that is
    /// what makes a `cancel-read` legal while parked and a second concurrent
    /// read trap.
    /// Give a stream's registered producer a chance to supply elements.
    ///
    /// Returns true if anything landed, in which case the caller retries the
    /// copy instead of parking. This is what makes a stream whose elements
    /// arrive over time — an inbound-connection stream — readable at all: the
    /// call that created it returned long ago, so the reader is the only one
    /// left to ask.
    ///
    /// The producer is named by the host that created the stream
    /// (`EventLoop::set_stream_producer`) and called through the ordinary host
    /// registry, so the VM carries no knowledge of what any particular stream
    /// produces.
    pub(crate) fn run_stream_producer(&mut self, stream_id: u64) -> bool {
        let Some((module, name)) = self.event_loop.borrow().stream_producer(stream_id) else {
            return false;
        };
        let Some(idx) = self.host_registry.get(&(module, name)).copied() else {
            return false;
        };
        let host_fn = self.host_fns[idx].clone();
        let args = [Value::F64(stream_id as f64)];
        {
            let mut ctx = self.make_host_context();
            host_fn(&mut ctx, &args);
        }
        // Ask the buffer, not the producer's return value: a producer pushes
        // through `HostContext`, and a stream that reached EOF has an answer
        // too (`DROPPED`), which is equally a reason not to park.
        let el = self.event_loop.borrow();
        el.stream_has_bytes(stream_id) || el.stream_has_item(stream_id) || el.stream_is_eof(stream_id)
    }

    fn park_sync_copy(&mut self, pending: crate::fiber::PendingCopy) -> VMError {
        let end_id = pending.end_id;
        let is_future = matches!(pending.kind, crate::fiber::PendingCopyKind::Future(_));
        let mut fiber = self.save_fiber();
        fiber.pending_copy = Some(pending);
        let mut el = self.event_loop.borrow_mut();
        if is_future {
            el.suspend_future_sync_reader(end_id, fiber);
        } else {
            el.suspend_stream_sync_reader(end_id, fiber);
        }
        drop(el);
        VMError::new(format!("__stream_read__:{}", end_id))
    }

    /// Top-level settled/plain-value await (no promising boundary, not inside
    /// a driven continuation): ECMA-262 §6.2.3.1 still requires one job tick.
    /// Save the whole fiber exactly like a pending top-level await and wake it
    /// immediately off the ready queue with the value (or rejection).
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
            el.immediate
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
    /// Does an await on an ALREADY-SETTLED value still cost a turn? That is
    /// the `eager` parameter — and it was decided at COMPILE time, not here:
    /// `jspi.await` (ECMA-262 §6.2.3.1 / §27.7.5.3: `Await` performs
    /// `PromiseResolve` and resumes as a JOB, so even `await 1` yields one
    /// tick) passes `eager = false`; `jspi.await_eager` (.NET's contract: a
    /// Task whose antecedent is already complete may run its continuation
    /// synchronously) passes `eager = true`. The walker normalized the
    /// language's `await` to one of two AST operations, the compiler lowered
    /// each to its own import, and the VM implements both — it never consults
    /// a per-module property to pick semantics.
    fn do_await(&mut self, val: Value, eager: bool) -> Result<(), VMError> {
        let mut val = val;
        loop {
            // Clone the Arc so `val` is free to be reassigned for the next
            // flatten iteration without a borrow conflict.
            let arc = match &val {
                Value::Object(o) => o.clone(),
                // Primitive: ECMA-262 §6.2.3.1 Await performs
                // PromiseResolve(v) and ALWAYS resumes as a job — one
                // turn even for plain values. Inside an async
                // boundary, suspend and schedule the immediate resume; at
                // top level (no boundary) keep the direct return.
                _ => {
                    if !self.async_floors.is_empty() && !eager {
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
                return self.await_task_object(arc, eager);
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
                if !self.async_floors.is_empty() && !eager {
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
                // fiber off the ready queue (its captured try/catch handlers fire
                // there). No synchronous shortcut.
                if !self.async_floors.is_empty() && !eager {
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
            // an immediately-ready resume with the fulfilled value — the
            // spec "await always yields one tick" ordering, no sync shortcut.
            if !self.async_floors.is_empty() && !eager {
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

    fn await_task_object(
        &mut self,
        task_obj: Arc<Mutex<Object>>,
        eager: bool,
    ) -> Result<(), VMError> {
        self.join_task_object_if_needed(&task_obj);

        let token_cancelled = self.task_token_cancelled(&task_obj);
        let (state, result, exception) = {
            let task = task_obj.lock().unwrap();
            let state = task
                .properties
                .get(TASK_STATE)
                .map(|v| format!("{}", v))
                .unwrap_or_default();
            let result = task
                .properties
                .get("result")
                .cloned()
                .unwrap_or(Value::Null);
            let exception = task.properties.get("exception").cloned();
            (state, result, exception)
        };

        // ⛔ SETTLED STATE, NOT A LANGUAGE'S STATUS NAMES. This branched on
        // `Faulted`/`Canceled` — `System.Threading.Tasks.TaskStatus` members —
        // which made the VM's await path .NET-shaped. `__state` is the same
        // pending/fulfilled/rejected vocabulary the promise objects already
        // carry, and a plugin maps its own status names onto it (dotnet's
        // `emit_task_status` does exactly that).
        let faulted = token_cancelled
            || state.eq_ignore_ascii_case(TASK_STATE_REJECTED)
            || exception
                .as_ref()
                .is_some_and(|v| !matches!(v, Value::Null | Value::Undefined));

        if faulted {
            let reason = exception.unwrap_or_else(|| {
                if token_cancelled {
                    make_operation_cancelled_error()
                } else {
                    Value::String(Arc::from("Task faulted"))
                }
            });
            if !self.async_floors.is_empty() && !eager {
                let id = self.event_loop.borrow_mut().next_promise_id();
                self.pending_settled_await = Some((id, reason, true));
                return Err(VMError::new(format!("__jspi__:{}", id)));
            }
            if self.top_level_await_ticks() {
                return Err(self.tick_top_level_await(reason, true));
            }
            return self.raise_exception_value(reason);
        }

        if !self.async_floors.is_empty() && !eager {
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

    /// Whether the task carries a cancellation token that has been triggered.
    ///
    /// The token is read LIVE: it can be cancelled after the task was created,
    /// which is the whole point of handing one to a delay. Both keys belong to
    /// the VM's task protocol, not to any language — a plugin that wants a task
    /// cancellable stamps them from its own adapter.
    fn task_token_cancelled(&self, task_obj: &Arc<Mutex<Object>>) -> bool {
        let token = {
            let task = task_obj.lock().unwrap();
            task.properties.get(TASK_CANCEL_TOKEN).cloned()
        };
        let Some(Value::Object(token_obj)) = token else {
            return false;
        };
        let token = token_obj.lock().unwrap();
        token
            .properties
            .get(TASK_CANCELLED)
            .is_some_and(|value| matches!(value, Value::Bool(true)))
    }

    fn join_task_object_if_needed(&mut self, task_obj: &Arc<Mutex<Object>>) {
        let (tid, futex) = {
            let task = task_obj.lock().unwrap();
            (
                task.properties
                    .get("__thread_id")
                    .map(|v| v.as_f64() as i32)
                    .unwrap_or(-1),
                task.properties
                    .get(TASK_FUTEX)
                    .map(|v| v.as_f64() as usize),
            )
        };
        if let Some(handle) = self.thread_handles.remove(&tid) {
            self.memory.mark_parked();
            let joined_ok = match handle.join() {
                Ok(result) => result.first().copied().unwrap_or(1) == 0,
                Err(_) => false,
            };
            self.memory.unmark_parked();
            // ⛔ THE STATUS IS THE `wasi:threads` STATUS WORD, NOT A SECOND
            // COPY OF IT. At thread exit the child stamps `__futex + 4` — 1
            // done, 2 faulted — and notifies joiners; that is the proposal's
            // own thread-exit contract and what `__stdlib_task_wait` waits on.
            // Reading it here leaves ONE status in the system. The OS handle
            // join only waits for the thread to be gone; a thread whose object
            // carries no futex (nothing spawned it through the wasi path)
            // falls back to what the join answered.
            let ok = futex
                .map(|base| self.memory.atomic_load_i32(base + 4) == 1)
                .unwrap_or(joined_ok);
            let mut task = task_obj.lock().unwrap();
            task.properties
                .insert("iscompleted".into(), Value::Bool(true));
            task.properties.insert("isalive".into(), Value::Bool(false));
            task.properties.insert(
                TASK_STATE.into(),
                Value::String(Arc::from(if ok {
                    TASK_STATE_FULFILLED
                } else {
                    TASK_STATE_REJECTED
                })),
            );
        }
    }

    /// The element type of the canonical built-in currently executing, from
    /// the `$t` immediate its `canon` definition carried.
    ///
    /// Refuses rather than defaults. A built-in that needs the type needs the
    /// REAL one: substituting a plausible width would move the wrong number of
    /// bytes into a peer component's memory, and nothing downstream could
    /// detect it. An error names the missing declaration instead.
    /// The four 🧵 compound handoffs — `CanonicalABI.md`
    /// `Thread.{suspend,yield}_then_{resume,promote}`.
    ///
    /// ```python
    /// def suspend_then_resume(self, cancellable, other):
    ///   assert(self.running() and other.suspended())
    ///   return self.switch_to_internal(cancellable, other)
    ///
    /// def yield_then_resume(self, cancellable, other):
    ///   self.start_waiting_internal(lambda: True)      # stay runnable
    ///   return self.switch_to_internal(cancellable, other)
    ///
    /// def suspend_then_promote(self, cancellable, other):
    ///   if other.ready():
    ///     other.stop_waiting_internal(cancelled = False)
    ///     return self.suspend_then_resume(cancellable, other)
    ///   else:
    ///     return self.suspend(cancellable)             # FALL BACK
    /// ```
    ///
    /// `yields` = stay runnable rather than park. `promotes` = only switch if
    /// the target is READY, otherwise degrade to a plain suspend/yield instead
    /// of trapping — that fallback is the whole difference between `resume`
    /// and `promote`, and it is why `promote` does NOT require the target to
    /// be suspended up front.
    fn exec_thread_handoff(
        &mut self,
        b: crate::vm::CanonBuiltin,
        yields: bool,
        promotes: bool,
    ) -> Result<(), VMError> {
        let who = b.spec_name();
        let cancellable = self.canon_cancellable();
        let target_idx = self.pop().as_i32() as u32;

        let me = self.current_thread.ok_or_else(|| {
            VMError::new(format!(
                "canon {who}: no current thread — a thread exists only inside a \
                 `canon lift`ed call (trap)"
            ))
        })?;
        // ⛔ NO explicit self-target check. The spec's two families disagree
        // about it, and BOTH already fall out of the checks below:
        //
        //   resume : `assert(self.running() and other.suspended())` — self is
        //            running, therefore not suspended, so `!other.suspended()`
        //            below traps. Same answer, from the spec's own condition.
        //   promote: `assert(self.running())` then `if other.ready()` — self is
        //            running, therefore not WAITING, therefore not ready, so it
        //            takes the else branch and plain suspends/yields.
        //
        // An unconditional trap here (what this did) was wrong for the two
        // promote rows: it refused a handoff the spec defines as a fallback.

        let other = self.cm_instance.threads.get(target_idx).ok_or_else(|| {
            VMError::new(format!("canon {who}: no thread at index {target_idx} (trap)"))
        })?;

        // `promote` consults READINESS; `resume` demands SUSPENDED.
        let switch = if promotes {
            let ready = other.ready();
            if ready {
                self.cm_instance
                    .threads
                    .get_mut(target_idx)
                    .expect("just read")
                    .stop_waiting_internal(false)
                    .map_err(|e| VMError::new(format!("canon {who}: {e} (trap)")))?;
            }
            ready
        } else {
            if !other.suspended() {
                return Err(VMError::new(format!(
                    "canon {who}: thread {target_idx} is {:?}, must be Suspended (trap)",
                    other.state()
                )));
            }
            true
        };

        // Cancellation is delivered HERE — after `promote` has already
        // de-waited its target, before this thread parks. The spec's ordering
        // is `other.ready()` -> `stop_waiting_internal` -> the nested
        // `suspend_then_resume`/`yield_then_resume`, which is where the
        // `deliver_pending_cancel` lives. Delivering earlier would leave a
        // promoted target still waiting; delivering later would park a thread
        // that should have returned.
        if self.deliver_pending_cancel_now(cancellable) {
            self.push(Value::I32(1))?;
            return Ok(());
        }

        if !switch {
            // `promote` on a target that is not ready degrades to plain
            // `suspend`/`yield_` — and those are NOT the same thing.
            if yields {
                // `yield_` is `wait_until(lambda: True)`, whose readiness
                // condition already holds, so returning without switching is
                // a choice the spec explicitly permits the embedder to make.
                self.push(Value::I32(0))?;
                return Ok(());
            }
            // `suspend` has no such early return: it BLOCKS. Returning here
            // (what this did) turned `suspend-then-promote` into a no-op
            // whenever the target happened not to be ready.
            return self.thread_block(who, me);
        }

        // `yield_then_*` leaves this thread READY rather than parked. After the
        // cancel check, per `yield_then_resume`.
        if yields {
            let me_t = self.cm_instance.threads.get_mut(me).expect("current thread");
            // `start_waiting_internal` asserts we are not already waiting; the
            // running thread never is, so a failure here is a real state bug.
            me_t.start_waiting_internal(crate::cm_thread::ReadyWhen::Always)
                .map_err(|e| VMError::new(format!("canon {who}: {e} (trap)")))?;
        }

        self.park_and_switch_to(who, me, target_idx)?;
        // Reached only when this thread is resumed again; it was not cancelled
        // on the way in (checked above), so `Cancelled.FALSE`.
        self.push(Value::I32(0))?;
        Ok(())
    }

    /// Resolve `$ftbl[fi]` for the table rows (`thread.new-indirect`,
    /// `thread.spawn-indirect`). `$ftbl` is an IMMEDIATE on the canonical
    /// definition, which is why these rows were unreachable before the canon
    /// section existed to carry it.
    fn thread_table_funcref(&mut self, who: &str, fi: i32) -> Result<Value, VMError> {
        let def = self.canon_def_required(who)?;
        let tableidx = def.table.ok_or_else(|| {
            VMError::new(format!(
                "canon {who}: definition carries no $ftbl immediate"
            ))
        })? as usize;
        let table = self.table_ref(tableidx).ok_or_else(|| {
            VMError::new(format!("canon {who}: unknown table {tableidx} (trap)"))
        })?;
        if fi < 0 || fi as usize >= table.len() {
            return Err(VMError::new(format!(
                "canon {who}: table index {fi} out of bounds (len {}) (trap)",
                table.len()
            )));
        }
        Ok(table[fi as usize].clone())
    }

    /// `canon_thread_resume_later(i)` — the second half of the spawn fusion.
    /// Kept as one call so `spawn-*` is literally `new-*` followed by
    /// `resume-later`, rather than a reimplementation that could drift.
    fn resume_thread_later(&mut self, who: &str, index: u32) -> Result<(), VMError> {
        self.cm_instance
            .threads
            .get_mut(index)
            .ok_or_else(|| {
                VMError::new(format!("canon {who}: no thread at index {index} (trap)"))
            })?
            .resume_later()
            .map_err(|e| VMError::new(format!("canon {who}: {e} (trap)")))
    }

    /// The `$cancellable?` immediate of the canonical definition this
    /// built-in was defined with. Absent definition ⇒ NOT cancellable, which
    /// is the safe default: a caller that never opted in is never told, and
    /// the request survives for a later `cancellable` call.
    fn canon_cancellable(&self) -> bool {
        self.canon_type_immediate
            .and_then(|i| self.canon_defs.get(i as usize))
            .map(|d| d.cancellable)
            .unwrap_or(false)
    }

    /// 🔀 `async` — `canonopt` 0x06 on this row's canon definition.
    ///
    /// ⛔ NOT a separate built-in. `Binary.md:310` is
    /// `0x0f t:<typeidx> opts:<opts> => (canon stream.read t opts)`, so the
    /// async/sync distinction rides in `opts` and there is exactly ONE
    /// `CanonBuiltin::StreamRead`. Adding a second built-in for it would put
    /// the same operation at two canonidx spaces.
    ///
    /// Absent immediates default to sync, which is the conservative answer:
    /// a sync copy suspends until it has a real payload, where an async one
    /// answers `BLOCKED` and obliges the caller to wait for the event.
    fn canon_async_opt(&self) -> bool {
        self.canon_type_immediate
            .and_then(|i| self.canon_defs.get(i as usize))
            .map(|d| d.opts.is_async)
            .unwrap_or(false)
    }

    /// `self.task.deliver_pending_cancel(cancellable)` for the current task.
    ///
    /// Returns `Cancelled.TRUE`'s bool. Every blocking thread built-in calls
    /// this FIRST — before parking, before switching — because a delivered
    /// cancellation means the call returns instead of blocking.
    fn deliver_pending_cancel_now(&mut self, cancellable: bool) -> bool {
        match self.cm_tasks.last_mut() {
            Some(task) => task.deliver_pending_cancel(cancellable),
            None => false,
        }
    }

    /// `$ft` must be `(func (param $c T))` — one parameter, no result.
    ///
    /// The VM is untyped, so this checks param/result COUNTS off the callee's
    /// chunk, which is exactly what `call_indirect` does for its `(type $sig)`
    /// check. It is not a full structural type comparison and does not pretend
    /// to be: `T` is `i32` in every non-🐘 profile, so the shape is determined.
    fn require_thread_func_shape(&self, who: &str, funcref: &Value) -> Result<(), VMError> {
        if let Value::Object(o) = funcref {
            let ob = o.lock().unwrap();
            if let ObjectKind::Function(f) = &ob.kind {
                let ch = &self.chunks[f.chunk_index];
                if ch.param_count != 1 || ch.result_arity != 0 {
                    return Err(VMError::new(format!(
                        "canon {who}: thread function has {} param(s) and {} result(s); \
                         `$ft` must be `(func (param $c i32))` (trap)",
                        ch.param_count, ch.result_arity
                    )));
                }
            }
        }
        Ok(())
    }

    /// Create a thread over `funcref` with `closure` bound, register it, and
    /// return its index. The thread is SUSPENDED — `thread.new-*` never runs
    /// the thread it creates.
    fn create_thread_over(
        &mut self,
        who: &str,
        funcref: Value,
        closure: Value,
    ) -> Result<u32, VMError> {
        if funcref.is_null_ref() {
            return Err(VMError::new(format!("canon {who}: null funcref (trap)")));
        }
        self.require_thread_func_shape(who, &funcref)?;
        let task_id = self.cm_tasks.last().map(|t| t.id).ok_or_else(|| {
            VMError::new(format!(
                "canon {who}: no current task — a thread belongs to the task of a \
                 `canon lift`ed call (trap)"
            ))
        })?;
        let cont = self.new_bound_continuation(funcref, closure);
        let thread = crate::cm_thread::Thread::new(task_id, cont);
        debug_assert!(thread.suspended(), "a new thread must not run");
        Ok(self.cm_instance.threads.register(thread))
    }

    /// `shared?` 🧵② — present means the spawned thread is PREEMPTIVE and runs
    /// in parallel with all other threads. Cooperative fibers cannot do that,
    /// and handing back a cooperative thread to a guest that asked for a
    /// parallel one would be a silent lie about the memory model it may then
    /// rely on. Refuse instead.
    fn refuse_shared_threads(&self, who: &str) -> Result<(), VMError> {
        let shared = self
            .canon_type_immediate
            .and_then(|i| self.canon_defs.get(i as usize))
            .map(|d| d.shared)
            .unwrap_or(false);
        if shared {
            return Err(VMError::new(format!(
                "canon {who}: `shared` requests a PREEMPTIVE thread able to execute in \
                 parallel; this runtime schedules cooperative fibers and has no \
                 preemptive threads to give (trap)"
            )));
        }
        Ok(())
    }

    /// Park the current thread and enter `target_idx`'s continuation.
    ///
    /// The running thread has `cont == None` by definition, so a continuation
    /// has to be MINTED to hold its fiber; that object is what a later resume
    /// of this thread restores.
    fn park_and_switch_to(&mut self, who: &str, me: u32, target_idx: u32) -> Result<(), VMError> {
        let target_cont = self
            .cm_instance
            .threads
            .get(target_idx)
            .and_then(|t| t.cont.clone())
            .ok_or_else(|| {
                VMError::new(format!(
                    "canon {who}: thread {target_idx} has no continuation to enter (trap)"
                ))
            })?;
        let my_cont = self.new_parked_continuation();
        self.switch_to_continuation(who, &target_cont, Value::Undefined, Some(&my_cont))?;
        if let Some(me_t) = self.cm_instance.threads.get_mut(me) {
            me_t.put_cont(my_cont);
        }
        Ok(())
    }

    /// `Thread.block_internal` — `block(switch_to = None)`.
    ///
    /// The spec's `Thread.resume` loop exits when nothing is handed to
    /// `switch_to` and control returns to the embedder's scheduler. We have
    /// three genuinely different situations here and they get three different
    /// answers, because collapsing them would report a deadlock for a program
    /// that is merely waiting on I/O:
    ///
    /// 1. another thread is READY — switch to it. That IS the scheduler's
    ///    choice, and is what the spec's loop does on the next iteration.
    /// 2. nothing ready, nothing pending — a real deadlock. Trap, in the shape
    ///    of the existing "parked with no writer left to wake them".
    /// 3. nothing ready but HOST work is pending — a legitimate block that
    ///    needs the fiber-suspension path (`SuspensionKind`) the event loop
    ///    drives. Not implemented; traps naming that gap rather than hanging
    ///    or pretending the call returned normally.
    ///
    /// Pushes the `Cancelled` result on the paths that return.
    fn thread_block(&mut self, who: &str, me: u32) -> Result<(), VMError> {
        let target = self
            .cm_instance
            .threads
            .ready_indices()
            .into_iter()
            .find(|&i| i != me);
        if let Some(t) = target {
            self.cm_instance
                .threads
                .get_mut(t)
                .expect("just listed")
                .stop_waiting_internal(false)
                .map_err(|e| VMError::new(format!("canon {who}: {e} (trap)")))?;
            self.park_and_switch_to(who, me, t)?;
            self.push(Value::I32(0))?;
            return Ok(());
        }

        let host_work_pending = {
            let el = self.event_loop.borrow();
            // The ready QUEUE counts: a job already queued will run and may be
            // exactly what wakes this thread. Leaving it out would report a
            // deadlock for a thread suspending one step ahead of its waker.
            !el.immediate.is_empty()
                || el.parked_sync_copies() > 0
                || !el.waiting_fibers.is_empty()
                || !el.future_waiting_fibers.is_empty()
                || !el.stream_waiting_fibers.is_empty()
        };
        if host_work_pending {
            return Err(VMError::new(format!(
                "canon {who}: blocking with host work still pending needs the fiber \
                 suspension path the event loop drives, which this built-in does not \
                 yet reach (trap)"
            )));
        }
        Err(VMError::new(format!(
            "canon {who}: no thread is ready and no host work is pending — this thread \
             blocks with nothing left that could ever wake it (trap)"
        )))
    }

    /// A `Ready` continuation over `entry` with `arg` bound, so entering it
    /// calls `entry(arg)`.
    ///
    /// Uses `__bound_args` — the same property `cont.bind` writes and
    /// `switch_to_continuation` reads — so a thread's closure parameter travels
    /// by the mechanism that already exists rather than a parallel one.
    fn new_bound_continuation(&mut self, entry: Value, arg: Value) -> Value {
        let mut bound = crate::value::Object::new();
        bound.kind = ObjectKind::Array(vec![arg]);
        let state = crate::value::ContinuationState {
            entry,
            saved: std::sync::Mutex::new(None),
            state: std::sync::Mutex::new(crate::value::ContinuationPhase::Ready),
        };
        let mut obj = crate::value::Object::new();
        obj.kind = ObjectKind::Continuation(state);
        obj.properties.insert(
            "__bound_args".into(),
            Value::Object(crate::heap::alloc(bound)),
        );
        Value::Object(crate::heap::alloc(obj))
    }

    /// A fresh `Suspended`-capable continuation object with no entry, used to
    /// hold the fiber of a thread that is being parked mid-flight.
    ///
    /// Its `entry` is never called: `switch_to_continuation` only consults
    /// `entry` for a `Ready` continuation, and this one is handed straight to
    /// the park path, which sets `Suspended` and stores the fiber.
    fn new_parked_continuation(&mut self) -> Value {
        let state = crate::value::ContinuationState {
            entry: Value::Undefined,
            saved: std::sync::Mutex::new(None),
            state: std::sync::Mutex::new(crate::value::ContinuationPhase::Ready),
        };
        let mut obj = crate::value::Object::new();
        obj.kind = ObjectKind::Continuation(state);
        Value::Object(crate::heap::alloc(obj))
    }

    /// Park the current fiber into `park_into` and enter `target` — the
    /// switch/resume core, with **no tag and no handler search**.
    ///
    /// `Op::SWITCH` is this plus the stack-switching proposal's tag matching.
    /// The Component Model's `thread.{suspend,yield}-then-{resume,promote}`
    /// need the same park-and-enter with no tag: `Thread.resume` is a plain
    ///
    ///     (thread.cont, switch_to) = resume(cont, cancelled, thread)
    ///
    /// loop with a `switch_to` target, and a CM thread has no handler to
    /// search. Keeping the tag in the OPCODE and the switch here is what lets
    /// both callers share one implementation instead of two that drift.
    ///
    /// `park_into` is where the CURRENT execution goes: `Some(cont)` saves the
    /// fiber into that continuation and marks it `Suspended`; `None` abandons
    /// it, which is what a thread that is ending does.
    ///
    /// The caller owns any handler-frame bookkeeping — on `Err` nothing here
    /// has been pushed, so a caller that popped a frame must restore it.
    pub(crate) fn switch_to_continuation(
        &mut self,
        who: &str,
        target: &Value,
        val: Value,
        park_into: Option<&Value>,
    ) -> Result<(), VMError> {
        let (phase, entry) = match target {
            Value::Object(obj) => {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Continuation(cs) => {
                        (*cs.state.lock().unwrap(), cs.entry.clone())
                    }
                    _ => return Err(VMError::new(format!("{who}: not a continuation"))),
                }
            }
            _ => return Err(VMError::new(format!("{who}: not a continuation"))),
        };
        if matches!(phase, crate::value::ContinuationPhase::Done) {
            return Err(VMError::new(format!(
                "trap: {who} to completed continuation"
            )));
        }

        // Park the current execution BEFORE entering the target: once the
        // target runs, `self` is no longer this fiber.
        let fiber = self.save_fiber();
        if let Some(Value::Object(obj)) = park_into {
            let o = obj.lock().unwrap();
            if let ObjectKind::Continuation(cs) = &o.kind {
                *cs.saved.lock().unwrap() = Some(fiber);
                *cs.state.lock().unwrap() = crate::value::ContinuationPhase::Suspended;
            }
        }

        match phase {
            // Never entered: call its entry function, prefixing whatever
            // `cont.bind` attached.
            crate::value::ContinuationPhase::Ready => {
                let bound: Vec<Value> = match target {
                    Value::Object(obj) => {
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
                    }
                    _ => Vec::new(),
                };
                let argc = bound.len() + 1;
                self.push(entry)?;
                for b in bound {
                    self.push(b)?;
                }
                self.push(val)?;
                self.call_value_direct(argc)?;
            }
            // Already running once: restore its saved fiber.
            crate::value::ContinuationPhase::Suspended => {
                let saved = match target {
                    Value::Object(obj) => {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Continuation(cs) = &o.kind {
                            cs.saved.lock().unwrap().take()
                        } else {
                            None
                        }
                    }
                    _ => None,
                };
                if let Some(fiber) = saved {
                    self.resume_fiber_with(fiber, Some(val))?;
                }
            }
            crate::value::ContinuationPhase::Done => unreachable!("checked above"),
        }
        Ok(())
    }

    /// `(current thread index, slot index)` for a `context.*` built-in.
    /// Pops the slot index.
    ///
    /// `thread = current_thread()` is unconditional in the spec, so there is no
    /// answer outside a lifted call — the implicit thread IS the context.
    fn canon_context_slot(&mut self, builtin: &str) -> Result<(u32, usize), VMError> {
        let index = self.pop().as_i32() as usize;
        // `Explainer.md:1679`: "Validation currently restricts `i` to be less
        // than 2". Checked here as well as by the array bound, so the message
        // names the SPEC limit rather than a Rust length.
        if index >= crate::vm::CONTEXT_STORAGE_SLOTS {
            return Err(VMError::new(format!(
                "canon {builtin}: slot {index} exceeds the {} slots the spec permits (trap)",
                crate::vm::CONTEXT_STORAGE_SLOTS
            )));
        }
        let thread_idx = self.current_thread.ok_or_else(|| {
            VMError::new(format!(
                "canon {builtin}: no current thread — thread-local storage exists \
                 only inside a `canon lift`ed call (trap)"
            ))
        })?;
        Ok((thread_idx, index))
    }

    /// The `$t` / `$rt` TYPE INDEX this built-in was defined with — the row's
    /// type, not the row's own index.
    ///
    /// `canon_type_immediate` holds a CANONIDX. `resource.{new,rep,drop}` want
    /// `rt:<typeidx>`, and reading the canonidx directly is only correct while
    /// the identity section is in force. The moment a real canon section exists
    /// those are different numbers, and comparing a canonidx against a
    /// `type_id` is the `GLOBAL_GET` defect with a new pair of tables.
    ///
    /// `None` means the import declared no immediate at all.
    pub(crate) fn canon_type_index(&self) -> Option<u32> {
        let idx = self.canon_type_immediate?;
        if self.canon_defs.is_empty() {
            // Identity section: row `i` declares `$t = i`.
            return Some(idx);
        }
        self.canon_defs.get(idx as usize)?.ty
    }

    fn canon_element_type(&self, builtin: &str) -> Result<crate::component::ValType, VMError> {
        let Some(idx) = self.canon_type_immediate else {
            return Err(VMError::new(format!(
                "canon {builtin}: no canonidx — import it as `canon`/`{builtin}@<canonidx>` \
                 naming a row of `VM::canon_defs` whose `$t` is registered in `VM::canon_types`"
            )));
        };
        // The canonidx names a ROW; the row's `$t` names the type. Two hops,
        // because they are two index spaces — reading `canon_types[canonidx]`
        // directly (what this did) is the `GLOBAL_GET` defect: one integer
        // silently serving whichever table the reader happened to pick.
        //
        // A module with NO canon section gets the IDENTITY section: row `i`
        // declares `$t = i`. Core WASM has no way to spell a canon section —
        // `(import "canon" "future.read@1")` in a `.wat` carries the immediate
        // and nothing else — so without this the only expressible spelling
        // would name a row that cannot exist. The identity mapping is what
        // `@N` already meant, now stated as a rule instead of assumed.
        //
        // It is a declared fallback, not a silent one: the moment a real canon
        // section exists, `canon_defs` is non-empty and the row wins.
        let ty = if self.canon_defs.is_empty() {
            idx
        } else {
            let def = self.canon_defs.get(idx as usize).ok_or_else(|| {
                VMError::new(format!(
                    "canon {builtin}: canonidx {idx} is not a row of `VM::canon_defs` \
                     (have {})",
                    self.canon_defs.len()
                ))
            })?;
            def.require_type(builtin)
                .map_err(|e| VMError::new(format!("{e} (canonidx {idx})")))?
        };
        match self.canon_types.get(ty as usize) {
            Some(Some(v)) => Ok(v.clone()),
            // The typeidx EXISTS but holds something else — a function type, or
            // a form the front end recorded without decomposing. Distinguished
            // from out-of-range because they are different mistakes: one is a
            // stale index, the other is naming the wrong kind of type.
            Some(None) => Err(VMError::new(format!(
                "canon {builtin}: $t {ty} (from canonidx {idx}) is a declared typeidx \
                 but does not hold a VALUE type"
            ))),
            None => Err(VMError::new(format!(
                "canon {builtin}: $t {ty} (from canonidx {idx}) is not registered in \
                 `VM::canon_types` (have {})",
                self.canon_types.len()
            ))),
        }
    }

    /// The canon-section row this built-in was defined by, or a trap naming
    /// what is missing. `lift`/`lower` cannot fall back to the identity
    /// section: identity answers "which TYPE", and these need `$callee`,
    /// `$opts` and `$ft` — three immediates no import name can carry.
    fn canon_def_required(
        &self,
        builtin: &str,
    ) -> Result<crate::canon_def::CanonDef, VMError> {
        let idx = self.canon_type_immediate.ok_or_else(|| {
            VMError::new(format!(
                "canon {builtin}: no canonidx — `{builtin}` is defined by a \
                 `(canon {builtin} ...)` row carrying its immediates, so it needs a \
                 row of the canon section; import it as `canon`/`{builtin}@<canonidx>`"
            ))
        })?;
        self.canon_defs.get(idx as usize).cloned().ok_or_else(|| {
            VMError::new(format!(
                "canon {builtin}: canonidx {idx} is not a row of `VM::canon_defs` (have {})",
                self.canon_defs.len()
            ))
        })
    }

    /// Resolve a row's `$ft` through the FUNCTION type index space.
    ///
    /// Deliberately not `canon_types`: `$ft` is a function type and `$t` is a
    /// value type. One table serving both is the `GLOBAL_GET` defect.
    fn canon_functype(
        &self,
        builtin: &str,
        def: &crate::canon_def::CanonDef,
    ) -> Result<crate::canon_def::CanonFuncType, VMError> {
        let idx = def.functype.ok_or_else(|| {
            VMError::new(format!("canon {builtin}: definition carries no $ft immediate"))
        })?;
        match self.canon_functypes.get(idx as usize) {
            Some(Some(ft)) => Ok(ft.clone()),
            Some(None) => Err(VMError::new(format!(
                "canon {builtin}: $ft {idx} is a declared typeidx but does not hold a \
                 FUNCTION type"
            ))),
            None => Err(VMError::new(format!(
                "canon {builtin}: $ft {idx} is not registered in `VM::canon_functypes` (have {})",
                self.canon_functypes.len()
            ))),
        }
    }

    /// Call a canon callee to completion and return its single result.
    ///
    /// `argc` values are already on the operand stack. The callee value is
    /// inserted BELOW them — the same shape `ImportTarget::ChunkFn` uses — and
    /// `execute_until` runs the nested frame to completion, which is what makes
    /// this a synchronous call rather than a frame push the caller never sees
    /// the result of.
    fn call_canon_callee(
        &mut self,
        builtin: &str,
        callee: crate::canon_def::CalleeRef,
        argc: usize,
    ) -> Result<Value, VMError> {
        let chunk_index = match callee {
            crate::canon_def::CalleeRef::Core(i) => i as usize,
            // A component funcidx indexes the COMPONENT's function index
            // space, which needs the component linker — deferred with the
            // export path (see cmplan.md). Refusing names that dependency
            // instead of silently reading the core space, where the same
            // integer means a different function.
            // A component function DEFINED HERE — by a `canon lift` — needs no
            // linker: `component_funcs` maps the funcidx to the canonidx of the
            // row that defines it, and calling it IS running that lift. The
            // linker is for component functions that arrive as IMPORTS.
            crate::canon_def::CalleeRef::Component(i) => {
                let slot = *self.component_funcs.get(i as usize).ok_or_else(|| {
                    VMError::new(format!(
                        "canon {builtin}: $callee {i} is not in the component function \
                         index space (have {})",
                        self.component_funcs.len()
                    ))
                })?;
                // ⛔ THE SLOT EXISTS AND IS EMPTY. That is an IMPORTED component
                // function: `(import "x" (func $x (type $ft)))` occupies an
                // index in declaration order without defining anything here, so
                // the slot has to be present and unfilled. This is the one case
                // that genuinely needs the component linker — every other
                // producer (`canon lift`, `(alias export …)`, `(export …)`)
                // resolves to a row in THIS component.
                //
                // Out-of-range and empty are deliberately different messages,
                // because they are different mistakes: a stale index versus a
                // callee nothing has supplied.
                let canonidx = slot.ok_or_else(|| {
                    VMError::new(format!(
                        "canon {builtin}: component func {i} is IMPORTED — it has no \
                         defining row in this component, so calling it needs the \
                         component linker (see cmplan.md §Deferred to export)"
                    ))
                })?;
                // `exec_canon_lift` reads its row through `canon_type_immediate`,
                // so the callee's canonidx is installed for the nested call and
                // restored after: a lifted call NESTS (a `realloc` is itself a
                // lift), and leaving it set would make the outer row read the
                // inner one's immediates.
                let saved = self.canon_type_immediate;
                self.canon_type_immediate = Some(canonidx);
                let outcome = self.exec_canon_lift();
                self.canon_type_immediate = saved;
                outcome?;
                // `exec_canon_lift` pushes exactly one value, per this VM's
                // canon-import ABI.
                return Ok(self.pop());
            }
        };
        if chunk_index >= self.chunks.len() {
            return Err(VMError::new(format!(
                "canon {builtin}: $callee core funcidx {chunk_index} is out of range (have {})",
                self.chunks.len()
            )));
        }
        let func = crate::value::Function {
            name: None,
            arity: self.chunks[chunk_index].arity,
            chunk_index,
            upvalues: Vec::new(),
        };
        let mut obj = crate::value::Object::new();
        obj.kind = crate::value::ObjectKind::Function(func);
        let func_val = Value::Object(crate::heap::alloc(obj));
        let args_start = self.stack.len() - argc;
        self.stack.insert(args_start, func_val);
        let depth = self.frames.len();
        self.call_value(argc)?;
        // `call_value` does not always push a frame — a host or callable shim
        // completes inline and leaves its result on the stack. Mirrors the
        // same guard in `vm.rs`'s callback path.
        if self.frames.len() == depth {
            return Ok(self.pop());
        }
        // ⛔ `depth + 1`, not `depth`. `execute_until_inner` exits a nested
        // loop when `frames.len() < min_depth`, and the callee's RETURN pops
        // its own frame FIRST — so frames are back to `depth` at the moment
        // the check runs, and `depth < depth` is false.
        //
        // The consequence was not a hang: the arm fell through, PUSHED the
        // results onto the stack and kept interpreting the caller's frames.
        // The value `canon lift` then lifted was `Null`, while the callee's
        // real result sat on the operand stack where the core caller read it —
        // which is why a lifted `(func (result bool))` handed its caller the
        // callee's raw 9 instead of 1, and why deleting the result type
        // changed nothing.
        //
        // `vm.rs` and `jspi.rs` both pass `saved_frame_depth + 1`.
        self.execute_until(depth + 1)
    }

    /// `canon lower` — `CanonicalABI.md §canon lower`, synchronous path.
    ///
    /// ```python
    /// flat_args = CoreValueIter(flat_args)
    /// args = lift_flat_values(cx, MAX_FLAT_PARAMS, flat_args, ft.param_types())
    /// result = callee(args)
    /// flat_results = lower_flat_values(cx, MAX_FLAT_RESULTS, result, ft.result_type())
    /// return flat_results
    /// ```
    ///
    /// A lowered import is what CORE wasm calls, so its arguments arrive FLAT
    /// on the operand stack and its results must leave flat. The conversion is
    /// the whole job: lift the flat args up to component values, call, lower
    /// the result back down.
    fn exec_canon_lower(&mut self) -> Result<(), VMError> {
        use crate::canon_flat::{CoreType, FlattenContext, MAX_FLAT_PARAMS, MAX_FLAT_RESULTS};
        let def = self.canon_def_required("lower")?;
        let ft = self.canon_functype("lower", &def)?;
        let callee = def
            .require_callee("lower")
            .map_err(VMError::new)?;

        let flat_ft = crate::canon_flat::flatten_functype(
            &ft.params,
            ft.result.as_ref(),
            FlattenContext::Lower,
            def.opts.is_async,
            def.opts.callback.is_some(),
            CoreType::I32,
        );
        let flat_args = self.pop_core_values(&flat_ft.params)?;

        let memory = self.memory.clone();
        let args = crate::canon_flat_values::lift_flat_values(
            &memory,
            MAX_FLAT_PARAMS,
            &flat_args,
            &ft.params,
            CoreType::I32,
        )
        .map_err(|e| VMError::new(format!("canon lower: lifting arguments: {e}")))?;

        let argc = args.len();
        for a in args {
            self.push(a)?;
        }
        let result = self.call_canon_callee("lower", callee, argc)?;

        let mut bump = self.canon_bump_start();
        let flat_results = {
            let mut alloc = Self::bump_realloc(&memory, &mut bump);
            // `Realloc<'a> = &'a mut dyn FnMut(u32,u32) -> Option<u32>` — the
            // coercion is explicit because an `impl FnMut` is not the trait
            // object the canonical-ABI helpers take.
            let mut realloc: crate::canon_value::Realloc<'_> = &mut alloc;
            let types: Vec<_> = ft.result.iter().cloned().collect();
            let values: Vec<_> = ft.result.iter().map(|_| result.clone()).collect();
            crate::canon_flat_values::lower_flat_values(
                &memory,
                &mut realloc,
                MAX_FLAT_RESULTS,
                &values,
                &types,
                CoreType::I32,
            )
            .map_err(|e| VMError::new(format!("canon lower: lowering result: {e}")))?
        };
        self.canon_bump_commit(bump);
        for v in flat_results {
            self.push(Self::core_value_to_value(v))?;
        }
        Ok(())
    }

    /// `canon lift` — `CanonicalABI.md §canon lift`, synchronous path.
    ///
    /// ```python
    /// flat_args = lower_flat_values(cx, MAX_FLAT_PARAMS, args, ft.param_types())
    /// flat_results = call_and_trap_on_throw(callee, flat_args)
    /// result = lift_flat_values(cx, MAX_FLAT_RESULTS, flat_results, ft.result_type())
    /// task.return_(result)
    /// if opts.post_return is not None:
    ///   inst.may_leave = False
    ///   [] = call_and_trap_on_throw(opts.post_return, flat_results)
    ///   inst.may_leave = True
    /// ```
    ///
    /// The mirror of `lower`: component values in, flat values to the core
    /// callee, and the core results lifted back up. The result goes to the
    /// TASK (`task.return_`), not to the operand stack.
    fn exec_canon_lift(&mut self) -> Result<(), VMError> {
        use crate::canon_flat::{CoreType, MAX_FLAT_PARAMS, MAX_FLAT_RESULTS};
        let def = self.canon_def_required("lift")?;
        let ft = self.canon_functype("lift", &def)?;
        let callee = def.require_callee("lift").map_err(VMError::new)?;

        // "a task is created for each call to a component export (in
        // `canon_lift`) ... starting with the IMPLICIT thread that is spawned
        // by `canon_lift`" — CanonicalABI.md §Tasks.
        //
        // This is the ONLY place either comes into existence, and nothing was
        // doing it: `CMTask::new` had no caller outside its own unit tests, so
        // `cm_tasks` was permanently empty and `task.return`'s guards could
        // never fire. With no thread, `thread.index` had nothing to answer and
        // every other 🧵 built-in had no table entry to address.
        let task_id = self.next_cm_task_id;
        self.next_cm_task_id += 1;
        let mut task = crate::cm_task::CMTask::new(task_id);
        task.start();
        self.cm_tasks.push(task);

        // The implicit thread is RUNNING, not suspended: it is the one
        // executing this lifted call. `Thread::new` starts a thread suspended
        // holding a continuation, which is right for `thread.new-indirect` and
        // wrong here, so the continuation is taken immediately — `running()`
        // is `cont is None`.
        let mut implicit = crate::cm_thread::Thread::new(task_id, Value::Undefined);
        implicit.take_cont();
        debug_assert!(implicit.running());
        let thread_index = self.cm_instance.threads.register(implicit);
        let prev_thread = self.current_thread.replace(thread_index);

        let mut args = Vec::with_capacity(ft.params.len());
        for _ in 0..ft.params.len() {
            args.push(self.pop());
        }
        args.reverse();

        let memory = self.memory.clone();
        let mut bump = self.canon_bump_start();
        let flat_args = {
            let mut alloc = Self::bump_realloc(&memory, &mut bump);
            let mut realloc: crate::canon_value::Realloc<'_> = &mut alloc;
            crate::canon_flat_values::lower_flat_values(
                &memory,
                &mut realloc,
                MAX_FLAT_PARAMS,
                &args,
                &ft.params,
                CoreType::I32,
            )
            .map_err(|e| VMError::new(format!("canon lift: lowering arguments: {e}")))?
        };
        self.canon_bump_commit(bump);

        let argc = flat_args.len();
        for v in &flat_args {
            self.push(Self::core_value_to_value(*v))?;
        }
        let raw = self.call_canon_callee("lift", callee, argc)?;

        // The core type comes from the SPEC'S FLATTENING of the declared
        // result, not from whatever variant the returned `Value` happens to
        // hold. Those are different questions: `flatten_type(u32)` is `i32`,
        // while a core function returning `(i32.const 5)` hands back a
        // `Value::F64` in this VM, because its numeric tower is doubles. Reading
        // the core type off the value made every integer-returning lift fail
        // with "wanted a core I32 and got a F64" — the callee was right and the
        // question was wrong.
        let flat_results = match &ft.result {
            Some(t) => {
                let want = crate::canon_flat::flatten_type(t, CoreType::I32);
                vec![Self::value_to_core_value_as(
                    &raw,
                    want.first().copied().unwrap_or(CoreType::I32),
                )]
            }
            None => Vec::new(),
        };
        let types: Vec<_> = ft.result.iter().cloned().collect();
        let lifted = crate::canon_flat_values::lift_flat_values(
            &memory,
            MAX_FLAT_RESULTS,
            &flat_results,
            &types,
            CoreType::I32,
        )
        .map_err(|e| VMError::new(format!("canon lift: lifting result: {e}")))?;

        // `task.return_(result)` — the result belongs to the TASK. `canon_lift`
        // has no return value in the spec: it hands the lifted result to the
        // task and a core caller reaches it through `canon lower`, never by
        // calling the lift directly.
        let result = lifted.into_iter().next().unwrap_or(Value::Undefined);
        if let Some(task) = self.cm_tasks.last_mut() {
            task.return_(result.clone())
                .map_err(|e| VMError::new(format!("canon lift: {e} (trap)")))?;
        }
        // …but THIS VM's canon-import ABI is one value per call: the emitter
        // reserves a stack slot for every `canon` import's result and drops it
        // when unused. Pushing nothing does not mean "no result", it means the
        // caller reads whatever sat BELOW — which is how a lifted
        // `(func (result bool))` handed its caller the callee's raw 9 instead
        // of 1, and how deleting the result type changed nothing.
        //
        // What a CORE caller must receive is the result LOWERED back to flat
        // core values, not the component value: `result` here is a
        // `Value::Bool`, and a core function cannot receive one. Importing
        // `canon`/`lift@N` into core wasm is standing in for lower ∘ lift, so
        // this composes them — which is also what makes the value observably
        // the LIFTED one: a `bool` lifted from a core 9 lowers back to 1, so
        // 9 arriving would prove the lift had been skipped.
        let pushed = match &ft.result {
            Some(t) => {
                let mut bump2 = self.canon_bump_start();
                let flat = {
                    let mut alloc = Self::bump_realloc(&memory, &mut bump2);
                    let mut realloc: crate::canon_value::Realloc<'_> = &mut alloc;
                    crate::canon_flat_values::lower_flat_values(
                        &memory,
                        &mut realloc,
                        MAX_FLAT_RESULTS,
                        std::slice::from_ref(&result),
                        std::slice::from_ref(t),
                        CoreType::I32,
                    )
                    .map_err(|e| VMError::new(format!("canon lift: lowering result: {e}")))?
                };
                self.canon_bump_commit(bump2);
                flat.first()
                    .map(|v| Self::core_value_to_value(*v))
                    .unwrap_or(Value::Undefined)
            }
            // A lift with no declared result still owes the caller its one
            // stack slot; the emitter drops it.
            None => Value::Undefined,
        };
        self.push(pushed)?;

        // `post-return` runs with `may_leave` CLEARED, which is what lets a
        // sync-lowered call to a sync-lifted function be a plain call: neither
        // it nor `realloc` may block, so no fiber is needed.
        if let Some(pr) = def.opts.post_return {
            self.cm_instance
                .enter_no_leave()
                .map_err(|e| VMError::new(format!("canon lift: post-return: {e}")))?;
            for v in &flat_args {
                self.push(Self::core_value_to_value(*v))?;
            }
            let outcome = self.call_canon_callee(
                "lift",
                crate::canon_def::CalleeRef::Core(pr),
                flat_args.len(),
            );
            self.cm_instance.exit_no_leave();
            outcome?;
        }

        // `task.exit_implicit_thread()` — thread and task end with the call.
        // `current_thread` is RESTORED, not cleared: a lifted call nests (a
        // `realloc` is itself a `canon_lift`), and clearing would strand the
        // outer thread with no index.
        self.cm_instance.threads.unregister(thread_index);
        self.current_thread = prev_thread;
        self.cm_tasks.pop();
        Ok(())
    }

    /// Pop the operand-stack values for `want` as flat core values, in stack
    /// order, each typed by the FLATTENED SIGNATURE.
    ///
    /// ⛔ This used to take a COUNT and infer each core type from the popped
    /// value's own variant — the same defect `exec_canon_lift`'s result path
    /// had. It is wrong for the same reason: `flatten_type(u32)` is `i32`, but
    /// a core function's `(i32.const 5)` is a `Value::F64` in this VM, so the
    /// inferred type disagreed with the signature and `lift_flat_values`
    /// rejected an argument the caller had passed correctly.
    ///
    /// The caller already has `flat_ft.params`; it was throwing it away and
    /// passing only its length.
    fn pop_core_values(
        &mut self,
        want: &[crate::canon_flat::CoreType],
    ) -> Result<Vec<crate::canon_flat_values::CoreValue>, VMError> {
        let n = want.len();
        if self.stack.len() < n {
            return Err(VMError::new(format!(
                "canon: expected {n} flat argument(s), stack holds {}",
                self.stack.len()
            )));
        }
        let mut out = Vec::with_capacity(n);
        // Popped in reverse, so the type for each pop comes from the tail.
        for t in want.iter().rev() {
            let v = self.pop();
            out.push(Self::value_to_core_value_as(&v, *t));
        }
        out.reverse();
        Ok(out)
    }

    /// A flat core value as an operand-stack `Value`.
    fn core_value_to_value(v: crate::canon_flat_values::CoreValue) -> Value {
        use crate::canon_flat_values::CoreValue as C;
        match v {
            C::I32(i) => Value::I32(i as i32),
            C::I64(i) => Value::I64(i as i64),
            C::F32(f) => Value::F64(f as f64),
            C::F64(f) => Value::F64(f),
        }
    }

    /// An operand-stack `Value` as a flat core value.
    ///
    /// The flat ABI has exactly four core types; anything else on the stack at
    /// a canon boundary is a value that was never lowered, so it is carried as
    /// its i32 form rather than silently reinterpreted as a float.
    /// A `Value` as the core type the canonical ABI asks for.
    ///
    /// Distinct from [`Self::value_to_core_value`], which infers the core type
    /// from the value's own variant. That inference is right where the flat
    /// types are genuinely unknown and wrong wherever the signature states
    /// them — a `u32` result flattens to `i32` no matter how the VM happens to
    /// be holding the number.
    fn value_to_core_value_as(
        v: &Value,
        want: crate::canon_flat::CoreType,
    ) -> crate::canon_flat_values::CoreValue {
        use crate::canon_flat::CoreType as T;
        use crate::canon_flat_values::CoreValue as C;
        match want {
            T::I32 => C::I32(v.as_i32() as u32),
            T::I64 => C::I64(v.as_i64() as u64),
            T::F32 => C::F32(v.as_f64() as f32),
            T::F64 => C::F64(v.as_f64()),
        }
    }

    fn value_to_core_value(v: &Value) -> crate::canon_flat_values::CoreValue {
        use crate::canon_flat_values::CoreValue as C;
        match v {
            Value::I64(i) => C::I64(*i as u64),
            Value::F64(f) => C::F64(*f),
            other => C::I32(other.as_i32() as u32),
        }
    }

    /// Look up a waitable set and take its ready event, as an `EventTuple`.
    ///
    ///     wset = inst.handles.get(si)
    ///     trap_if(not isinstance(wset, WaitableSet))
    ///
    /// The trap is the point. Returning `NONE` for a handle that is not a
    /// waitable set is indistinguishable from "the set exists and nothing is
    /// ready", so a guest polling a bogus handle would spin forever instead of
    /// failing at the call that was wrong.
    fn poll_waitable_set(
        &self,
        builtin: &str,
        set_handle: u32,
    ) -> Result<(crate::waitable::EventCode, u32, u32), VMError> {
        let el = self.event_loop.borrow();
        let Some(set) = self.waitable_sets.get(set_handle) else {
            return Err(VMError::new(format!(
                "canon {builtin}: handle {set_handle} is not a waitable set (trap)"
            )));
        };
        // `EventTuple = tuple[EventCode, int, int]`. `p2` is event-kind
        // specific — `subtask.state` for SUBTASK, the packed
        // `result | (count << 4)` for the stream/future codes. `poll_ready`
        // reports only `(code, id)`, so p2 is 0 until it carries the payload;
        // recorded in cmplan.md rather than left to look intentional.
        Ok(match set.poll_ready(&el) {
            Some((code, id)) => (code, id as u32, 0),
            None => (crate::waitable::EventCode::None, 0, 0),
        })
    }

    /// `unpack_event` — `CanonicalABI.md`:
    ///
    ///     def unpack_event(mem, inst, ptr, e: EventTuple):
    ///       event, p1, p2 = e
    ///       store(cx, p1, U32Type(), ptr)
    ///       store(cx, p2, U32Type(), ptr + 4)
    ///       return [event]
    ///
    /// **EIGHT bytes, `[p1, p2]`** — the event CODE is the return value and is
    /// never written to memory. This previously wrote `[code, p1, 0]`: twelve
    /// bytes, with the code sitting in the slot the guest reads `p1` from, so
    /// every reader got an event code where it expected a waitable index and
    /// four bytes past the payload were clobbered.
    ///
    /// `store` traps out of bounds. Silently skipping the write while still
    /// returning a code (what this did) hands back a plausible event with
    /// nothing behind it.
    fn unpack_event(
        &mut self,
        ptr: usize,
        (code, p1, p2): (crate::waitable::EventCode, u32, u32),
    ) -> Result<i32, VMError> {
        if ptr + 8 > self.memory.len() {
            return Err(VMError::new(format!(
                "canon waitable-set: event pointer {ptr} + 8 is out of bounds (trap)"
            )));
        }
        self.memory.store_i32(ptr, p1 as i32)?;
        self.memory.store_i32(ptr + 4, p2 as i32)?;
        Ok(code as i32)
    }

    /// `cancel_copy` — `CanonicalABI.md` §`canon
    /// {stream,future}.cancel-{read,write}`. One function for all four, as the
    /// spec factors it; `EndKind` is the only thing that differs.
    ///
    ///     $f : (func (param i32) (result i32))
    ///
    /// Cancel RECLAIMS THE BUFFER from an in-flight copy. It does not close
    /// anything: after cancelling, both ends are still usable and the end
    /// returns to `Idle`. The `CANCELLED` code exists precisely to tell wasm
    /// that ownership of the memory buffer has come back to it.
    fn canon_cancel_copy(&mut self, want: crate::canon_copy::EndKind) -> Result<(), VMError> {
        use crate::canon_copy::EndKind;
        use crate::handle_table::{CopyState, HandleEntry};
        let handle = self.pop().as_i32() as u32;
        let entry = self.handle_table.get(handle);
        let end = match (want, entry) {
            (EndKind::ReadableStream, Some(HandleEntry::ReadableStreamEnd(e)))
            | (EndKind::WritableStream, Some(HandleEntry::WritableStreamEnd(e)))
            | (EndKind::ReadableFuture, Some(HandleEntry::ReadableFutureEnd(e)))
            | (EndKind::WritableFuture, Some(HandleEntry::WritableFutureEnd(e))) => *e,
            _ => {
                return Err(VMError::new(format!(
                    "canon {}: handle is not a {} end",
                    want.builtin_name(),
                    want.describe()
                )));
            }
        };
        // `trap_if(e.state != CopyState.COPYING …)` — there is nothing to
        // cancel unless a copy is in flight, and silently succeeding would let
        // a guest believe it had reclaimed a buffer it never lent out.
        if end.state != CopyState::Copying {
            return Err(VMError::new(format!(
                "canon {}: end is not COPYING — there is no copy in flight to cancel",
                want.builtin_name()
            )));
        }
        if let Some(e) = self.handle_table.get_mut(handle) {
            let state = match e {
                HandleEntry::ReadableStreamEnd(s)
                | HandleEntry::WritableStreamEnd(s)
                | HandleEntry::ReadableFutureEnd(s)
                | HandleEntry::WritableFutureEnd(s) => &mut s.state,
                _ => unreachable!("kind was matched above"),
            };
            *state = CopyState::Idle;
        }
        // Nothing had been copied into the buffer — the copy blocked before
        // making progress — so the count is 0.
        self.push(Value::I32(crate::canon_copy::pack(
            crate::canon_copy::CopyResult::Cancelled,
            0,
        ) as i32))?;
        Ok(())
    }

    /// Execute a Component Model canonical built-in (VM-implemented import
    /// under module "canon" — see `ImportTarget::Canon`). Args and results
    /// ride the operand stack; each builtin pops exactly its own args.
    pub(crate) fn exec_canon_builtin(&mut self, b: crate::vm::CanonBuiltin) -> Result<(), VMError> {
        use crate::vm::CanonBuiltin as B;

        // `trap_if(not inst.may_leave)` — 32 occurrences in `CanonicalABI.md`,
        // and it had NO implementation at all: `may_leave` did not exist, so
        // the guard on nearly every built-in was simply absent.
        //
        // Stated once here rather than restated per arm, because it is one
        // rule. The delegating rows are covered through their helpers exactly
        // as the spec covers them: `stream.read`/`write` via `stream_copy`,
        // the four `cancel-*` via `cancel_copy`, `future.read`/`write` via
        // `future_copy`, and all four `drop-*` via `drop` — each of which
        // opens with the guard.
        //
        // The exemptions are the spec's, verified row by row against
        // `CanonicalABI.md`, not assumed:
        //   backpressure.inc/dec — must work precisely WHILE entry is blocked
        //   context.get/set      — thread-local storage, never leaves
        //   resource.rep         — guarded by `isinstance` + `h.rt is rt`
        //   thread.available-parallelism 🧵② — `canon_thread_available_
        //                          parallelism()` is two lines returning a
        //                          count and carries NO `trap_if` at all
        //
        // ⛔ That last row used to read "is also exempt, but has no arm yet".
        // The arm landed and the comment did not, so the row was running a
        // guard the spec does not give it.
        let exempt = matches!(
            b,
            B::BackpressureInc
                | B::BackpressureDec
                | B::ContextGet
                | B::ContextSet
                | B::ResourceRep
                | B::ThreadAvailableParallelism
        );
        if !exempt {
            self.cm_instance
                .require_may_leave(b.spec_name())
                .map_err(VMError::new)?;
        }

        match b {
            B::Lift => self.exec_canon_lift()?,
            B::Lower => self.exec_canon_lower()?,
            B::TaskReturn => {
                // canon task.return — `CanonicalABI.md §canon task.return`:
                //
                //     task.return_(result)
                //     return []
                //
                // The result belongs to the TASK. It is NOT pushed back:
                // `canon_task_return` returns the empty list, so a value left
                // on the operand stack here is one the spec says is not there,
                // and the result itself would be discarded.
                //
                // Still missing (they need the `rs`/`opts` immediates, which
                // an import name cannot carry — see cmplan.md):
                //   trap_if(not task.opts.async_)
                //   trap_if(result_type != task.ft.result)
                //   trap_if(not LiftOptions.equal(opts, task.opts))
                let result = self.pop();
                if let Some(task) = self.cm_tasks.last_mut() {
                    task.return_(result)
                        .map_err(|e| VMError::new(format!("canon task.return: {e} (trap)")))?;
                }
            }
            B::TaskCancel => {
                // canon task.cancel — `CanonicalABI.md §Task.cancel`:
                //
                //     trap_if(self.state != Task.State.CANCEL_DELIVERED)
                //     trap_if(self.num_borrows > 0)
                //     self.on_resolve(None)
                //
                // Both guards are the point. Cancelling a task that was never
                // TOLD to cancel is a trap, not a no-op — the callee has not
                // unwound, so resolving it with no value strands whatever it
                // still holds. Writing `phase = Resolved` directly (as this
                // did) skipped both checks and made a cancel indistinguishable
                // from a return.
                if let Some(task) = self.cm_tasks.last_mut() {
                    task.cancel()
                        .map_err(|e| VMError::new(format!("canon task.cancel: {e} (trap)")))?;
                }
            }
            B::SubtaskCancel => {
                // canon subtask.cancel — pops subtask handle (i32), cancels the subtask.
                let handle = self.pop().as_i32() as u32;
                let fid =
                    if let Some(crate::handle_table::HandleEntry::Subtask { future_id, .. }) =
                        self.handle_table.get(handle)
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
                        el.immediate
                            .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                    }
                }
            }
            B::SubtaskDrop => {
                // canon subtask.drop — pops subtask handle (i32), removes from handle table.
                let handle = self.pop().as_i32() as u32;
                self.handle_table.remove(handle);
            }
            B::WaitableSetNew => {
                // canon waitable-set.new — create a new waitable set, push its handle (i32).
                let set_id = self.waitable_sets.create();
                self.push(Value::I32(set_id as i32))?;
            }
            B::WaitableSetWait => {
                // canon waitable-set.wait — `CanonicalABI.md`:
                //
                //     wset = inst.handles.get(si)
                //     trap_if(not isinstance(wset, WaitableSet))
                //     event = wset.wait_for_event(cancellable)
                //     return unpack_event(mem, inst, ptr, event)
                //
                // If nothing is ready this returns NONE immediately; true
                // blocking needs `cancellable`, which no import name can carry.
                let memory_ptr = self.pop().as_i32() as usize;
                let set_handle = self.pop().as_i32() as u32;
                let event = self.poll_waitable_set("waitable-set.wait", set_handle)?;
                let code = self.unpack_event(memory_ptr, event)?;
                self.push(Value::I32(code))?;
            }
            B::WaitableSetPoll => {
                // canon waitable-set.poll — identical to `wait` but never
                // blocks; the spec factors both through `unpack_event`.
                let memory_ptr = self.pop().as_i32() as usize;
                let set_handle = self.pop().as_i32() as u32;
                let event = self.poll_waitable_set("waitable-set.poll", set_handle)?;
                let code = self.unpack_event(memory_ptr, event)?;
                self.push(Value::I32(code))?;
            }
            B::WaitableJoin => {
                // canon waitable.join — pops [waitable_handle_i32, set_handle_i32];
                // looks up waitable in handle table, adds to set.
                let set_handle = self.pop().as_i32() as u32;
                let waitable_handle = self.pop().as_i32() as u32;
                let waitable = match self.handle_table.get(waitable_handle) {
                    Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e)) => {
                        Some(crate::waitable::Waitable::Stream(e.id))
                    }
                    Some(crate::handle_table::HandleEntry::ReadableFutureEnd(e)) => {
                        Some(crate::waitable::Waitable::Future(e.id))
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
            B::StreamNew => {
                // canon stream.new — `CanonicalABI.md` §`canon {stream,future}.new`:
                //
                //     (canon stream.new $stream_t (core func $f))
                //     $f : (func (result i64))
                //     return [ ri | (wi << 32) ]
                //
                // ONE i64, not two i32s: readable end in the LOW 32 bits,
                // writable end in the HIGH 32. This pushed two separate stack
                // values, which is not a signature any conforming module can
                // call — and it made the built-in unusable from WAT, because a
                // frontend that sees `(result i64)` reads one value while a
                // frontend told `(result i32 i32)` goes looking for the
                // multi-value pack convention. Neither finds two bare pushes.
                let stream_id = self.event_loop.borrow_mut().create_stream();
                let rd =
                    self.handle_table
                        .insert(crate::handle_table::HandleEntry::ReadableStreamEnd(
                            crate::handle_table::StreamEnd::new(stream_id),
                        ));
                let wr =
                    self.handle_table
                        .insert(crate::handle_table::HandleEntry::WritableStreamEnd(
                            crate::handle_table::StreamEnd::new(stream_id),
                        ));
                self.push(Value::I64((rd as i64) | ((wr as i64) << 32)))?;
            }
            B::StreamWrite => {
                // canon stream.write — the mirror of `stream.read`, and the
                // spec implements both through ONE `stream_copy`:
                //
                //     $f : (func (param i32 T T) (result T))
                //     (handle, ptr, n) -> packed CopyResult
                //
                // Elements come FROM linear memory. This used to take the item
                // as a stack VALUE, which no conforming component could
                // produce — it worked only because both ends were ours, and it
                // was the last shape mismatch in the stream family.
                let n = self.pop().as_i32();
                let ptr = self.pop().as_i32();
                let handle = self.pop().as_i32() as u32;

                let end = match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::WritableStreamEnd(e)) => *e,
                    _ => {
                        return Err(VMError::new(
                            "canon stream.write: handle is not a writable stream end",
                        ));
                    }
                };
                if end.state != crate::handle_table::CopyState::Idle {
                    return Err(VMError::new(
                        "canon stream.write: end is not IDLE — pipelined copies are not permitted",
                    ));
                }
                if n < 0 || ptr < 0 || n as u32 > crate::canon_copy::MAX_LENGTH {
                    return Err(VMError::new(
                        "canon stream.write: buffer length out of range",
                    ));
                }
                let (ptr, n) = (ptr as usize, n as usize);

                // `n` counts ELEMENTS, so a typed stream lifts `n` values at
                // the element stride. Reading `n` BYTES and pushing them as
                // items would corrupt the stream silently: the reader's typed
                // path would then lower those bytes as if each were a whole
                // element. `stream<u8>` keeps the byte path below unchanged.
                let typed_elem = self.event_loop.borrow().stream_elem(end.id);
                if let Some(elem) = typed_elem {
                    return self.stream_write_typed(end, ptr, n, &elem);
                }

                if ptr.saturating_add(n) > self.memory.len() {
                    return Err(VMError::new(
                        "canon stream.write: buffer is out of bounds of linear memory",
                    ));
                }

                let mut bytes = vec![0u8; n];
                self.memory.read_bytes(ptr, &mut bytes);
                // One item per copy, carrying the whole buffer: the reader
                // flattens items back to bytes, so the item boundary is
                // invisible downstream (`EventLoop::stream_read_bytes`).
                let item = Value::Object(crate::heap::alloc(crate::value::Object::new_array(
                    bytes.iter().map(|b| Value::I32(*b as i32)).collect(),
                )));
                let mut el = self.event_loop.borrow_mut();
                if let Some(fiber) = el.stream_push(end.id, item) {
                    el.immediate
                        .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                }
                drop(el);
                // Data is available now, so a reader parked in COPYING by a
                // BLOCKED read may copy again — the reset the spec performs on
                // event delivery. Without it, BLOCKED is a dead end.
                self.handle_table.release_copying(end.id, true);
                self.push(Value::I32(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Completed,
                    n as u32,
                ) as i32))?;
            }
            B::StreamDropReadable => {
                // canon stream.drop-readable — pops readable stream handle.
                let handle = self.pop().as_i32() as u32;
                if let Some(crate::handle_table::HandleEntry::ReadableStreamEnd(end)) =
                    self.handle_table.remove(handle)
                {
                    // Close the stream so waiting writers don't block forever.
                    let mut el = self.event_loop.borrow_mut();
                    if let Some(fiber) = el.stream_close(end.id) {
                        el.immediate
                            .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                    }
                }
            }
            B::StreamDropWritable => {
                // canon stream.drop-writable — pops writable stream handle.
                let handle = self.pop().as_i32() as u32;
                if let Some(crate::handle_table::HandleEntry::WritableStreamEnd(end)) =
                    self.handle_table.remove(handle)
                {
                    // Closing the write end signals EOF to the reader.
                    let mut el = self.event_loop.borrow_mut();
                    if let Some(fiber) = el.stream_close(end.id) {
                        el.immediate
                            .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                    }
                }
            }
            B::StreamRead => {
                // canon stream.read — `CanonicalABI.md` §`canon stream.{read,write}`:
                //
                //     (canon stream.read $stream_t $opts (core func $f))
                //     $f : (func (param i32 T T) (result T))     T = i32
                //
                // i.e. `(handle, ptr, n) -> packed`, copying ELEMENTS INTO
                // LINEAR MEMORY and answering a packed `CopyResult` + count.
                //
                // ⚠ This built-in previously took the stream as a stack VALUE
                // and pushed the item itself. That shape cannot interoperate:
                // a conforming runtime hands core wasm a handle and a buffer.
                // It also could not read what `stream.new` produced, because
                // `stream.new` yields i32 handles and this arm only accepted
                // the object form — an i32 fell through to `push(Null)`, which
                // a guest reads as EOF. Silent-empty, never an error.
                //
                // ASYNC variant (`opts.async_`): a copy that cannot complete
                // synchronously answers `BLOCKED` and the real result arrives
                // later as an `EventCode.STREAM_READ` event. The SYNC variant
                // instead suspends until an event exists; that needs a resume
                // path which re-enters the copy (the current suspension
                // mechanism resumes by pushing exactly one value, which cannot
                // both fill a buffer and return a count), so it is not wired
                // yet and this is the async shape.
                let n = self.pop().as_i32();
                let ptr = self.pop().as_i32();
                let handle = self.pop().as_i32() as u32;

                // §stream_copy trap conditions, in spec order.
                let end = match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e)) => *e,
                    _ => {
                        return Err(VMError::new(
                            "canon stream.read: handle is not a readable stream end",
                        ));
                    }
                };
                if end.state != crate::handle_table::CopyState::Idle {
                    return Err(VMError::new(
                        "canon stream.read: end is not IDLE — pipelined copies are not permitted",
                    ));
                }
                if end.in_waitable_set {
                    return Err(VMError::new(
                        "canon stream.read: synchronous read on an end already awaited via a waitable set",
                    ));
                }
                // `Buffer`'s constructor eagerly checks the bounds of (ptr, n),
                // and MAX_LENGTH is fixed independently of the address type.
                if n < 0 || ptr < 0 || n as u32 > crate::canon_copy::MAX_LENGTH {
                    return Err(VMError::new("canon stream.read: buffer length out of range"));
                }
                let (ptr, n) = (ptr as usize, n as usize);
                if ptr.saturating_add(n) > self.memory.len() {
                    return Err(VMError::new(
                        "canon stream.read: buffer is out of bounds of linear memory",
                    ));
                }

                // `n` is a count of ELEMENTS. For `stream<u8>` that is also a
                // byte count, which is why the byte path below was correct for
                // every stream that existed before typed elements did.
                let typed_elem = self.event_loop.borrow().stream_elem(end.id);
                if let Some(elem) = typed_elem {
                    return self.stream_read_typed(handle, end, ptr, n, &elem);
                }

                // Same as the typed path: the producer is a last resort the
                // event loop drives, never an inline blocking call.
                let bytes = self.event_loop.borrow_mut().stream_read_bytes(end.id, n);
                if bytes.is_empty() {
                    // Nothing copied: either the far end is gone for good, or
                    // it simply has not written yet.
                    if self.event_loop.borrow().stream_is_eof(end.id) {
                        // DROPPED means no further copies are possible, so the
                        // end goes to DONE and anything but `drop-*` now traps.
                        if let Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e)) =
                            self.handle_table.get_mut(handle)
                        {
                            e.state = crate::handle_table::CopyState::Done;
                        }
                        self.push(Value::I32(crate::canon_copy::pack(
                            crate::canon_copy::CopyResult::Dropped,
                            0,
                        ) as i32))?;
                    } else {
                        // §stream_copy sets `e.state = CopyState.COPYING`
                        // BEFORE the copy, and only the delivered event resets
                        // it. So a parked read leaves the end COPYING: that is
                        // what makes a subsequent `cancel-read` legal (it traps
                        // unless COPYING) and what makes a second concurrent
                        // read trap (it traps unless IDLE). Staying IDLE would
                        // quietly permit both.
                        if let Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e)) =
                            self.handle_table.get_mut(handle)
                        {
                            e.state = crate::handle_table::CopyState::Copying;
                        }
                        // 🔀 ASYNC (`opts.async_`): answer BLOCKED now and let
                        // the real result arrive later as an
                        // `EventCode.STREAM_READ` event. Only the async form
                        // may do this — `CanonicalABI.md` §canon
                        // stream.{read,write} — which is exactly why the sync
                        // form below suspends instead.
                        //
                        // ⚠ The end stays COPYING, deliberately: the copy IS in
                        // flight. A caller that wants POSIX `EAGAIN` retry
                        // semantics must issue `stream.cancel-read` before
                        // reading again, which is the only thing that returns
                        // an end to IDLE; a bare retry traps on "not IDLE".
                        if self.canon_async_opt() {
                            self.push(Value::I32(crate::canon_copy::BLOCKED as i32))?;
                            return Ok(());
                        }
                        // SUSPEND — this is the synchronous variant. Answering
                        // `BLOCKED` here is what every reader in the tree was
                        // written against, and it is why they all had to break
                        // out of their drain loop on it: on a file that reads
                        // as one short answer, but on a socket, where nothing
                        // ready is the ordinary case, it is silent truncation.
                        return Err(self.park_sync_copy(crate::fiber::PendingCopy {
                            handle,
                            end_id: end.id,
                            ptr,
                            n,
                            kind: crate::fiber::PendingCopyKind::StreamBytes,
                        }));
                    }
                } else {
                    self.write_memory_bytes(0, ptr, &bytes)?;
                    self.push(Value::I32(crate::canon_copy::pack(
                        crate::canon_copy::CopyResult::Completed,
                        bytes.len() as u32,
                    ) as i32))?;
                }
            }
            // canon {stream,future}.cancel-{read,write} —
            // `CanonicalABI.md` §`canon {stream,future}.cancel-{read,write}`:
            //
            //     $f : (func (param i32) (result i32))
            //
            // All four funnel through one `cancel_copy`, so they do here too.
            //
            // ⚠ This used to take the stream as a stack VALUE and CLOSE it,
            // returning nothing. Cancel is not close: it reclaims a buffer from
            // an in-flight copy and leaves the stream usable. Closing the far
            // end on a cancel would give every waiting reader a spurious EOF.
            B::StreamCancelRead => {
                self.canon_cancel_copy(crate::canon_copy::EndKind::ReadableStream)?;
            }
            B::StreamCancelWrite => {
                self.canon_cancel_copy(crate::canon_copy::EndKind::WritableStream)?;
            }
            B::FutureCancelRead => {
                self.canon_cancel_copy(crate::canon_copy::EndKind::ReadableFuture)?;
            }
            B::FutureCancelWrite => {
                self.canon_cancel_copy(crate::canon_copy::EndKind::WritableFuture)?;
            }
            // canon future.{read,write} — `CanonicalABI.md` §`canon
            // future.{read,write}`:
            //
            //     $f : (func (param i32 T) (result i32))
            //
            // `(handle, ptr)` with NO count, because a future carries exactly
            // one element — the spec fixes the buffer length to 1. Its SIZE
            // comes from the `$t` immediate, which is why this could not exist
            // until canon imports could carry one.
            B::FutureRead => {
                let ptr = self.pop().as_i32();
                let handle = self.pop().as_i32() as u32;
                let t = self.canon_element_type("future.read")?;
                let end = match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::ReadableFutureEnd(e)) => *e,
                    _ => {
                        return Err(VMError::new(
                            "canon future.read: handle is not a readable future end",
                        ));
                    }
                };
                if end.state != crate::handle_table::CopyState::Idle {
                    return Err(VMError::new(
                        "canon future.read: end is not IDLE — pipelined copies are not permitted",
                    ));
                }
                if end.in_waitable_set {
                    return Err(VMError::new(
                        "canon future.read: synchronous read on an end already awaited via a waitable set",
                    ));
                }
                let settled = {
                    let el = self.event_loop.borrow();
                    el.future_states.get(&end.id).map(|r| (r.phase, r.value.clone()))
                };
                match settled {
                    Some((crate::event_loop::FuturePhase::Resolved, Some(v))) => {
                        // `store_with`, not `store`: a future's element is a
                        // whole component type, and `future<result<_,
                        // error-code>>` — what every 0.3.1 stream-producing
                        // call answers as tuple element 1 — needs a realloc
                        // for its payload. Scalar-only `store` reported
                        // OutOfMemory for exactly the shape futures carry most.
                        let memory = self.memory.clone();
                        let mut bump = self.canon_bump_start();
                        crate::canon_value::store_with(
                            &memory,
                            &mut Self::bump_realloc(&memory, &mut bump),
                            &v,
                            &t,
                            ptr as u32,
                        )
                        .map_err(|e| VMError::new(format!("canon future.read: {e}")))?;
                        self.canon_bump_commit(bump);
                        self.push(Value::I32(crate::canon_copy::pack(
                            crate::canon_copy::CopyResult::Completed,
                            1,
                        ) as i32))?;
                    }
                    // A rejected future is the writable end going away without
                    // ever producing a value: no further copies are possible.
                    Some((crate::event_loop::FuturePhase::Rejected, _)) | None => {
                        if let Some(crate::handle_table::HandleEntry::ReadableFutureEnd(e)) =
                            self.handle_table.get_mut(handle)
                        {
                            e.state = crate::handle_table::CopyState::Done;
                        }
                        self.push(Value::I32(crate::canon_copy::pack(
                            crate::canon_copy::CopyResult::Dropped,
                            0,
                        ) as i32))?;
                    }
                    _ => {
                        if let Some(crate::handle_table::HandleEntry::ReadableFutureEnd(e)) =
                            self.handle_table.get_mut(handle)
                        {
                            e.state = crate::handle_table::CopyState::Copying;
                        }
                        // 🔀 ASYNC (`opts.async_`) — the third and last park
                        // site. All three branch on the same option, so a
                        // future, a typed stream and a `stream<u8>` cannot
                        // disagree about what `async` means.
                        if self.canon_async_opt() {
                            self.push(Value::I32(crate::canon_copy::BLOCKED as i32))?;
                            return Ok(());
                        }
                        // Pending, and this is the synchronous variant: suspend
                        // until it settles rather than answering `BLOCKED`.
                        // Every 0.3.1 write path ends in a
                        // `future<result<_, error-code>>` whose whole job is to
                        // carry the failure — a reader that gave up on BLOCKED
                        // read "no error" off a future that had not answered.
                        return Err(self.park_sync_copy(crate::fiber::PendingCopy {
                            handle,
                            end_id: end.id,
                            ptr: ptr as usize,
                            n: 1,
                            kind: crate::fiber::PendingCopyKind::Future(t.clone()),
                        }));
                    }
                }
            }
            B::FutureWrite => {
                let ptr = self.pop().as_i32();
                let handle = self.pop().as_i32() as u32;
                let t = self.canon_element_type("future.write")?;
                let end = match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::WritableFutureEnd(e)) => *e,
                    _ => {
                        return Err(VMError::new(
                            "canon future.write: handle is not a writable future end",
                        ));
                    }
                };
                if end.state != crate::handle_table::CopyState::Idle {
                    return Err(VMError::new(
                        "canon future.write: end is not IDLE — pipelined copies are not permitted",
                    ));
                }
                let v = crate::canon_value::load(&self.memory, &t, ptr as u32)
                    .map_err(|e| VMError::new(format!("canon future.write: {e}")))?;
                let mut el = self.event_loop.borrow_mut();
                if let Some(fiber) = el.resolve_future(end.id, v) {
                    el.immediate
                        .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                }
                drop(el);
                // The value is available, so a reader parked in COPYING by a
                // BLOCKED read can copy again — this is the reset the spec
                // performs when it delivers the event.
                self.handle_table.release_copying(end.id, true);
                // A future takes exactly one value, so a successful write is
                // always one element.
                self.push(Value::I32(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Completed,
                    1,
                ) as i32))?;
            }
            // canon resource.new — `CanonicalABI.md` §`canon resource.new`:
            //
            //     $f : (func (param $rt.rep) (result i32))     $rt.rep = i32
            //
            // Wraps a REPRESENTATION (an opaque i32 the component chose) in an
            // owning handle. The rep is the component's private business; the
            // handle is what crosses a boundary, which is the whole point of
            // the indirection — a peer never sees the representation.
            // 📝 canon error-context.new — `CanonicalABI.md:5147`:
            //
            //     def canon_error_context_new(opts, ptr, tagged_code_units):
            //       trap_if(not inst.may_leave)
            //       if DETERMINISTIC_PROFILE or random.randint(0,1):
            //         s = String(('', 'utf8', 0))
            //       else:
            //         s = host_defined_transformation(load_string_from_range(...))
            //       i = inst.handles.add(ErrorContext(s))
            //
            // ⛔ The spec permits the host to DISCARD the message — that branch
            // exists so a production host can skip the cost. We PRESERVE it,
            // which is equally conformant and is the only choice that makes the
            // feature worth having: an error-context whose message is always
            // empty is a handle that aids no debugging.
            //
            // Not the deterministic profile either. That flag also gates NaN
            // scrambling, so claiming it here would assert a whole execution
            // profile the rest of this VM does not implement.
            B::ErrorContextNew => {
                let len = self.pop().as_i32() as usize;
                let ptr = self.pop().as_i32() as usize;
                let memory = self.memory.clone();
                let msg = crate::canon_value::read_utf8(&memory, ptr, len)
                    .map_err(|e| VMError::new(format!("canon error-context.new: {e}")))?;
                let h = self
                    .handle_table
                    .insert(crate::handle_table::HandleEntry::ErrorContext {
                        debug_message: msg,
                    });
                self.push(Value::I32(h as i32))?;
            }
            // 📝 canon error-context.debug-message — `CanonicalABI.md:5189`.
            // `store_string(cx, errctx.debug_message, ptr)` — the (ptr, length)
            // pair goes at `ptr`, the bytes into freshly `realloc`ed memory,
            // which is exactly what storing a `ValType::String` does.
            B::ErrorContextDebugMessage => {
                let ptr = self.pop().as_i32() as u32;
                let handle = self.pop().as_i32() as u32;
                let msg = match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::ErrorContext { debug_message }) => {
                        debug_message.clone()
                    }
                    // `trap_if(not isinstance(errctx, ErrorContext))` — naming
                    // what the handle IS, because "wrong handle" and "no such
                    // handle" are different mistakes to chase.
                    Some(other) => {
                        return Err(VMError::new(format!(
                            "canon error-context.debug-message: handle {handle} is a {} , \
                             not an error-context (trap)",
                            crate::handle_table::HandleEntry::kind_name(other)
                        )))
                    }
                    None => {
                        return Err(VMError::new(format!(
                            "canon error-context.debug-message: handle {handle} is not in the \
                             instance handle table (trap)"
                        )))
                    }
                };
                let memory = self.memory.clone();
                let mut bump = self.canon_bump_start();
                {
                    let mut alloc = Self::bump_realloc(&memory, &mut bump);
                    let mut realloc: crate::canon_value::Realloc<'_> = &mut alloc;
                    crate::canon_value::store_with(
                        &memory,
                        &mut realloc,
                        &Value::String(msg.into()),
                        &crate::component::ValType::String,
                        ptr,
                    )
                    .map_err(|e| {
                        VMError::new(format!("canon error-context.debug-message: {e}"))
                    })?;
                }
                self.canon_bump_commit(bump);
            }
            // 📝 canon error-context.drop — `CanonicalABI.md:5215`.
            // `remove` then `trap_if(not isinstance(...))`: the handle must be
            // CHECKED before it is released, or a mistyped drop has already
            // freed someone else's entry by the time it traps.
            B::ErrorContextDrop => {
                let handle = self.pop().as_i32() as u32;
                match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::ErrorContext { .. }) => {}
                    Some(other) => {
                        return Err(VMError::new(format!(
                            "canon error-context.drop: handle {handle} is a {}, not an \
                             error-context (trap)",
                            crate::handle_table::HandleEntry::kind_name(other)
                        )))
                    }
                    None => {
                        return Err(VMError::new(format!(
                            "canon error-context.drop: handle {handle} is not in the instance \
                             handle table (trap)"
                        )))
                    }
                }
                self.handle_table.remove(handle);
            }
            B::ResourceNew => {
                let rep = self.pop().as_i32();
                let type_id = self.canon_type_index().unwrap_or(0);
                let h = self
                    .handle_table
                    .insert(crate::handle_table::HandleEntry::OwnedResource {
                        type_id,
                        value: Value::I32(rep),
                    });
                self.push(Value::I32(h as i32))?;
            }
            // canon resource.rep — §`canon resource.rep`:
            //
            //     $f : (func (param i32) (result $rt.rep))
            //
            // The inverse, and only valid for a handle of the SAME resource
            // type: `trap_if(h.rt is not rt)`. Without that check a component
            // could read another type's representation through a handle it
            // legitimately holds.
            B::ResourceRep => {
                let handle = self.pop().as_i32() as u32;
                let want = self.canon_type_index().unwrap_or(0);
                match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::OwnedResource { type_id, value })
                    | Some(crate::handle_table::HandleEntry::BorrowedResource {
                        type_id, value, ..
                    }) => {
                        if self.canon_type_index().is_some() && *type_id != want {
                            return Err(VMError::new(format!(
                                "canon resource.rep: handle is resource type {type_id}, not {want}"
                            )));
                        }
                        let rep = value.as_i32();
                        self.push(Value::I32(rep))?;
                    }
                    _ => {
                        return Err(VMError::new(
                            "canon resource.rep: handle is not a resource handle",
                        ));
                    }
                }
            }
            // canon resource.drop — §`canon resource.drop`:
            //
            //     $f : (func (param i32))
            //
            // Removes the handle and, IF IT WAS OWNING, calls the resource's
            // destructor. A borrow is dropped without one — that asymmetry is
            // the difference between `own` and `borrow`, and getting it wrong
            // means either a leak or a double free.
            B::ResourceDrop => {
                let handle = self.pop().as_i32() as u32;
                let want = self.canon_type_index();
                match self.handle_table.get(handle) {
                    Some(crate::handle_table::HandleEntry::OwnedResource { type_id, .. })
                    | Some(crate::handle_table::HandleEntry::BorrowedResource {
                        type_id, ..
                    }) => {
                        if let Some(want) = want {
                            if *type_id != want {
                                return Err(VMError::new(format!(
                                    "canon resource.drop: handle is resource type {type_id}, not {want}"
                                )));
                            }
                        }
                    }
                    _ => {
                        return Err(VMError::new(
                            "canon resource.drop: handle is not a resource handle",
                        ));
                    }
                }
                let removed = self.handle_table.remove(handle);
                // ⚠ The destructor is NOT called yet. `component_model.rs`
                // registers one as a `[resource-drop]` ResourceMethod, but
                // invoking it from here means re-entering the VM mid-builtin,
                // which the current dispatch cannot do. Left explicit rather
                // than silently skipped: an owning drop that runs no destructor
                // leaks whatever the representation owned.
                let _ = removed;
            }
            B::WaitableSetDrop => {
                // canon waitable-set.drop — pops the set handle and releases it.
                // Without this a set could be created and waited on but never
                // freed, so a long-running component leaks one per wait site.
                let handle = self.pop().as_i32() as u32;
                self.waitable_sets.remove(handle);
            }
            B::ThreadYield => {
                // canon thread.yield — `CanonicalABI.md`:
                //
                //     def canon_thread_yield(cancellable):
                //       thread = current_thread()
                //       trap_if(not thread.task.inst.may_leave)
                //       cancelled = thread.yield_(cancellable)
                //       return [cancelled]
                //
                // `current_thread()` is UNCONDITIONAL, exactly as in
                // `thread.index`. An earlier comment here claimed yield needs
                // no current thread; that is not what the spec says, and core
                // wasm inside a real component is always inside a lifted call.
                let cancellable = self.canon_cancellable();
                let _me = self.current_thread.ok_or_else(|| {
                    VMError::new(
                        "canon thread.yield: no current thread — a thread exists only \
                         inside a `canon lift`ed call (trap)",
                    )
                })?;
                // `yield_` is `wait_until(lambda: True)`, and `wait_until` may
                // return early when the readiness condition already holds —
                // the embedder's choice of whether to switch. We keep running.
                let cancelled = self.deliver_pending_cancel_now(cancellable);
                self.push(Value::I32(i32::from(cancelled)))?;
            }
            B::FutureNew => {
                // canon future.new — same shape as `stream.new` above and from
                // the same spec paragraph: `(func (result i64))`,
                // `return [ ri | (wi << 32) ]`.
                let future_id = self.event_loop.borrow_mut().create_future();
                let rd =
                    self.handle_table
                        .insert(crate::handle_table::HandleEntry::ReadableFutureEnd(
                            crate::handle_table::StreamEnd::new(future_id),
                        ));
                let wr =
                    self.handle_table
                        .insert(crate::handle_table::HandleEntry::WritableFutureEnd(
                            crate::handle_table::StreamEnd::new(future_id),
                        ));
                self.push(Value::I64((rd as i64) | ((wr as i64) << 32)))?;
            }
            B::FutureDropReadable => {
                // canon future.drop-readable — pops readable future handle (i32).
                let handle = self.pop().as_i32() as u32;
                self.handle_table.remove(handle);
            }
            B::FutureDropWritable => {
                // canon future.drop-writable — pops writable future handle (i32).
                let handle = self.pop().as_i32() as u32;
                if let Some(crate::handle_table::HandleEntry::WritableFutureEnd(end)) =
                    self.handle_table.remove(handle)
                {
                    // Dropping the write end without resolving rejects the future.
                    let mut el = self.event_loop.borrow_mut();
                    if let Some(fiber) =
                        el.reject_future(end.id, Value::String(Arc::from("future dropped")))
                    {
                        el.immediate
                            .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                    }
                }
            }
            // canon backpressure.inc / backpressure.dec — CM3 replaced the
            // boolean `backpressure.set` (retired) with a counter: the
            // instance resists new calls while > 0. No args, no results.
            B::ThreadIndex => {
                // canon thread.index — `CanonicalABI.md`:
                //     thread = current_thread()
                //     assert(thread.index is not None)
                //     return [thread.index]
                // `assert`, not a default: outside a lifted call there IS no
                // current thread, and answering 0 would name another slot.
                let idx = self.current_thread.ok_or_else(|| {
                    VMError::new(
                        "canon thread.index: no current thread — a thread exists only \
                         inside a `canon lift`ed call (trap)",
                    )
                })?;
                self.push(Value::I32(idx as i32))?;
            }
            B::ThreadResumeLater => {
                // canon thread.resume-later — `CanonicalABI.md`:
                //     other_thread = inst.threads.get(i)
                //     trap_if(not other_thread.suspended())
                //     other_thread.resume_later()
                // Never suspends the CURRENT thread, which is why this row
                // carries no `cancellable` immediate: there is no suspension
                // point at which a cancellation could be delivered.
                let i = self.pop().as_i32() as u32;
                let thread = self.cm_instance.threads.get_mut(i).ok_or_else(|| {
                    VMError::new(format!(
                        "canon thread.resume-later: no thread at index {i} (trap)"
                    ))
                })?;
                thread.resume_later().map_err(|e| {
                    VMError::new(format!("canon thread.resume-later: {e} (trap)"))
                })?;
            }
            // The four compound handoffs are a 2x2 in the spec and get one
            // implementation here for the same reason: **what happens to me**
            // (suspend = park, yield = stay runnable) x **what happens to
            // them** (resume = switch unconditionally, promote = switch only
            // if they are ready, else fall back to plain suspend/yield).
            B::ThreadSuspend => {
                // canon thread.suspend — `CanonicalABI.md:4962`:
                //
                //     def canon_thread_suspend(cancellable):
                //       thread = current_thread()
                //       trap_if(not thread.task.inst.may_leave)
                //       cancelled = thread.suspend(cancellable)
                //       return [cancelled]
                //
                // and `Thread.suspend` is `deliver_pending_cancel` then
                // `block_internal` — a block with NO `switch_to`, which is the
                // one thing separating this row from the four handoffs.
                let cancellable = self.canon_cancellable();
                let me = self.current_thread.ok_or_else(|| {
                    VMError::new(
                        "canon thread.suspend: no current thread — a thread exists only \
                         inside a `canon lift`ed call (trap)",
                    )
                })?;
                if self.deliver_pending_cancel_now(cancellable) {
                    self.push(Value::I32(1))?;
                } else {
                    self.thread_block("thread.suspend", me)?;
                }
            }
            B::ThreadSpawnRef => {
                // canon thread.spawn-ref 🧵② — `CanonicalABI.md`:
                //
                //     [i] = canon_thread_new_ref(shared, ft, f, c)
                //     []  = canon_thread_resume_later(shared, i)
                //     return [i]
                //
                // "fuses thread.new-ref and thread.resume-later, allowing
                // thread-creation to skip the intermediate suspended state".
                // `canon_thread_new_ref` is not itself defined in the spec yet
                // (it arrives with the GC ABI option) but is specified as
                // "like canon_thread_new_indirect minus the table access and
                // type check" — so the funcref arrives as a VALUE, not an
                // index, and that is the only difference here.
                self.refuse_shared_threads("thread.spawn-ref")?;
                let closure = self.pop();
                let funcref = self.pop();
                let index = self.create_thread_over("thread.spawn-ref", funcref, closure)?;
                self.resume_thread_later("thread.spawn-ref", index)?;
                self.push(Value::I32(index as i32))?;
            }
            B::ThreadSpawnIndirect => {
                // canon thread.spawn-indirect 🧵② — the same fusion over
                // `thread.new-indirect`, so it takes the table path: `$ftbl`
                // is an immediate, `fi` and `c` are runtime args.
                self.refuse_shared_threads("thread.spawn-indirect")?;
                let closure = self.pop();
                let fi = self.pop().as_i32();
                let funcref = self.thread_table_funcref("thread.spawn-indirect", fi)?;
                let index = self.create_thread_over("thread.spawn-indirect", funcref, closure)?;
                self.resume_thread_later("thread.spawn-indirect", index)?;
                self.push(Value::I32(index as i32))?;
            }
            B::ThreadAvailableParallelism => {
                // canon thread.available-parallelism 🧵② — "the number of
                // threads the underlying hardware can be expected to execute in
                // PARALLEL", and "not allowed to change over the lifetime of a
                // component instance".
                //
                // Cooperative fibers execute exactly one thread at a time, so
                // the true answer is 1. This is not the deterministic profile's
                // `return [1]` standing in for a real number — it IS the real
                // number for this scheduler, and it is constant by construction.
                //
                // `shared?` is not refused here: this row spawns nothing, it
                // only reports a count, and the count it reports is honest.
                self.push(Value::I32(1))?;
            }
            B::ThreadSuspendThenResume => self.exec_thread_handoff(b, false, false)?,
            B::ThreadYieldThenResume => self.exec_thread_handoff(b, true, false)?,
            B::ThreadSuspendThenPromote => self.exec_thread_handoff(b, false, true)?,
            B::ThreadYieldThenPromote => self.exec_thread_handoff(b, true, true)?,
            B::ThreadNewIndirect => {
                // canon thread.new-indirect — `CanonicalABI.md`:
                //
                //     f = ftbl.get(fi)
                //     trap_if(f.t != ft)
                //     def thread_func(): call_and_trap_on_throw(f.callee, [c])
                //     new_thread = Thread(task, thread_func)
                //     task.register_thread(new_thread)
                //     return [new_thread.index]
                //
                // Runtime args are `(fi, c)`; `$ft` and `$ftbl` are IMMEDIATES
                // on the canon definition — which is why this row was
                // unreachable until the canon section existed to carry them.
                //
                // The new thread starts SUSPENDED holding a continuation it has
                // not entered. That is the whole contract: "core wasm must call
                // one of the other `thread.*` built-ins" to start it, so
                // creating a thread must never run it. `thread.spawn-indirect`
                // is exactly this row plus `resume-later` and shares both
                // helpers, so the two cannot drift apart.
                let closure = self.pop();
                let fi = self.pop().as_i32();
                let funcref = self.thread_table_funcref("thread.new-indirect", fi)?;
                let index = self.create_thread_over("thread.new-indirect", funcref, closure)?;
                self.push(Value::I32(index as i32))?;
            }
            B::BackpressureInc => {
                // canon backpressure.inc — `CanonicalABI.md`:
                //
                //     inst.backpressure += 1
                //     trap_if(inst.backpressure == 2**16)
                //
                // `saturating_add` was a SILENT failure: at the ceiling the
                // counter stops moving while `dec` keeps decrementing, so the
                // pairing is lost and backpressure releases early — the exact
                // situation the trap exists to prevent.
                //
                // The counter is per INSTANCE (`current_instance()`), which is
                // why it lives on `cm_instance` and not on the task: two tasks
                // in one instance share the backpressure that gates entry to
                // that instance.
                self.cm_instance
                    .backpressure_inc()
                    .map_err(|e| VMError::new(format!("canon backpressure.inc: {e}")))?;
            }
            B::BackpressureDec => {
                // canon backpressure.dec — `trap_if(inst.backpressure < 0)`.
                // `saturating_sub` turned an unbalanced `dec` into a no-op, so
                // a missing `inc` never surfaced.
                self.cm_instance
                    .backpressure_dec()
                    .map_err(|e| VMError::new(format!("canon backpressure.dec: {e}")))?;
            }
            B::ContextGet => {
                // canon context.get — `CanonicalABI.md`:
                //
                //     thread = current_thread()
                //     assert(i < len(thread.storage))
                //     return [thread.storage[i]]
                //
                // Storage is per-THREAD. It lived in a process-wide
                // `VM::context_slots`, so two threads shared one context —
                // the opposite of what thread-local storage is for. Now that
                // `canon lift` creates the implicit thread, the array has its
                // real owner and the VM-wide field is gone.
                let (thread_idx, index) = self.canon_context_slot("context.get")?;
                let val = self
                    .cm_instance
                    .threads
                    .get(thread_idx)
                    .and_then(|t| t.storage.get(index).cloned())
                    .ok_or_else(|| {
                        VMError::new(format!(
                            "canon context.get: no thread at index {thread_idx} (trap)"
                        ))
                    })?;
                self.push(val)?;
            }
            B::ContextSet => {
                // canon context.set — same owner, same bound:
                //
                //     thread.storage[i] = v
                let val = self.pop();
                let (thread_idx, index) = self.canon_context_slot("context.set")?;
                let slot = self
                    .cm_instance
                    .threads
                    .get_mut(thread_idx)
                    .and_then(|t| t.storage.get_mut(index))
                    .ok_or_else(|| {
                        VMError::new(format!(
                            "canon context.set: no thread at index {thread_idx} (trap)"
                        ))
                    })?;
                *slot = val;
            }
        }
        Ok(())
    }

    // (The optional `0xEE 0x00 <memidx u16 BE>` selector reader is deleted:
    // memory.size/grow/fill/copy/init carry fixed u16 memidx immediates now,
    // declared in their OperandFormat so format-driven walks stay in sync.)

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
    /// Where host-side canonical lowering allocates from — `cx.opts.realloc`.
    ///
    /// The SAME bump global compiler-emitted marshalling uses
    /// (`vybe_compiler::primitives::canon_marshal`). Two independent bump
    /// pointers over one linear memory would eventually hand the same address
    /// to a guest string and a host-stored one, so every canonical lowering in
    /// this file goes through this pair.
    fn canon_bump_start(&self) -> u32 {
        self.global(canon_marshal_bump())
            .map(|v| v.as_i32() as u32)
            .unwrap_or(0)
    }

    fn canon_bump_commit(&mut self, bump: u32) {
        self.set_global_owned(canon_marshal_bump().to_string(), Value::I32(bump as i32));
    }

    /// A `Realloc` over `bump`, growing linear memory a page at a time.
    fn bump_realloc<'a>(
        memory: &'a crate::shared_memory::SharedMemory,
        bump: &'a mut u32,
    ) -> impl FnMut(u32, u32) -> Option<u32> + 'a {
        move |size: u32, align: u32| -> Option<u32> {
            if *bump == 0 {
                let pages = memory.grow(1);
                if pages == usize::MAX {
                    return None;
                }
                *bump = (pages * 65536) as u32;
            }
            let at = crate::canon_layout::align_to(*bump, align.max(1));
            let end = at.checked_add(size)?;
            while (end as usize) > memory.len() {
                if memory.grow(1) == usize::MAX {
                    return None;
                }
            }
            // Left 8-ALIGNED, not at `end`. This global is shared with
            // `canon_marshal::emit_alloc`/`emit_store_utf8`, which take the
            // value as their buffer address DIRECTLY and only guarantee the
            // alignment of what they leave behind. Committing a raw `end` here
            // therefore hands the next compiler-side allocation an arbitrary
            // address — and `canon stream.read` traps on a misaligned element
            // buffer rather than writing it crooked, so `os.scandir` died on
            // "buffer at 829 is not aligned to 4 bytes" the moment a string
            // allocation left the pointer odd.
            //
            // 8 rather than `align`: it is the invariant the other two clients
            // already document ("so a later i64/f64 store lands legally"), and
            // an allocator with one rule is an allocator whose clients cannot
            // disagree about it.
            *bump = crate::canon_layout::align_to(end, 8);
            Some(at)
        }
    }

    /// `canon stream.write` for a `stream<T>` where `T` is not `u8`.
    ///
    /// The mirror of [`stream_read_typed`]: `n` ELEMENTS are lifted out of
    /// linear memory at the canonical stride and pushed as whole items.
    ///
    /// An element type `canon_value::load` cannot lift yet is REFUSED rather
    /// than approximated — the same rule that module already states, and the
    /// reason matters more on this side: a wrong lift here puts a malformed
    /// value into a stream some other component will later read as if it were
    /// well-formed.
    fn stream_write_typed(
        &mut self,
        end: crate::handle_table::StreamEnd,
        ptr: usize,
        n: usize,
        elem: &crate::component::ValType,
    ) -> Result<(), VMError> {
        let stride = crate::canon_layout::elem_size(elem) as usize;
        let align = crate::canon_layout::alignment(elem);
        if ptr as u32 != crate::canon_layout::align_to(ptr as u32, align) {
            return Err(VMError::new(format!(
                "canon stream.write: buffer at {ptr} is not aligned to {align} bytes for this element type"
            )));
        }
        if ptr.saturating_add(stride.saturating_mul(n)) > self.memory.len() {
            return Err(VMError::new(
                "canon stream.write: element buffer is out of bounds of linear memory",
            ));
        }

        let mut items = Vec::with_capacity(n);
        for i in 0..n {
            match crate::canon_value::load(&self.memory, elem, (ptr + stride * i) as u32) {
                Ok(value) => items.push(value),
                Err(e) => return Err(VMError::new(format!("canon stream.write: {e}"))),
            }
        }

        let written = items.len() as u32;
        {
            let mut el = self.event_loop.borrow_mut();
            for item in items {
                if let Some(fiber) = el.stream_push(end.id, item) {
                    el.immediate
                        .push_back(crate::event_loop::Task::ResumeFiber(fiber));
                }
            }
        }
        self.handle_table.release_copying(end.id, true);
        self.push(Value::I32(crate::canon_copy::pack(
            crate::canon_copy::CopyResult::Completed,
            written,
        ) as i32))?;
        Ok(())
    }

    /// `canon stream.read` for a `stream<T>` where `T` is not `u8`.
    ///
    /// `CanonicalABI.md` §`stream_copy`: `n` is a count of ELEMENTS, and each
    /// is lowered into linear memory at its canonical stride. The byte path
    /// this splits from is not a special case of it — a `stream<u8>` copies a
    /// byte RUN, which need not land on an item boundary, whereas a typed
    /// element is all-or-nothing.
    /// Redo the copy a fiber parked on, now that a producer has appeared.
    ///
    /// `Ok(Some(packed))` — the copy happened (or the stream is at EOF and the
    /// answer is `DROPPED`); `Ok(None)` — still nothing to copy, so the caller
    /// re-parks. The three arms mirror the three read paths exactly, because
    /// this IS those paths, run one wake-up later.
    pub(crate) fn perform_pending_copy(
        &mut self,
        p: &crate::fiber::PendingCopy,
    ) -> Result<Option<i32>, VMError> {
        use crate::canon_copy::{pack, CopyResult};
        use crate::fiber::PendingCopyKind as K;

        match &p.kind {
            K::StreamBytes => {
                let bytes = self.event_loop.borrow_mut().stream_read_bytes(p.end_id, p.n);
                if !bytes.is_empty() {
                    self.write_memory_bytes(0, p.ptr, &bytes)?;
                    // The copy completed, so the end leaves COPYING — the reset
                    // the spec performs when it delivers the event.
                    self.handle_table.release_copying(p.end_id, true);
                    return Ok(Some(pack(CopyResult::Completed, bytes.len() as u32) as i32));
                }
                if self.event_loop.borrow().stream_is_eof(p.end_id) {
                    self.mark_end_done(p.handle);
                    return Ok(Some(pack(CopyResult::Dropped, 0) as i32));
                }
                Ok(None)
            }
            K::StreamTyped(elem) => {
                let items = self.event_loop.borrow_mut().stream_read_items(p.end_id, p.n);
                if items.is_empty() {
                    if self.event_loop.borrow().stream_is_eof(p.end_id) {
                        self.mark_end_done(p.handle);
                        return Ok(Some(pack(CopyResult::Dropped, 0) as i32));
                    }
                    return Ok(None);
                }
                let stride = crate::canon_layout::elem_size(elem) as usize;
                let memory = self.memory.clone();
                let mut bump = self.canon_bump_start();
                let mut realloc = Self::bump_realloc(&memory, &mut bump);
                let mut copied = 0u32;
                for (i, item) in items.iter().enumerate() {
                    let at = (p.ptr + stride * i) as u32;
                    if let Err(e) = crate::canon_value::store_with(&memory, &mut realloc, item, elem, at)
                    {
                        return Err(VMError::new(format!("canon stream.read: {e}")));
                    }
                    copied += 1;
                }
                drop(realloc);
                self.canon_bump_commit(bump);
                self.handle_table.release_copying(p.end_id, true);
                Ok(Some(pack(CopyResult::Completed, copied) as i32))
            }
            K::Future(t) => {
                let settled = {
                    let el = self.event_loop.borrow();
                    el.future_states
                        .get(&p.end_id)
                        .map(|r| (r.phase, r.value.clone()))
                };
                match settled {
                    Some((crate::event_loop::FuturePhase::Resolved, Some(v))) => {
                        let memory = self.memory.clone();
                        let mut bump = self.canon_bump_start();
                        crate::canon_value::store_with(
                            &memory,
                            &mut Self::bump_realloc(&memory, &mut bump),
                            &v,
                            t,
                            p.ptr as u32,
                        )
                        .map_err(|e| VMError::new(format!("canon future.read: {e}")))?;
                        self.canon_bump_commit(bump);
                        self.handle_table.release_copying(p.end_id, true);
                        Ok(Some(pack(CopyResult::Completed, 1) as i32))
                    }
                    // Rejected is the writable end going away without ever
                    // producing a value: no further copies are possible.
                    Some((crate::event_loop::FuturePhase::Rejected, _)) | None => {
                        self.mark_end_done(p.handle);
                        Ok(Some(pack(CopyResult::Dropped, 0) as i32))
                    }
                    _ => Ok(None),
                }
            }
        }
    }

    /// Move a readable end to DONE — no further copy is possible on it, so
    /// anything but `drop-*` now traps.
    fn mark_end_done(&mut self, handle: u32) {
        match self.handle_table.get_mut(handle) {
            Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e))
            | Some(crate::handle_table::HandleEntry::ReadableFutureEnd(e)) => {
                e.state = crate::handle_table::CopyState::Done;
            }
            _ => {}
        }
    }

    fn stream_read_typed(
        &mut self,
        handle: u32,
        end: crate::handle_table::StreamEnd,
        ptr: usize,
        n: usize,
        elem: &crate::component::ValType,
    ) -> Result<(), VMError> {
        let stride = crate::canon_layout::elem_size(elem) as usize;
        let align = crate::canon_layout::alignment(elem);
        if ptr as u32 != crate::canon_layout::align_to(ptr as u32, align) {
            return Err(VMError::new(format!(
                "canon stream.read: buffer at {ptr} is not aligned to {align} bytes for this element type"
            )))?;
        }
        if ptr.saturating_add(stride.saturating_mul(n)) > self.memory.len() {
            return Err(VMError::new(
                "canon stream.read: element buffer is out of bounds of linear memory",
            ));
        }

        // The producer is NOT called here. It is a LAST RESORT, driven by the
        // event loop once nothing else can run — see `drive_parked_producers`.
        //
        // Calling it inline deadlocks the ordinary case. `accept-into-stream`
        // waits for a connection, and the code that will make that connection
        // is very often another fiber in this same program: a python
        // `Thread.start()` DEFERS its body, so a server that accepts inline
        // blocks the one execution thread that would have run the client. The
        // read has to park FIRST so the loop can run everyone else; only when
        // it has run out of work is blocking for a peer the right thing to do.
        let items = self.event_loop.borrow_mut().stream_read_items(end.id, n);
        if items.is_empty() {
            if self.event_loop.borrow().stream_is_eof(end.id) {
                if let Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e)) =
                    self.handle_table.get_mut(handle)
                {
                    e.state = crate::handle_table::CopyState::Done;
                }
                self.push(Value::I32(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Dropped,
                    0,
                ) as i32))?;
            } else {
                // Was `COMPLETED` with a count of zero — a copy that reports
                // success without copying anything. A reader cannot tell that
                // from a real short read, so a not-yet-ready typed stream read
                // as an empty one. The synchronous variant SUSPENDS instead.
                if let Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e)) =
                    self.handle_table.get_mut(handle)
                {
                    e.state = crate::handle_table::CopyState::Copying;
                }
                // 🔀 ASYNC — same rule as the byte path: only `opts.async_`
                // may answer BLOCKED. Both paths branch, so a typed stream and
                // a `stream<u8>` cannot disagree about what `async` means.
                if self.canon_async_opt() {
                    self.push(Value::I32(crate::canon_copy::BLOCKED as i32))?;
                    return Ok(());
                }
                return Err(self.park_sync_copy(crate::fiber::PendingCopy {
                    handle,
                    end_id: end.id,
                    ptr,
                    n,
                    kind: crate::fiber::PendingCopyKind::StreamTyped(elem.clone()),
                }));
            }
            return Ok(());
        }

        // Shared with `future.read` and with compiler-emitted marshalling —
        // see `bump_realloc`.
        let memory = self.memory.clone();
        let mut bump = self.canon_bump_start();
        let mut realloc = Self::bump_realloc(&memory, &mut bump);

        let mut copied = 0u32;
        for (i, item) in items.iter().enumerate() {
            let at = (ptr + stride * i) as u32;
            if let Err(e) = crate::canon_value::store_with(&memory, &mut realloc, item, elem, at) {
                return Err(VMError::new(format!("canon stream.read: {e}")))?;
            }
            copied += 1;
        }
        drop(realloc);
        self.canon_bump_commit(bump);

        self.push(Value::I32(crate::canon_copy::pack(
            crate::canon_copy::CopyResult::Completed,
            copied,
        ) as i32))?;
        Ok(())
    }

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
            return Err(VMError::new("trap: out of bounds array access"));
        }
        Ok(elems[start..end]
            .iter()
            .map(|v| v.as_i32() as u32)
            .collect())
    }

    pub(crate) fn read_optional_memarg(&mut self) -> (usize, usize) {
        // OPTIONAL marker-tagged memarg (`SimdMemArg` — the same shape the
        // v128 loads/stores use): present iff the first LEB carries the 0x80
        // marker. The peek is unambiguous — instruction group-hi bytes are
        // always 0x00 — so there is NO opcode-decode guessing. 0x100 =
        // memory64 (u64 offset), 0x40 = explicit memidx LEB follows (the
        // spec multi-memory bit). Absent means align natural, offset 0,
        // memory 0.
        let chunk_idx = self.frame().chunk_index;
        let code = &self.chunks[chunk_idx].code;
        let mut ip = self.frame().ip;
        let align = read_leb_u32(code, &mut ip);
        if align & 0x80 == 0 {
            return (0, 0);
        }
        let offset = if align & 0x100 != 0 {
            read_leb_u64(code, &mut ip) as usize
        } else {
            read_leb_u32(code, &mut ip) as usize
        };
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

    /// Pop a table index/count operand, widening per the table's index type.
    /// Used by table.grow/fill/copy/init.
    ///
    /// WASM reads these operands as UNSIGNED: `0xffff_fff0` is 4294967280, not
    /// -16 and not 0. Clamping with `.max(0)` made an out-of-bounds
    /// `table.fill` silently fill at 0 and `table.grow` report success for an
    /// impossible delta (table_grow.wast:49).
    ///
    /// This NEVER traps. `table.grow` must REPORT -1 for a delta it cannot
    /// satisfy, so the old `table64_index` (which trapped on a negative i64)
    /// was wrong here even ignoring the sign — fill/copy/init get their trap
    /// from their own bounds checks instead. Saturating the widen keeps a
    /// table64 count honest on a 32-bit host: `usize::MAX` fails every bounds
    /// check and every grow limit, which is what a 2^64-scale count must do.
    fn pop_table_count(&mut self, is64: bool) -> usize {
        if is64 {
            usize::try_from(self.pop().as_i64() as u64).unwrap_or(usize::MAX)
        } else {
            self.pop().as_i32() as u32 as usize
        }
    }

    /// The spec's `call_indirect` runtime type check (§4.4.8, step 10): the
    /// funcref taken from the table must have THE TYPE the instruction names.
    ///
    /// ⛔ ARITY IS NOT A TYPE, AND WASM IS NOT UNTYPED HERE. This used to
    /// compare `param_count`/`result_arity` only — two `u8` counts — so
    /// `(func (result i32))` satisfied a call declared `(func (result i64))`:
    /// same 0→1 shape, different type, no trap. `Comptype_sub/func` compares
    /// the parameter and result TYPES, which is exactly what `test_concrete`
    /// already does for `ref.test`/`ref.cast` on a funcref. The counts stay as
    /// a cheap first test; the declared signature settles it.
    ///
    /// The expected signature is the DECLARED functype the wast walker emits
    /// as a fourth argument and the compiler files under this opcode's own
    /// offset (`Chunk::call_indirect_sigs`); the callee's is the structural
    /// one `__wast_register_func_sig` records on its chunk. Both are built by
    /// the same spelling normalisation, so they are comparable as written.
    ///
    /// ⛔ A MISSING SIGNATURE MUST NOT TRAP. `func_sig` is recorded for every
    /// function a module DEFINES; an imported or host function may have none,
    /// and trapping there would break every legitimate host call. `ref.test`
    /// answers `false` in that case because a failed TEST is harmless — a
    /// failed CALL is not, so this falls back to the count check instead.
    ///
    /// ⛔ ONE HELPER, BOTH SPELLINGS. `call_indirect` and
    /// `return_call_indirect` had drifted: only the former trapped on a null
    /// slot, and only the former spelled its out-of-bounds trap the way the
    /// spec does. Two near-identical arms is how that happens.
    fn indirect_call_type_check(
        &self,
        funcref: &Value,
        argc: usize,
        expected_results: usize,
        opcode_start: usize,
        tail: bool,
    ) -> Result<(), VMError> {
        let Value::Object(o) = funcref else { return Ok(()) };
        let ob = o.lock().unwrap();
        let crate::value::ObjectKind::Function(f) = &ob.kind else { return Ok(()) };
        let ch = &self.chunks[f.chunk_index];
        let how = if tail { " (return_call_indirect)" } else { "" };
        if ch.param_count as usize != argc || ch.result_arity as usize != expected_results {
            return Err(VMError::new(format!(
                "trap: indirect call type mismatch{} (callee {}→{}, expected {}→{})",
                how, ch.param_count, ch.result_arity, argc, expected_results
            )));
        }
        let ci = self.frame().chunk_index;
        // ⛔ IDENTITY FIRST, SHAPE SECOND. Iso-recursive equivalence is not
        // structural equality: `type-rec.wast` declares `$f1` and `$f2` both as
        // `(func)` in DIFFERENT rec groups, and the call must trap even though
        // the signatures are character-for-character equal. Only the
        // CANONICALISED name separates them, and it is canonical on both sides
        // — `qualify_type_name` at the call site, `declared_func_type` on the
        // callee — so the comparison is meaningful.
        //
        // Both must be present: a function declared with an inline signature
        // has no type name, and a name compared against nothing is a guess.
        //
        // ⛔ ONE MODULE ONLY. Canonicalisation is per-module here (names carry
        // an `m#<seq>#` prefix and the map is rebuilt for each module), while
        // the spec canonicalises across the whole store — so two modules that
        // declare the SAME type still get different names, and comparing them
        // would trap a call the spec says succeeds. Across that boundary the
        // structural check below is the answer; a name means nothing there.
        //
        // ⛔ AND A NAME MISMATCH IS NOT ENOUGH ON ITS OWN. Canonicalisation
        // merges by the composite's source TEXT, so it splits types that are
        // equal — `(param f32 f32)` from `(param $x f32) (param $y f32)`, and
        // `(ref $r1)` from `(ref $r2)` when `$r1` and `$r2` are themselves
        // equal. Trapping on the name alone therefore rejects valid programs.
        // The REC GROUP's size and the member's position come from the tree
        // rather than the text, and under iso-recursive equivalence a type is
        // identified by its whole group plus its position in it — so a
        // difference there is a real difference, and it is the only part of
        // the identity precise enough to trap on.
        let same_module = |a: &str, b: &str| {
            let seq =
                |n: &str| n.strip_prefix("m#").and_then(|r| r.split_once('#')).map(|(s, _)| s.to_string());
            matches!((seq(a), seq(b)), (Some(x), Some(y)) if x == y)
        };
        if let (Some(want), Some(got)) = (
            self.chunks[ci].call_indirect_canon.get(&opcode_start),
            &ch.declared_func_type,
        ) {
            let shape = |n: &String| self.chunks[0].type_rec_shape.get(n);
            let shapes_differ = match (shape(want), shape(got)) {
                (Some(a), Some(b)) => a != b,
                _ => false,
            };
            if same_module(want, got) && want != got && shapes_differ {
                return Err(VMError::new(format!(
                    "trap: indirect call type mismatch{how} (callee {got}, expected {want})"
                )));
            }
        }
        let declared = self.chunks[ci].call_indirect_sigs.get(&opcode_start);
        if let (Some(want), Some((params, results))) = (declared, &ch.func_sig) {
            // ⛔ ONLY TRUST A SIGNATURE THAT AGREES WITH THE COUNTS. The
            // declared string and the three count immediates are produced by
            // DIFFERENT readers of the same typeuse, and they do not always
            // read it the same way — a plain `call_indirect (type $t)` reaches
            // one of them wrapped in an `instr_arg` and comes back empty while
            // the counts are right. An empty `"->"` compared against a real
            // `->i32` callee then trapped a perfectly valid call.
            //
            // So the counts, which are already checked above and known good,
            // gate the string: if the two disagree about arity, this reader did
            // not see the same typeuse and abstains. That makes an over-fire
            // impossible — the check can only ever add traps the counts missed.
            let want_arity = (
                want.split("->").next().unwrap_or("").split(',').filter(|s| !s.is_empty()).count(),
                want.split("->").nth(1).unwrap_or("").split(',').filter(|s| !s.is_empty()).count(),
            );
            let got = format!("{params}->{results}");
            if want_arity == (argc, expected_results) && *want != got {
                return Err(VMError::new(format!(
                    "trap: indirect call type mismatch{how} (callee {got}, expected {want})"
                )));
            }
        }
        Ok(())
    }

    /// Pop a memory-op count/index operand, widening per the memory's index
    /// type: i64 for a 64-bit memory, i32 otherwise. Used by
    /// `memory.size/grow/copy/fill` — all standard opcodes; memory64 adds none.
    ///
    /// Unsigned at BOTH widths, for the same reason as `pop_table_count`: the
    /// doc here already said "unsigned i32" while the code clamped signed, so
    /// `memory.fill` at address `0xffff_ffff` wrote at 0 instead of trapping.
    fn pop_mem_index(&mut self, is64: bool) -> usize {
        if is64 {
            usize::try_from(self.pop().as_i64() as u64).unwrap_or(usize::MAX)
        } else {
            self.pop().as_i32() as u32 as usize
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
                "trap: out of bounds table access ({context}: negative index)"
            )));
        }
        usize::try_from(idx).map_err(|_| VMError::new(format!("trap: {context} index too large")))
    }

    pub(crate) fn execute_until(&mut self, min_depth: usize) -> Result<Value, VMError> {
        // Track this loop's floor so exception unwinding defers instead of
        // crossing it (see `raise_exception_value`). Pop on every exit.
        self.exec_floors.push(min_depth);
        let mut result = self.execute_until_inner(min_depth);
        // TRAP → HOST BOUNDARY. A trap leaves the dispatch loop as a plain
        // `Err`, but `frames`, `stack` and `exception_handlers` are VM fields,
        // so the machine state is still fully intact here — this is the one
        // place a trap can be offered to a host-level handler without touching
        // the interpreter loop at all. `raise_trap` unwinds to that handler if
        // one exists (and returns Ok), and we resume; otherwise the trap keeps
        // escaping exactly as before. Cold path only: zero cost per instruction.
        while let Err(e) = &result {
            if !e.is_trap() {
                break;
            }
            let message = e.message.clone();
            if self.raise_trap(&message).is_err() {
                break; // no host-level handler — the trap escapes
            }
            result = self.execute_until_inner(min_depth);
        }
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
            if self.dbg_ac && self.frames.is_empty() {
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
            // Trusted construction: every chunk was validated instruction-by-
            // instruction at LOAD (`VM::validate_chunk_code` — run_linked,
            // nested eval, and reload all gate on it), so the per-dispatch
            // `Op::decode` → `wasm_name_opt` name-table probe this replaces
            // re-proved a fact millions of times (a top-3 profile sample on a
            // pure-arithmetic loop). An op that somehow escaped validation
            // still lands in this match's final `Unhandled opcode` arm — an
            // error, never undefined behaviour.
            let op = Op::new(group as u16, sub as u16);
            if self.dbg_ac {
                dbg_last_op = Some(op);
            }

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
                Op::UNREACHABLE => {
                    return Err(VMError::new("trap: unreachable executed"));
                }
                Op::NOP => { /* no-op */ }

                Op::DROP => {
                    if self.stack.len() > self.stack_floor() {
                        self.pop();
                    }
                }

                // -- Variables --
                Op::LOCAL_GET => {
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
                Op::LOCAL_SET => {
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
                        // Locals exist (zero-initialized) from function entry
                        // (spec §4.4.9) — grow to the declared local frame, as
                        // LOCAL_TEE does. A slot still out of range after that
                        // is an out-of-bounds local index: trap rather than
                        // silently discarding the write (LOCAL_GET/LOCAL_TEE
                        // both trap, and a dropped store would corrupt the
                        // frame's value semantics invisibly).
                        let ci = self.frame().chunk_index;
                        let need = base + self.chunks[ci].local_count as usize;
                        if self.stack.len() < need {
                            self.stack.resize(need, Value::Null);
                        }
                        let dst = self
                            .stack
                            .get_mut(idx)
                            .ok_or_else(|| VMError::new("trap: local index out of bounds"))?;
                        *dst = val;
                    }
                }
                Op::LOCAL_TEE => {
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
                Op::GLOBAL_GET => {
                    // A globalidx over `global_imports ++ defined`, exactly as
                    // WASM's `global.get`. No name is consulted here.
                    let idx = self.read_u16() as usize;
                    // An UNASSIGNED slot reads exactly as it did when globals
                    // were a map and the key was simply absent: Undefined.
                    // That matters because emitted code tests "unset" two
                    // different ways — `ref.is_null` (which accepts Undefined)
                    // and `js-undefined:test` (which does NOT accept Null).
                    // Storing Null and returning it satisfies the first and
                    // breaks the second; storing Undefined breaks neither, but
                    // then `GlobalInit` cannot tell unset from a real
                    // undefined. `globals_assigned` separates the two
                    // questions, so the READ can answer this one correctly.
                    let val = if self.globals_assigned.get(idx).copied().unwrap_or(false) {
                        self.globals.get(idx).cloned().unwrap_or(Value::Undefined)
                    } else {
                        Value::Undefined
                    };
                    self.push(val)?;
                }
                Op::GLOBAL_SET => {
                    // See GLOBAL_GET: a globalidx, not a name.
                    let idx = self.read_u16() as usize;
                    let val = self.pop();
                    // Grow each vector against ITS OWN length. Testing only
                    // `globals.len()` assumed the two were already in step, so
                    // a path that populated one and not the other skipped the
                    // guard entirely and panicked on the second index instead
                    // of self-correcting.
                    if idx >= self.globals.len() {
                        self.globals.resize(idx + 1, Value::Null);
                    }
                    if idx >= self.globals_assigned.len() {
                        self.globals_assigned.resize(idx + 1, false);
                    }
                    self.globals[idx] = val;
                    self.globals_assigned[idx] = true;
                }

                // -- Properties --
                Op::STRUCT_GET => {
                    let typeidx = self.read_u16() as usize;
                    let idx = self.read_u16();
                    if typeidx != 0 {
                        // Spec `struct.get $t i` — indexed read of the
                        // instance's field storage, where a typed
                        // `struct.new` / `struct.new_default` put its values.
                        let obj = self.pop();
                        if obj.is_null_ref() {
                            return Err(VMError::new("trap: null structure reference (struct.get)"));
                        }
                        let val = match &obj {
                            Value::Object(o) => o
                                .lock()
                                .unwrap()
                                .fields
                                .get(idx as usize)
                                .cloned()
                                .ok_or_else(|| {
                                    VMError::new("trap: struct.get field index out of range")
                                })?,
                            _ => return Err(VMError::new("trap: struct.get on a non-struct")),
                        };
                        self.push(val)?;
                        continue;
                    }
                    let name = self.constant_str(idx);
                    let obj = self.pop();
                    // WASM GC `struct.get` traps on a null ref. Only a TYPED null
                    // (a GC reference) traps; a plain null — a dynamic-language
                    // `obj.field` on null — stays lenient (handled below).
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null structure reference (struct.get)"));
                    }
                    // Auto-join thread when accessing .result on a Task/Thread object
                    if let Value::Object(ref o) = obj {
                        let needs_join = {
                            let o_ref = o.lock().unwrap();
                            // ⛔ ONE MEMBER, AND IT IS THE PROTOCOL'S. This also
                            // fired on `exitcode` — `System.Diagnostics.Process`'s
                            // spelling — which no object reaching this guard can
                            // even have: `__thread_id` is written ONLY by
                            // `primitives/channels.rs`, and a `Process` built by
                            // the dotnet adapter carries neither. The arm named a
                            // framework member the VM cannot see.
                            name == "result"
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
                                self.memory.mark_parked();
                                let _ = handle.join();
                                self.memory.unmark_parked();
                                // Task object was updated by child thread
                            }
                        }
                        // §10.1.8.1 OrdinaryGet: an accessor is found by
                        // walking the PROTOTYPE CHAIN, not just own properties,
                        // and its getter is called with the ORIGINAL receiver
                        // as `this` (step 7: Call(getter, Receiver)).
                        //
                        // This used to consult own properties only, which made
                        // every inherited accessor unreachable — `Map.prototype.size`
                        // (§24.1.3.10), `RegExp.prototype.flags`, the `%TypedArray%`
                        // family. The host compensated by giving instances an own
                        // data property and the JS prelude re-wrapped constructors
                        // to install real getters; both are the shape that predates
                        // having an ECMA host at all.
                        let getter_key = format!("__get_{}", name);
                        let mut current = Some(obj.clone());
                        let mut getter = None;
                        // Bounded like `proto_chain_has`: a corrupt cyclic chain
                        // must not spin forever.
                        for _ in 0..1024 {
                            let Some(Value::Object(ref node)) = current else {
                                break;
                            };
                            let found = {
                                let n = node.lock().unwrap();
                                if let Some(g) = n.properties.get(&getter_key) {
                                    Some(g.clone())
                                } else {
                                    // A DATA property at this level shadows an
                                    // accessor further up — stop looking.
                                    if n.properties.contains_key(&name) {
                                        break;
                                    }
                                    None
                                }
                            };
                            if let Some(g) = found {
                                getter = Some(g);
                                break;
                            }
                            let next = node.lock().unwrap().properties.get("__proto__").cloned();
                            current = match next {
                                Some(Value::Object(p)) => Some(Value::Object(p)),
                                _ => None,
                            };
                        }
                        if let Some(getter_fn) = getter {
                            self.push(getter_fn)?;
                            self.push(obj)?;
                            self.call_value(1)?;
                            continue;
                        }
                    }
                    self.push(self.resolve_property(&obj, &name)?)?;
                }
                Op::STRUCT_SET => {
                    let typeidx = self.read_u16() as usize;
                    let idx = self.read_u16();
                    if typeidx != 0 {
                        // Spec `struct.set $t i`. This form did not exist —
                        // `struct.set` was name-keyed only, so nothing could
                        // write the indexed storage that `struct.get_s`/`_u`
                        // read.
                        let val = self.pop();
                        let obj = self.pop();
                        if obj.is_null_ref() {
                            return Err(VMError::new("trap: null structure reference (struct.set)"));
                        }
                        match &obj {
                            Value::Object(o) => {
                                let mut o = o.lock().unwrap();
                                // ⚠ TRANSITIONAL TOLERANCE — grow, do not trap.
                                //
                                // In real WASM a struct's arity is fixed at
                                // allocation and the type system makes a
                                // short-vec write unrepresentable. Here the
                                // class model is mid-conversion: a class can
                                // have a registered type (so its FIELD ACCESS
                                // is indexed) while some of its instances are
                                // still allocated dynamically by a language
                                // emitter's own path (so their storage is
                                // empty). python's `@dataclass` and
                                // `__slots__` classes are exactly that shape.
                                //
                                // The READ side already tolerates this —
                                // `calls.rs::resolve_property` step 1b falls
                                // through to the property bag rather than
                                // failing — so trapping on the write was an
                                // ASYMMETRY, not a spec guarantee.
                                //
                                // ⛔ REMOVE THIS once every allocation of a
                                // registered type goes through
                                // `struct.new_default <typeidx>`; at that point
                                // a short vec is a real bug and should trap
                                // again.
                                let i = idx as usize;
                                if i >= o.fields.len() {
                                    o.fields.resize(i + 1, Value::Null);
                                }
                                o.fields[i] = val;
                            }
                            _ => return Err(VMError::new("trap: struct.set on a non-struct")),
                        }
                        continue;
                    }
                    let name = self.constant_str(idx);
                    let val = self.pop();
                    let obj = self.pop();
                    // WASM GC `struct.set` traps on a typed null (GC ref); a plain
                    // null (dynamic-language write) stays lenient.
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null structure reference (struct.set)"));
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
                            // ⛔ ONE RECEIVER, PASSED ONCE. Passing the pair
                            // `[obj, val]` to `invoke_callback` is a receiver
                            // in the ARGUMENT LIST, and `invoke_callback`
                            // ALSO prepends one for any callee whose chunk
                            // declares `takes_receiver` — which every
                            // receiver-first accessor does, in every language,
                            // not only under `ReceiverAbi::Parameter`. The
                            // setter then ran with `this` = the prepended
                            // receiver and its VALUE parameter = the receiver
                            // again. Measured on dart: `c.count = 5` reached
                            // the setter as `v = [object Counter]` and the
                            // backing field kept its initial value, across 36
                            // getter/setter tests.
                            let _result =
                                self.invoke_with_receiver(&setter_fn, obj.clone(), &[val.clone()]);
                            self.stack.truncate(stack_save);
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
                        }
                    }
                    // Spec `struct.set`: pops [obj, val], pushes NOTHING. The
                    // name-keyed addressing is this VM's dynamic-object
                    // extension; the stack contract is not. (It used to push
                    // `val` back on every path — the reason ~500 emit sites
                    // carried a compensating DROP.)
                }
                Op::ARRAY_GET => {
                    let key = self.pop();
                    let obj = self.pop();
                    // WASM GC `array.get` traps on a typed null (GC array ref);
                    // a plain null (dynamic subscript) stays lenient.
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null array reference (array.get)"));
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
                                                "trap: out of bounds array access (array.get)",
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
                Op::ARRAY_SET => {
                    let val = self.pop();
                    let key = self.pop();
                    let obj = self.pop();
                    // WASM GC `array.set` traps on a typed null (GC array ref);
                    // a plain null (dynamic subscript) stays lenient.
                    if matches!(obj, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null array reference (array.set)"));
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
                                        continue;
                                    }
                                    _ => {
                                        return Err(VMError::new("trap: out of bounds array access (array.set)"));
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
                                continue;
                            }
                        }
                        let k = format!("{}", key);
                        o.lock().unwrap().set(k, val.clone());
                    }
                    // Spec `array.set`: pops [array, index, value], pushes
                    // NOTHING — same contract flip as name-keyed `struct.set`.
                }

                // -- F32 arithmetic (f32 precision, stored as F64) --
                Op::F32_ADD => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a + b))?;
                }
                Op::F32_SUB => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a - b))?;
                }
                Op::F32_MUL => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a * b))?;
                }
                Op::F32_DIV => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a / b))?;
                }
                // -- Float arithmetic --
                Op::F64_ADD => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a + b))?;
                }
                Op::F64_SUB => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a - b))?;
                }
                Op::F64_MUL => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a * b))?;
                }
                Op::F64_DIV => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a / b))?;
                }
                // f64_mod: removed (non-WASM, use __stdlib_fmod)
                Op::F64_NEG => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(-a))?;
                }

                // -- Integer arithmetic --
                Op::I32_ADD => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_add(b)))?;
                }
                Op::I32_SUB => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_sub(b)))?;
                }
                Op::I32_MUL => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(Value::I32(a.wrapping_mul(b)))?;
                }
                Op::I32_DIV_S => {
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
                Op::I32_DIV_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I32((a / b) as i32))?;
                }
                Op::I32_REM_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I32(a.wrapping_rem(b)))?;
                }
                Op::I32_REM_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I32((a % b) as i32))?;
                }

                // -- i64 arithmetic --
                Op::I64_ADD => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.wrapping_add(b)))?;
                }
                Op::I64_SUB => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.wrapping_sub(b)))?;
                }
                Op::I64_MUL => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.wrapping_mul(b)))?;
                }
                Op::I64_DIV_S => {
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
                Op::I64_DIV_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I64((a / b) as i64))?;
                }
                Op::I64_REM_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I64(a.wrapping_rem(b)))?;
                }
                Op::I64_REM_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    if b == 0 {
                        return Err(VMError::new("trap: integer divide by zero"));
                    }
                    self.push(Value::I64((a % b) as i64))?;
                }
                Op::I64_AND => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a & b))?;
                }
                Op::I64_OR => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a | b))?;
                }
                Op::I64_XOR => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a ^ b))?;
                }
                Op::I64_SHL => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a << (b & 0x3f)))?;
                }
                Op::I64_SHR_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a >> (b & 0x3f)))?;
                }
                Op::I64_SHR_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::I64((a >> (b & 0x3f)) as i64))?;
                }
                Op::I64_ROTL => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::I64(a.rotate_left((b & 0x3f) as u32) as i64))?;
                }
                Op::I64_ROTR => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::I64(a.rotate_right((b & 0x3f) as u32) as i64))?;
                }
                Op::I64_CLZ => {
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.leading_zeros() as i64))?;
                }
                Op::I64_CTZ => {
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.trailing_zeros() as i64))?;
                }
                Op::I64_POPCNT => {
                    let a = self.pop().as_i64();
                    self.push(Value::I64(a.count_ones() as i64))?;
                }

                // -- f64 math --
                Op::F64_ABS => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.abs()))?;
                }
                Op::F64_CEIL => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.ceil()))?;
                }
                Op::F64_FLOOR => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.floor()))?;
                }
                Op::F64_TRUNC => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.trunc()))?;
                }
                Op::F64_NEAREST => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.round_ties_even()))?;
                }
                Op::F64_SQRT => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.sqrt()))?;
                }
                Op::F64_MIN => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(if a.is_nan() || b.is_nan() {
                        f64::NAN
                    } else {
                        a.min(b)
                    }))?;
                }
                Op::F64_MAX => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(if a.is_nan() || b.is_nan() {
                        f64::NAN
                    } else {
                        a.max(b)
                    }))?;
                }
                Op::F64_COPYSIGN => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a.copysign(b)))?;
                }

                // -- f32 (promoted to f64) --
                Op::F32_ABS => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.abs()))?;
                }
                Op::F32_NEG => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(-a))?;
                }
                Op::F32_CEIL => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.ceil()))?;
                }
                Op::F32_FLOOR => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.floor()))?;
                }
                Op::F32_TRUNC => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.trunc()))?;
                }
                Op::F32_NEAREST => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.round_ties_even()))?;
                }
                Op::F32_SQRT => {
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.sqrt()))?;
                }
                Op::F32_MIN => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(if a.is_nan() || b.is_nan() {
                        f32::NAN
                    } else {
                        a.min(b)
                    }))?;
                }
                Op::F32_MAX => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(if a.is_nan() || b.is_nan() {
                        f32::NAN
                    } else {
                        a.max(b)
                    }))?;
                }
                Op::F32_COPYSIGN => {
                    let b = self.pop().as_f32();
                    let a = self.pop().as_f32();
                    self.push(Value::F32(a.copysign(b)))?;
                }

                // -- WASM select --
                Op::SELECT => {
                    let cond = self.pop().as_i32();
                    let val2 = self.pop();
                    let val1 = self.pop();
                    self.push(if cond != 0 { val1 } else { val2 })?;
                }
                // Typed select (`select t`): same runtime semantics as
                // untyped select; the result-type vec is a validation-time
                // hint. The emitter writes `0x1C <count> <valtype>*`; VM
                // side just pops and picks.
                Op::SELECT_T => {
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
                Op::TABLE_GET => {
                    let table_idx = self.read_u16() as usize;
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
                        .ok_or_else(|| VMError::new("trap: out of bounds table access (table.get)"))?;
                    self.push(val)?;
                }
                // `table.set tbl` — pop value + index, write into table.
                // Trap on out-of-bounds index per spec.
                Op::TABLE_SET => {
                    let table_idx = self.read_u16() as usize;
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
                        return Err(VMError::new("trap: out of bounds table access (table.set)"));
                    }
                    table[idx] = val;
                }

                // -- i32 rotation and bit counting --
                Op::I32_ROTL => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.rotate_left(b & 0x1f) as i32))?;
                }
                Op::I32_ROTR => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.rotate_right(b & 0x1f) as i32))?;
                }
                Op::I32_CLZ => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.leading_zeros() as i32))?;
                }
                Op::I32_CTZ => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.trailing_zeros() as i32))?;
                }
                Op::I32_POPCNT => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I32(a.count_ones() as i32))?;
                }

                // -- eqz --
                Op::I32_EQZ => {
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a == 0))?;
                }
                Op::I64_EQZ => {
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a == 0))?;
                }

                // -- String --

                // -- Bitwise --
                Op::I32_AND => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a & b))?;
                }
                Op::I32_OR => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a | b))?;
                }
                Op::I32_XOR => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a ^ b))?;
                }
                // i32_not: removed (non-WASM, use i32.const -1 + i32.xor)
                Op::I32_SHL => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a.wrapping_shl((b as u32) & 0x1f)))?;
                }
                Op::I32_SHR_S => {
                    let b = self.pop().to_ecma_int32();
                    let a = self.pop().to_ecma_int32();
                    self.push(Value::I32(a >> (b & 0x1f)))?;
                }
                Op::I32_SHR_U => {
                    let b = self.pop().to_ecma_int32() as u32;
                    let a = self.pop().to_ecma_int32() as u32;
                    self.push(Value::I32((a >> (b & 0x1f)) as i32))?;
                }

                // -- Comparison --
                // i32 comparisons (WASM MVP 0x46–0x4F)
                Op::I32_EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(wasm_bool(a.eq(&b)))?;
                }
                Op::I32_NE => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(wasm_bool(!a.eq(&b)))?;
                }
                Op::I32_LT_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a < b))?;
                }
                Op::I32_LT_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a < b))?;
                }
                Op::I32_GT_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a > b))?;
                }
                Op::I32_GT_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a > b))?;
                }
                Op::I32_LE_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a <= b))?;
                }
                Op::I32_LE_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a <= b))?;
                }
                Op::I32_GE_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    self.push(wasm_bool(a >= b))?;
                }
                Op::I32_GE_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    self.push(wasm_bool(a >= b))?;
                }
                // i64 comparisons (WASM MVP 0x51–0x5A)
                Op::I64_EQ => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a == b))?;
                }
                Op::I64_NE => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a != b))?;
                }
                Op::I64_LT_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a < b))?;
                }
                Op::I64_LT_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a < b))?;
                }
                Op::I64_GT_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a > b))?;
                }
                Op::I64_GT_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a > b))?;
                }
                Op::I64_LE_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a <= b))?;
                }
                Op::I64_LE_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a <= b))?;
                }
                Op::I64_GE_S => {
                    let b = self.pop().as_i64();
                    let a = self.pop().as_i64();
                    self.push(wasm_bool(a >= b))?;
                }
                Op::I64_GE_U => {
                    let b = self.pop().as_i64() as u64;
                    let a = self.pop().as_i64() as u64;
                    self.push(wasm_bool(a >= b))?;
                }
                // f32 comparisons (WASM MVP 0x5B–0x60) — operate on f32 precision
                Op::F32_EQ => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a == b))?;
                }
                Op::F32_NE => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a != b))?;
                }
                Op::F32_LT => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a < b))?;
                }
                Op::F32_GT => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a > b))?;
                }
                Op::F32_LE => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a <= b))?;
                }
                Op::F32_GE => {
                    let b = self.pop().as_f64() as f32;
                    let a = self.pop().as_f64() as f32;
                    self.push(wasm_bool(a >= b))?;
                }
                // f64 comparisons (WASM MVP 0x61–0x66)
                Op::F64_EQ => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a == b))?;
                }
                Op::F64_NE => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a != b))?;
                }
                Op::F64_LT => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a < b))?;
                }
                Op::F64_GT => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a > b))?;
                }
                Op::F64_LE => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a <= b))?;
                }
                Op::F64_GE => {
                    let b = self.pop().as_f64();
                    let a = self.pop().as_f64();
                    self.push(wasm_bool(a >= b))?;
                }
                // str_lt, str_gt: removed (non-WASM, were unused)

                // -- Logical --
                // bool_not: removed (non-WASM, use dyn_to_bool + i32_eqz)

                // -- Control flow --
                Op::BR => {
                    let ci = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let depth = read_leb_u32(&self.chunks[ci].code, &mut ip) as usize;
                    self.frame_mut().ip = ip;
                    if let Some(entry) = self.label_stack.iter().rev().nth(depth).copied() {
                        self.branch_to_label(depth, entry);
                    }
                }
                Op::BR_IF => {
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
                // The old callee-on-stack `Op::CALL` arm (byte-identical to
                // CALL_REF) is deleted; spec `call` (0x00 0x10) is being
                // redefined as a static import call — see callimportretirement.md.
                Op::CALL_REF => {
                    // Direct call through a function reference — same as call
                    // but the func ref is already on the stack (no table lookup).
                    let argc = self.read_byte() as usize;
                    // Result count: carried for the writer's exact functype
                    // annotation; the callee chunk's own result_arity drives
                    // execution, so it is not read here.
                    let _results = self.read_byte();
                    self.call_value(argc)?;
                }
                // ── Call Tags proposal ───────────────────────────────────
                //
                // `call_with_tag $call_tag : [ti* funcref] -> [to*]`
                //
                // The decision is the CALLEE's, which is what separates this
                // from `call_indirect`: the caller names a tag, and the funcref
                // either handles it or does not. The Overview:
                //
                //   If the call tag is not recognized, then the code
                //   jumps-to/tail-calls the fall-back handler pointed to by the
                //   call tag, leaving all the arguments in their place but
                //   replacing the call-tag value with the value of the current
                //   `funcref`.
                //
                // and, for a tag with no handler declared:
                //
                //   For canonical call tags, the answer is simply that the
                //   program traps.
                //
                // Trapping is the point. An unhandled convention is a MISTAKE,
                // and the alternative — calling anyway under the wrong shape —
                // is a silent wrong answer.
                Op::CALL_WITH_TAG => {
                    // The immediate NAMES the tag (a constant-pool index); the
                    // entity id is resolved from it, as a `throw`'s tagidx is.
                    let name_idx = self.read_u16();
                    let argc = self.read_byte() as usize;
                    let ci = self.frame().chunk_index;
                    let tag = self.resolve_chunk_call_tag(ci, name_idx)?;
                    self.call_with_tag(tag, argc)?;
                }
                // `call_indirect_with_tag $table $call_tag : [ti* i32] -> [to*]`
                // is shorthand for `(call_with_tag $call_tag (table.get $table))`,
                // so it resolves the element and then runs the identical path —
                // one implementation, no second set of rules to drift.
                Op::CALL_INDIRECT_WITH_TAG => {
                    let table = self.read_u16() as usize;
                    let local = self.read_u16();
                    let argc = self.read_byte() as usize;
                    let ci = self.frame().chunk_index;
                    let tag = self.resolve_chunk_call_tag(ci, local)?;
                    let elem_idx = match self.stack.pop() {
                        Some(Value::I32(i)) => i as usize,
                        // wast integers reach the stack as f64 through this
                        // pipeline, exactly as they do for plain
                        // `call_indirect`; a whole f64 IS the i32 index.
                        Some(Value::F64(f)) if f.fract() == 0.0 && f >= 0.0 => f as usize,
                        Some(other) => {
                            return Err(VMError::new(format!(
                                "call_indirect_with_tag: table index must be i32, got {other:?}"
                            )));
                        }
                        None => {
                            return Err(VMError::new(
                                "call_indirect_with_tag: missing table index".to_string(),
                            ));
                        }
                    };
                    // `table.get $table` — out of bounds traps, as it does for
                    // the plain `call_indirect` this is shorthand for.
                    let funcref = self
                        .wasm_tables
                        .get(table)
                        .and_then(|t| t.get(elem_idx))
                        .cloned()
                        .ok_or_else(|| {
                            VMError::new(format!(
                                "call_indirect_with_tag: undefined element {elem_idx} in table {table}"
                            ))
                        })?;
                    self.stack.push(funcref);
                    self.call_with_tag(tag, argc)?;
                }
                // The tail-call form. The Overview defers the tail behaviour to
                // the Tail Call proposal ("for engines that support
                // `call_return`"); the tag semantics are identical, so it shares
                // the same resolution and differs only in not growing the frame
                // where the engine can avoid it.
                Op::CALL_RETURN_WITH_TAG => {
                    let name_idx = self.read_u16();
                    let argc = self.read_byte() as usize;
                    let ci = self.frame().chunk_index;
                    let tag = self.resolve_chunk_call_tag(ci, name_idx)?;
                    // "…where `[to*]` is a SUBTYPE OF THE RESULT TYPE OF THE
                    // FUNCTION CONTAINING THIS INSTRUCTION" — Design
                    // §Instructions. A tail call hands the callee's results
                    // straight to THIS function's caller, so they have to be
                    // results this function was allowed to return. Nothing
                    // checked it: a func declaring no results tail-called a tag
                    // yielding `[i32]` and the module was accepted.
                    //
                    // ⚠ HONEST LIMIT: runtime types are erased here — a
                    // `CallTagDef`'s signature IS its arity (`params`/`results`,
                    // and the struct says so) — so this enforces the ARITY half
                    // of the subtype relation, not element-wise subtyping. It
                    // catches the shape the proposal names and would not catch
                    // `[i32]` against `[i64]`. A full check belongs at load,
                    // with declared types, and is a bigger change than the
                    // instruction it would guard.
                    if let Some(tag_def) = self.call_tags.get(tag as usize) {
                        let declared = self.chunks[ci].result_arity;
                        if tag_def.results != declared {
                            return Err(VMError::new(format!(
                                "call_return_with_tag: tag '{}' returns {} result(s), \
                                 but the containing function declares {declared} — \
                                 [to*] must be a subtype of the function's result type",
                                tag_def.debug_name, tag_def.results
                            )));
                        }
                    }
                    // ⛔ THE "RETURN" HALF WAS NOT IMPLEMENTED. This arm was
                    // byte-identical to `CALL_WITH_TAG` — it dispatched and kept
                    // the frame, so the instruction the proposal defines as
                    // "TAIL calls the given funcref" accumulated one frame per
                    // call. 200_000 self-calls exhausted the stack; the limit
                    // sat between 5_000 and 20_000. A genuine tail call has no
                    // limit, which is what makes the depth the assertion.
                    //
                    // Same shape as `RETURN_CALL_REF` above: relocate
                    // `[args… funcref]` down to this frame's base, drop the
                    // frame — `pop_frame_for_tail_call` also disarms its
                    // `try_table` handlers, which a bare `frames.pop()` does not
                    // — and only then dispatch. Dispatch stays `call_with_tag`
                    // so the tag check, `func_switch` resolution and the
                    // fall-back path are the ones already proven.
                    let old_base = self.frame().base;
                    let operand_idx = self.stack.len() - argc - 1;
                    for i in 0..=argc {
                        self.stack[old_base + i] = self.stack[operand_idx + i].clone();
                    }
                    self.stack.truncate(old_base + 1 + argc);
                    self.pop_frame_for_tail_call();
                    self.call_with_tag(tag, argc)?;
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
                Op::RETURN => {
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
                        if self.dbg_ac {
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
                Op::REF_FUNC => {
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
                        properties: indexmap::IndexMap::new(),
                        kind: ObjectKind::Function(func),
                        type_id: 0,
                        fields: Vec::new(),
                    };
                    // Add to function table for call_indirect
                    let table_idx = self.func_table.len();
                    obj.properties
                        .insert("__table_idx".into(), Value::F64(table_idx as f64));
                    let func_val = Value::Object(crate::heap::alloc(obj));
                    self.func_table.push(func_val.clone());
                    // Intern the canonical capture-free funcref for reuse.
                    if uv_count == 0 {
                        self.funcref_cache.insert(func_idx, func_val.clone());
                    }
                    self.push(func_val)?;
                }

                // -- Host functions --
                // Spec `call` (0x00 0x10): u16 chunk-scoped import index +
                // VM-internal u8 argc. Resolution goes through the frame
                // chunk's import table, falling back to the linked module
                // table. (The retired 0xFF CALL_IMPORT alias carried the
                // identical immediates and body.)
                Op::CALL => {
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

                            if self.dbg_ac {
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
                            // back to the embedder — like the end-of-code
                            // top-frame path, but without a process exit.
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
                            let func_val = Value::Object(crate::heap::alloc(obj));
                            let args_start = self.stack.len() - argc;
                            self.stack.insert(args_start, func_val);
                            self.call_value(argc)?;
                        }
                        ImportTarget::StdlibRedirect(ref global_name) => {
                            if let Some(func_val) = self.global(global_name).cloned() {
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
                        ImportTarget::JspiSuspend | ImportTarget::JspiSuspendEager => {
                            let val = if argc == 0 {
                                Value::Undefined
                            } else {
                                self.pop()
                            };
                            for _ in 1..argc {
                                self.pop();
                            }
                            let eager = matches!(target, ImportTarget::JspiSuspendEager);
                            self.do_await(val, eager)?;
                        }
                        ImportTarget::JspiYield => {
                            for _ in 0..argc {
                                self.pop();
                            }
                            // One full ready-queue turn: whole-fiber save +
                            // back-of-queue requeue (never synchronous).
                            return Err(self.tick_top_level_await(Value::Null, false));
                        }
                        ImportTarget::WasiThreadSpawn => {
                            // wasi-threads `thread-spawn(start_arg) -> tid`.
                            for _ in 1..argc {
                                self.pop();
                            }
                            let start_arg = self.pop().as_i32();
                            let tid = self.wasi_thread_spawn(start_arg)?;
                            self.push(Value::I32(tid))?;
                        }
                        ImportTarget::StringConst(ref s) => {
                            for _ in 0..argc {
                                self.pop();
                            }
                            self.push(Value::String(s.clone()))?;
                        }
                        ImportTarget::Canon(b, type_idx) => {
                            // CM canonical built-in: args/results ride the
                            // stack; the builtin pops exactly its own args
                            // (the emitter's argc matches by construction).
                            // `type_idx` is the `$t` immediate the `canon`
                            // definition carried — what tells a typed copy how
                            // wide one element is.
                            self.canon_type_immediate = type_idx;
                            self.exec_canon_builtin(b)?;
                        }
                    }
                }

                // -- Object/Array --
                // `struct.new` — two forms behind one opcode, discriminated by
                // the typeidx exactly like `array.new_fixed`:
                //
                //   typeidx == 0  dynamic object literal. `count` key/value
                //                 pairs on the stack become named properties.
                //                 This is what every language front end emits.
                //   typeidx != 0  spec `struct.new $t`. The field count comes
                //                 from $t (its `field_defs`), NOT from an
                //                 immediate, so the values land in indexed
                //                 storage and the instance is stamped with
                //                 $t's rtt — which is what makes `ref.test` /
                //                 `ref.cast` answer from the type registry
                //                 instead of a `__type` string.
                Op::STRUCT_NEW => {
                    let typeidx = self.read_u16() as usize;
                    let count = self.read_u16() as usize;
                    if typeidx != 0 {
                        let type_id = self.resolve_gc_rtt(typeidx);
                        let arity = self
                            .type_registry
                            .get(type_id)
                            .map_or(0, |td| td.field_defs.len())
                            .min(self.stack.len());
                        let start = self.stack.len() - arity;
                        let fields: Vec<Value> = self.stack[start..].to_vec();
                        self.stack.truncate(start);
                        let mut obj = Object::new();
                        obj.fields = fields;
                        obj.type_id = type_id;
                        self.push(Value::Object(crate::heap::alloc(obj)))?;
                    } else {
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
                        self.push(Value::Object(crate::heap::alloc(obj)))?;
                    }
                }
                // `array.new_fixed $t N` — pops N values off the stack
                // and allocates an N-element array initialised from them.
                // `array.new_fixed $t N` — [v1..vN] -> [array].
                //
                // BOTH immediates, matching `array.new` / `array.new_default`:
                // the type index stamps the rtt, and that stamp is what makes
                // the array bounds-check per spec. This used to read only `N`,
                // so a fixed array could never be a GC array and `array.get` /
                // `set` / `fill` / `copy` could never trap on one — the trap
                // code existed but was unreachable for anything built here.
                // Type index 0 = dynamic-language array literal (lenient).
                Op::ARRAY_NEW_FIXED => {
                    let typeidx = self.read_u16() as usize;
                    let count = self.read_u16() as usize;
                    let count = count.min(self.stack.len());
                    let start = self.stack.len() - count;
                    let elems: Vec<Value> = self.stack[start..].to_vec();
                    self.stack.truncate(start);
                    let mut obj = Object::new_array(elems);
                    obj.type_id = self.resolve_gc_rtt(typeidx);
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
                }
                // `array.new $t` — [value, length] -> [array of length,
                // every lane = value].
                Op::ARRAY_NEW => {
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
                    obj.type_id = self.resolve_gc_rtt(typeidx);
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
                }
                // `array.new_default $t` — [length] -> [array of length,
                // zero-initialised]. We use `Value::Null` as the default
                // for externref lanes (the only lane type we actually
                // support) per the "null is the default for refs" rule.
                Op::ARRAY_NEW_DEFAULT => {
                    // Immediate is a 1-based script-chunk type-table index (0 =
                    // dynamic); resolved to the instance rtt so a defaulted GC
                    // array traps per spec, matching `array.new`.
                    let typeidx = self.read_u16() as usize;
                    let len = self.pop().as_i32().max(0) as usize;
                    let elems = vec![Value::Null; len];
                    let mut obj = Object::new_array(elems);
                    obj.type_id = self.resolve_gc_rtt(typeidx);
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
                }
                // `array.new_data $t $d` — allocate a new array of `size`
                // ELEMENTS read from a data segment at a byte offset.
                Op::ARRAY_NEW_DATA => {
                    let typeidx = self.read_u16();
                    let dataidx = self.read_u16() as u32;
                    // ⚠ A DROPPED segment is an EMPTY one, not an error of
                    // its own. The spec drops the payload and leaves the
                    // segment in place, so the bounds check is what decides:
                    // a zero-length copy off a dropped segment SUCCEEDS, and
                    // only a non-zero one traps. Returning early here made
                    // `(array_init_data 0 0 0)` after `drop_segs` trap, which
                    // the fixture asserts must return. `MEMORY_INIT` already
                    // modelled it this way; these did not.
                    let dropped = self.dropped_data.contains(&dataidx);
                    let size = self.pop_u32_operand();
                    let offset = self.pop_u32_operand();
                    // ⛔ THE COUNT IS IN ELEMENTS, THE SEGMENT IS IN BYTES.
                    // The value model stores i32/f32/f64 all as f64, so the
                    // width cannot come from the value — it comes from the
                    // array type's element storage type, which the typeidx
                    // immediate names. Treating the count as bytes made
                    // `(array i32)` read one byte per element: the bound
                    // passed where 4 bytes were needed, and a segment of 1
                    // byte satisfied a request the spec traps.
                    //
                    // An unregistered type keeps the byte-wide reading; a
                    // width invented for it would be a guess.
                    let type_id = self.resolve_gc_rtt(typeidx as usize);
                    let (elem_size, kind) = self
                        .type_registry
                        .get(type_id)
                        .and_then(|td| td.field_defs.first())
                        .and_then(|f| array_elem_storage_kind(&f.name))
                        .unwrap_or((1, 4));
                    let data = self
                        .data_segments
                        .get(dataidx as usize)
                        .ok_or_else(|| VMError::new("trap: array.new_data: missing data segment"))?;
                    let seg_len = if dropped { 0 } else { data.len() };
                    let end = offset.saturating_add(size.saturating_mul(elem_size));
                    if end > seg_len {
                        return Err(VMError::new("trap: out of bounds memory access (array.new_data)"));
                    }
                    let elems: Vec<Value> = (0..size)
                        .map(|i| {
                            let base = offset + i * elem_size;
                            decode_le_numeric(kind, &data[base..base + elem_size])
                        })
                        .collect();
                    let mut obj = Object::new_array(elems);
                    obj.type_id = type_id;
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
                }
                Op::ARRAY_NEW_ELEM => {
                    let _typeidx = self.read_u16();
                    let elemidx = self.read_u16() as u32;
                    // ⚠ A DROPPED segment is an EMPTY one, not an error of
                    // its own. The spec drops the payload and leaves the
                    // segment in place, so the bounds check is what decides:
                    // a zero-length copy off a dropped segment SUCCEEDS, and
                    // only a non-zero one traps. Returning early here made
                    // `(array_init_data 0 0 0)` after `drop_segs` trap, which
                    // the fixture asserts must return. `MEMORY_INIT` already
                    // modelled it this way; these did not.
                    let dropped = self.dropped_elems.contains(&elemidx);
                    let size = self.pop_u32_operand();
                    let offset = self.pop_u32_operand();
                    let elems = self
                        .elem_segments
                        .get(elemidx as usize)
                        .ok_or_else(|| VMError::new("trap: array.new_elem: missing element segment"))?;
                    let seg_len = if dropped { 0 } else { elems.len() };
                    let end = offset.saturating_add(size);
                    if end > seg_len {
                        return Err(VMError::new("trap: out of bounds table access (array.new_elem)"));
                    }
                    self.push(Value::Object(crate::heap::alloc(Object::new_array(
                        elems[offset..end].to_vec(),
                    ))))?;
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
                Op::ARRAY_GET_S | Op::ARRAY_GET_U => {
                    let _typeidx = self.read_u16();
                    let is_signed = op == Op::ARRAY_GET_S;
                    let raw_idx = self.pop().as_i32();
                    let arr = self.pop();
                    // Same trap discipline as `array.get`: a GC array reference
                    // is bounds-checked and a typed null traps, while dynamic
                    // (JS-shaped) arrays stay lenient. These two used to clamp a
                    // negative index to 0 and answer 0 past the end, so a
                    // packed-field read could never fail — unlike `array.get`
                    // directly above, which has always trapped.
                    if matches!(arr, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null array reference (array.get)"));
                    }
                    let gc_len = match &arr {
                        Value::Object(o) if self.is_gc_array_obj(o) => {
                            match &o.lock().unwrap().kind {
                                ObjectKind::Array(a) => Some(a.len()),
                                _ => None,
                            }
                        }
                        _ => None,
                    };
                    if let Some(len) = gc_len {
                        if raw_idx < 0 || raw_idx as usize >= len {
                            return Err(VMError::new("trap: out of bounds array access (array.get)"));
                        }
                    }
                    let idx = raw_idx.max(0) as usize;
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
                Op::ARRAY_INIT_DATA => {
                    let _typeidx = self.read_u16();
                    let dataidx = self.read_u16() as u32;
                    // A dropped segment is an EMPTY one — see ARRAY_NEW_DATA.
                    let dropped = self.dropped_data.contains(&dataidx);
                    let size = self.pop_u32_operand();
                    let src_offset = self.pop_u32_operand();
                    let dst_offset = self.pop_u32_operand();
                    let array = self.pop();
                    let data = self
                        .data_segments
                        .get(dataidx as usize)
                        .ok_or_else(|| VMError::new("trap: array.init_data: missing data segment"))?
                        .clone();
                    let seg_len = if dropped { 0 } else { data.len() };
                    let check_src = |elem_size: usize| -> Result<(), VMError> {
                        let end = src_offset.saturating_add(size.saturating_mul(elem_size));
                        if end > seg_len {
                            return Err(VMError::new("trap: out of bounds memory access (array.init_data source)"));
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
                                // ⚠ DESTINATION FIRST. When both ends
                                // overrun, the fixtures name the ARRAY:
                                // `array_init_data (0,0,13)` into a 12-element
                                // array off a 12-byte segment expects "out of
                                // bounds array access", not the memory one.
                                let dst_end = dst_offset.saturating_add(size);
                                if dst_end > elems.len() {
                                    return Err(VMError::new(
                                        "trap: out of bounds array access (array.init_data destination)",
                                    ));
                                }
                                check_src(elem_size)?;
                                for i in 0..size {
                                    let base = src_offset + i * elem_size;
                                    elems[dst_offset + i] =
                                        decode_le_numeric(kind, &data[base..base + elem_size]);
                                }
                            }
                            ObjectKind::TypedArray(ta) => {
                                let elem_size = ta.elem.bytes_per_element();
                                let dst_end = dst_offset.saturating_add(size);
                                if dst_end > typed_array_live_length(ta) {
                                    return Err(VMError::new(
                                        "trap: out of bounds array access (array.init_data destination)",
                                    ));
                                }
                                check_src(elem_size)?;
                                for i in 0..size {
                                    let base = src_offset + i * elem_size;
                                    let v = decode_typed_le(ta.elem, &data[base..base + elem_size]);
                                    typed_array_write(ta, dst_offset + i, &v);
                                }
                            }
                            _ => return Err(VMError::new("trap: null array reference (array.init_data)")),
                        }
                    } else {
                        return Err(VMError::new("trap: null array reference (array.init_data)"));
                    }
                }
                Op::ARRAY_INIT_ELEM => {
                    let _typeidx = self.read_u16();
                    let elemidx = self.read_u16() as u32;
                    // A dropped segment is an EMPTY one — see ARRAY_NEW_DATA.
                    let dropped = self.dropped_elems.contains(&elemidx);
                    let size = self.pop_u32_operand();
                    let src_offset = self.pop_u32_operand();
                    let dst_offset = self.pop_u32_operand();
                    let array = self.pop();
                    let source = self
                        .elem_segments
                        .get(elemidx as usize)
                        .ok_or_else(|| VMError::new("trap: array.init_elem: missing element segment"))?;
                    let seg_len = if dropped { 0 } else { source.len() };
                    let src_end = src_offset.saturating_add(size);
                    // ⚠ The DESTINATION check is inside the object match below
                    // and runs FIRST when both ends overrun — see ARRAY_INIT_DATA.
                    let src_oob = src_end > seg_len;
                    if let Value::Object(obj) = array {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(elems) = &mut o.kind {
                            let dst_end = dst_offset.saturating_add(size);
                            if dst_end > elems.len() {
                                return Err(VMError::new(
                                    "trap: out of bounds array access (array.init_elem destination)",
                                ));
                            }
                            if src_oob {
                                return Err(VMError::new(
                                    "trap: out of bounds table access (array.init_elem source)",
                                ));
                            }
                            elems[dst_offset..dst_end]
                                .clone_from_slice(&source[src_offset..src_end]);
                        } else {
                            return Err(VMError::new("trap: null array reference (array.init_elem)"));
                        }
                    } else {
                        return Err(VMError::new("trap: null array reference (array.init_elem)"));
                    }
                }
                // `struct.new_default $t` — no operands on the stack; the
                // instance is $t's declared fields at their default value.
                // Same typeidx discipline as `struct.new`: 0 is the dynamic
                // "empty object" form, non-zero allocates $t's field slots and
                // stamps its rtt.
                //
                // Defaults are `Null` for every slot because field TYPES are
                // not modelled — `FieldDef` carries a name, an index and a
                // property descriptor, but no value type, so there is nothing
                // to derive a typed zero from. Recovering the real per-type
                // defaults needs the reader to build field types into the
                // TypeEntry.
                Op::STRUCT_NEW_DEFAULT => {
                    let typeidx = self.read_u16() as usize;
                    let mut obj = Object::new();
                    if typeidx != 0 {
                        let type_id = self.resolve_gc_rtt(typeidx);
                        let arity = self
                            .type_registry
                            .get(type_id)
                            .map_or(0, |td| td.field_defs.len());
                        obj.fields = vec![Value::Null; arity];
                        obj.type_id = type_id;
                    }
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
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
                // These are `struct.new` / `struct.new_default` with a
                // descriptor operand — NOT a lightweight side form. A type
                // carrying a `(descriptor $d)` clause can ONLY be allocated
                // with these (Overview.md §"Allocation With Descriptors":
                // *"`struct.new` and `struct.new_default` cannot be used to
                // allocate types with descriptors"*), so they must do
                // everything the plain forms do — size the field vector from
                // the declared type and stamp the rtt — and then attach the
                // descriptor.
                //
                // The descriptor is the LAST operand, so it is on top of the
                // stack, above the field values. Per the typing rule
                //   struct.new_desc x : t* (ref null (exact y)) -> (ref (exact x))
                // the field count comes from `C.types[x]`, exactly as it does
                // for `struct.new` — hence the same `field_defs.len()` source
                // rather than the bytecode's count immediate, which is only a
                // fallback for a type the registry does not know.
                Op::STRUCT_NEW_DESC => {
                    let typeidx = self.read_u16() as usize;
                    let count = self.read_u16() as usize;
                    let descriptor = self.pop();
                    if descriptor.is_null_ref() {
                        return Err(VMError::new("trap: null descriptor reference"));
                    }
                    let type_id = if typeidx != 0 {
                        self.resolve_gc_rtt(typeidx)
                    } else {
                        0
                    };
                    let arity = if type_id != 0 {
                        self.type_registry
                            .get(type_id)
                            .map_or(count, |td| td.field_defs.len())
                    } else {
                        count
                    }
                    .min(self.stack.len());
                    let start = self.stack.len() - arity;
                    let fields: Vec<Value> = self.stack[start..].to_vec();
                    self.stack.truncate(start);
                    let mut obj = Object::new();
                    obj.fields = fields;
                    obj.type_id = type_id;
                    set_descriptor(&mut obj, descriptor);
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
                }
                Op::STRUCT_NEW_DEFAULT_DESC => {
                    let typeidx = self.read_u16() as usize;
                    let descriptor = self.pop();
                    if descriptor.is_null_ref() {
                        return Err(VMError::new("trap: null descriptor reference"));
                    }
                    let mut obj = Object::new();
                    if typeidx != 0 {
                        let type_id = self.resolve_gc_rtt(typeidx);
                        let arity = self
                            .type_registry
                            .get(type_id)
                            .map_or(0, |td| td.field_defs.len());
                        obj.fields = vec![Value::Null; arity];
                        obj.type_id = type_id;
                    }
                    set_descriptor(&mut obj, descriptor);
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
                }
                // `ref.get_desc x : (ref null (exact_1 x)) -> (ref (exact_1 y))`
                //
                // The RESULT is non-nullable, so a null input cannot be passed
                // through — it traps. `test/core/custom-descriptors/
                // ref_get_desc.wast:400-406` asserts exactly that, six ways.
                Op::REF_GET_DESC => {
                    let _typeidx = self.read_u16();
                    let val = self.pop();
                    if val.is_null_ref() {
                        return Err(VMError::new("trap: null reference"));
                    }
                    let desc = descriptor_of(&val);
                    self.push(desc)?;
                }
                // `ref.cast_desc_eq (ref ht)` / `(ref null ht)` — cast keyed on
                // descriptor identity rather than on the type hierarchy.
                //
                //   [ref, descriptor] -> [ref]
                //
                // The descriptor operand is popped; the reference stays on the
                // stack, exactly like `ref.cast`. Per the proposal a NULL
                // descriptor traps unconditionally — before the reference is
                // even looked at — so the null check comes first for all four
                // of these instructions.
                Op::REF_CAST_DESC_EQ | Op::REF_CAST_DESC_EQ_NULL => {
                    let _typeidx = self.read_u16();
                    let expected = self.pop();
                    if expected.is_null_ref() {
                        return Err(VMError::new("trap: null descriptor reference"));
                    }
                    let val = self.peek(0).clone();
                    if val.is_null_ref() {
                        // The `(ref ht)` form does not admit null; the
                        // `(ref null ht)` form passes it through untouched.
                        //
                        // ⚠ The message is "descriptor cast failure", NOT
                        // "null reference". A null reference here is an
                        // ordinary failed cast, and the proposal's suite
                        // distinguishes the two:
                        //   (assert_trap (invoke "self-nonnullable-null-desc")
                        //                "descriptor cast failure")
                        // while `ref.get_desc`, whose null check is about its
                        // own non-nullable RESULT rather than a cast, keeps
                        // "null reference" (ref_get_desc.wast:400).
                        if op == Op::REF_CAST_DESC_EQ {
                            return Err(VMError::new("trap: descriptor cast failure"));
                        }
                    } else if !ref_eq(&descriptor_of(&val), &expected) {
                        return Err(VMError::new(
                            "trap: descriptor cast failure",
                        ));
                    }
                }
                // `br_on_cast_desc_eq $l ht ht` / `..._fail` — the branching
                // forms. Same operand discipline as `br_on_cast`: (u16
                // type-name-idx, u8 label depth). The descriptor is consumed
                // either way; the reference stays for both the branch and the
                // fallthrough.
                Op::BR_ON_CAST_DESC_EQ | Op::BR_ON_CAST_DESC_EQ_FAIL => {
                    // `ht_2` (target) then `ht_1` (source), then the label
                    // depth. `ht_1` is carried for the WRITER — it is one of
                    // this instruction's spec immediates and the binary is
                    // wrong without it — and is not consulted here, because
                    // nothing at run time depends on the static source type.
                    let to_idx = self.read_u16();
                    let _from_idx = self.read_u16();
                    let depth = self.read_byte() as usize;
                    // ⛔ A NULL REFERENCE MATCHES A NULLABLE TARGET.
                    //
                    // This read `matched = !val.is_null_ref() && …`, so a null
                    // reference never matched even against `(ref null ht_2)`,
                    // which the spec requires it to. That is exactly why the
                    // wast walker LOWERED this instruction into
                    // `ref.is_null` + `if`/`else` instead of emitting it — the
                    // lowering was the only correct implementation available,
                    // not laziness. Nullability rides in the target's spelling.
                    let target_nullable = self
                        .frames
                        .last()
                        .and_then(|f| self.chunks.get(f.chunk_index))
                        .and_then(|c| c.constants.get(to_idx as usize))
                        .map(|v| v.to_string())
                        .is_some_and(|n| {
                            let n = n.trim_start_matches('(').trim();
                            n.starts_with("ref null") || n.starts_with("null ")
                        });
                    let expected = self.pop();
                    if expected.is_null_ref() {
                        return Err(VMError::new("trap: null descriptor reference"));
                    }
                    let val = self.peek(0).clone();
                    let matched = if val.is_null_ref() {
                        target_nullable
                    } else {
                        ref_eq(&descriptor_of(&val), &expected)
                    };
                    let take = if op == Op::BR_ON_CAST_DESC_EQ {
                        matched
                    } else {
                        !matched
                    };
                    if take {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth).copied() {
                            self.branch_to_label(depth, entry);
                        }
                    }
                }
                // `struct.get_s $t i` / `struct.get_u $t i` — packed field
                // variants. Our structs have externref fields only, so
                // there's no sign extension to do — both behave like
                // `struct.get`.
                Op::STRUCT_GET_S | Op::STRUCT_GET_U => {
                    let _typeidx = self.read_u16();
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
                Op::REF_TEST_NULL => {
                    let ht = HeapType::from_sleb(self.read_leb_i32());
                    let val = self.pop();
                    let result = val.is_null_ref() || self.ref_test_or_declared_name(&val, ht);
                    self.push(Value::I32(if result { 1 } else { 0 }))?;
                }
                Op::REF_CAST_NULL => {
                    let ht = HeapType::from_sleb(self.read_leb_i32());
                    let val = self.peek(0).clone();
                    if !val.is_null_ref() && !self.ref_test_or_declared_name(&val, ht) {
                        return Err(VMError::new(&format!(
                            "trap: cast failure: value is not {}",
                            self.heaptype_label(ht)
                        )));
                    }
                }
                // `any.convert_extern` / `extern.convert_any` — identity at
                // runtime for us: our value ABI is a universal externref.
                // Spec says composing the two yields the original value, so
                // emitting them as nops is semantically correct.
                Op::ANY_CONVERT_EXTERN | Op::EXTERN_CONVERT_ANY => {}
                // `ref.as_non_null` — trap if the operand is null, otherwise
                // pass the value through unchanged.
                Op::REF_AS_NON_NULL => {
                    if self.stack.last().map_or(false, |v| v.is_null_ref()) {
                        return Err(VMError::new("trap: ref.as_non_null on null reference"));
                    }
                }
                // `br_on_null $l` — if TOS is null, pop it and branch;
                // otherwise leave the value on the stack and fall through.
                // `br_on_non_null $l` — if TOS is non-null, branch with
                // the value; otherwise pop and fall through.
                Op::BR_ON_NULL => {
                    let offset = self.read_i16();
                    let is_null = self.stack.last().map_or(false, |v| v.is_null_ref());
                    if is_null {
                        self.pop();
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                Op::BR_ON_NON_NULL => {
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
                // `ref.null <heaptype>`. A GC-heap null is a WASM GC **typed
                // null** — the GC accessors trap on it per spec — while
                // `ref.null extern`/`func` is the lenient null the dynamic
                // languages use. That distinction is the immediate, not a
                // second opcode.
                Op::NULL => {
                    let ht = self.read_byte();
                    if crate::opcode::heaptype::is_gc_heap(ht) {
                        self.push(Value::TypedNull(0))?
                    } else {
                        self.push(Value::Null)?
                    }
                }

                Op::I32_CONST => {
                    let v = self.read_leb_i32();
                    self.push(Value::I32(v))?;
                }
                Op::I64_CONST => {
                    let v = self.read_leb_i64();
                    self.push(Value::I64(v))?;
                }
                Op::F32_CONST => {
                    let v = self.read_f32();
                    self.push(Value::F32(v))?;
                }
                Op::F64_CONST => {
                    let v = self.read_f64();
                    self.push(Value::F64(v))?;
                }

                // ref.eq (GC proposal) — reference identity equality.
                // Two references are equal iff they point at the same
                // underlying object. Null-null is also true. Used by JS
                // `===` for object identity.
                Op::REF_EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(wasm_bool(ref_eq(&a, &b)))?;
                }

                // -- Type checks --
                // `ref.test ht` / `ref.cast ht` — the immediate is a heaptype,
                // decoded from the spec's signed LEB. Abstract types answer
                // from the value's shape; a concrete one is an index walk over
                // declared supertypes.
                Op::REF_TEST => {
                    let ht = HeapType::from_sleb(self.read_leb_i32());
                    let val = self.pop();
                    let result = self.ref_test_or_declared_name(&val, ht);
                    self.push(wasm_bool(result))?;
                }
                // The EXACT forms — `(ref null? (exact $t))`. Identical operand
                // to the inexact opcodes; only the comparison narrows, from
                // "is a subtype of" to "is". Null is admitted by the `_null`
                // spelling exactly as it is for `ref.test`/`ref.cast`.
                Op::REF_TEST_EXACT | Op::REF_TEST_EXACT_NULL => {
                    let ht = HeapType::from_sleb(self.read_leb_i32());
                    let val = self.pop();
                    let result = if val.is_null_ref() {
                        op == Op::REF_TEST_EXACT_NULL
                    } else {
                        self.ref_test_exact(&val, ht)
                    };
                    self.push(wasm_bool(result))?;
                }
                // VM-internal (`desc.set_proto`). The descriptor singleton is a
                // WASM-GC construct that only exists in the emitted binary —
                // this interpreter resolves prototypes through its own object
                // model and never reads descriptor field 0. So the VM's whole
                // job here is to keep the stack honest: consume the prototype
                // the class path pushed, and read the immediate so the
                // instruction stream stays in sync.
                //
                // ⛔ That means field 0 is NOT observable from this
                // interpreter, and no test run under it can check the write.
                // The check is at the BINARY level (`-w`, then decode) for
                // exactly this reason.
                Op::DESC_SET_PROTO => {
                    // U16, not a LEB: the compiler writes it with `emit_u16`
                    // and the writer reads it with `read_u16`. All three sides
                    // of the operand contract must agree or the instruction
                    // stream desynchronises — which it did, as
                    // "Unhandled opcode: unknown".
                    let _name_idx = self.read_u16();
                    let _proto = self.pop();
                }
                Op::REF_CAST_EXACT | Op::REF_CAST_EXACT_NULL => {
                    let ht = HeapType::from_sleb(self.read_leb_i32());
                    let val = self.peek(0).clone();
                    let ok = if val.is_null_ref() {
                        op == Op::REF_CAST_EXACT_NULL
                    } else {
                        self.ref_test_exact(&val, ht)
                    };
                    if !ok {
                        // `trap: ` prefixes every trap — `VMError::is_trap`
                        // classifies on it alone. "cast failure" is the
                        // proposal's own wording (`exact-casts.wast`).
                        return Err(VMError::new("trap: cast failure"));
                    }
                }
                Op::REF_CAST => {
                    let ht = HeapType::from_sleb(self.read_leb_i32());
                    let val = self.peek(0).clone();
                    if !self.ref_test_or_declared_name(&val, ht) {
                        // `trap: ` is not decoration — `VMError::is_trap`
                        // classifies solely on that prefix, and only a message
                        // carrying it is offered to a host-level handler at the
                        // trap/host boundary. Without it a failed cast escaped
                        // uncaught: the spec says `ref.cast` TRAPS, and a trap
                        // has to be catchable like every other one.
                        return Err(VMError::new(&format!(
                            "trap: cast failure: value is not {}",
                            self.heaptype_label(ht)
                        )));
                    }
                    // Value stays on stack (cast is a no-op if it passes)
                }
                // `br_on_cast l ht` / `br_on_cast_fail l ht` — structured
                // branch keyed off a runtime type test. Operand is
                // (u16 type-name-idx, u8 label-depth), matching core `br`'s
                // label-stack discipline so the VM can honour the branch
                // without a parallel byte-offset table.
                Op::BR_ON_CAST => {
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
                Op::BR_ON_CAST_FAIL => {
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
                Op::I31_NEW => {
                    // Box i32 as i31ref. In our VM, I32 is already unboxed,
                    // so this is a no-op identity. The optimization is that
                    // the VM can use I32 directly without heap allocation.
                    let v = self.pop().as_i32();
                    self.push(Value::I32(v & 0x7FFF_FFFF))?; // mask to 31 bits
                }
                Op::I31_GET_S => {
                    // Both getters take `(ref null i31)` and trap on null; a
                    // null used to read back as 0, indistinguishable from a
                    // genuine `ref.i31 0`.
                    let raw = self.pop();
                    if raw.is_null_ref() {
                        return Err(VMError::new("trap: null i31 reference (i31.get_s)"));
                    }
                    let v = raw.as_i32();
                    // Sign extend from 31 bits
                    let extended = if v & 0x4000_0000 != 0 {
                        v | !0x7FFF_FFFF_u32 as i32
                    } else {
                        v
                    };
                    self.push(Value::I32(extended))?;
                }
                Op::I31_GET_U => {
                    let raw = self.pop();
                    if raw.is_null_ref() {
                        return Err(VMError::new("trap: null i31 reference (i31.get_u)"));
                    }
                    self.push(Value::I32(raw.as_i32() & 0x7FFF_FFFF))?;
                }

                // ── Stringref proposal ────────────────────────────────────
                // Strings are `Value::String`. The `$mem` immediate defaults to
                // memory 0 (no immediate bytes read — matches operand_format).
                Op::STRING_NEW_UTF8 | Op::STRING_NEW_WTF8 => {
                    let len = self.pop().as_i32() as u32 as usize;
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let bytes = self.read_memory_bytes(0, ptr, len)?;
                    // string.new_utf8 traps on invalid UTF-8; new_wtf8 accepts
                    // valid UTF-8 identically (WTF-8 surrogate forms unsupported).
                    let s = String::from_utf8(bytes)
                        .map_err(|_| VMError::new("trap: invalid UTF-8"))?;
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                Op::STRING_NEW_LOSSY_UTF8 => {
                    let len = self.pop().as_i32() as u32 as usize;
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let bytes = self.read_memory_bytes(0, ptr, len)?;
                    let s = String::from_utf8_lossy(&bytes).into_owned();
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                Op::STRING_NEW_UTF8_ARRAY | Op::STRING_NEW_WTF16_ARRAY => {
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
                Op::STRING_MEASURE_UTF8 | Op::STRING_MEASURE_WTF8 => {
                    let s = self.pop_stringref()?;
                    self.push(Value::I32(s.len() as i32))?;
                }
                Op::STRING_MEASURE_WTF16 => {
                    let s = self.pop_stringref()?;
                    self.push(Value::I32(s.encode_utf16().count() as i32))?;
                }
                Op::STRING_ENCODE_UTF8 => {
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let bytes = s.as_bytes().to_vec();
                    self.write_memory_bytes(0, ptr, &bytes)?;
                    self.push(Value::I32(bytes.len() as i32))?;
                }
                Op::STRING_ENCODE_WTF16 => {
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
                Op::STRING_ENCODE_UTF8_ARRAY | Op::STRING_ENCODE_WTF16_ARRAY => {
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
                        return Err(VMError::new("trap: out of bounds array access"));
                    }
                    for (i, u) in units.iter().enumerate() {
                        elems[start + i] = Value::I32(*u as i32);
                    }
                    self.push(Value::I32(units.len() as i32))?;
                }
                Op::STRING_CONCAT => {
                    let b = self.pop_stringref()?;
                    let a = self.pop_stringref()?;
                    let mut s = String::with_capacity(a.len() + b.len());
                    s.push_str(&a);
                    s.push_str(&b);
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                Op::STRING_EQ => {
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
                Op::STRING_AS_WTF8 | Op::STRING_AS_WTF16 => {
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
                Op::STRING_NEW_WTF16 => {
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
                Op::STRING_ENCODE_WTF8 | Op::STRING_ENCODE_LOSSY_UTF8 => {
                    // (str, ptr): write the UTF-8 bytes, return the byte count.
                    let ptr = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let bytes = s.as_bytes().to_vec();
                    self.write_memory_bytes(0, ptr, &bytes)?;
                    self.push(Value::I32(bytes.len() as i32))?;
                }
                Op::STRING_NEW_WTF8_ARRAY | Op::STRING_NEW_LOSSY_UTF8_ARRAY => {
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
                Op::STRING_ENCODE_WTF8_ARRAY | Op::STRING_ENCODE_LOSSY_UTF8_ARRAY => {
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
                        return Err(VMError::new("trap: out of bounds array access"));
                    }
                    for (i, u) in units.iter().enumerate() {
                        elems[start + i] = Value::I32(*u as i32);
                    }
                    self.push(Value::I32(units.len() as i32))?;
                }
                Op::STRING_IS_USV_SEQUENCE => {
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
                Op::STRINGVIEW_WTF16_LENGTH => {
                    let s = self.pop_stringref()?;
                    self.push(Value::I32(s.encode_utf16().count() as i32))?;
                }
                Op::STRINGVIEW_WTF16_GET_CODEUNIT => {
                    let pos = self.pop().as_i32() as u32 as usize;
                    let s = self.pop_stringref()?;
                    let units: Vec<u16> = s.encode_utf16().collect();
                    let u = *units.get(pos).ok_or_else(|| {
                        VMError::new("trap: stringview_wtf16 index out of bounds")
                    })?;
                    self.push(Value::I32(u as i32))?;
                }
                Op::STRINGVIEW_WTF16_SLICE => {
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
                Op::STRINGVIEW_WTF16_ENCODE => {
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
                Op::STRINGVIEW_WTF8_ADVANCE => {
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
                Op::STRINGVIEW_WTF8_SLICE => {
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
                Op::STRINGVIEW_WTF8_ENCODE_UTF8 => {
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
                Op::STRING_AS_ITER => {
                    let s = self.pop_stringref()?;
                    let mut obj = crate::value::Object::new();
                    obj.properties
                        .insert("__iter_str".to_string(), Value::String(s));
                    obj.properties
                        .insert("__iter_pos".to_string(), Value::I32(0));
                    self.push(Value::Object(Arc::new(std::sync::Mutex::new(obj))))?;
                }
                Op::STRINGVIEW_ITER_NEXT => {
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
                Op::STRINGVIEW_ITER_ADVANCE => {
                    let n = self.pop().as_i32() as u32 as usize;
                    let view = self.pop();
                    let (s, pos) = self.read_string_iter(&view)?;
                    let total = s.chars().count();
                    let new_pos = pos.saturating_add(n).min(total);
                    self.write_string_iter_pos(&view, new_pos)?;
                    self.push(Value::I32((new_pos - pos) as i32))?;
                }
                Op::STRINGVIEW_ITER_REWIND => {
                    let n = self.pop().as_i32() as u32 as usize;
                    let view = self.pop();
                    let (_s, pos) = self.read_string_iter(&view)?;
                    let new_pos = pos.saturating_sub(n);
                    self.write_string_iter_pos(&view, new_pos)?;
                    self.push(Value::I32((pos - new_pos) as i32))?;
                }
                Op::STRINGVIEW_ITER_SLICE => {
                    // Substring of up to `n` codepoints from the cursor; does NOT
                    // advance the iterator.
                    let n = self.pop().as_i32() as u32 as usize;
                    let view = self.pop();
                    let (s, pos) = self.read_string_iter(&view)?;
                    let out: String = s.chars().skip(pos).take(n).collect();
                    self.push(Value::String(Arc::from(out.as_str())))?;
                }

                Op::REF_IS_NULL => {
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
                Op::F64_FROM_I32 => {
                    let v = self.pop();
                    self.push(Value::F64(v.as_f64()))?;
                }
                Op::F64_CONVERT_I32_U => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::F64(a as f64))?;
                }
                Op::F64_CONVERT_I64_S => {
                    let a = self.pop().as_i64();
                    self.push(Value::F64(a as f64))?;
                }
                Op::F64_CONVERT_I64_U => {
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::F64(a as f64))?;
                }
                Op::F32_CONVERT_I32_S => {
                    let a = self.pop().as_i32();
                    self.push(Value::F32(a as f32))?;
                }
                Op::F32_CONVERT_I32_U => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::F32(a as f32))?;
                }
                Op::F32_CONVERT_I64_S => {
                    let a = self.pop().as_i64();
                    self.push(Value::F32(a as f32))?;
                }
                Op::F32_CONVERT_I64_U => {
                    let a = self.pop().as_i64() as u64;
                    self.push(Value::F32(a as f32))?;
                }
                // `i32.trunc_f64_s` (0xAA). The SIGNED domain is the open
                // interval (-2^31 - 1, 2^31): the result of `trunc` has to
                // land in [-2^31, 2^31), and -2147483648.9 truncates to
                // -2147483648, which does. Guarding on `v < -2^31` rejected
                // the whole of (-2^31 - 1, -2^31] — `conversions.wast:123`.
                //
                // The f32 forms below keep `< -2^31` deliberately: -2147483649
                // is not representable in f32 (it rounds to -2^31, which would
                // then trap a legal value), and the f32 spacing near 2^31 is
                // 256, so no representable f32 falls in the gap anyway. Same
                // reasoning for the i64 forms against f64.
                Op::I32_FROM_F64 => {
                    let v = self.pop().as_f64();
                    if v.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if v >= 2147483648.0 || v <= -2147483649.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as i32))?;
                }
                // The unsigned truncations' domain is the OPEN interval
                // (-1, 2^N): `trunc` rounds toward zero, so every value
                // strictly greater than -1 lands on 0 and is legal — tiny
                // negatives, -0.0, -0.9. Only -1.0 and below are out of range.
                // Rejecting `v < 0.0` trapped on the whole of (-1, 0), which
                // is what `conversions.wast:90` catches.
                //
                // NaN is a DIFFERENT trap from being out of range (spec
                // §4.3.3; `conversions.wast:101` vs `:104`), so the two carry
                // different messages.
                Op::I32_TRUNC_F64_U => {
                    let v = self.pop().as_f64();
                    if v.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if v <= -1.0 || v >= 4294967296.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as u32 as i32))?;
                }
                Op::I32_TRUNC_F32_S => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if v >= 2147483648.0f32 || v < -2147483648.0f32 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as i32))?;
                }
                Op::I32_TRUNC_F32_U => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if v <= -1.0f32 || v >= 4294967296.0f32 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I32(v as u32 as i32))?;
                }
                Op::I64_TRUNC_F32_S => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if v >= 9223372036854775808.0f32 || v < -9223372036854775808.0f32 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(v as i64))?;
                }
                Op::I64_TRUNC_F32_U => {
                    let v = self.pop().as_f64() as f32;
                    if v.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if v <= -1.0f32 || v >= 18446744073709551616.0f32 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(v as u64 as i64))?;
                }

                // nontrapping-float-to-int-conversions proposal (0xFC 0x00–0x07).
                // Rust `as` casts saturate since 1.45: NaN → 0, overflow → min/max.
                Op::I32_TRUNC_SAT_F32_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32(v as i32))?;
                }
                Op::I32_TRUNC_SAT_F32_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32((v as u32) as i32))?;
                }
                Op::I32_TRUNC_SAT_F64_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32(v as i32))?;
                }
                Op::I32_TRUNC_SAT_F64_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I32((v as u32) as i32))?;
                }
                Op::I64_TRUNC_SAT_F32_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64(v as i64))?;
                }
                Op::I64_TRUNC_SAT_F32_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64((v as u64) as i64))?;
                }
                Op::I64_TRUNC_SAT_F64_S => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64(v as i64))?;
                }
                Op::I64_TRUNC_SAT_F64_U => {
                    let v = self.pop().as_f64();
                    self.push(Value::I64((v as u64) as i64))?;
                }

                // -- Async (await) --
                // r#await: removed (duplicate of promise_suspend, use JSPI proposal name)

                // -- Exceptions (WASM exception-handling proposal, final) --
                // Normal exit from a try block is handled by the structural
                // `end` (Op::END, `is_try` label) — see the END dispatch. The
                // custom TRY_END opcode has been retired.
                Op::THROW => {
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
                Op::THROW_REF => {
                    // Spec `throw_ref`: rethrow the exception an exnref
                    // refers to — same tag identity, same payload.
                    let val = self.pop();
                    // A NULL exnref is the spec's own trap case, with its own
                    // wording (`throw_ref.wast` asserts "null exception
                    // reference"). Reporting it as "not an exnref" both misses
                    // that assertion and describes the wrong fault: a null
                    // `exnref` IS an exnref.
                    if val.is_null_ref() || matches!(val, Value::Undefined) {
                        return Err(VMError::new("trap: null exception reference"));
                    }
                    let (entity, payload) = Self::unpack_exnref(&val)
                        .ok_or_else(|| VMError::new("throw_ref: operand is not an exnref"))?;
                    self.raise_exception(entity, payload, 0)?;
                }
                Op::RETHROW => {
                    // Legacy EH rethrow — carries the exception object as a
                    // value; re-raises through the vybe:exception tag.
                    let chunk_idx = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let _depth = read_leb_u32(&self.chunks[chunk_idx].code, &mut ip);
                    self.frame_mut().ip = ip;
                    let val = self.pop();
                    self.raise_exception_value(val)?;
                }
                Op::DELEGATE => {
                    let chunk_idx = self.frame().chunk_index;
                    let mut ip = self.frame().ip;
                    let depth = read_leb_u32(&self.chunks[chunk_idx].code, &mut ip);
                    self.frame_mut().ip = ip;
                    let val = self.pop();
                    self.raise_exception_value_skipping(val, depth as usize)?;
                }
                Op::TRY_TABLE => {
                    // Spec try_table. Internal fixed-width encoding:
                    //   [try_table, u8 clause_count, per clause:
                    //    u8 kind (0=catch 1=catch_ref 2=catch_all 3=catch_all_ref),
                    //    u16 tag_idx (ignored for catch_all kinds),
                    //    u16 labelidx (relative block depth, as `br`)]
                    // Matching is TAG IDENTITY only — clauses are tried in
                    // order (pushed reversed so the first clause is on top).
                    // Spec blocktype `bt` — `try_table` IS a block and may take
                    // and produce values. Encoded as BLOCK's (params, results).
                    let params = self.read_byte() as usize;
                    let try_results = self.read_byte();
                    // Base below the params, as for BLOCK. The label and the
                    // handlers must agree on it: `stack_depth` is where a
                    // caught exception unwinds to before the payload is
                    // pushed, and that is the same base a `br` truncates to.
                    let param_base = self.stack.len().saturating_sub(params);
                    let clause_count = self.read_u16() as usize;
                    let chunk_index = self.frame().chunk_index;
                    self.try_group_counter += 1;
                    let group = self.try_group_counter;
                    let mut handlers = Vec::with_capacity(clause_count);
                    for _ in 0..clause_count {
                        let kind = self.read_byte();
                        let tag_idx = self.read_u16();
                        let catch_label = self.read_u16();
                        let tag_entity = if kind == crate::vm::CATCH_KIND_CATCH
                            || kind == crate::vm::CATCH_KIND_CATCH_REF
                        {
                            self.resolve_chunk_tag(chunk_index, tag_idx)?
                        } else {
                            0 // unused for catch_all kinds
                        };
                        handlers.push(ExceptionHandler {
                            catch_label,
                            stack_depth: param_base,
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
                        // Names the clause group this label protects, so `end`
                        // and a `br` out of the region dispose of the same set.
                        try_group: group,
                        // The try_table's OWN blocktype results. This was
                        // hardcoded 0, so a `try_table (result i32)` dropped its
                        // value on normal completion and on any `br` to it.
                        result_arity: try_results,
                        stack_height: param_base,
                    });
                }

                // -- Tail call --
                Op::RETURN_CALL => {
                    let argc = self.read_byte() as usize;
                    let _results = self.read_byte(); // writer-facing functype info
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
                    self.pop_frame_for_tail_call();
                    self.call_value(argc)?;
                }
                Op::RETURN_CALL_INDIRECT => {
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
                        // ⛔ THE SAME TRAP AS `call_indirect`, SPELLED THE SAME
                        // WAY. An out-of-range table index is "undefined
                        // element" in the spec, and this arm said "invalid
                        // table index" — so three fixtures asserting the spec
                        // wording failed against a module that trapped
                        // correctly, for the wrong reason.
                        if raw_idx < 0.0 || raw_idx.is_nan() || raw_idx >= table.len() as f64 {
                            return Err(VMError::new(format!(
                                "trap: undefined element: table index {} out of bounds",
                                raw_idx
                            )));
                        }
                        table[raw_idx as usize].clone()
                    };
                    // A NULL slot traps before the call, exactly as in
                    // `call_indirect` — this arm never had the check.
                    if funcref.is_null_ref() || matches!(funcref, Value::Undefined) {
                        return Err(VMError::new(format!(
                            "trap: uninitialized element {} (table slot is null)",
                            raw_idx
                        )));
                    }
                    self.indirect_call_type_check(
                        &funcref,
                        argc,
                        expected_results,
                        opcode_start,
                        true,
                    )?;
                    // Splice the funcref in below the args, then reuse the frame.
                    let callee_idx = self.stack.len() - argc;
                    self.stack.insert(callee_idx, funcref);
                    let old_base = self.frame().base;
                    for i in 0..=argc {
                        self.stack[old_base + i] = self.stack[callee_idx + i].clone();
                    }
                    self.stack.truncate(old_base + 1 + argc);
                    self.pop_frame_for_tail_call();
                    self.call_value(argc)?;
                }
                Op::RETURN_CALL_REF => {
                    let argc = self.read_byte() as usize;
                    let _results = self.read_byte(); // writer-facing functype info
                    let old_base = self.frame().base;
                    let callee_idx = self.stack.len() - argc - 1;
                    for i in 0..=argc {
                        self.stack[old_base + i] = self.stack[callee_idx + i].clone();
                    }
                    self.stack.truncate(old_base + 1 + argc);
                    self.pop_frame_for_tail_call();
                    self.call_value(argc)?;
                }

                // -- Linear memory --
                Op::MEMORY_SIZE => {
                    let memidx = self.read_u16() as usize;
                    let pages = self.mem_len(memidx) / 65536;
                    // memory64: page count is i64; 32-bit memory: i32.
                    if self.mem_is_64(memidx) {
                        self.push(Value::I64(pages as i64))?;
                    } else {
                        self.push(Value::I32(pages as i32))?;
                    }
                }
                Op::MEMORY_GROW => {
                    let memidx = self.read_u16() as usize;
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
                Op::I32_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    self.push(Value::I32(i32::from_le_bytes(read_le(&bytes))))?;
                }
                Op::I32_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i32();
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                Op::I64_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    self.push(Value::I64(i64::from_le_bytes(read_le(&bytes))))?;
                }
                Op::I64_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64();
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                Op::F64_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    self.push(Value::F64(f64::from_le_bytes(read_le(&bytes))))?;
                }
                Op::F64_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_f64();
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                Op::I32_LOAD8_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I32(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i32,
                    ))?;
                }
                Op::I32_STORE8 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i32() as u8;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &[val])?;
                }
                Op::F32_LOAD => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    let val = f32::from_le_bytes(read_le(&bytes));
                    self.push(Value::F32(val))?;
                }
                Op::F32_STORE => {
                    let (offset, memidx) = self.read_optional_memarg();
                    // Bit-preserving: memory stores the operand's ENCODING, so
                    // this must not round-trip through f64 (that quiets a
                    // signalling NaN). `float_memory.wast` exists to check
                    // exactly this — "load and store do not canonicalize NaNs".
                    let val = self.pop().as_f32();
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                Op::I32_LOAD8_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I32(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i8 as i32,
                    ))?;
                }
                Op::I32_LOAD16_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = i16::from_le_bytes(read_le(&bytes)) as i32;
                    self.push(Value::I32(val))?;
                }
                Op::I32_LOAD16_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = u16::from_le_bytes(read_le(&bytes)) as i32;
                    self.push(Value::I32(val))?;
                }
                Op::I32_STORE16 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i32() as i16;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                Op::I64_LOAD8_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I64(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i8 as i64,
                    ))?;
                }
                Op::I64_LOAD8_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    self.push(Value::I64(
                        self.read_memory_bytes(memidx, addr, 1)?[0] as i64,
                    ))?;
                }
                Op::I64_LOAD16_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = i16::from_le_bytes(read_le(&bytes)) as i64;
                    self.push(Value::I64(val))?;
                }
                Op::I64_LOAD16_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 2)?;
                    let val = u16::from_le_bytes(read_le(&bytes)) as i64;
                    self.push(Value::I64(val))?;
                }
                Op::I64_LOAD32_S => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    self.push(Value::I64(i32::from_le_bytes(read_le(&bytes)) as i64))?;
                }
                Op::I64_LOAD32_U => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let addr = self.effective_addr(memidx, offset);
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    self.push(Value::I64(i32::from_le_bytes(read_le(&bytes)) as u32 as i64))?;
                }
                Op::I64_STORE8 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64() as u8;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &[val])?;
                }
                Op::I64_STORE16 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64() as i16;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }
                Op::I64_STORE32 => {
                    let (offset, memidx) = self.read_optional_memarg();
                    let val = self.pop().as_i64() as i32;
                    let addr = self.effective_addr(memidx, offset);
                    self.write_memory_bytes(memidx, addr, &val.to_le_bytes())?;
                }

                // -- Conversions --
                Op::I32_WRAP_I64 => {
                    let a = self.pop().as_i64();
                    self.push(Value::I32(a as i32))?;
                }
                Op::I64_EXTEND_I32_S => {
                    let a = self.pop().as_i32();
                    self.push(Value::I64(a as i64))?;
                }
                Op::I64_EXTEND_I32_U => {
                    let a = self.pop().as_i32() as u32;
                    self.push(Value::I64(a as i64))?;
                }
                Op::I64_TRUNC_F64_S => {
                    let a = self.pop().as_f64();
                    if a.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if a >= 9223372036854775808.0 || a < -9223372036854775808.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(a as i64))?;
                }
                Op::I64_TRUNC_F64_U => {
                    let a = self.pop().as_f64();
                    if a.is_nan() {
                        return Err(VMError::new("trap: invalid conversion to integer"));
                    }
                    if a <= -1.0 || a >= 18446744073709551616.0 {
                        return Err(VMError::new("trap: integer overflow"));
                    }
                    self.push(Value::I64(a as u64 as i64))?;
                }
                Op::F64_PROMOTE_F32 => {
                    let a = self.pop().as_f64();
                    self.push(Value::F64(a))?;
                }
                Op::F32_DEMOTE_F64 => {
                    let a = self.pop().as_f64();
                    self.push(Value::F32(a as f32))?;
                }
                Op::I32_REINTERPRET_F32 => {
                    // `as_f32()`, NOT `as_f64() as f32`. Reinterpret is a pure
                    // bit copy, so it must not round-trip through f64: widening
                    // a SIGNALLING NaN to f64 and narrowing back sets the quiet
                    // bit (x86 `cvtss2sd` quiets an SNaN), so `0x7fa00000` came
                    // back as `0x7fc00000` — a bit pattern the operand never
                    // had. `as_f32` returns a `Value::F32` untouched.
                    let a = self.pop().as_f32();
                    self.push(Value::I32(a.to_bits() as i32))?;
                }
                Op::I64_REINTERPRET_F64 => {
                    let a = self.pop().as_f64();
                    self.push(Value::I64(a.to_bits() as i64))?;
                }
                Op::F32_REINTERPRET_I32 => {
                    let a = self.pop().as_i32();
                    self.push(Value::F32(f32::from_bits(a as u32)))?;
                }
                Op::F64_REINTERPRET_I64 => {
                    let a = self.pop().as_i64();
                    self.push(Value::F64(f64::from_bits(a as u64)))?;
                }

                // -- Sign extension --
                Op::I32_EXTEND8_S => {
                    let a = self.pop().as_i32() as i8;
                    self.push(Value::I32(a as i32))?;
                }
                Op::I32_EXTEND16_S => {
                    let a = self.pop().as_i32() as i16;
                    self.push(Value::I32(a as i32))?;
                }
                Op::I64_EXTEND8_S => {
                    let a = self.pop().as_i64() as i8;
                    self.push(Value::I64(a as i64))?;
                }
                Op::I64_EXTEND16_S => {
                    let a = self.pop().as_i64() as i16;
                    self.push(Value::I64(a as i64))?;
                }
                Op::I64_EXTEND32_S => {
                    let a = self.pop().as_i64() as i32;
                    self.push(Value::I64(a as i64))?;
                }

                // -- Multi-value --
                // pack, unpack: removed (non-WASM, were unused by compilers)

                // -- Block/loop/if structured control (WASM-compliant) --
                Op::BLOCK => {
                    // Blocktype = (param_count, result_count). Params are
                    // already on the stack (no runtime action); the label's
                    // branch arity for a BLOCK is its RESULT count (spec).
                    let param_arity = self.read_byte() as usize;
                    let result_arity = self.read_byte();
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
                        try_group: 0,
                        result_arity,
                        // The label's base is BELOW the params (spec §2.4.7):
                        // a `br` keeps the result and discards everything down
                        // to it, params included. Recording `stack.len()` with
                        // the params still on the stack put the base one slot
                        // per param too high, so a `br` left them stranded
                        // under the result.
                        stack_height: self.stack.len().saturating_sub(param_arity),
                    });
                }
                Op::LOOP => {
                    // Blocktype = (param_count, result_count). Spec: a `br`
                    // to a LOOP label carries the loop's PARAMS — so the
                    // label arity recorded here is the param count, not the
                    // result count.
                    let param_arity = self.read_byte();
                    let _result_arity = self.read_byte();
                    let param_base = param_arity as usize;
                    // Loop target is the ip right after the blocktype bytes —
                    // that is where `br 0` restarts (the loop body start).
                    let loop_body_start = self.frame().ip;
                    self.label_stack.push(LabelEntry {
                        target: loop_body_start,
                        is_loop: true,
                        is_try: false,
                        try_group: 0,
                        result_arity: param_arity,
                        // Base below the params, as for BLOCK. A `br 0` here
                        // keeps `param_arity` values and truncates to this
                        // base — with the params counted in, each iteration
                        // re-pushed them and the stack grew without bound.
                        stack_height: self.stack.len().saturating_sub(param_base),
                    });
                }
                Op::IF => {
                    // Blocktype = (param_count, result_count); label arity
                    // for an IF is its RESULT count, like BLOCK.
                    let param_arity = self.read_byte() as usize;
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
                            try_group: 0,
                            result_arity,
                            // Base below the params. The i32 condition was
                            // already popped above, so `len()` here is the
                            // post-condition height and the params are the
                            // top `param_arity` slots of it.
                            stack_height: self.stack.len().saturating_sub(param_arity),
                        });
                    } else if let Some(else_ip) = targets.else_ip {
                        // Condition false, ELSE exists — push label and jump into else-body.
                        // The else-body ends at END which pops the label.
                        // (We jump past the ELSE opcode itself to reach the else-body start.)
                        self.label_stack.push(LabelEntry {
                            target: targets.end_ip,
                            is_loop: false,
                            is_try: false,
                            try_group: 0,
                            result_arity,
                            // Same base as the then-arm: both arms of an `if`
                            // share one blocktype, so both see the params.
                            stack_height: self.stack.len().saturating_sub(param_arity),
                        });
                        self.frame_mut().ip = else_ip + 4; // +4 skips the ELSE opcode bytes
                    } else {
                        // Condition false, no ELSE — skip the block entirely.
                        // No label push needed; jump past END directly.
                        self.frame_mut().ip = targets.end_ip;
                    }
                }
                Op::ELSE => {
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
                Op::END => {
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
                            // THIS label's group, not whatever group happens to
                            // be on top of the handler stack.
                            self.exception_handlers
                                .retain(|h| h.group != label.try_group);
                        }
                    }
                }
                Op::BR_TABLE => {
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
                Op::CALL_INDIRECT => {
                    let argc = self.read_byte() as usize;
                    let tableidx = self.read_byte() as usize;
                    let expected_results = self.read_byte() as usize;
                    // Spec `call_indirect`: `[t* i32] → [t'*]` — the table index
                    // is on TOP of the stack, above the `argc` call arguments.
                    // Pop it, resolve the funcref, then splice the funcref in
                    // below the args so `call_value` sees `[funcref, args…]`.
                    //
                    // A table64 (`(table i64 …)`) is addressed with an i64, and
                    // `i64.const` lowers to `Literal::BigInt` — whose `as_f64`
                    // is NaN, which surfaced as "table index NaN out of
                    // bounds". `pop_table_count` reads either width through
                    // `as_i64`, the same route the neighbouring table ops take.
                    let raw_idx = self.pop_table_count(self.tbl_is_64(tableidx)) as f64;
                    let funcref = {
                        let table = self
                            .table_ref(tableidx)
                            .ok_or_else(|| VMError::new("trap: call_indirect unknown table"))?;
                        if raw_idx < 0.0 || raw_idx.is_nan() || raw_idx >= table.len() as f64 {
                            return Err(VMError::new(format!(
                                "trap: undefined element: table index {} out of bounds",
                                raw_idx
                            )));
                        }
                        table[raw_idx as usize].clone()
                    };
                    // Spec: `call_indirect` reads the table slot and calls the
                    // reference in it; a NULL reference traps (`call_ref` step
                    // 3a). Without this the null fell through to `call_value`
                    // and surfaced as the language-level "null is not
                    // callable" — a trap either way, but not this one, and the
                    // fixture asserting "uninitialized element" passed only
                    // because `assert_trap` used to ignore its message.
                    if funcref.is_null_ref() || matches!(funcref, Value::Undefined) {
                        return Err(VMError::new(format!(
                            "trap: uninitialized element {} (table slot is null)",
                            raw_idx
                        )));
                    }
                    // Spec §4.4.8 step 10: the funcref's TYPE must match the
                    // call's static `(type $sig)` — see
                    // `indirect_call_type_check` for why arity alone is not it.
                    self.indirect_call_type_check(
                        &funcref,
                        argc,
                        expected_results,
                        opcode_start,
                        false,
                    )?;
                    // ▶▶ CALL TAGS, Design §Instructions: "call_indirect
                    // $table $functype is now shorthand for (call_with_tag
                    // (call_tag.canon $functype) (table.get $table))".
                    //
                    // Without this, plain `call_indirect` bypassed tag checking
                    // entirely, so a func declaring `(call_tag $t1)` — which by
                    // the proposal handles EXACTLY $t1 and therefore NOT the
                    // canonical tag — stayed reachable through the front door.
                    // That defeats the property the proposal exists for:
                    // "a funcref called under a convention it does not handle
                    // STOPS, rather than being called anyway under the wrong
                    // shape."
                    //
                    // The canonical tag is taken from the CALL SITE's static
                    // `$functype` (argc → expected_results), exactly as the
                    // shorthand says. An undeclared func handles the canonical
                    // tag of its own shape, and the type check just above has
                    // already established that shape equals this one — so
                    // nothing that works today changes.
                    let callee_chunk = match &funcref {
                        Value::Object(o) => match &o.lock().unwrap().kind {
                            crate::value::ObjectKind::Function(f) => Some(f.chunk_index),
                            _ => None,
                        },
                        _ => None,
                    };
                    if let Some(ci) = callee_chunk {
                        // Shape key: this instruction's immediates are counts, not types.
                        let canon = self.call_tag_canon(argc as u8, expected_results as u8, "");
                        if !self.func_handles_call_tag(ci, canon) {
                            return Err(VMError::new(format!(
                                "trap: call_indirect: funcref does not handle the canonical \
                                 call tag for [{argc}->{expected_results}]"
                            )));
                        }
                    }
                    let insert_pos = self.stack.len() - argc;
                    self.stack.insert(insert_pos, funcref);
                    self.call_value(argc)?;
                }

                // -- Component Model --
                // No opcodes: canon built-ins are (core func) DEFINITIONS in
                // the CM spec, not instructions. They resolve as imports under
                // module "canon" (ImportTarget::Canon → exec_canon_builtin)
                // and are reached via spec `call`. Prefix 0xF0 is empty.

                // -- Shared-Everything Threads (shared GC objects) --

                // -- Weak References & Finalizers --

                // -- Multi-Memory --
                Op::MEMORY_INIT => {
                    let data_idx = self.read_u16() as u32;
                    let memidx = self.read_u16() as usize;
                    // Spec (bulk-memory Overview §data.drop): a dropped
                    // segment SHRINKS TO ZERO LENGTH — it may still be used
                    // by memory.init, "but only a zero-length access at
                    // offset zero will not trap". A missing segment (raw
                    // chunks that skipped validation) behaves the same.
                    // Operands are UNSIGNED (dst is the memory's index
                    // type), and a zero count still bounds-checks both ends.
                    let is64 = self.mem_is_64(memidx);
                    let count = self.pop().as_i32() as u32 as usize;
                    let src = self.pop().as_i32() as u32 as usize;
                    let dst = self.pop_mem_index(is64);
                    let seg_len = if self.dropped_data.contains(&data_idx) {
                        0
                    } else {
                        self.data_segments
                            .get(data_idx as usize)
                            .map(|d| d.len())
                            .unwrap_or(0)
                    };
                    if src.saturating_add(count) > seg_len {
                        return Err(VMError::new("trap: out of bounds memory access (memory.init source)"));
                    }
                    if dst.saturating_add(count) > self.mem_len(memidx) {
                        return Err(VMError::new("trap: out of bounds memory access (memory.init destination)"));
                    }
                    if count > 0 {
                        let bytes =
                            self.data_segments[data_idx as usize][src..src + count].to_vec();
                        self.write_memory_bytes(memidx, dst, &bytes)?;
                    }
                }
                // ── reference-types: table operations ─────────────────
                // Each op reads a `u8 table_idx` operand per spec. Tables
                // route through `table_ref`/`table_mut` so the multi-table
                // proposal works: index 0 maps to `func_table`, indexes
                // indexed directly in `wasm_tables`.
                Op::TABLE_SIZE => {
                    let tidx = self.read_u16() as usize;
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
                Op::TABLE_GROW => {
                    let tidx = self.read_u16() as usize;
                    let is64 = self.tbl_is_64(tidx);
                    let delta = self.pop_table_count(is64);
                    let init = self.pop();
                    // WASM spec: growing past the declared max fails, returning -1
                    // (as the index type) without resizing.
                    let max = self.wasm_table_maxes.get(tidx).copied().flatten();
                    let table = self
                        .table_mut(tidx)
                        .ok_or_else(|| VMError::new("trap: table.grow unknown table"))?;
                    let old_size = table.len();
                    let new_size = old_size.saturating_add(delta);
                    // Three ways a grow reports -1 without resizing: the
                    // DECLARED max, the INDEX-TYPE bound (a table32 cannot
                    // exceed 2^32-1 elements), and the host's own allocation
                    // limit — the spec permits refusing for any reason, and a
                    // spec-legal 2^32-1-element table is a multi-GB `Vec`.
                    let exceeds_max = max.is_some_and(|m| new_size > m)
                        || (!is64 && new_size as u64 > crate::vm::MAX_TABLE32_ELEMS)
                        || new_size as u64 > crate::vm::TABLE_ALLOC_LIMIT;
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
                Op::TABLE_FILL => {
                    let tidx = self.read_u16() as usize;
                    let is64 = self.tbl_is_64(tidx);
                    let count = self.pop_table_count(is64);
                    let value = self.pop();
                    let dst = self.pop_table_count(is64);
                    let table = self
                        .table_mut(tidx)
                        .ok_or_else(|| VMError::new("trap: table.fill unknown table"))?;
                    let end = dst.saturating_add(count);
                    if end > table.len() {
                        return Err(VMError::new("trap: out of bounds table access (table.fill)"));
                    }
                    for i in dst..end {
                        table[i] = value.clone();
                    }
                }
                Op::TABLE_COPY => {
                    let dst_table_idx = self.read_u16() as usize;
                    let src_table_idx = self.read_u16() as usize;
                    // table64: operands are i64 if either table is 64-bit.
                    let is64 = self.tbl_is_64(dst_table_idx) || self.tbl_is_64(src_table_idx);
                    let count = self.pop_table_count(is64);
                    let src = self.pop_table_count(is64);
                    let dst = self.pop_table_count(is64);
                    let source = self
                        .table_ref(src_table_idx)
                        .ok_or_else(|| VMError::new("trap: table.copy unknown table"))?;
                    if src.saturating_add(count) > source.len() {
                        return Err(VMError::new("trap: out of bounds table access (table.copy)".to_string()));
                    }
                    let values: Vec<Value> = source[src..src + count].to_vec();
                    let destination = self
                        .table_mut(dst_table_idx)
                        .ok_or_else(|| VMError::new("trap: table.copy unknown table"))?;
                    if dst.saturating_add(count) > destination.len() {
                        return Err(VMError::new("trap: out of bounds table access (table.copy)".to_string()));
                    }
                    destination[dst..dst + count].clone_from_slice(&values);
                }
                Op::TABLE_INIT => {
                    let elem_idx = self.read_u16() as u32;
                    let table_idx = self.read_u16() as usize;
                    // ⚠ A DROPPED segment is an EMPTY one, not an error of
                    // its own. The spec drops the payload and leaves the
                    // segment in place, so the bounds check is what decides:
                    // a zero-length copy off a dropped segment SUCCEEDS, and
                    // only a non-zero one traps. Returning early here made
                    // `(array_init_data 0 0 0)` after `drop_segs` trap, which
                    // the fixture asserts must return. `MEMORY_INIT` already
                    // modelled it this way; these did not.
                    let dropped = self.dropped_elems.contains(&elem_idx);
                    let is64 = self.tbl_is_64(table_idx);
                    let count = self.pop_table_count(is64);
                    let src = self.pop_table_count(is64);
                    let dst = self.pop_table_count(is64);
                    let elems = self
                        .elem_segments
                        .get(elem_idx as usize)
                        .ok_or_else(|| VMError::new("trap: table.init: missing element segment"))?;
                    let seg_len = if dropped { 0 } else { elems.len() };
                    if src.saturating_add(count) > seg_len {
                        return Err(VMError::new("trap: out of bounds table access (table.init source)"));
                    }
                    let values: Vec<Value> = elems[src..src + count].to_vec();
                    let table = self
                        .table_mut(table_idx)
                        .ok_or_else(|| VMError::new("trap: table.init unknown table"))?;
                    if dst.saturating_add(count) > table.len() {
                        return Err(VMError::new("trap: out of bounds table access (table.init destination)"));
                    }
                    table[dst..dst + count].clone_from_slice(&values);
                }
                Op::ELEM_DROP => {
                    let elem_idx = self.read_u16() as u32;
                    self.dropped_elems.insert(elem_idx);
                }
                Op::DATA_DROP => {
                    let data_idx = self.read_u16() as u32;
                    self.dropped_data.insert(data_idx);
                }
                Op::MEMORY_COPY => {
                    let dst_mem = self.read_u16() as usize;
                    let src_mem = self.read_u16() as usize;
                    // memory64: operands are i64 if either memory is 64-bit.
                    let is64 = self.mem_is_64(dst_mem) || self.mem_is_64(src_mem);
                    let count = self.pop_mem_index(is64);
                    let src = self.pop_mem_index(is64);
                    let dst = self.pop_mem_index(is64);
                    let buf = self.read_memory_bytes(src_mem, src, count)?;
                    self.write_memory_bytes(dst_mem, dst, &buf)?;
                }
                Op::MEMORY_FILL => {
                    let memidx = self.read_u16() as usize;
                    let is64 = self.mem_is_64(memidx);
                    let count = self.pop_mem_index(is64);
                    let val = self.pop().as_i32() as u8;
                    let dst = self.pop_mem_index(is64);
                    // BOUNDS FIRST, THEN THE BUFFER. The spec traps when
                    // `dst + count` exceeds the memory and writes nothing, so
                    // the check has to precede the fill anyway — and building
                    // `count` bytes before it turns a 4294967280-byte fill into
                    // a 4GiB allocation instead of a trap. `write_memory_bytes`
                    // still checks; this one only ensures we never MATERIALIZE
                    // a buffer the memory could not hold.
                    let limit = self.mem_len(memidx);
                    if dst.saturating_add(count) > limit {
                        return Err(VMError::new(format!(
                            "trap: out of bounds memory access: addr={dst} size={count} limit={limit}"
                        )));
                    }
                    let buf = vec![val; count];
                    self.write_memory_bytes(memidx, dst, &buf)?;
                }

                // Type discrimination opcodes

                // -- Array builtins --
                Op::ARRAY_LENGTH => {
                    let arr = self.pop();
                    // `array.len` takes `(ref null array)` and traps on null.
                    if matches!(arr, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null array reference (array.len)"));
                    }
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
                Op::ARRAY_FILL => {
                    // Spec `array.fill $t`: stack `[arrayref, index, value, count]`,
                    // so popping off the top yields count, value, index, arrayref.
                    let count = self.pop().as_i32().max(0) as usize;
                    let val = self.pop();
                    let start = self.pop().as_i32().max(0) as usize;
                    let arr = self.pop();
                    if matches!(arr, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null array reference (array.fill)"));
                    }
                    // `array.fill` traps when the filled region leaves the
                    // array, exactly as `array.copy` does — clamping to the end
                    // silently wrote less than asked. Dynamic arrays stay
                    // lenient; only a stamped GC array is bounds-checked.
                    if let Value::Object(obj) = &arr {
                        if self.is_gc_array_obj(obj) {
                            let len = match &obj.lock().unwrap().kind {
                                ObjectKind::Array(a) => a.len(),
                                _ => 0,
                            };
                            if start + count > len {
                                return Err(VMError::new("trap: out of bounds array access (array.fill)"));
                            }
                        }
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
                Op::ARRAY_COPY => {
                    let len = self.pop().as_i32().max(0) as usize;
                    let src_off = self.pop().as_i32().max(0) as usize;
                    let src = self.pop();
                    let dst_off = self.pop().as_i32().max(0) as usize;
                    let dst = self.pop();
                    if matches!(src, Value::TypedNull(_)) || matches!(dst, Value::TypedNull(_)) {
                        return Err(VMError::new("trap: null array reference (array.copy)"));
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
                            return Err(VMError::new("trap: out of bounds array access (array.copy)"));
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
                Op::CONT_NEW => {
                    let func_val = self.pop();
                    if func_val.is_null_ref() {
                        return Err(VMError::new("trap: cont.new: null function reference"));
                    }
                    let state = crate::value::ContinuationState {
                        entry: func_val,
                        saved: std::sync::Mutex::new(None),
                        state: std::sync::Mutex::new(crate::value::ContinuationPhase::Ready),
                    };
                    let obj = Object {
                        properties: indexmap::IndexMap::new(),
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
                    let cg = self.continuation_globals();
                    crate::calls::attach_continuation_protocols(
                        &mut obj.properties,
                        cg,
                        entry_async,
                    );
                    self.push(Value::Object(crate::heap::alloc(obj)))?;
                }
                Op::SUSPEND => {
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
                Op::RESUME => {
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
                Op::SWITCH => {
                    // `switch $tag` — the stack-switching proposal's symmetric
                    // swap. The TAG SEARCH is what belongs to this opcode; the
                    // park-and-enter underneath it is shared with the Component
                    // Model's `thread.{suspend,yield}-then-{resume,promote}`,
                    // which have no tag at all, so it lives in
                    // `switch_to_continuation`.
                    let tag = self.read_u16();
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
                    // On any failure the handler frame must go back, or the
                    // next `switch` in this coroutine reports "no active
                    // continuation handler" for a completely unrelated reason.
                    if let Err(e) =
                        self.switch_to_continuation("switch", &target, val, Some(&current.cont))
                    {
                        self.active_continuations.push(current);
                        return Err(e);
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
                Op::CONT_BIND => {
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
                                properties: indexmap::IndexMap::new(),
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
                            let cg = self.continuation_globals();
                            crate::calls::attach_continuation_protocols(
                                &mut new_obj.properties,
                                cg,
                                entry_async,
                            );
                            // Store the bound args as an array property
                            // keyed `__bound_args`; RESUME sees this on
                            // first fire.
                            let bound = Object {
                                properties: indexmap::IndexMap::new(),
                                kind: ObjectKind::Array(args),
                                type_id: 0,
                                fields: Vec::new(),
                            };
                            new_obj.properties.insert(
                                "__bound_args".into(),
                                Value::Object(crate::heap::alloc(bound)),
                            );
                            Value::Object(crate::heap::alloc(new_obj))
                        } else {
                            return Err(VMError::new("cont.bind: not a continuation"));
                        }
                    } else if cont_val.is_null_ref() {
                        return Err(VMError::new("trap: cont.bind: null continuation"));
                    } else {
                        return Err(VMError::new("cont.bind: not a continuation"));
                    };
                    self.push(new_cont)?;
                }
                // `resume_throw $ct $tag handlers` — resume a continuation
                // by throwing an exception into it. Stack:
                // [cont, exn_value] → control transfers into the cont's
                // nearest try_table matching the throw tag.
                Op::RESUME_THROW => {
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

                Op::RESUME_THROW_REF => {
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
                        return Err(VMError::new("trap: resume_throw_ref: null exception reference"));
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
                // THREAD_SPAWN / THREAD_JOIN opcodes RETIRED 2026-08-06:
                // spawning is the `wasi:threads/thread-spawn` IMPORT
                // (ImportTarget::WasiThreadSpawn — the VM is the embedder
                // implementation), and join is helper BYTECODE futex-waiting
                // the task's status word. No thread opcodes exist.
                Op::V128_LOAD => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let mut b = [0u8; 16];
                    b.copy_from_slice(&self.read_memory_bytes(memidx, addr, 16)?);
                    self.push(Value::V128(b))?;
                }
                Op::V128_LOAD8X8_S => {
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
                Op::V128_LOAD8X8_U => {
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
                Op::V128_LOAD16X4_S => {
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
                Op::V128_LOAD16X4_U => {
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
                Op::V128_LOAD32X2_S => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..2 {
                        let v = i32::from_le_bytes(read_le(&bytes[i * 4..i * 4 + 4])) as i64;
                        out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                Op::V128_LOAD32X2_U => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let mut out = [0u8; 16];
                    for i in 0..2 {
                        let v = u32::from_le_bytes(read_le(&bytes[i * 4..i * 4 + 4])) as u64;
                        out[i * 8..i * 8 + 8].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                Op::V128_LOAD8_SPLAT => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let b = self.read_memory_bytes(memidx, addr, 1)?[0];
                    self.push(Value::V128([b; 16]))?;
                }
                Op::V128_LOAD16_SPLAT => {
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
                Op::V128_LOAD32_SPLAT => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    let v = i32::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                Op::V128_LOAD64_SPLAT => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let v = i64::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    out[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                Op::V128_STORE => {
                    let val = self.pop();
                    let (memidx, addr) = self.pop_simd_addr()?;
                    if let Value::V128(b) = val {
                        self.write_memory_bytes(memidx, addr, &b)?;
                    }
                }
                Op::V128_CONST => {
                    let mut b = [0u8; 16];
                    for i in 0..16 {
                        b[i] = self.read_byte();
                    }
                    self.push(Value::V128(b))?;
                }
                Op::V128_LOAD8_LANE => {
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
                Op::V128_LOAD16_LANE => {
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
                Op::V128_LOAD32_LANE => {
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
                Op::V128_LOAD64_LANE => {
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
                // `v128.storeN_lane memarg laneidx : [i32 v128] -> []` — the
                // ADDRESS is the deeper operand and the vector is on top, the
                // same order as `loadN_lane`. These arms popped the address
                // first, so a spec-ordered module (anything read back from a
                // conforming `.wasm`, and any correctly written `.wat`) stored
                // from the address and addressed with the vector.
                Op::V128_STORE8_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 15;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = val {
                        self.write_memory_bytes(memidx, addr, &[v[lane]])?;
                    }
                }
                Op::V128_STORE16_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 7;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = val {
                        self.write_memory_bytes(memidx, addr, &v[lane * 2..lane * 2 + 2])?;
                    }
                }
                Op::V128_STORE32_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 3;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = val {
                        self.write_memory_bytes(memidx, addr, &v[lane * 4..lane * 4 + 4])?;
                    }
                }
                Op::V128_STORE64_LANE => {
                    let (offset, memidx, memory64) = self.read_optional_simd_memarg();
                    let lane = self.read_byte() as usize & 1;
                    let val = self.pop();
                    let base = self.pop();
                    let addr = self.simd_effective_addr(base, offset, memory64)?;
                    if let Value::V128(v) = val {
                        self.write_memory_bytes(memidx, addr, &v[lane * 8..lane * 8 + 8])?;
                    }
                }
                Op::V128_LOAD32_ZERO => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 4)?;
                    let v = i32::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    out[0..4].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                Op::V128_LOAD64_ZERO => {
                    let (memidx, addr) = self.pop_simd_addr()?;
                    let bytes = self.read_memory_bytes(memidx, addr, 8)?;
                    let v = i64::from_le_bytes(read_le(&bytes));
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                // Shuffle / swizzle
                Op::I8X16_SHUFFLE => {
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
                Op::I8X16_SWIZZLE => {
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
                Op::I8X16_SPLAT => {
                    let v = self.pop().as_i32() as u8;
                    self.push(Value::V128([v; 16]))?;
                }
                Op::I16X8_SPLAT => {
                    let v = self.pop().as_i32() as i16;
                    let b = v.to_le_bytes();
                    let mut out = [0u8; 16];
                    for i in 0..8 {
                        out[i * 2..i * 2 + 2].copy_from_slice(&b);
                    }
                    self.push(Value::V128(out))?;
                }
                Op::I32X4_SPLAT => {
                    let v = self.pop().as_i32();
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        out[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
                    }
                    self.push(Value::V128(out))?;
                }
                Op::I64X2_SPLAT => {
                    let v = self.pop().as_i64();
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    out[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                Op::F32X4_SPLAT => {
                    // Lanes hold the operand's ENCODING — bit-preserving, so no
                    // f64 round-trip (it would quiet a signalling NaN).
                    let v = self.pop().as_f32();
                    let b = v.to_le_bytes();
                    let mut out = [0u8; 16];
                    for i in 0..4 {
                        out[i * 4..i * 4 + 4].copy_from_slice(&b);
                    }
                    self.push(Value::V128(out))?;
                }
                Op::F64X2_SPLAT => {
                    let v = self.pop().as_f64();
                    let mut out = [0u8; 16];
                    out[0..8].copy_from_slice(&v.to_le_bytes());
                    out[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(out))?;
                }
                // Extract / replace lane
                Op::I8X16_EXTRACT_LANE_S => {
                    let l = self.read_byte() as usize & 15;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(a[l] as i8 as i32))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                Op::I8X16_EXTRACT_LANE_U => {
                    let l = self.read_byte() as usize & 15;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(a[l] as i32))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                Op::I8X16_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 15;
                    let v = self.pop().as_i32() as u8;
                    if let Value::V128(mut a) = self.pop() {
                        a[l] = v;
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                Op::I16X8_EXTRACT_LANE_S => {
                    let l = self.read_byte() as usize & 7;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(
                            i16::from_le_bytes([a[l * 2], a[l * 2 + 1]]) as i32
                        ))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                Op::I16X8_EXTRACT_LANE_U => {
                    let l = self.read_byte() as usize & 7;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(
                            u16::from_le_bytes([a[l * 2], a[l * 2 + 1]]) as i32
                        ))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                Op::I16X8_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 7;
                    let v = self.pop().as_i32() as i16;
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 2..l * 2 + 2].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                Op::I32X4_EXTRACT_LANE => {
                    let l = self.read_byte() as usize & 3;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(i32::from_le_bytes(
                            a[l * 4..l * 4 + 4].try_into().unwrap(),
                        )))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                Op::I32X4_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 3;
                    let v = self.pop().as_i32();
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 4..l * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                Op::I64X2_EXTRACT_LANE => {
                    let l = self.read_byte() as usize & 1;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I64(i64::from_le_bytes(
                            a[l * 8..l * 8 + 8].try_into().unwrap(),
                        )))?;
                    } else {
                        self.push(Value::I64(0))?;
                    }
                }
                Op::I64X2_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 1;
                    let v = self.pop().as_i64();
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 8..l * 8 + 8].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                Op::F32X4_EXTRACT_LANE => {
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
                Op::F32X4_REPLACE_LANE => {
                    let l = self.read_byte() as usize & 3;
                    // Bit-preserving into a lane — see `F32X4_SPLAT`.
                    let v = self.pop().as_f32();
                    if let Value::V128(mut a) = self.pop() {
                        a[l * 4..l * 4 + 4].copy_from_slice(&v.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else {
                        self.push(Value::V128([0; 16]))?;
                    }
                }
                Op::F64X2_EXTRACT_LANE => {
                    let l = self.read_byte() as usize & 1;
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::F64(f64::from_le_bytes(
                            a[l * 8..l * 8 + 8].try_into().unwrap(),
                        )))?;
                    } else {
                        self.push(Value::F64(0.0))?;
                    }
                }
                Op::F64X2_REPLACE_LANE => {
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
                Op::I8X16_EQ => {
                    self.simd_i8x16_binop(|a, b| if a == b { 0xFF } else { 0 })?;
                }
                Op::I8X16_NE => {
                    self.simd_i8x16_binop(|a, b| if a != b { 0xFF } else { 0 })?;
                }
                Op::I8X16_LT_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) < (b as i8) { 0xFF } else { 0 })?;
                }
                Op::I8X16_LT_U => {
                    self.simd_i8x16_binop(|a, b| if a < b { 0xFF } else { 0 })?;
                }
                Op::I8X16_GT_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) > (b as i8) { 0xFF } else { 0 })?;
                }
                Op::I8X16_GT_U => {
                    self.simd_i8x16_binop(|a, b| if a > b { 0xFF } else { 0 })?;
                }
                Op::I8X16_LE_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) <= (b as i8) { 0xFF } else { 0 })?;
                }
                Op::I8X16_LE_U => {
                    self.simd_i8x16_binop(|a, b| if a <= b { 0xFF } else { 0 })?;
                }
                Op::I8X16_GE_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) >= (b as i8) { 0xFF } else { 0 })?;
                }
                Op::I8X16_GE_U => {
                    self.simd_i8x16_binop(|a, b| if a >= b { 0xFF } else { 0 })?;
                }
                // i16x8 comparisons
                Op::I16X8_EQ => {
                    self.simd_i16x8_binop(|a, b| if a == b { -1 } else { 0 })?;
                }
                Op::I16X8_NE => {
                    self.simd_i16x8_binop(|a, b| if a != b { -1 } else { 0 })?;
                }
                Op::I16X8_LT_S => {
                    self.simd_i16x8_binop(|a, b| if a < b { -1 } else { 0 })?;
                }
                Op::I16X8_LT_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) < (b as u16) { -1 } else { 0 })?;
                }
                Op::I16X8_GT_S => {
                    self.simd_i16x8_binop(|a, b| if a > b { -1 } else { 0 })?;
                }
                Op::I16X8_GT_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) > (b as u16) { -1 } else { 0 })?;
                }
                Op::I16X8_LE_S => {
                    self.simd_i16x8_binop(|a, b| if a <= b { -1 } else { 0 })?;
                }
                Op::I16X8_LE_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) <= (b as u16) { -1 } else { 0 })?;
                }
                Op::I16X8_GE_S => {
                    self.simd_i16x8_binop(|a, b| if a >= b { -1 } else { 0 })?;
                }
                Op::I16X8_GE_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) >= (b as u16) { -1 } else { 0 })?;
                }
                // i32x4 comparisons
                Op::I32X4_EQ => {
                    self.simd_i32x4_binop(|a, b| if a == b { -1 } else { 0 })?;
                }
                Op::I32X4_NE => {
                    self.simd_i32x4_binop(|a, b| if a != b { -1 } else { 0 })?;
                }
                Op::I32X4_LT_S => {
                    self.simd_i32x4_binop(|a, b| if a < b { -1 } else { 0 })?;
                }
                Op::I32X4_LT_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) < (b as u32) { -1 } else { 0 })?;
                }
                Op::I32X4_GT_S => {
                    self.simd_i32x4_binop(|a, b| if a > b { -1 } else { 0 })?;
                }
                Op::I32X4_GT_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) > (b as u32) { -1 } else { 0 })?;
                }
                Op::I32X4_LE_S => {
                    self.simd_i32x4_binop(|a, b| if a <= b { -1 } else { 0 })?;
                }
                Op::I32X4_LE_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) <= (b as u32) { -1 } else { 0 })?;
                }
                Op::I32X4_GE_S => {
                    self.simd_i32x4_binop(|a, b| if a >= b { -1 } else { 0 })?;
                }
                Op::I32X4_GE_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) >= (b as u32) { -1 } else { 0 })?;
                }
                // f32x4 comparisons
                Op::F32X4_EQ => {
                    self.simd_f32x4_cmp(|a, b| a == b)?;
                }
                Op::F32X4_NE => {
                    self.simd_f32x4_cmp(|a, b| a != b)?;
                }
                Op::F32X4_LT => {
                    self.simd_f32x4_cmp(|a, b| a < b)?;
                }
                Op::F32X4_GT => {
                    self.simd_f32x4_cmp(|a, b| a > b)?;
                }
                Op::F32X4_LE => {
                    self.simd_f32x4_cmp(|a, b| a <= b)?;
                }
                Op::F32X4_GE => {
                    self.simd_f32x4_cmp(|a, b| a >= b)?;
                }
                // f64x2 comparisons
                Op::F64X2_EQ => {
                    self.simd_f64x2_cmp(|a, b| a == b)?;
                }
                Op::F64X2_NE => {
                    self.simd_f64x2_cmp(|a, b| a != b)?;
                }
                Op::F64X2_LT => {
                    self.simd_f64x2_cmp(|a, b| a < b)?;
                }
                Op::F64X2_GT => {
                    self.simd_f64x2_cmp(|a, b| a > b)?;
                }
                Op::F64X2_LE => {
                    self.simd_f64x2_cmp(|a, b| a <= b)?;
                }
                Op::F64X2_GE => {
                    self.simd_f64x2_cmp(|a, b| a >= b)?;
                }
                // v128 bitwise
                Op::V128_NOT => {
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
                Op::V128_AND => {
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
                Op::V128_ANDNOT => {
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
                Op::V128_OR => {
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
                Op::V128_XOR => {
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
                Op::V128_BITSELECT => {
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
                Op::V128_ANY_TRUE => {
                    if let Value::V128(a) = self.pop() {
                        self.push(Value::I32(if a.iter().any(|&b| b != 0) { 1 } else { 0 }))?;
                    } else {
                        self.push(Value::I32(0))?;
                    }
                }
                // Promote / demote
                Op::F32X4_DEMOTE_F64X2_ZERO => {
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
                Op::F64X2_PROMOTE_LOW_F32X4 => {
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
                Op::I8X16_ABS => {
                    self.simd_i8x16_unop(|a| (a as i8).unsigned_abs())?;
                }
                Op::I8X16_NEG => {
                    self.simd_i8x16_unop(|a| (a as i8).wrapping_neg() as u8)?;
                }
                Op::I8X16_POPCNT => {
                    self.simd_i8x16_unop(|a| a.count_ones() as u8)?;
                }
                Op::I8X16_ALL_TRUE => {
                    self.simd_i8x16_testop(|a| a != 0)?;
                }
                Op::I8X16_BITMASK => {
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
                Op::I8X16_NARROW_I16X8_S => {
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
                Op::I8X16_NARROW_I16X8_U => {
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
                Op::F32X4_CEIL => {
                    self.simd_f32x4_unop(|a| a.ceil())?;
                }
                Op::F32X4_FLOOR => {
                    self.simd_f32x4_unop(|a| a.floor())?;
                }
                Op::F32X4_TRUNC => {
                    self.simd_f32x4_unop(|a| a.trunc())?;
                }
                Op::F32X4_NEAREST => {
                    self.simd_f32x4_unop(|a| a.round_ties_even())?;
                }
                // i8x16 shifts
                Op::I8X16_SHL => {
                    let sh = self.pop().as_i32() as u32 & 7;
                    self.simd_i8x16_unop(|a| a.wrapping_shl(sh))?;
                }
                Op::I8X16_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 7;
                    self.simd_i8x16_unop(|a| ((a as i8).wrapping_shr(sh)) as u8)?;
                }
                Op::I8X16_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 7;
                    self.simd_i8x16_unop(|a| a.wrapping_shr(sh))?;
                }
                // i8x16 arithmetic
                Op::I8X16_ADD => {
                    self.simd_i8x16_binop(|a, b| a.wrapping_add(b))?;
                }
                Op::I8X16_ADD_SAT_S => {
                    self.simd_i8x16_binop(|a, b| ((a as i8).saturating_add(b as i8)) as u8)?;
                }
                Op::I8X16_ADD_SAT_U => {
                    self.simd_i8x16_binop(|a, b| a.saturating_add(b))?;
                }
                Op::I8X16_SUB => {
                    self.simd_i8x16_binop(|a, b| a.wrapping_sub(b))?;
                }
                Op::I8X16_SUB_SAT_S => {
                    self.simd_i8x16_binop(|a, b| ((a as i8).saturating_sub(b as i8)) as u8)?;
                }
                Op::I8X16_SUB_SAT_U => {
                    self.simd_i8x16_binop(|a, b| a.saturating_sub(b))?;
                }
                Op::I8X16_MIN_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) < (b as i8) { a } else { b })?;
                }
                Op::I8X16_MIN_U => {
                    self.simd_i8x16_binop(|a, b| a.min(b))?;
                }
                Op::I8X16_MAX_S => {
                    self.simd_i8x16_binop(|a, b| if (a as i8) > (b as i8) { a } else { b })?;
                }
                Op::I8X16_MAX_U => {
                    self.simd_i8x16_binop(|a, b| a.max(b))?;
                }
                Op::I8X16_AVGR_U => {
                    self.simd_i8x16_binop(|a, b| ((a as u16 + b as u16 + 1) / 2) as u8)?;
                }
                // f64x2 unary
                Op::F64X2_CEIL => {
                    self.simd_f64x2_unop(|a| a.ceil())?;
                }
                Op::F64X2_FLOOR => {
                    self.simd_f64x2_unop(|a| a.floor())?;
                }
                Op::F64X2_TRUNC => {
                    self.simd_f64x2_unop(|a| a.trunc())?;
                }
                Op::F64X2_NEAREST => {
                    self.simd_f64x2_unop(|a| a.round_ties_even())?;
                }
                // extadd pairwise
                Op::I16X8_EXTADD_PAIRWISE_I8X16_S => {
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
                Op::I16X8_EXTADD_PAIRWISE_I8X16_U => {
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
                Op::I32X4_EXTADD_PAIRWISE_I16X8_S => {
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
                Op::I32X4_EXTADD_PAIRWISE_I16X8_U => {
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
                Op::I16X8_ABS => {
                    self.simd_i16x8_unop(|a| a.unsigned_abs() as i16)?;
                }
                Op::I16X8_NEG => {
                    self.simd_i16x8_unop(|a| a.wrapping_neg())?;
                }
                Op::I16X8_Q15MULR_SAT_S => {
                    self.simd_i16x8_binop(|a, b| {
                        let r = (a as i32 * b as i32 + 0x4000) >> 15;
                        r.clamp(i16::MIN as i32, i16::MAX as i32) as i16
                    })?;
                }
                Op::I16X8_ALL_TRUE => {
                    self.simd_i16x8_testop(|a| a != 0)?;
                }
                Op::I16X8_BITMASK => {
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
                Op::I16X8_NARROW_I32X4_S => {
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
                Op::I16X8_NARROW_I32X4_U => {
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
                Op::I16X8_EXTEND_LOW_I8X16_S => {
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
                Op::I16X8_EXTEND_HIGH_I8X16_S => {
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
                Op::I16X8_EXTEND_LOW_I8X16_U => {
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
                Op::I16X8_EXTEND_HIGH_I8X16_U => {
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
                Op::I16X8_SHL => {
                    let sh = self.pop().as_i32() as u32 & 15;
                    self.simd_i16x8_unop(|a| a.wrapping_shl(sh))?;
                }
                Op::I16X8_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 15;
                    self.simd_i16x8_unop(|a| a.wrapping_shr(sh))?;
                }
                Op::I16X8_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 15;
                    self.simd_i16x8_unop(|a| (a as u16).wrapping_shr(sh) as i16)?;
                }
                Op::I16X8_ADD => {
                    self.simd_i16x8_binop(|a, b| a.wrapping_add(b))?;
                }
                Op::I16X8_ADD_SAT_S => {
                    self.simd_i16x8_binop(|a, b| a.saturating_add(b))?;
                }
                Op::I16X8_ADD_SAT_U => {
                    self.simd_i16x8_binop(|a, b| ((a as u16).saturating_add(b as u16)) as i16)?;
                }
                Op::I16X8_SUB => {
                    self.simd_i16x8_binop(|a, b| a.wrapping_sub(b))?;
                }
                Op::I16X8_SUB_SAT_S => {
                    self.simd_i16x8_binop(|a, b| a.saturating_sub(b))?;
                }
                Op::I16X8_SUB_SAT_U => {
                    self.simd_i16x8_binop(|a, b| ((a as u16).saturating_sub(b as u16)) as i16)?;
                }
                Op::I16X8_MUL => {
                    self.simd_i16x8_binop(|a, b| a.wrapping_mul(b))?;
                }
                Op::I16X8_MIN_S => {
                    self.simd_i16x8_binop(|a, b| a.min(b))?;
                }
                Op::I16X8_MIN_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) < (b as u16) { a } else { b })?;
                }
                Op::I16X8_MAX_S => {
                    self.simd_i16x8_binop(|a, b| a.max(b))?;
                }
                Op::I16X8_MAX_U => {
                    self.simd_i16x8_binop(|a, b| if (a as u16) > (b as u16) { a } else { b })?;
                }
                Op::I16X8_AVGR_U => {
                    self.simd_i16x8_binop(|a, b| {
                        (((a as u16 as u32) + (b as u16 as u32) + 1) / 2) as i16
                    })?;
                }
                Op::I16X8_EXTMUL_LOW_I8X16_S => {
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
                Op::I16X8_EXTMUL_HIGH_I8X16_S => {
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
                Op::I16X8_EXTMUL_LOW_I8X16_U => {
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
                Op::I16X8_EXTMUL_HIGH_I8X16_U => {
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
                Op::I32X4_ABS => {
                    self.simd_i32x4_unop(|a| a.unsigned_abs() as i32)?;
                }
                Op::I32X4_NEG => {
                    self.simd_i32x4_unop(|a| a.wrapping_neg())?;
                }
                Op::I32X4_ALL_TRUE => {
                    self.simd_i32x4_testop(|a| a != 0)?;
                }
                Op::I32X4_BITMASK => {
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
                Op::I32X4_EXTEND_LOW_I16X8_S => {
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
                Op::I32X4_EXTEND_HIGH_I16X8_S => {
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
                Op::I32X4_EXTEND_LOW_I16X8_U => {
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
                Op::I32X4_EXTEND_HIGH_I16X8_U => {
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
                Op::I32X4_SHL => {
                    let sh = self.pop().as_i32() as u32 & 31;
                    self.simd_i32x4_unop(|a| a.wrapping_shl(sh))?;
                }
                Op::I32X4_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 31;
                    self.simd_i32x4_unop(|a| a.wrapping_shr(sh))?;
                }
                Op::I32X4_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 31;
                    self.simd_i32x4_unop(|a| (a as u32).wrapping_shr(sh) as i32)?;
                }
                Op::I32X4_ADD => {
                    self.simd_i32x4_binop(|a, b| a.wrapping_add(b))?;
                }
                Op::I32X4_SUB => {
                    self.simd_i32x4_binop(|a, b| a.wrapping_sub(b))?;
                }
                Op::I32X4_MUL => {
                    self.simd_i32x4_binop(|a, b| a.wrapping_mul(b))?;
                }
                Op::I32X4_MIN_S => {
                    self.simd_i32x4_binop(|a, b| a.min(b))?;
                }
                Op::I32X4_MIN_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) < (b as u32) { a } else { b })?;
                }
                Op::I32X4_MAX_S => {
                    self.simd_i32x4_binop(|a, b| a.max(b))?;
                }
                Op::I32X4_MAX_U => {
                    self.simd_i32x4_binop(|a, b| if (a as u32) > (b as u32) { a } else { b })?;
                }
                Op::I32X4_DOT_I16X8_S => {
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
                Op::I32X4_EXTMUL_LOW_I16X8_S => {
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
                Op::I32X4_EXTMUL_HIGH_I16X8_S => {
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
                Op::I32X4_EXTMUL_LOW_I16X8_U => {
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
                Op::I32X4_EXTMUL_HIGH_I16X8_U => {
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
                Op::I64X2_ABS => {
                    self.simd_i64x2_unop(|a| a.unsigned_abs() as i64)?;
                }
                Op::I64X2_NEG => {
                    self.simd_i64x2_unop(|a| a.wrapping_neg())?;
                }
                Op::I64X2_ALL_TRUE => {
                    self.simd_i64x2_testop(|a| a != 0)?;
                }
                Op::I64X2_BITMASK => {
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
                Op::I64X2_EXTEND_LOW_I32X4_S => {
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
                Op::I64X2_EXTEND_HIGH_I32X4_S => {
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
                Op::I64X2_EXTEND_LOW_I32X4_U => {
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
                Op::I64X2_EXTEND_HIGH_I32X4_U => {
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
                Op::I64X2_SHL => {
                    let sh = self.pop().as_i32() as u32 & 63;
                    self.simd_i64x2_unop(|a| a.wrapping_shl(sh))?;
                }
                Op::I64X2_SHR_S => {
                    let sh = self.pop().as_i32() as u32 & 63;
                    self.simd_i64x2_unop(|a| a.wrapping_shr(sh))?;
                }
                Op::I64X2_SHR_U => {
                    let sh = self.pop().as_i32() as u32 & 63;
                    self.simd_i64x2_unop(|a| (a as u64).wrapping_shr(sh) as i64)?;
                }
                Op::I64X2_ADD => {
                    self.simd_i64x2_binop(|a, b| a.wrapping_add(b))?;
                }
                Op::I64X2_SUB => {
                    self.simd_i64x2_binop(|a, b| a.wrapping_sub(b))?;
                }
                Op::I64X2_MUL => {
                    self.simd_i64x2_binop(|a, b| a.wrapping_mul(b))?;
                }
                Op::I64X2_EQ => {
                    self.simd_i64x2_cmp(|a, b| a == b)?;
                }
                Op::I64X2_NE => {
                    self.simd_i64x2_cmp(|a, b| a != b)?;
                }
                Op::I64X2_LT_S => {
                    self.simd_i64x2_cmp(|a, b| a < b)?;
                }
                Op::I64X2_GT_S => {
                    self.simd_i64x2_cmp(|a, b| a > b)?;
                }
                Op::I64X2_LE_S => {
                    self.simd_i64x2_cmp(|a, b| a <= b)?;
                }
                Op::I64X2_GE_S => {
                    self.simd_i64x2_cmp(|a, b| a >= b)?;
                }
                Op::I64X2_EXTMUL_LOW_I32X4_S => {
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
                Op::I64X2_EXTMUL_HIGH_I32X4_S => {
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
                Op::I64X2_EXTMUL_LOW_I32X4_U => {
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
                Op::I64X2_EXTMUL_HIGH_I32X4_U => {
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
                Op::F32X4_ABS => {
                    self.simd_f32x4_unop(|a| a.abs())?;
                }
                Op::F32X4_NEG => {
                    self.simd_f32x4_unop(|a| -a)?;
                }
                Op::F32X4_SQRT => {
                    self.simd_f32x4_unop(|a| a.sqrt())?;
                }
                Op::F32X4_ADD => {
                    self.simd_f32x4_binop(|a, b| a + b)?;
                }
                Op::F32X4_SUB => {
                    self.simd_f32x4_binop(|a, b| a - b)?;
                }
                Op::F32X4_MUL => {
                    self.simd_f32x4_binop(|a, b| a * b)?;
                }
                Op::F32X4_DIV => {
                    self.simd_f32x4_binop(|a, b| a / b)?;
                }
                Op::F32X4_MIN => {
                    self.simd_f32x4_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f32::NAN
                        } else {
                            a.min(b)
                        }
                    })?;
                }
                Op::F32X4_MAX => {
                    self.simd_f32x4_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f32::NAN
                        } else {
                            a.max(b)
                        }
                    })?;
                }
                Op::F32X4_PMIN => {
                    self.simd_f32x4_binop(|a, b| if b < a { b } else { a })?;
                }
                Op::F32X4_PMAX => {
                    self.simd_f32x4_binop(|a, b| if a < b { b } else { a })?;
                }
                // f64x2
                Op::F64X2_ABS => {
                    self.simd_f64x2_unop(|a| a.abs())?;
                }
                Op::F64X2_NEG => {
                    self.simd_f64x2_unop(|a| -a)?;
                }
                Op::F64X2_SQRT => {
                    self.simd_f64x2_unop(|a| a.sqrt())?;
                }
                Op::F64X2_ADD => {
                    self.simd_f64x2_binop(|a, b| a + b)?;
                }
                Op::F64X2_SUB => {
                    self.simd_f64x2_binop(|a, b| a - b)?;
                }
                Op::F64X2_MUL => {
                    self.simd_f64x2_binop(|a, b| a * b)?;
                }
                Op::F64X2_DIV => {
                    self.simd_f64x2_binop(|a, b| a / b)?;
                }
                Op::F64X2_MIN => {
                    self.simd_f64x2_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f64::NAN
                        } else {
                            a.min(b)
                        }
                    })?;
                }
                Op::F64X2_MAX => {
                    self.simd_f64x2_binop(|a, b| {
                        if a.is_nan() || b.is_nan() {
                            f64::NAN
                        } else {
                            a.max(b)
                        }
                    })?;
                }
                Op::F64X2_PMIN => {
                    self.simd_f64x2_binop(|a, b| if b < a { b } else { a })?;
                }
                Op::F64X2_PMAX => {
                    self.simd_f64x2_binop(|a, b| if a < b { b } else { a })?;
                }
                // Conversions
                Op::I32X4_TRUNC_SAT_F32X4_S => {
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
                Op::I32X4_TRUNC_SAT_F32X4_U => {
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
                Op::F32X4_CONVERT_I32X4_S => {
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
                Op::F32X4_CONVERT_I32X4_U => {
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
                Op::I32X4_TRUNC_SAT_F64X2_S_ZERO => {
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
                Op::I32X4_TRUNC_SAT_F64X2_U_ZERO => {
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
                Op::F64X2_CONVERT_LOW_I32X4_S => {
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
                Op::F64X2_CONVERT_LOW_I32X4_U => {
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
                Op::I8X16_RELAXED_SWIZZLE => {
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
                Op::I32X4_RELAXED_TRUNC_F32X4_S => {
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
                Op::I32X4_RELAXED_TRUNC_F32X4_U => {
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
                Op::I32X4_RELAXED_TRUNC_F64X2_S_ZERO => {
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
                Op::I32X4_RELAXED_TRUNC_F64X2_U_ZERO => {
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
                Op::F32X4_RELAXED_MADD => {
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
                Op::F32X4_RELAXED_NMADD => {
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
                Op::F64X2_RELAXED_MADD => {
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
                Op::F64X2_RELAXED_NMADD => {
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
                Op::I8X16_RELAXED_LANESELECT => {
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
                Op::I16X8_RELAXED_LANESELECT => {
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
                Op::I32X4_RELAXED_LANESELECT => {
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
                Op::I64X2_RELAXED_LANESELECT => {
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
                Op::F32X4_RELAXED_MIN => {
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
                Op::F32X4_RELAXED_MAX => {
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
                Op::F64X2_RELAXED_MIN => {
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
                Op::F64X2_RELAXED_MAX => {
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
                Op::I16X8_RELAXED_Q15MULR_S => {
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
                Op::I16X8_RELAXED_DOT_I8X16_I7X16_S => {
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
                Op::I32X4_RELAXED_DOT_I8X16_I7X16_ADD_S => {
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

                // -- CM3 canon built-ins (stream/future/task/waitable/
                // backpressure/context) -- no opcodes; they are "canon"-module
                // imports executed by exec_canon_builtin via the CALL arm.

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

/// The bump-allocator global that compiler-emitted canonical marshalling uses
/// (`vybe_compiler::primitives::canon_marshal`).
///
/// Host-side lowering shares it deliberately: two independent bump pointers
/// over one linear memory would eventually hand the same address to a guest
/// string and a host-stored one.
pub(crate) fn canon_marshal_bump() -> &'static str {
    "__vybe_chan_futex_next"
}

#[cfg(test)]
mod thread_block_tests {
    use super::*;
    use crate::vm::VM;

    /// The two non-switching outcomes of `block(switch_to = None)` must be
    /// DISTINGUISHABLE, and both must be reachable.
    ///
    /// The whole reason `thread_block` splits three ways is that a thread
    /// blocking while host work is outstanding is *not* deadlocked — it is
    /// waiting, and the machinery to park it properly is the gap. Report those
    /// two as one message and the gap reads as a program bug forever after.
    ///
    /// If this test ever finds the two messages equal, or finds the host-work
    /// branch unreachable, the three-way split has collapsed into a two-way one
    /// with dead code — which is precisely what it was introduced to avoid.
    #[test]
    fn a_deadlock_and_a_pending_host_wait_are_different_answers() {
        let mut vm = VM::new();
        // No threads registered, nothing queued: nothing can ever wake us.
        let deadlock = vm
            .thread_block("thread.suspend", 0)
            .expect_err("blocking with no waker must not succeed")
            .message;
        assert!(
            deadlock.contains("no host work is pending"),
            "the deadlock message must say WHY it is a deadlock: {deadlock}"
        );

        // One queued job is enough to make this a wait rather than a deadlock.
        vm.event_loop
            .borrow_mut()
            .immediate
            .push_back(crate::event_loop::Task::Callback {
                callback: Value::Undefined,
                value: Value::Undefined,
            });
        let waiting = vm
            .thread_block("thread.suspend", 0)
            .expect_err("the fiber-suspension path is not implemented yet")
            .message;
        assert!(
            waiting.contains("fiber"),
            "a legitimate block must name the MISSING MACHINERY, not blame the guest: {waiting}"
        );
        assert_ne!(
            deadlock, waiting,
            "collapsing these two would report a deadlock for a thread that is merely waiting"
        );
    }

    /// A READY thread is a switch target, so `thread_block` must not reach
    /// either error path — it hands control over instead. Asserted by the
    /// ABSENCE of the deadlock message, since actually switching needs a live
    /// fiber this test has no way to mint.
    #[test]
    fn a_ready_thread_is_preferred_over_reporting_a_deadlock() {
        let mut vm = VM::new();
        let cont = vm.new_parked_continuation();
        let mut t = crate::cm_thread::Thread::new(0, cont);
        t.resume_later().expect("a fresh thread is suspended");
        assert!(t.ready(), "resume_later must leave it READY or this proves nothing");
        let idx = vm.cm_instance.threads.register(t);

        let err = vm.thread_block("thread.suspend", idx + 1).unwrap_err().message;
        assert!(
            !err.contains("no host work is pending"),
            "a ready thread must be switched to, never reported as a deadlock: {err}"
        );
    }
}
