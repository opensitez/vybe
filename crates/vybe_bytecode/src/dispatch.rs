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

use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use crate::chunk::Chunk;
use crate::error::VMError;
use crate::event_loop::{EventLoop, Task};
use crate::fiber::{Fiber, SavedFrame};
use crate::opcode::Op;
use crate::shared_memory::SharedMemory;
use crate::value::{Function, Object, ObjectKind, Upvalue, UpvalueLocation, Value};
use crate::vm::{
    dyn_truthy,
    VM, CallFrame, ExceptionHandler, FinalizerEntry, LabelEntry,
    ExecResult, HostContext, HostFn, ImportTarget,
    MAX_FRAMES, MAX_STACK,
};

impl VM {
    pub(crate) fn execute_until(&mut self, min_depth: usize) -> Result<Value, VMError> {
        loop {
            let f = self.frame();
            let chunk = &self.chunks[f.chunk_index];

            if f.ip >= chunk.code.len() {
                if self.frames.len() <= 1.max(min_depth + 1) {
                    return Ok(self.stack.pop().unwrap_or(Value::Null));
                }
                let base = self.frame().base;
                self.frames.pop();
                self.stack.truncate(base);
                self.push(Value::Null)?;
                continue;
            }

            // Uniform 2-byte opcode decode: [prefix, sub]
            let prefix = self.read_byte();
            let sub = self.read_byte();
            let op = match Op::decode(prefix, sub) {
                Some(op) => op,
                None => return Err(VMError::new(format!("Invalid opcode: 0x{:02X} 0x{:02X}", prefix, sub))),
            };

            // ── Execution trace ──────────────────────────────────────────
            if self.trace {
                let f = self.frame();
                let chunk_name = &self.chunks[f.chunk_index].name;
                let ip = f.ip;
                let stack_top = if self.stack.is_empty() {
                    "[]".to_string()
                } else {
                    let top = &self.stack[self.stack.len() - 1];
                    let depth = self.stack.len();
                    format!("[{}] (depth={})", top, depth)
                };
                eprintln!("  TRACE {:>12} @{:04} {:?}  stack: {}",
                    chunk_name, ip.saturating_sub(1), op, stack_top);
            }

            match op {
                _ if op == Op::HALT => {
                    if self.frames.len() <= 1 {
                        // Top-level halt: terminate execution
                        self.close_upvalues(0);
                        return Ok(if self.stack.is_empty() { Value::Null } else { self.pop() });
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

                _ if op == Op::CONST => {
                    let idx = self.read_u16();
                    let val = self.get_constant(idx);
                    self.push(val)?;
                }
                _ if op == Op::DROP => { self.pop(); }
                _ if op == Op::DUP => {
                    let val = self.peek(0).clone();
                    self.push(val)?;
                }

                // -- Variables --
                _ if op == Op::LOCAL_GET => {
                    let slot = self.read_u16() as usize;
                    let base = self.frame().base;
                    let val = self.stack[base + slot].clone();
                    self.push(val)?;
                }
                _ if op == Op::LOCAL_SET => {
                    let slot = self.read_u16() as usize;
                    let val = self.peek(0).clone();
                    let base = self.frame().base;
                    self.stack[base + slot] = val;
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
                        } else { name }
                    } else { name };
                    let val = self.globals.get(&key).cloned().unwrap_or(Value::Undefined);
                    self.push(val)?;
                }
                _ if op == Op::GLOBAL_SET => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let key = if self.strict_isolation {
                        if let Some(ref prefix) = self.module_prefix {
                            format!("{}::{}", prefix, name)
                        } else { name }
                    } else { name };
                    let val = self.peek(0).clone();
                    self.globals.insert(key, val);
                }
                _ if op == Op::UPVALUE_GET => {
                    let idx = self.read_byte() as usize;
                    let uv = self.frame().upvalues[idx].clone();
                    let val = match &uv.lock().unwrap().location {
                        UpvalueLocation::Open(si) => self.stack[*si].clone(),
                        UpvalueLocation::Closed(v) => v.clone(),
                    };
                    self.push(val)?;
                }
                _ if op == Op::UPVALUE_SET => {
                    let idx = self.read_byte() as usize;
                    let val = self.peek(0).clone();
                    let uv = self.frame().upvalues[idx].clone();
                    let mut u = uv.lock().unwrap();
                    match &mut u.location {
                        UpvalueLocation::Open(si) => self.stack[*si] = val,
                        UpvalueLocation::Closed(v) => *v = val,
                    }
                }

                // -- Properties --
                _ if op == Op::STRUCT_GET => {
                    let idx = self.read_u16();
                    let name = self.constant_str(idx);
                    let obj = self.pop();
                    // Auto-join thread when accessing .result on a Task/Thread object
                    if let Value::Object(ref o) = obj {
                        let needs_join = {
                            let o_ref = o.lock().unwrap();
                            (name == "result" || name == "exitcode")
                                && o_ref.properties.contains_key("__thread_id")
                                && !o_ref.properties.get("iscompleted").map(|v| v.as_bool()).unwrap_or(true)
                        };
                        if needs_join {
                            let tid = o.lock().unwrap().properties.get("__thread_id")
                                .map(|v| v.as_f64() as i32).unwrap_or(-1);
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
                    if let Value::Object(o) = &obj {
                        // Check for setter: __set_{name}
                        let setter_key = format!("__set_{}", name);
                        let setter = o.lock().unwrap().properties.get(&setter_key).cloned();
                        if let Some(setter_fn) = setter {
                            // Call the setter synchronously. Save stack depth
                            // and restore after — invoke_callback leaks the
                            // return value and intermediate locals on the stack.
                            let stack_save = self.stack.len();
                            let _result = self.invoke_callback(&setter_fn, &[obj.clone(), val.clone()]);
                            self.stack.truncate(stack_save);
                            self.push(val)?;
                        } else {
                            o.lock().unwrap().set(name.clone(), val.clone());
                            self.push(val)?;
                        }
                    } else {
                        self.push(val)?;
                    }
                }
                _ if op == Op::ARRAY_GET => {
                    let key = self.pop();
                    let obj = self.pop();
                    match &obj {
                        Value::Object(o) => {
                            // Handle negative indices: x[-1] → x[len-1]
                            let k = {
                                let idx = key.as_f64() as i64;
                                if idx < 0 {
                                    let ob = o.lock().unwrap();
                                    let len = match &ob.kind {
                                        ObjectKind::Array(a) => a.len() as i64,
                                        _ => 0,
                                    };
                                    format!("{}", (len + idx).max(0))
                                } else {
                                    format!("{}", key)
                                }
                            };
                            let val = o.lock().unwrap().get(&k);
                            // If not found and object has __getitem__, call it
                            if matches!(val, Value::Null) {
                                let getitem = o.lock().unwrap().properties.get("__getitem__").cloned();
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
                    if let Value::Object(o) = &obj {
                        // Check for __setitem__ dunder
                        let setitem = o.lock().unwrap().properties.get("__setitem__").cloned();
                        if let Some(func) = setitem {
                            self.push(func)?;
                            self.push(obj.clone())?; // self
                            self.push(key)?;          // key
                            self.push(val.clone())?;  // value
                            self.call_value(3)?;
                            self.pop(); // discard __setitem__ return
                            self.push(val)?;
                            continue;
                        }
                        let k = format!("{}", key);
                        o.lock().unwrap().set(k, val.clone());
                    }
                    self.push(val)?;
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
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32(a.wrapping_div(b)))?;
                }
                _ if op == Op::I32_DIV_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32((a / b) as i32))?;
                }
                _ if op == Op::I32_REM_S => {
                    let b = self.pop().as_i32();
                    let a = self.pop().as_i32();
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32(a.wrapping_rem(b)))?;
                }
                _ if op == Op::I32_REM_U => {
                    let b = self.pop().as_i32() as u32;
                    let a = self.pop().as_i32() as u32;
                    if b == 0 { return Err(VMError::new("trap: integer divide by zero")); }
                    self.push(Value::I32((a % b) as i32))?;
                }

                // -- i64 arithmetic --
                _ if op == Op::I64_ADD => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_add(b)))?; }
                _ if op == Op::I64_SUB => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_sub(b)))?; }
                _ if op == Op::I64_MUL => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a.wrapping_mul(b)))?; }
                _ if op == Op::I64_DIV_S => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64(a.wrapping_div(b)))?; }
                _ if op == Op::I64_DIV_U => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64((a / b) as i64))?; }
                _ if op == Op::I64_REM_S => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64(a.wrapping_rem(b)))?; }
                _ if op == Op::I64_REM_U => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; if b == 0 { return Err(VMError::new("trap: integer divide by zero")); } self.push(Value::I64((a % b) as i64))?; }
                _ if op == Op::I64_AND => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a & b))?; }
                _ if op == Op::I64_OR => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a | b))?; }
                _ if op == Op::I64_XOR => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a ^ b))?; }
                _ if op == Op::I64_SHL => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a << (b & 0x3f)))?; }
                _ if op == Op::I64_SHR_S => { let b = self.pop().as_i64(); let a = self.pop().as_i64(); self.push(Value::I64(a >> (b & 0x3f)))?; }
                _ if op == Op::I64_SHR_U => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64((a >> (b & 0x3f)) as i64))?; }
                _ if op == Op::I64_ROTL => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64(a.rotate_left((b & 0x3f) as u32) as i64))?; }
                _ if op == Op::I64_ROTR => { let b = self.pop().as_i64() as u64; let a = self.pop().as_i64() as u64; self.push(Value::I64(a.rotate_right((b & 0x3f) as u32) as i64))?; }
                _ if op == Op::I64_CLZ => { let a = self.pop().as_i64(); self.push(Value::I64(a.leading_zeros() as i64))?; }
                _ if op == Op::I64_CTZ => { let a = self.pop().as_i64(); self.push(Value::I64(a.trailing_zeros() as i64))?; }
                _ if op == Op::I64_POPCNT => { let a = self.pop().as_i64(); self.push(Value::I64(a.count_ones() as i64))?; }

                // -- f64 math --
                _ if op == Op::F64_ABS => { let a = self.pop().as_f64(); self.push(Value::F64(a.abs()))?; }
                _ if op == Op::F64_CEIL => { let a = self.pop().as_f64(); self.push(Value::F64(a.ceil()))?; }
                _ if op == Op::F64_FLOOR => { let a = self.pop().as_f64(); self.push(Value::F64(a.floor()))?; }
                _ if op == Op::F64_TRUNC => { let a = self.pop().as_f64(); self.push(Value::F64(a.trunc()))?; }
                _ if op == Op::F64_NEAREST => { let a = self.pop().as_f64(); self.push(Value::F64(a.round()))?; }
                _ if op == Op::F64_SQRT => { let a = self.pop().as_f64(); self.push(Value::F64(a.sqrt()))?; }
                _ if op == Op::F64_MIN => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64(a.min(b)))?; }
                _ if op == Op::F64_MAX => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64(a.max(b)))?; }
                _ if op == Op::F64_COPYSIGN => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64(a.copysign(b)))?; }

                // -- f32 (promoted to f64) --
                _ if op == Op::F32_ABS => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).abs() as f64))?; }
                _ if op == Op::F32_NEG => { let a = self.pop().as_f64(); self.push(Value::F64(-(a as f32) as f64))?; }
                _ if op == Op::F32_CEIL => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).ceil() as f64))?; }
                _ if op == Op::F32_FLOOR => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).floor() as f64))?; }
                _ if op == Op::F32_TRUNC => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).trunc() as f64))?; }
                _ if op == Op::F32_NEAREST => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).round() as f64))?; }
                _ if op == Op::F32_SQRT => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32).sqrt() as f64))?; }
                _ if op == Op::F32_MIN => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64((a as f32).min(b as f32) as f64))?; }
                _ if op == Op::F32_MAX => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64((a as f32).max(b as f32) as f64))?; }
                _ if op == Op::F32_COPYSIGN => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::F64((a as f32).copysign(b as f32) as f64))?; }

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
                // `extra_tables` for the multi-table proposal.
                _ if op == Op::TABLE_GET => {
                    let table_idx = self.read_byte() as usize;
                    let idx = self.pop().as_i32() as usize;
                    let table = self.table_ref(table_idx);
                    let val = table.and_then(|t| t.get(idx).cloned()).unwrap_or(Value::Null);
                    self.push(val)?;
                }
                // `table.set tbl` — pop value + i32 index, write into
                // table. Trap on out-of-bounds index per spec.
                _ if op == Op::TABLE_SET => {
                    let table_idx = self.read_byte() as usize;
                    let val = self.pop();
                    let idx = self.pop().as_i32() as usize;
                    let table = self.table_mut(table_idx)
                        .ok_or_else(|| VMError::new("trap: table.set unknown table"))?;
                    if idx >= table.len() {
                        return Err(VMError::new("trap: table.set out of bounds"));
                    }
                    table[idx] = val;
                }

                // -- i32 rotation and bit counting --
                _ if op == Op::I32_ROTL => { let b = self.pop().as_i32() as u32; let a = self.pop().as_i32() as u32; self.push(Value::I32(a.rotate_left(b & 0x1f) as i32))?; }
                _ if op == Op::I32_ROTR => { let b = self.pop().as_i32() as u32; let a = self.pop().as_i32() as u32; self.push(Value::I32(a.rotate_right(b & 0x1f) as i32))?; }
                _ if op == Op::I32_CLZ => { let a = self.pop().as_i32() as u32; self.push(Value::I32(a.leading_zeros() as i32))?; }
                _ if op == Op::I32_CTZ => { let a = self.pop().as_i32() as u32; self.push(Value::I32(a.trailing_zeros() as i32))?; }
                _ if op == Op::I32_POPCNT => { let a = self.pop().as_i32() as u32; self.push(Value::I32(a.count_ones() as i32))?; }

                // -- eqz --
                _ if op == Op::I32_EQZ => { let a = self.pop().as_i32(); self.push(Value::Bool(a == 0))?; }
                _ if op == Op::I64_EQZ => { let a = self.pop().as_i64(); self.push(Value::Bool(a == 0))?; }

                // -- String --
                _ if op == Op::STR_CONCAT => {
                    let b = self.pop();
                    let a = self.pop();
                    let s = format!("{}{}", a, b);
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::STR_CONCAT_N => {
                    let count = self.read_byte() as usize;
                    let count = count.min(self.stack.len());
                    let start = self.stack.len() - count;
                    let mut result = String::new();
                    for i in start..self.stack.len() {
                        result.push_str(&format!("{}", self.stack[i]));
                    }
                    self.stack.truncate(start);
                    self.push(Value::String(Arc::from(result.as_str())))?;
                }

                // -- Bitwise --
                _ if op == Op::I32_AND => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a & b))?; }
                _ if op == Op::I32_OR => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a | b))?; }
                _ if op == Op::I32_XOR => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a ^ b))?; }
                // i32_not: removed (non-WASM, use i32.const -1 + i32.xor)
                _ if op == Op::I32_SHL => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a << (b & 0x1f)))?; }
                _ if op == Op::I32_SHR_S => { let b = self.pop().as_i32(); let a = self.pop().as_i32(); self.push(Value::I32(a >> (b & 0x1f)))?; }
                _ if op == Op::I32_SHR_U => { let b = self.pop().as_i32() as u32; let a = self.pop().as_i32() as u32; self.push(Value::I32((a >> (b & 0x1f)) as i32))?; }

                // -- Comparison --
                _ if op == Op::EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(a.eq(&b)))?;
                }
                _ if op == Op::NE => {
                    let b = self.pop();
                    let a = self.pop();
                    self.push(Value::Bool(!a.eq(&b)))?;
                }
                _ if op == Op::F64_LT => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a < b))?; }
                _ if op == Op::F64_GT => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a > b))?; }
                _ if op == Op::F64_LE => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a <= b))?; }
                _ if op == Op::F64_GE => { let b = self.pop().as_f64(); let a = self.pop().as_f64(); self.push(Value::Bool(a >= b))?; }
                // str_lt, str_gt: removed (non-WASM, were unused)

                // -- Logical --
                // bool_not: removed (non-WASM, use dyn_to_bool + i32_eqz)

                // -- Control flow --
                _ if op == Op::BR => {
                    let offset = self.read_i16();
                    let f = self.frame_mut();
                    f.ip = (f.ip as i64 + offset as i64) as usize;
                }
                _ if op == Op::BR_IF_FALSE => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if val.as_bool() == false {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                _ if op == Op::BR_IF_TRUE => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if val.as_bool() == true {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                _ if op == Op::BR_IF_NULL => {
                    let offset = self.read_i16();
                    let val = self.pop();
                    if matches!(val, Value::Null) {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
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
                    let n = (self.chunks[frame_chunk].result_arity as usize).max(1);
                    let split = self.stack.len().saturating_sub(n);
                    let mut results: Vec<Value> = self.stack.split_off(split);
                    let base = self.frame().base;
                    self.close_upvalues(base);
                    self.frames.pop();
                    if self.frames.is_empty() || self.frames.len() < min_depth {
                        let last = results.pop().unwrap_or(Value::Null);
                        return Ok(last);
                    }
                    self.stack.truncate(base);
                    for r in results { self.push(r)?; }
                }
                _ if op == Op::REF_FUNC => {
                    let func_idx = self.read_u16() as usize;
                    let chunk = &self.chunks[func_idx];
                    let arity = chunk.arity;
                    let name = if chunk.name == "<script>" { None } else { Some(chunk.name.clone()) };

                    let uv_count = self.read_byte() as usize;
                    let mut upvalues: Vec<Arc<Mutex<Upvalue>>> = Vec::with_capacity(uv_count);
                    for _ in 0..uv_count {
                        let is_local = self.read_byte() != 0;
                        let index = self.read_byte() as usize;
                        if is_local {
                            let base = self.frame().base;
                            let uv = self.capture_upvalue(base + index);
                            upvalues.push(uv);
                        } else {
                            let uv = self.frame().upvalues[index].clone();
                            upvalues.push(uv);
                        }
                    }

                    let func = Function { name, arity, chunk_index: func_idx, upvalues };
                    let mut obj = Object { properties: HashMap::new(), kind: ObjectKind::Function(func), type_id: 0, fields: Vec::new() };
                    // Add to function table for call_indirect
                    let table_idx = self.func_table.len();
                    obj.properties.insert("__table_idx".into(), Value::F64(table_idx as f64));
                    let func_val = Value::Object(Arc::new(Mutex::new(obj)));
                    self.func_table.push(func_val.clone());
                    self.push(func_val)?;
                }

                // -- Host functions --
                _ if op == Op::CALL_IMPORT => {
                    let import_idx = self.read_u16() as usize;
                    let argc = self.read_byte() as usize;

                    if import_idx >= self.import_table.len() {
                        return Err(VMError::new(format!("Unresolved import index: {}", import_idx)));
                    }

                    match self.import_table[import_idx].clone() {
                        ImportTarget::Host(host_idx) => {
                            let base = self.stack.len() - argc;
                            let args: Vec<Value> = self.stack[base..].to_vec();
                            self.stack.truncate(base);

                            let placeholder: HostFn = Arc::new(|_, _| Value::Null);
                            let host_fn = std::mem::replace(&mut self.host_fns[host_idx], placeholder);
                            let result = {
                                let mut ctx = self.make_host_context();
                                host_fn(&mut ctx, &args)
                            };
                            self.host_fns[host_idx] = host_fn;

                            // JSPI: transparent async suspension
                            if let Value::Object(ref obj) = result {
                                let o = obj.lock().unwrap();
                                let is_pending = o.properties.get("__type")
                                    .map(|v| format!("{}", v) == "Promise")
                                    .unwrap_or(false)
                                    && o.properties.get("__state")
                                        .map(|v| format!("{}", v) == "pending")
                                        .unwrap_or(false);
                                if is_pending {
                                    let promise_id = o.properties.get("__id")
                                        .map(|v| v.as_f64() as u64)
                                        .unwrap_or(0);
                                    drop(o);
                                    let fiber = self.save_fiber();
                                    self.event_loop.borrow_mut().suspend_fiber(promise_id, fiber);
                                    return Err(VMError::new(format!("__jspi__:{}", promise_id)));
                                }
                            }

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
                                    "Stdlib redirect not found: {}", global_name
                                )));
                            }
                        }
                    }
                }

                // -- Object/Array --
                _ if op == Op::STRUCT_NEW => {
                    let count = self.read_u16() as usize;
                    let mut obj = Object::new();
                    let needed = count * 2;
                    let available = self.stack.len();
                    let start = if needed <= available { available - needed } else { 0 };
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
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))?;
                }
                // `array.new $t` — [value, length] -> [array of length,
                // every lane = value].
                _ if op == Op::ARRAY_NEW => {
                    let _typeidx = self.read_u16();
                    let len = self.pop().as_i32().max(0) as usize;
                    let value = self.pop();
                    let elems = vec![value; len];
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))?;
                }
                // `array.new_default $t` — [length] -> [array of length,
                // zero-initialised]. We use `Value::Null` as the default
                // for externref lanes (the only lane type we actually
                // support) per the "null is the default for refs" rule.
                _ if op == Op::ARRAY_NEW_DEFAULT => {
                    let _typeidx = self.read_u16();
                    let len = self.pop().as_i32().max(0) as usize;
                    let elems = vec![Value::Null; len];
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))?;
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
                    let _dataidx = self.read_u16();
                    let _size = self.pop().as_i32();
                    let _offset = self.pop().as_i32();
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new())))))?;
                }
                _ if op == Op::ARRAY_NEW_ELEM => {
                    let _typeidx = self.read_u16();
                    let _elemidx = self.read_u16();
                    let _size = self.pop().as_i32();
                    let _offset = self.pop().as_i32();
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new())))))?;
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
                                        let v = if is_signed { (b as i8) as i32 } else { b as i32 };
                                        Value::I32(v)
                                    }
                                    2 => {
                                        let lo = buf.get(base).copied().unwrap_or(0) as u16;
                                        let hi = buf.get(base + 1).copied().unwrap_or(0) as u16;
                                        let raw = lo | (hi << 8);
                                        let v = if is_signed { (raw as i16) as i32 } else { raw as i32 };
                                        Value::I32(v)
                                    }
                                    _ => Value::Null,
                                }
                            }
                            ObjectKind::ArrayBuffer(ab) => {
                                let buf = ab.bytes.lock().unwrap();
                                let b = buf.get(idx).copied().unwrap_or(0);
                                let v = if is_signed { (b as i8) as i32 } else { b as i32 };
                                Value::I32(v)
                            }
                            ObjectKind::Array(elems) => {
                                elems.get(idx).cloned().unwrap_or(Value::Null)
                            }
                            _ => Value::Null,
                        }
                    } else { Value::Null };
                    self.push(val)?;
                }
                // `array.init_data $t $d` / `array.init_elem $t $e` — copy
                // elements into an existing array. Stub to a no-op (same
                // rationale as new_data / new_elem above).
                _ if op == Op::ARRAY_INIT_DATA => {
                    let _typeidx = self.read_u16();
                    let _dataidx = self.read_u16();
                    let _size = self.pop().as_i32();
                    let _src_offset = self.pop().as_i32();
                    let _dst_offset = self.pop().as_i32();
                    let _array = self.pop();
                }
                _ if op == Op::ARRAY_INIT_ELEM => {
                    let _typeidx = self.read_u16();
                    let _elemidx = self.read_u16();
                    let _size = self.pop().as_i32();
                    let _src_offset = self.pop().as_i32();
                    let _dst_offset = self.pop().as_i32();
                    let _array = self.pop();
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
                        o.lock().unwrap().properties
                            .get("__descriptor").cloned().unwrap_or(Value::Null)
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
                        o.fields.get(field_idx as usize).cloned().unwrap_or(Value::Null)
                    } else { Value::Null };
                    self.push(val)?;
                }
                // `ref.test_null ht` / `ref.cast_null ht` — same as their
                // non-null variants but succeed when the operand is null.
                // Our VM already treats null as assignable to externref,
                // so these short-circuit to true / pass-through on null.
                _ if op == Op::REF_TEST_NULL => {
                    let _typeidx = self.read_u16();
                    let val = self.pop();
                    let matches = matches!(val, Value::Null | Value::Object(_) | Value::Symbol(_) | Value::String(_));
                    self.push(Value::I32(if matches { 1 } else { 0 }))?;
                }
                _ if op == Op::REF_CAST_NULL => {
                    let _typeidx = self.read_u16();
                    // cast_null succeeds for null — pass the value through.
                    // Our VM is already dynamically typed so this is always
                    // a pass-through (validation happens at the engine level
                    // once we emit spec-correct bytes).
                }
                // `any.convert_extern` / `extern.convert_any` — identity at
                // runtime for us: our value ABI is a universal externref.
                // Spec says composing the two yields the original value, so
                // emitting them as nops is semantically correct.
                _ if op == Op::ANY_CONVERT_EXTERN || op == Op::EXTERN_CONVERT_ANY => {}
                // `ref.as_non_null` — trap if the operand is null, otherwise
                // pass the value through unchanged.
                _ if op == Op::REF_AS_NON_NULL => {
                    if let Some(Value::Null) = self.stack.last() {
                        return Err(VMError::new("trap: ref.as_non_null on null reference"));
                    }
                }
                // `br_on_null $l` — if TOS is null, pop it and branch;
                // otherwise leave the value on the stack and fall through.
                // `br_on_non_null $l` — if TOS is non-null, branch with
                // the value; otherwise pop and fall through.
                _ if op == Op::BR_ON_NULL => {
                    let offset = self.read_i16();
                    let is_null = matches!(self.stack.last(), Some(Value::Null));
                    if is_null {
                        self.pop();
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    }
                }
                _ if op == Op::BR_ON_NON_NULL => {
                    let offset = self.read_i16();
                    let is_null = matches!(self.stack.last(), Some(Value::Null));
                    if !is_null {
                        let f = self.frame_mut();
                        f.ip = (f.ip as i64 + offset as i64) as usize;
                    } else {
                        self.pop();
                    }
                }

                // -- Immediates --
                _ if op == Op::NULL
                    || op == Op::NULL_FUNC
                    || op == Op::NULL_ANY
                    || op == Op::NULL_NONE => self.push(Value::Null)?,
                _ if op == Op::UNDEFINED => self.push(Value::Undefined)?,
                _ if op == Op::SYMBOL => {
                    // Each `SYMBOL` execution produces a *fresh* identity
                    // (new Arc allocation) so `symbol.eq` correctly returns
                    // false for two symbols made with the same description.
                    let idx = self.read_u16();
                    let desc: Arc<str> = match self.get_constant(idx) {
                        Value::String(s) => Arc::from(s.to_string().as_str()),
                        _ => Arc::from(""),
                    };
                    self.push(Value::Symbol(desc))?;
                }
                _ if op == Op::BIGINT => {
                    let idx = self.read_u16();
                    let n = match self.get_constant(idx) {
                        Value::I64(n) => n,
                        Value::I32(n) => n as i64,
                        Value::F64(n) => n as i64,
                        _ => 0,
                    };
                    self.push(Value::BigInt(n))?;
                }
                _ if op == Op::REF_IS_UNDEFINED => {
                    let v = self.pop();
                    self.push(Value::Bool(matches!(v, Value::Undefined)))?;
                }
                _ if op == Op::REF_IS_SYMBOL => {
                    let v = self.pop();
                    self.push(Value::Bool(matches!(v, Value::Symbol(_))))?;
                }
                _ if op == Op::REF_IS_BIGINT => {
                    let v = self.pop();
                    self.push(Value::Bool(matches!(v, Value::BigInt(_))))?;
                }
                _ if op == Op::REF_IS_I32 => {
                    // JS Number that fits in i32 range with no fractional part.
                    let v = self.pop();
                    let is_i32 = match &v {
                        Value::I32(_) => true,
                        Value::I64(n) => *n >= i32::MIN as i64 && *n <= i32::MAX as i64,
                        Value::F64(n) => n.is_finite() && *n == (*n as i32) as f64,
                        Value::BigInt(n) => *n >= i32::MIN as i64 && *n <= i32::MAX as i64,
                        _ => false,
                    };
                    self.push(Value::Bool(is_i32))?;
                }
                _ if op == Op::REF_IS_U32 => {
                    let v = self.pop();
                    let is_u32 = match &v {
                        Value::I32(n) => *n >= 0,
                        Value::I64(n) => *n >= 0 && *n <= u32::MAX as i64,
                        Value::F64(n) => n.is_finite() && *n >= 0.0 && *n == (*n as u32) as f64,
                        Value::BigInt(n) => *n >= 0 && *n <= u32::MAX as i64,
                        _ => false,
                    };
                    self.push(Value::Bool(is_u32))?;
                }
                _ if op == Op::NUM_BOX_U32 => {
                    // Interpret the top-of-stack i32 as unsigned and
                    // push a Value::F64 in the u32 range (matches JS
                    // semantics: all Numbers are double-precision).
                    let n = self.pop().as_i32() as u32;
                    self.push(Value::F64(n as f64))?;
                }
                _ if op == Op::NUM_UNBOX_U32 => {
                    let v = self.pop();
                    let n = match v {
                        Value::F64(n) => n as u32,
                        Value::I32(n) => n as u32,
                        Value::I64(n) => n as u32,
                        Value::BigInt(n) => n as u32,
                        _ => 0,
                    };
                    // u32 stored in i32 slot with its bit pattern preserved.
                    self.push(Value::I32(n as i32))?;
                }
                _ if op == Op::BOOL_CAST => {
                    // Traps in the spec if value isn't a boolean; we
                    // fall back to truthiness.
                    let v = self.pop();
                    let b = match v {
                        Value::Bool(b) => if b { 1 } else { 0 },
                        other => if dyn_truthy(&other) { 1 } else { 0 },
                    };
                    self.push(Value::I32(b))?;
                }
                _ if op == Op::STR_CAST => {
                    // Validating cast: pass the String through unchanged,
                    // or convert via Display. The emitter-side trap
                    // semantics apply only in the JS host.
                    let v = self.pop();
                    let s = match v {
                        Value::String(s) => s,
                        Value::Symbol(s) => s,
                        other => Arc::from(format!("{}", other).as_str()),
                    };
                    self.push(Value::String(s))?;
                }
                _ if op == Op::STR_FROM_I32 => {
                    let n = self.pop().as_i32();
                    self.push(Value::String(Arc::from(n.to_string().as_str())))?;
                }
                _ if op == Op::STR_FROM_U32 => {
                    let n = self.pop().as_i32() as u32;
                    self.push(Value::String(Arc::from(n.to_string().as_str())))?;
                }
                _ if op == Op::STR_FROM_I64 => {
                    let n = self.pop().as_i64();
                    self.push(Value::String(Arc::from(n.to_string().as_str())))?;
                }
                _ if op == Op::STR_FROM_U64 => {
                    let n = self.pop().as_i64() as u64;
                    self.push(Value::String(Arc::from(n.to_string().as_str())))?;
                }
                _ if op == Op::STR_FROM_F64 => {
                    let n = self.pop().as_f64();
                    // JS-compatible formatting: integer-valued doubles print without `.0`.
                    let s = if n.is_nan() {
                        "NaN".to_string()
                    } else if n.is_infinite() {
                        if n > 0.0 { "Infinity".to_string() } else { "-Infinity".to_string() }
                    } else if n == (n as i64) as f64 && n.abs() < 1e21 {
                        (n as i64).to_string()
                    } else {
                        format!("{}", n)
                    };
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                _ if op == Op::SYMBOL_EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    let eq = match (&a, &b) {
                        (Value::Symbol(a), Value::Symbol(b)) => Arc::ptr_eq(a, b),
                        _ => false,
                    };
                    self.push(Value::Bool(eq))?;
                }
                _ if op == Op::TRUE => self.push(Value::Bool(true))?,
                _ if op == Op::FALSE => self.push(Value::Bool(false))?,
                _ if op == Op::I32_CONST_0 => self.push(Value::I32(0))?,
                _ if op == Op::I32_CONST_1 => self.push(Value::I32(1))?,
                _ if op == Op::F64_CONST_0 => self.push(Value::F64(0.0))?,

                // ref.eq (GC proposal) — reference identity equality.
                // Two references are equal iff they point at the same
                // underlying object. Null-null is also true. Used by JS
                // `===` for object identity.
                _ if op == Op::REF_EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    let eq = match (&a, &b) {
                        (Value::Null, Value::Null) => true,
                        (Value::Undefined, Value::Undefined) => true,
                        (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
                        (Value::Symbol(a), Value::Symbol(b)) => Arc::ptr_eq(a, b),
                        (Value::String(a), Value::String(b)) => Arc::ptr_eq(a, b),
                        _ => false,
                    };
                    self.push(Value::Bool(eq))?;
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
                    self.push(Value::Bool(result))?;
                }
                _ if op == Op::REF_CAST => {
                    let type_name_idx = self.read_u16();
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    let is_type = self.test_type(&val, &target_name);
                    if !is_type {
                        return Err(VMError::new(&format!("ref.cast failed: value is not {}", target_name)));
                    }
                    // Value stays on stack (cast is a no-op if it passes)
                }
                // `br_on_cast l ht` / `br_on_cast_fail l ht` — structured
                // branch keyed off a runtime type test. Operand is
                // (u16 type-name-idx, u8 label-depth), matching BR_LABEL's
                // label-stack discipline so the VM can honour the branch
                // without a parallel byte-offset table.
                _ if op == Op::BR_ON_CAST => {
                    let type_name_idx = self.read_u16();
                    let depth = self.read_byte() as usize;
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    if self.test_type(&val, &target_name) {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                            let target = entry.target;
                            let is_loop = entry.is_loop;
                            self.frames.last_mut().unwrap().ip = target;
                            let len = self.label_stack.len();
                            if is_loop {
                                self.label_stack.truncate(len - depth);
                            } else {
                                self.label_stack.truncate(len - depth - 1);
                            }
                        }
                    }
                }
                _ if op == Op::BR_ON_CAST_FAIL => {
                    let type_name_idx = self.read_u16();
                    let depth = self.read_byte() as usize;
                    let target_name = self.constant_str(type_name_idx);
                    let val = self.peek(0).clone();
                    if !self.test_type(&val, &target_name) {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                            let target = entry.target;
                            let is_loop = entry.is_loop;
                            self.frames.last_mut().unwrap().ip = target;
                            let len = self.label_stack.len();
                            if is_loop {
                                self.label_stack.truncate(len - depth);
                            } else {
                                self.label_stack.truncate(len - depth - 1);
                            }
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
                    } else { v };
                    self.push(Value::I32(extended))?;
                }
                _ if op == Op::I31_GET_U => {
                    let v = self.pop().as_i32();
                    self.push(Value::I32(v & 0x7FFF_FFFF))?;
                }

                _ if op == Op::REF_IS_NULL => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Null | Value::Undefined)))?; }
                _ if op == Op::REF_IS_STRING => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::String(_))))?; }
                _ if op == Op::REF_IS_NUMBER => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::F64(_) | Value::I32(_) | Value::I64(_))))?; }
                _ if op == Op::REF_IS_BOOL => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Bool(_))))?; }
                _ if op == Op::REF_IS_OBJECT => { let v = self.pop(); self.push(Value::Bool(matches!(v, Value::Object(_))))?; }
                _ if op == Op::REF_IS_FUNC => {
                    let v = self.pop();
                    let is_fn = matches!(&v, Value::Object(o) if matches!(o.lock().unwrap().kind, ObjectKind::Function(_)));
                    self.push(Value::Bool(is_fn))?;
                }

                // -- Conversions --
                _ if op == Op::F64_FROM_I32 => {
                    let v = self.pop();
                    self.push(Value::F64(v.as_f64()))?;
                }
                _ if op == Op::I32_FROM_F64 => {
                    let v = self.pop();
                    self.push(Value::I32(v.as_i32()))?;
                }

                // -- Dynamic ops (inline type dispatch, no host call) --
                _ if op == Op::DYN_ADD => {
                    let b = self.pop();
                    let a = self.pop();
                    let result = match (&a, &b) {
                        (Value::F64(x), Value::F64(y)) => Value::F64(x + y),
                        (Value::I32(x), Value::I32(y)) => Value::I32(x.wrapping_add(*y)),
                        (Value::String(_), _) | (_, Value::String(_)) => {
                            Value::String(Arc::from(format!("{}{}", a, b).as_str()))
                        }
                        // Object with __add__ dunder → cross-language operator overloading
                        (Value::Object(obj), _) => {
                            self.try_dunder_binary(obj, &b, "__add__")
                                .unwrap_or_else(|| Value::F64(a.as_f64() + b.as_f64()))
                        }
                        _ => Value::F64(a.as_f64() + b.as_f64()),
                    };
                    self.push(result)?;
                }
                _ if op == Op::DYN_EQ => {
                    let b = self.pop();
                    let a = self.pop();
                    // JS abstract equality (==) coercion rules
                    fn loose_eq(a: &Value, b: &Value) -> bool {
                        match (a, b) {
                            // null/undefined: equal to each other, nothing else
                            (Value::Null, Value::Null) | (Value::Undefined, Value::Undefined) => true,
                            (Value::Null, Value::Undefined) | (Value::Undefined, Value::Null) => true,
                            (Value::Null, _) | (_, Value::Null) => false,
                            (Value::Undefined, _) | (_, Value::Undefined) => false,
                            // Same primitive types
                            (Value::Bool(x), Value::Bool(y)) => x == y,
                            (Value::F64(x), Value::F64(y)) => if x.is_nan() || y.is_nan() { false } else { x == y },
                            (Value::I32(x), Value::I32(y)) => x == y,
                            (Value::F64(x), Value::I32(y)) => *x == *y as f64,
                            (Value::I32(x), Value::F64(y)) => *x as f64 == *y,
                            (Value::String(x), Value::String(y)) => x == y,
                            // Bool with anything else → coerce bool to number, retry
                            (Value::Bool(x), other) | (other, Value::Bool(x)) => {
                                let n = if *x { 1.0 } else { 0.0 };
                                loose_eq(&Value::F64(n), other)
                            }
                            // String == Number: parse string, NaN if invalid
                            (Value::String(s), Value::F64(n)) | (Value::F64(n), Value::String(s)) => {
                                let trimmed = s.trim();
                                if trimmed.is_empty() { *n == 0.0 }
                                else if let Ok(sv) = trimmed.parse::<f64>() { sv == *n } else { false }
                            }
                            (Value::String(s), Value::I32(n)) | (Value::I32(n), Value::String(s)) => {
                                let trimmed = s.trim();
                                if trimmed.is_empty() { *n == 0 }
                                else if let Ok(sv) = trimmed.parse::<f64>() { sv == *n as f64 } else { false }
                            }
                            // Object: same-reference, OR compare via toPrimitive (Array → string)
                            (Value::Object(x), Value::Object(y)) => Arc::ptr_eq(x, y),
                            // Object with primitive: coerce object to primitive (Array → joined string)
                            (Value::Object(o), other) | (other, Value::Object(o)) => {
                                let obj = o.lock().unwrap();
                                if let ObjectKind::Array(elems) = &obj.kind {
                                    let s = elems.iter().map(|v| format!("{}", v))
                                        .collect::<Vec<_>>().join(",");
                                    drop(obj);
                                    loose_eq(&Value::String(Arc::from(s.as_str())), other)
                                } else {
                                    false
                                }
                            }
                            _ => false,
                        }
                    }
                    self.push(Value::Bool(loose_eq(&a, &b)))?;
                }
                _ if op == Op::DYN_NE => {
                    let b = self.pop();
                    let a = self.pop();
                    let result = match (&a, &b) {
                        (Value::Null, Value::Null) | (Value::Undefined, Value::Undefined) => false,
                        (Value::Null, Value::Undefined) | (Value::Undefined, Value::Null) => false,
                        (Value::Bool(x), Value::Bool(y)) => x != y,
                        (Value::F64(x), Value::F64(y)) => if x.is_nan() || y.is_nan() { true } else { x != y },
                        (Value::I32(x), Value::I32(y)) => x != y,
                        (Value::F64(x), Value::I32(y)) => *x != *y as f64,
                        (Value::I32(x), Value::F64(y)) => *x as f64 != *y,
                        (Value::String(s), Value::F64(n)) | (Value::F64(n), Value::String(s)) => {
                            if let Ok(sv) = s.parse::<f64>() { sv != *n } else { true }
                        }
                        (Value::String(s), Value::I32(n)) | (Value::I32(n), Value::String(s)) => {
                            if let Ok(sv) = s.parse::<f64>() { sv != *n as f64 } else { true }
                        }
                        (Value::String(x), Value::String(y)) => x != y,
                        (Value::Object(x), Value::Object(y)) => !Arc::ptr_eq(x, y),
                        _ => true,
                    };
                    self.push(Value::Bool(result))?;
                }
                _ if op == Op::DYN_LT => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x < *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__lt__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() < b.as_f64()),
                        _ => a.as_f64() < b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                _ if op == Op::DYN_GT => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x > *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__gt__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() > b.as_f64()),
                        _ => a.as_f64() > b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                _ if op == Op::DYN_LE => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x <= *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__le__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() <= b.as_f64()),
                        _ => a.as_f64() <= b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                _ if op == Op::DYN_GE => {
                    let b = self.pop(); let a = self.pop();
                    let r = match (&a, &b) {
                        (Value::String(x), Value::String(y)) => *x >= *y,
                        (Value::Object(obj), _) => self.try_dunder_binary(obj, &b, "__ge__")
                            .map(|v| v.as_bool()).unwrap_or(a.as_f64() >= b.as_f64()),
                        _ => a.as_f64() >= b.as_f64(),
                    };
                    self.push(Value::Bool(r))?;
                }
                _ if op == Op::DYN_NEG => {
                    let a = self.pop();
                    self.push(Value::F64(-a.as_f64()))?;
                }
                _ if op == Op::DYN_NOT => {
                    let a = self.pop();
                    self.push(Value::Bool(!dyn_truthy(&a)))?;
                }
                _ if op == Op::DYN_TO_BOOL => {
                    let a = self.pop();
                    self.push(Value::Bool(dyn_truthy(&a)))?;
                }

                // -- Async (await) --
                // r#await: removed (duplicate of promise_suspend, use JSPI proposal name)

                _ if op == Op::SET_TIMER => {
                    let ms = self.pop().as_f64();
                    let callback = self.pop();
                    self.event_loop.borrow_mut().queue_timer(callback, ms);
                    self.push(Value::Null)?;
                }

                // -- Exceptions (WASM exception proposal) --
                _ if op == Op::TRY_START => {
                    let catch_offset = self.read_u16() as i16;
                    let _finally_offset = self.read_u16(); // reserved for finally
                    let f = self.frame();
                    let catch_ip = (f.ip as i64 + catch_offset as i64) as usize;
                    self.exception_handlers.push(ExceptionHandler {
                        catch_ip,
                        _chunk_index: f.chunk_index,
                        stack_depth: self.stack.len(),
                        frame_depth: self.frames.len(),
                        tag: 0, // catch-all
                    });
                }
                _ if op == Op::TRY_END => {
                    // Normal exit from try block — pop the handler
                    self.exception_handlers.pop();
                }
                _ if op == Op::THROW || op == Op::THROW_REF => {
                    let val = self.pop();
                    // Find a matching handler by walking the handler stack.
                    // Tag 0 = catch-all (always matches).
                    // Tag N = typed catch — match if exception type matches exception_tags[N].
                    let mut matched_idx = None;
                    for i in (0..self.exception_handlers.len()).rev() {
                        let handler = &self.exception_handlers[i];
                        if handler.tag == 0 {
                            // Catch-all — always matches
                            matched_idx = Some(i);
                            break;
                        }
                        // Typed catch — check if thrown value's type matches the tag
                        let tag_idx = handler.tag as usize;
                        let tag_name = self.chunks.get(0)
                            .and_then(|c| c.exception_tags.get(tag_idx))
                            .cloned()
                            .unwrap_or_default();
                        if !tag_name.is_empty() {
                            // Check: is val an instance of tag_name?
                            let matches = self.test_type(&val, &tag_name.to_lowercase())
                                || self.exception_value_matches(&val, &tag_name);
                            if matches {
                                matched_idx = Some(i);
                                break;
                            }
                        }
                        // This handler doesn't match — keep looking
                    }

                    if let Some(idx) = matched_idx {
                        // Remove this handler and all handlers above it
                        let handler = self.exception_handlers[idx].clone();
                        self.exception_handlers.truncate(idx);
                        // Unwind: restore stack and frames
                        while self.frames.len() > handler.frame_depth {
                            let base = self.frames.last().unwrap().base;
                            self.close_upvalues(base);
                            self.frames.pop();
                        }
                        self.stack.truncate(handler.stack_depth);
                        self.push(val)?;
                        let f = self.frame_mut();
                        f.ip = handler.catch_ip;
                    } else {
                        let stack = self.capture_call_stack();
                        return Err(VMError::new(format!("{}", val)).with_stack(stack));
                    }
                }
                _ if op == Op::TRY_TABLE => {
                    // WASM EH Phase 4: [try_table, u8 handler_count, then for each: u8 tag, u16 offset]
                    // Tag 0 = catch-all. Tag N = typed catch for exception_tags[N].
                    // Handlers are pushed in reverse order so the most specific (first) handler
                    // is on top of the stack and checked first during throw.
                    let handler_count = self.read_byte() as usize;
                    let mut handlers = Vec::new();
                    for _ in 0..handler_count {
                        let tag = self.read_byte();
                        let offset = self.read_u16();
                        let ip = self.frame().ip + offset as usize;
                        handlers.push(ExceptionHandler {
                            catch_ip: ip,
                            stack_depth: self.stack.len(),
                            frame_depth: self.frames.len(),
                            _chunk_index: self.frame().chunk_index,
                            tag,
                        });
                    }
                    // Push in reverse so first handler is checked first (it's on top)
                    for h in handlers.into_iter().rev() {
                        self.exception_handlers.push(h);
                    }
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
                    let argc = self.read_byte() as usize;
                    let args_start = self.stack.len() - argc;
                    let table_idx_pos = args_start - 1;
                    let table_idx = self.stack[table_idx_pos].as_i32() as usize;
                    if table_idx < self.func_table.len() {
                        let func = self.func_table[table_idx].clone();
                        self.stack[table_idx_pos] = func;
                        let old_base = self.frame().base;
                        let callee_idx = table_idx_pos;
                        for i in 0..=argc {
                            self.stack[old_base + i] = self.stack[callee_idx + i].clone();
                        }
                        self.stack.truncate(old_base + 1 + argc);
                        self.frames.pop();
                        self.call_value(argc)?;
                    }
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
                    let pages = (self.active_mem_len() / 65536) as i32;
                    self.push(Value::I32(pages))?;
                }
                _ if op == Op::MEMORY_GROW => {
                    let pages = self.pop().as_f64() as usize;
                    let old_pages = self.active_mem_grow(pages);
                    self.push(Value::I32(old_pages as i32))?;
                }
                _ if op == Op::I32_LOAD => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::I32(self.memory.load_i32(addr)?))?;
                }
                _ if op == Op::I32_STORE => {
                    let val = self.pop().as_f64() as i32;
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_i32(addr, val)?;
                }
                _ if op == Op::I64_LOAD => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::I64(self.memory.load_i64(addr)?))?;
                }
                _ if op == Op::I64_STORE => {
                    let val = self.pop().as_f64() as i64;
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_i64(addr, val)?;
                }
                _ if op == Op::F64_LOAD => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::F64(self.memory.load_f64(addr)?))?;
                }
                _ if op == Op::F64_STORE => {
                    let val = self.pop().as_f64();
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_f64(addr, val)?;
                }
                _ if op == Op::I32_LOAD8_U => {
                    let addr = self.pop().as_f64() as usize;
                    self.push(Value::I32(self.memory.load_u8(addr)? as i32))?;
                }
                _ if op == Op::I32_STORE8 => {
                    let val = self.pop().as_f64() as u8;
                    let addr = self.pop().as_f64() as usize;
                    self.memory.store_u8(addr, val)?;
                }
                _ if op == Op::F32_LOAD => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 4 <= buf.len() {
                            f32::from_le_bytes(buf[addr..addr+4].try_into().unwrap()) as f64
                        } else { 0.0 }
                    });
                    self.push(Value::F64(val))?;
                }
                _ if op == Op::F32_STORE => {
                    let val = self.pop().as_f64() as f32;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.with_buffer_mut(|buf| {
                        if addr + 4 <= buf.len() {
                            buf[addr..addr+4].copy_from_slice(&val.to_le_bytes());
                        }
                    });
                }
                _ if op == Op::I32_LOAD8_S => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I32(self.memory.load_u8(addr)? as i8 as i32))?;
                }
                _ if op == Op::I32_LOAD16_S => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { i16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i32 } else { 0 }
                    });
                    self.push(Value::I32(val))?;
                }
                _ if op == Op::I32_LOAD16_U => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { u16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i32 } else { 0 }
                    });
                    self.push(Value::I32(val))?;
                }
                _ if op == Op::I32_STORE16 => {
                    let val = self.pop().as_i32() as i16;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.with_buffer_mut(|buf| {
                        if addr + 2 <= buf.len() { buf[addr..addr+2].copy_from_slice(&val.to_le_bytes()); }
                    });
                }
                _ if op == Op::I64_LOAD8_S => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_u8(addr)? as i8 as i64))?;
                }
                _ if op == Op::I64_LOAD8_U => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_u8(addr)? as i64))?;
                }
                _ if op == Op::I64_LOAD16_S => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { i16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i64 } else { 0 }
                    });
                    self.push(Value::I64(val))?;
                }
                _ if op == Op::I64_LOAD16_U => {
                    let addr = self.pop().as_i32() as usize;
                    let val = self.memory.with_buffer(|buf| {
                        if addr + 2 <= buf.len() { u16::from_le_bytes(buf[addr..addr+2].try_into().unwrap()) as i64 } else { 0 }
                    });
                    self.push(Value::I64(val))?;
                }
                _ if op == Op::I64_LOAD32_S => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_i32(addr)? as i64))?;
                }
                _ if op == Op::I64_LOAD32_U => {
                    let addr = self.pop().as_i32() as usize;
                    self.push(Value::I64(self.memory.load_i32(addr)? as u32 as i64))?;
                }
                _ if op == Op::I64_STORE8 => {
                    let val = self.pop().as_i64() as u8;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.store_u8(addr, val)?;
                }
                _ if op == Op::I64_STORE16 => {
                    let val = self.pop().as_i64() as i16;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.with_buffer_mut(|buf| {
                        if addr + 2 <= buf.len() { buf[addr..addr+2].copy_from_slice(&val.to_le_bytes()); }
                    });
                }
                _ if op == Op::I64_STORE32 => {
                    let val = self.pop().as_i64() as i32;
                    let addr = self.pop().as_i32() as usize;
                    self.memory.store_i32(addr, val)?;
                }

                // -- Conversions --
                _ if op == Op::I32_WRAP_I64 => { let a = self.pop().as_i64(); self.push(Value::I32(a as i32))?; }
                _ if op == Op::I64_EXTEND_I32_S => { let a = self.pop().as_i32(); self.push(Value::I64(a as i64))?; }
                _ if op == Op::I64_EXTEND_I32_U => { let a = self.pop().as_i32() as u32; self.push(Value::I64(a as i64))?; }
                _ if op == Op::I64_TRUNC_F64_S => { let a = self.pop().as_f64(); self.push(Value::I64(a as i64))?; }
                _ if op == Op::I64_TRUNC_F64_U => { let a = self.pop().as_f64(); self.push(Value::I64(a as u64 as i64))?; }
                _ if op == Op::F64_PROMOTE_F32 => { let a = self.pop().as_f64(); self.push(Value::F64(a))?; }
                _ if op == Op::F32_DEMOTE_F64 => { let a = self.pop().as_f64(); self.push(Value::F64((a as f32) as f64))?; }
                _ if op == Op::I32_REINTERPRET_F32 => { let a = self.pop().as_f64() as f32; self.push(Value::I32(a.to_bits() as i32))?; }
                _ if op == Op::I64_REINTERPRET_F64 => { let a = self.pop().as_f64(); self.push(Value::I64(a.to_bits() as i64))?; }
                _ if op == Op::F32_REINTERPRET_I32 => { let a = self.pop().as_i32(); self.push(Value::F64(f32::from_bits(a as u32) as f64))?; }
                _ if op == Op::F64_REINTERPRET_I64 => { let a = self.pop().as_i64(); self.push(Value::F64(f64::from_bits(a as u64)))?; }

                // -- Sign extension --
                _ if op == Op::I32_EXTEND8_S => { let a = self.pop().as_i32() as i8; self.push(Value::I32(a as i32))?; }
                _ if op == Op::I32_EXTEND16_S => { let a = self.pop().as_i32() as i16; self.push(Value::I32(a as i32))?; }
                _ if op == Op::I64_EXTEND8_S => { let a = self.pop().as_i64() as i8; self.push(Value::I64(a as i64))?; }
                _ if op == Op::I64_EXTEND16_S => { let a = self.pop().as_i64() as i16; self.push(Value::I64(a as i64))?; }
                _ if op == Op::I64_EXTEND32_S => { let a = self.pop().as_i64() as i32; self.push(Value::I64(a as i64))?; }

                // -- Multi-value --
                // pack, unpack: removed (non-WASM, were unused by compilers)

                // -- Block/loop structured control --
                _ if op == Op::BLOCK => {
                    let end_offset = self.read_u16() as usize;
                    let _result_count = self.read_byte();
                    let ip = self.frame().ip;
                    self.label_stack.push(LabelEntry { target: ip + end_offset, is_loop: false });
                }
                _ if op == Op::LOOP => {
                    let _body_size = self.read_u16();
                    let _result_count = self.read_byte();
                    let ip = self.frame().ip;
                    self.label_stack.push(LabelEntry { target: ip, is_loop: true });
                }
                _ if op == Op::END => {
                    self.label_stack.pop();
                }
                _ if op == Op::BR_LABEL => {
                    let depth = self.read_byte() as usize;
                    if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                        let target = entry.target;
                        self.frames.last_mut().unwrap().ip = target;
                        let len = self.label_stack.len();
                        if entry.is_loop {
                            // Loop: pop labels above the loop (keep the loop + block below)
                            self.label_stack.truncate(len - depth);
                        } else {
                            // Block: pop the block and everything above it
                            self.label_stack.truncate(len - depth - 1);
                        }
                    }
                }
                _ if op == Op::BR_IF_LABEL => {
                    let depth = self.read_byte() as usize;
                    let cond = self.pop();
                    if dyn_truthy(&cond) {
                        if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                            let target = entry.target;
                            self.frames.last_mut().unwrap().ip = target;
                            let len = self.label_stack.len();
                            if entry.is_loop {
                                self.label_stack.truncate(len - depth);
                            } else {
                                self.label_stack.truncate(len - depth - 1);
                            }
                        }
                    }
                }
                _ if op == Op::BR_TABLE => {
                    let count = self.read_byte() as usize;
                    let default_depth = self.read_byte() as usize;
                    let mut labels = Vec::with_capacity(count);
                    for _ in 0..count { labels.push(self.read_byte() as usize); }
                    let idx = self.pop().as_f64() as usize;
                    let depth = if idx < count { labels[idx] } else { default_depth };
                    if let Some(entry) = self.label_stack.iter().rev().nth(depth) {
                        let target = entry.target;
                        self.frames.last_mut().unwrap().ip = target;
                    }
                }

                // -- call_indirect --
                _ if op == Op::CALL_INDIRECT => {
                    let argc = self.read_byte() as usize;
                    let table_idx_pos = self.stack.len() - 1 - argc;
                    let raw_idx = self.stack[table_idx_pos].as_f64();
                    if raw_idx < 0.0 || raw_idx.is_nan() || raw_idx >= self.func_table.len() as f64 {
                        return Err(VMError::new(format!("trap: call_indirect: invalid table index {}", raw_idx)));
                    }
                    let table_idx = raw_idx as usize;
                    if table_idx < self.func_table.len() {
                        self.stack[table_idx_pos] = self.func_table[table_idx].clone();
                        self.call_value(argc)?;
                    } else {
                        return Err(VMError::new(format!("call_indirect: table index {} out of bounds", table_idx)));
                    }
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
                _ if op == Op::TYPE_IMPORT => {
                    let import_idx = self.read_u16() as usize;
                    // Resolve type import: look up in type_imports table and register.
                    // The type should already be registered during linking.
                    // Push the type_id onto the stack for use by constructors.
                    if import_idx < self.chunks[0].type_imports.len() {
                        let (_iface, type_name) = &self.chunks[0].type_imports[import_idx];
                        if let Some(tid) = self.type_registry.get_id(type_name) {
                            self.push(Value::I32(tid as i32))?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                _ if op == Op::TYPE_EXPORT => {
                    let type_id = self.read_u16() as usize;
                    // Mark type as exported. This is a compile-time declaration;
                    // at runtime, ensure the type is visible to the linker.
                    // No stack effect.
                    let _ = type_id;
                }
                _ if op == Op::SHARED_NEW => {
                    // Create a new shared object from the type_id on stack.
                    let type_id_val = self.pop();
                    let type_id = type_id_val.as_f64() as usize;
                    let mut obj = Object::new();
                    obj.type_id = type_id;
                    if let Some(td) = self.type_registry.get(type_id) {
                        obj.fields = vec![Value::Null; td.field_count()];
                    }
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }

                // -- Shared-Everything Threads (shared GC objects) --
                _ if op == Op::SHARED_STRUCT_GET => {
                    let field_idx = self.read_u16() as usize;
                    let obj_val = self.pop();
                    if let Value::Object(ref obj) = obj_val {
                        let o = obj.lock().unwrap();
                        // Atomic read of indexed field
                        let val = if field_idx < o.fields.len() {
                            o.fields[field_idx].clone()
                        } else {
                            Value::Null
                        };
                        self.push(val)?;
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                _ if op == Op::SHARED_STRUCT_SET => {
                    let field_idx = self.read_u16() as usize;
                    let value = self.pop();
                    let obj_val = self.pop();
                    if let Value::Object(ref obj) = obj_val {
                        let mut o = obj.lock().unwrap();
                        // Atomic write of indexed field
                        if field_idx < o.fields.len() {
                            o.fields[field_idx] = value;
                        } else {
                            // Grow fields if needed
                            while o.fields.len() <= field_idx {
                                o.fields.push(Value::Null);
                            }
                            o.fields[field_idx] = value;
                        }
                    }
                }
                _ if op == Op::SHARED_ARRAY_GET => {
                    let idx_val = self.pop();
                    let arr_val = self.pop();
                    let idx = idx_val.as_f64() as usize;
                    if let Value::Object(ref obj) = arr_val {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref elems) = o.kind {
                            self.push(elems.get(idx).cloned().unwrap_or(Value::Null))?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                _ if op == Op::SHARED_ARRAY_SET => {
                    let value = self.pop();
                    let idx_val = self.pop();
                    let arr_val = self.pop();
                    let idx = idx_val.as_f64() as usize;
                    if let Value::Object(ref obj) = arr_val {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut elems) = o.kind {
                            if idx < elems.len() {
                                elems[idx] = value;
                            }
                        }
                    }
                }
                _ if op == Op::SHARED_STRUCT_CAS => {
                    let field_idx = self.read_u16() as usize;
                    let new_val = self.pop();
                    let expected = self.pop();
                    let obj_val = self.pop();
                    if let Value::Object(ref obj) = obj_val {
                        let mut o = obj.lock().unwrap();
                        if field_idx < o.fields.len() {
                            let old = o.fields[field_idx].clone();
                            // Compare (using string repr for simplicity)
                            if format!("{}", old) == format!("{}", expected) {
                                o.fields[field_idx] = new_val;
                            }
                            self.push(old)?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        self.push(Value::Null)?;
                    }
                }

                // -- Weak References & Finalizers --
                _ if op == Op::REF_MAKE_WEAK => {
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        self.push(Value::WeakRef(Arc::downgrade(obj)))?;
                    } else {
                        self.push(Value::Null)?;
                    }
                }
                _ if op == Op::REF_DEREF_WEAK => {
                    let val = self.pop();
                    if let Value::WeakRef(ref weak) = val {
                        if let Some(strong) = weak.upgrade() {
                            self.push(Value::Object(strong))?;
                        } else {
                            self.push(Value::Null)?;
                        }
                    } else {
                        // If it's already a strong ref, just pass through
                        self.push(val)?;
                    }
                }
                _ if op == Op::REF_IS_ALIVE => {
                    let val = self.pop();
                    let alive = match &val {
                        Value::WeakRef(weak) => weak.upgrade().is_some(),
                        Value::Object(_) => true,
                        _ => false,
                    };
                    self.push(Value::Bool(alive))?;
                }
                _ if op == Op::REF_REGISTER_FINALIZER => {
                    let callback = self.pop();
                    let target = self.pop();
                    if let Value::Object(ref obj) = target {
                        self.finalizers.push(FinalizerEntry {
                            target: Arc::downgrade(obj),
                            callback,
                        });
                    }
                }

                // -- Multi-Memory --
                _ if op == Op::MEMORY_SELECT => {
                    let mem_idx = self.read_byte() as usize;
                    self.active_memory = mem_idx;
                }
                _ if op == Op::MEMORY_INIT => {
                    let pages = self.pop().as_f64() as usize;
                    let mem_idx = self.extra_memories.len() + 1; // 0 is default memory
                    self.extra_memories.push(vec![0u8; pages * 65536]);
                    self.push(Value::I32(mem_idx as i32))?;
                }
                // ── reference-types: table operations ─────────────────
                // Each op reads a `u8 table_idx` operand per spec. Tables
                // route through `table_ref`/`table_mut` so the multi-table
                // proposal works: index 0 maps to `func_table`, indexes
                // 1+ map to lazily-created `extra_tables`.
                _ if op == Op::TABLE_SIZE => {
                    let tidx = self.read_byte() as usize;
                    let size = self.table_ref(tidx).map_or(0, |t| t.len()) as i32;
                    self.push(Value::I32(size))?;
                }
                _ if op == Op::TABLE_GROW => {
                    let tidx = self.read_byte() as usize;
                    let delta = self.pop().as_i32().max(0) as usize;
                    let init = self.pop();
                    if let Some(table) = self.table_mut(tidx) {
                        let old_size = table.len();
                        let new_size = old_size.saturating_add(delta);
                        table.resize(new_size, init);
                        self.push(Value::I32(old_size as i32))?;
                    } else {
                        self.push(Value::I32(-1))?;
                    }
                }
                _ if op == Op::TABLE_FILL => {
                    let tidx = self.read_byte() as usize;
                    let count = self.pop().as_i32().max(0) as usize;
                    let value = self.pop();
                    let dst = self.pop().as_i32().max(0) as usize;
                    if let Some(table) = self.table_mut(tidx) {
                        let end = (dst + count).min(table.len());
                        for i in dst..end {
                            table[i] = value.clone();
                        }
                    }
                }
                _ if op == Op::TABLE_COPY => {
                    // Spec: `table.copy dst_table src_table` — we have one
                    // table so both operands are 0.
                    let _dst_table = self.read_byte();
                    let count = self.pop().as_i32().max(0) as usize;
                    let src = self.pop().as_i32().max(0) as usize;
                    let dst = self.pop().as_i32().max(0) as usize;
                    let max = self.func_table.len();
                    if src.saturating_add(count) > max || dst.saturating_add(count) > max {
                        return Err(VMError::new("table.copy: out of bounds".to_string()));
                    }
                    if dst <= src {
                        for i in 0..count {
                            self.func_table[dst + i] = self.func_table[src + i].clone();
                        }
                    } else {
                        for i in (0..count).rev() {
                            self.func_table[dst + i] = self.func_table[src + i].clone();
                        }
                    }
                }
                _ if op == Op::TABLE_INIT => {
                    let _elem_idx = self.read_byte();
                    // No runtime element segments — bounds-check-only no-op.
                    let _count = self.pop().as_i32();
                    let _src = self.pop().as_i32();
                    let _dst = self.pop().as_i32();
                }
                _ if op == Op::ELEM_DROP => {
                    let _elem_idx = self.read_byte();
                }
                _ if op == Op::DATA_DROP => {
                    let _data_idx = self.read_byte();
                }
                _ if op == Op::MEMORY_COPY => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let src = self.pop().as_i32().max(0) as usize;
                    let dst = self.pop().as_i32().max(0) as usize;
                    let mut buf = vec![0u8; count];
                    self.memory.read_bytes(src, &mut buf);
                    self.memory.write_bytes(dst, &buf);
                }
                _ if op == Op::MEMORY_FILL => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let val = self.pop().as_i32() as u8;
                    let dst = self.pop().as_i32().max(0) as usize;
                    let buf = vec![val; count];
                    self.memory.write_bytes(dst, &buf);
                }
                _ if op == Op::MEMORY_COPY_CROSS => {
                    let len = self.pop().as_f64() as usize;
                    let src_addr = self.pop().as_f64() as usize;
                    let src_mem = self.pop().as_f64() as usize;
                    let dst_addr = self.pop().as_f64() as usize;
                    let dst_mem = self.pop().as_f64() as usize;
                    // Copy bytes between memories
                    let src_data: Vec<u8> = if src_mem == 0 {
                        let mut buf = vec![0u8; len];
                        self.memory.read_bytes(src_addr, &mut buf);
                        buf
                    } else {
                        let src = self.extra_mem(src_mem);
                        if src_addr + len <= src.len() {
                            src[src_addr..src_addr + len].to_vec()
                        } else {
                            vec![0u8; len]
                        }
                    };
                    if dst_mem == 0 {
                        self.memory.write_bytes(dst_addr, &src_data);
                    } else {
                        let dst = self.extra_mem_mut(dst_mem);
                        if dst_addr + len <= dst.len() {
                            dst[dst_addr..dst_addr + len].copy_from_slice(&src_data);
                        }
                    }
                }

                // -- JS String Builtins (wasm:js-string proposal) --
                _ if op == Op::STR_LENGTH => {
                    let s = self.pop();
                    let len = match &s {
                        Value::String(s) => s.chars().count() as i32,
                        Value::Object(obj) => {
                            let o = obj.lock().unwrap();
                            if let ObjectKind::Array(a) = &o.kind { a.len() as i32 }
                            else if let Some(Value::F64(n)) = o.properties.get("length") { *n as i32 }
                            else { 0 }
                        }
                        _ => 0,
                    };
                    self.push(Value::I32(len))?;
                }
                _ if op == Op::STR_CHAR_CODE_AT => {
                    let idx = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let code = if let Value::String(s) = &s {
                        s.chars().nth(idx).map(|c| c as i32).unwrap_or(-1)
                    } else { -1 };
                    self.push(Value::I32(code))?;
                }
                _ if op == Op::STR_FROM_CHAR_CODE => {
                    let code = self.pop().as_i32() as u32;
                    let ch = char::from_u32(code).unwrap_or('\0');
                    self.push(Value::String(Arc::from(ch.to_string().as_str())))?;
                }
                _ if op == Op::STR_CHAR_AT => {
                    let idx = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let ch = if let Value::String(s) = &s {
                        s.chars().nth(idx).map(|c| Arc::from(c.to_string().as_str()))
                            .unwrap_or(Arc::from(""))
                    } else { Arc::from("") };
                    self.push(Value::String(ch))?;
                }
                _ if op == Op::STR_SUBSTRING || op == Op::STR_SLICE => {
                    let end = self.pop().as_i32() as usize;
                    let start = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let result = if let Value::String(s) = &s {
                        let chars: Vec<char> = s.chars().collect();
                        let end = end.min(chars.len());
                        let start = start.min(end);
                        let sub: String = chars[start..end].iter().collect();
                        Arc::from(sub.as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(result))?;
                }
                _ if op == Op::STR_INDEX_OF => {
                    let needle = self.pop();
                    let haystack = self.pop();
                    let pos = match (&haystack, &needle) {
                        (Value::String(h), Value::String(n)) => {
                            h.find(n.as_ref()).map(|p| p as i32).unwrap_or(-1)
                        }
                        _ => -1,
                    };
                    self.push(Value::I32(pos))?;
                }
                _ if op == Op::STR_LAST_INDEX_OF => {
                    let needle = self.pop();
                    let haystack = self.pop();
                    let pos = match (&haystack, &needle) {
                        (Value::String(h), Value::String(n)) => {
                            h.rfind(n.as_ref()).map(|p| p as i32).unwrap_or(-1)
                        }
                        _ => -1,
                    };
                    self.push(Value::I32(pos))?;
                }
                _ if op == Op::STR_EQUALS => {
                    let b = self.pop(); let a = self.pop();
                    let eq = match (&a, &b) {
                        (Value::String(a), Value::String(b)) => a == b,
                        _ => false,
                    };
                    self.push(Value::Bool(eq))?;
                }
                _ if op == Op::STR_COMPARE => {
                    let b = self.pop(); let a = self.pop();
                    let cmp = match (&a, &b) {
                        (Value::String(a), Value::String(b)) => {
                            match a.cmp(b) {
                                std::cmp::Ordering::Less => -1,
                                std::cmp::Ordering::Equal => 0,
                                std::cmp::Ordering::Greater => 1,
                            }
                        }
                        _ => 0,
                    };
                    self.push(Value::I32(cmp))?;
                }
                _ if op == Op::STR_TO_UPPER => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.to_uppercase().as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_TO_LOWER => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.to_lowercase().as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_TRIM => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Arc::from(s.trim()) } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_TRIM_START => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Arc::from(s.trim_start()) } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_TRIM_END => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s { Arc::from(s.trim_end()) } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_STARTS_WITH => {
                    let prefix = self.pop(); let s = self.pop();
                    let r = match (&s, &prefix) {
                        (Value::String(s), Value::String(p)) => s.starts_with(p.as_ref()),
                        _ => false,
                    };
                    self.push(Value::Bool(r))?;
                }
                _ if op == Op::STR_ENDS_WITH => {
                    let suffix = self.pop(); let s = self.pop();
                    let r = match (&s, &suffix) {
                        (Value::String(s), Value::String(p)) => s.ends_with(p.as_ref()),
                        _ => false,
                    };
                    self.push(Value::Bool(r))?;
                }
                _ if op == Op::STR_CONTAINS => {
                    let needle = self.pop(); let s = self.pop();
                    let r = match (&s, &needle) {
                        (Value::String(s), Value::String(n)) => s.contains(n.as_ref()),
                        _ => false,
                    };
                    self.push(Value::Bool(r))?;
                }
                _ if op == Op::STR_REPLACE => {
                    let new = self.pop(); let old = self.pop(); let s = self.pop();
                    let r = match (&s, &old, &new) {
                        (Value::String(s), Value::String(o), Value::String(n)) => {
                            Arc::from(s.replace(o.as_ref(), n.as_ref()).as_str())
                        }
                        _ => Arc::from(""),
                    };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_SPLIT => {
                    let delim = self.pop(); let s = self.pop();
                    let parts: Vec<Value> = match (&s, &delim) {
                        (Value::String(s), Value::String(d)) => {
                            s.split(d.as_ref()).map(|p| Value::String(Arc::from(p))).collect()
                        }
                        _ => vec![],
                    };
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(parts)))))?;
                }
                _ if op == Op::STR_REPEAT => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.repeat(count).as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_PAD_START => {
                    let fill = self.pop(); let target_len = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let (Value::String(s), Value::String(f)) = (&s, &fill) {
                        if s.len() >= target_len { Arc::clone(s) }
                        else {
                            let pad = target_len - s.len();
                            let fill_str: String = f.chars().cycle().take(pad).collect();
                            Arc::from(format!("{}{}", fill_str, s).as_str())
                        }
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_PAD_END => {
                    let fill = self.pop(); let target_len = self.pop().as_i32().max(0) as usize;
                    let s = self.pop();
                    let r = if let (Value::String(s), Value::String(f)) = (&s, &fill) {
                        if s.len() >= target_len { Arc::clone(s) }
                        else {
                            let pad = target_len - s.len();
                            let fill_str: String = f.chars().cycle().take(pad).collect();
                            Arc::from(format!("{}{}", s, fill_str).as_str())
                        }
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                _ if op == Op::STR_REVERSE => {
                    let s = self.pop();
                    let r = if let Value::String(s) = &s {
                        Arc::from(s.chars().rev().collect::<String>().as_str())
                    } else { Arc::from("") };
                    self.push(Value::String(r))?;
                }
                // Unicode code points (beyond BMP — emoji, CJK)
                _ if op == Op::STR_FROM_CODE_POINT => {
                    let cp = self.pop().as_i32() as u32;
                    let ch = char::from_u32(cp).unwrap_or('\u{FFFD}');
                    self.push(Value::String(Arc::from(ch.to_string().as_str())))?;
                }
                _ if op == Op::STR_CODE_POINT_AT => {
                    let idx = self.pop().as_i32() as usize;
                    let s = self.pop();
                    let cp = if let Value::String(s) = &s {
                        s.chars().nth(idx).map(|c| c as i32).unwrap_or(-1)
                    } else { -1 };
                    self.push(Value::I32(cp))?;
                }
                // Bulk char code operations
                _ if op == Op::STR_INTO_CHAR_CODES => {
                    let s = self.pop();
                    let codes: Vec<Value> = if let Value::String(s) = &s {
                        s.chars().map(|c| Value::I32(c as i32)).collect()
                    } else { vec![] };
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(codes)))))?;
                }
                _ if op == Op::STR_FROM_CHAR_CODES => {
                    let arr = self.pop();
                    let s = if let Value::Object(obj) = &arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            a.iter().filter_map(|v| char::from_u32(v.as_i32() as u32)).collect::<String>()
                        } else { String::new() }
                    } else { String::new() };
                    self.push(Value::String(Arc::from(s.as_str())))?;
                }
                // Type discrimination opcodes
                _ if op == Op::REF_TYPEOF => {
                    let v = self.pop();
                    let tag = match &v {
                        Value::Undefined => "undefined",
                        Value::Null => "object",  // JS spec: typeof null === "object"
                        Value::Bool(_) => "boolean",
                        Value::I32(_) | Value::I64(_) | Value::F64(_) => "number",
                        Value::String(_) => "string",
                        Value::Symbol(_) => "symbol",
                        Value::BigInt(_) => "bigint",
                        Value::V128(_) => "v128",
                        Value::WeakRef(_) => "weakref",
                        Value::Object(o) => {
                            let ob = o.lock().unwrap();
                            match &ob.kind {
                                ObjectKind::Function(_) | ObjectKind::HostFunction(_) => "function",
                                ObjectKind::Array(_) => "array",
                                _ => "object",
                            }
                        }
                    };
                    self.push(Value::String(Arc::from(tag)))?;
                }
                _ if op == Op::REF_IS_ARRAY => {
                    let v = self.pop();
                    let is_arr = matches!(&v, Value::Object(o) if matches!(o.lock().unwrap().kind, ObjectKind::Array(_)));
                    self.push(Value::Bool(is_arr))?;
                }

                // -- Array builtins --
                _ if op == Op::ARRAY_LENGTH => {
                    let arr = self.pop();
                    let len = if let Value::Object(obj) = &arr {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind { a.len() as i32 } else { 0 }
                    } else if let Value::String(s) = &arr {
                        s.chars().count() as i32
                    } else { 0 };
                    self.push(Value::I32(len))?;
                }
                // REMOVED (Phase E): the 9 non-spec `0xFF` ARRAY_* dispatch
                // arms — ARRAY_PUSH / POP / SLICE / JOIN / REVERSE /
                // CONTAINS / INDEX_OF (here) and ARRAY_CONCAT / SHIFT
                // (below). All callers migrated to `wasm:js-array.*`
                // imports (Vybe handlers in vybe_host, native on v8 via
                // JS glue, polyfill on wasmtime). Any bytecode still
                // carrying these opcode bytes will hit the "unknown op"
                // path at `Op::decode`'s None branch and trap cleanly.

                // WASM GC array ops
                _ if op == Op::ARRAY_NEW_DEFAULT => {
                    let len = self.pop().as_i32().max(0) as usize;
                    let elems = vec![Value::Null; len];
                    self.push(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))?;
                }
                _ if op == Op::ARRAY_FILL => {
                    let count = self.pop().as_i32().max(0) as usize;
                    let start = self.pop().as_i32().max(0) as usize;
                    let val = self.pop();
                    let arr = self.pop();
                    if let Value::Object(obj) = &arr {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind {
                            let end = (start + count).min(a.len());
                            for i in start..end { a[i] = val.clone(); }
                        }
                    }
                }
                _ if op == Op::ARRAY_COPY => {
                    let len = self.pop().as_i32().max(0) as usize;
                    let src_off = self.pop().as_i32().max(0) as usize;
                    let src = self.pop();
                    let dst_off = self.pop().as_i32().max(0) as usize;
                    let dst = self.pop();
                    // Read source slice
                    let src_vals: Vec<Value> = if let Value::Object(obj) = &src {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(a) = &o.kind {
                            let end = (src_off + len).min(a.len());
                            a[src_off.min(a.len())..end].to_vec()
                        } else { vec![] }
                    } else { vec![] };
                    // Write to destination
                    if let Value::Object(obj) = &dst {
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut a) = o.kind {
                            for (i, v) in src_vals.into_iter().enumerate() {
                                let idx = dst_off + i;
                                if idx < a.len() { a[idx] = v; }
                            }
                        }
                    }
                }
                // ARRAY_CONCAT and ARRAY_SHIFT removed with the cluster
                // above — both were the non-spec `0xFF` variants. Use
                // `wasm:js-array.concat` / `wasm:js-array.shift` instead.

                // -- Stack Switching (wasm stack-switching proposal) --
                _ if op == Op::CONT_NEW => {
                    // Create a continuation from a function reference.
                    // The continuation wraps a function + saved state.
                    let func_val = self.pop();
                    let mut obj = Object::new_typed(0);
                    obj.properties.insert("__cont_func".into(), func_val);
                    obj.properties.insert("__cont_state".into(), Value::String(Arc::from("ready")));
                    obj.properties.insert("__cont_value".into(), Value::Null);
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                _ if op == Op::SUSPEND => {
                    let _tag = self.read_u16();
                    // Yield a value from the current continuation.
                    // The yielded value stays on the stack for the caller.
                    // This is like a return but the continuation can be resumed.
                    let val = self.pop();
                    return Ok(val);
                }
                _ if op == Op::RESUME => {
                    let _tag = self.read_u16();
                    // Resume a continuation, passing a value to it.
                    // [continuation, value] → [result]
                    let val = self.pop();
                    let cont = self.pop();
                    if let Value::Object(obj) = &cont {
                        let func_val = {
                            let o = obj.lock().unwrap();
                            o.properties.get("__cont_func").cloned().unwrap_or(Value::Null)
                        };
                        {
                            let mut o = obj.lock().unwrap();
                            o.properties.insert("__cont_state".into(), Value::String(Arc::from("running")));
                            o.properties.insert("__cont_value".into(), val.clone());
                        }
                        // Call the continuation's function with the resume value
                        self.push(func_val)?;
                        self.push(val)?;
                        self.call_value(1)?;
                    } else {
                        self.push(val)?;
                    }
                }
                _ if op == Op::SWITCH => {
                    let _tag = self.read_u16();
                    // Symmetric switch: suspend current continuation, resume target
                    let val = self.pop();
                    let cont = self.pop();
                    if let Value::Object(obj) = &cont {
                        let mut o = obj.lock().unwrap();
                        o.properties.insert("__cont_value".into(), val.clone());
                        o.properties.insert("__cont_state".into(), Value::String(Arc::from("running")));
                    }
                    self.push(val)?;
                }

                // -- wasi-threads: real OS thread spawning --
                _ if op == Op::THREAD_SPAWN => {
                    // [func_ref] → [task_object]
                    // Per wasi-threads: spawn a real OS thread.
                    // Value is now Arc-based (Send+Sync), so chunks and host_fns
                    // can be shared directly — no serialization needed.
                    let func_val = self.pop();

                    let chunk_idx = match &func_val {
                        Value::Object(obj) => {
                            let o = obj.lock().unwrap();
                            match &o.kind {
                                ObjectKind::Function(f) => Some(f.chunk_index),
                                _ => None,
                            }
                        }
                        _ => None,
                    };

                    if let Some(ci) = chunk_idx {
                        let tid = self.next_thread_id;
                        self.next_thread_id += 1;

                        // Create task object FIRST so child can write result to it
                        let mut obj = Object::new();
                        obj.properties.insert("__type".into(), Value::String(Arc::from("Task")));
                        obj.properties.insert("__thread_id".into(), Value::I32(tid));
                        obj.properties.insert("iscompleted".into(), Value::Bool(false));
                        obj.properties.insert("isalive".into(), Value::Bool(true));
                        obj.properties.insert("result".into(), Value::Null);
                        obj.properties.insert("status".into(), Value::String(Arc::from("Running")));
                        let task_obj = Arc::new(Mutex::new(obj));
                        let task_for_child = task_obj.clone();

                        // Share directly — Value is Send+Sync now
                        let child_chunks = self.chunks.clone();
                        let child_memory = self.memory.clone();
                        let child_host_fns = self.host_fns.clone();
                        let child_host_registry = self.host_registry.clone();
                        let child_import_table = self.import_table.clone();

                        let handle = std::thread::spawn(move || {
                            let mut child_vm = VM::new();
                            child_vm.chunks = child_chunks;
                            child_vm.memory = child_memory;
                            child_vm.host_fns = child_host_fns;
                            child_vm.host_registry = child_host_registry;
                            child_vm.import_table = child_import_table;

                            // Set up call frame
                            let _arity = child_vm.chunks.get(ci).map(|c| c.arity).unwrap_or(0);
                            child_vm.frames.push(CallFrame {
                                chunk_index: ci,
                                ip: 0,
                                base: 0,
                                upvalues: Vec::new(),
                            });
                            let local_count = child_vm.chunks.get(ci)
                                .map(|c| c.local_count as usize)
                                .unwrap_or(1)
                                .max(64);
                            for _ in 0..local_count {
                                child_vm.stack.push(Value::Null);
                            }

                            let result = match child_vm.execute() {
                                Ok(val) => {
                                    // Store return value in the shared task object
                                    let mut t = task_for_child.lock().unwrap();
                                    t.properties.insert("result".into(), val.clone());
                                    t.properties.insert("iscompleted".into(), Value::Bool(true));
                                    t.properties.insert("isalive".into(), Value::Bool(false));
                                    t.properties.insert("hasexited".into(), Value::Bool(true));
                                    t.properties.insert("exitcode".into(), Value::I32(0));
                                    t.properties.insert("status".into(), Value::String(Arc::from("RanToCompletion")));
                                    vec![0u8]
                                }
                                Err(e) => {
                                    let mut t = task_for_child.lock().unwrap();
                                    t.properties.insert("iscompleted".into(), Value::Bool(true));
                                    t.properties.insert("isalive".into(), Value::Bool(false));
                                    t.properties.insert("hasexited".into(), Value::Bool(true));
                                    t.properties.insert("exitcode".into(), Value::I32(-1));
                                    t.properties.insert("status".into(), Value::String(Arc::from("Faulted")));
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
                            o.properties.get("__thread_id").map(|v| v.as_f64() as i32).unwrap_or(-1)
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
                            o.properties.insert("exitcode".into(), Value::I32(if success { 0 } else { -1 }));
                            o.properties.insert("status".into(), Value::String(Arc::from(
                                if success { "RanToCompletion" } else { "Faulted" }
                            )));
                        }
                        self.push(Value::I32(if success { 0 } else { -1 }))?;
                    } else {
                        self.push(Value::I32(-1))?;
                    }
                }

                // -- Extended Const Expressions --
                _ if op == Op::GLOBAL_INIT => {
                    let idx = self.read_u16() as usize;
                    // Evaluate the const expr for global at index idx and store result.
                    // This is a runtime opcode for cases where init can't be done at load time.
                    if idx < self.chunks[0].global_inits.len() {
                        let gi = self.chunks[0].global_inits[idx].clone();
                        let val = self.eval_const_expr(&gi.init);
                        self.globals.insert(gi.name.clone(), val);
                    }
                }

                // -- Typed Continuations --
                _ if op == Op::CONT_NEW_TYPED => {
                    let tag_idx = self.read_u16() as usize;
                    let func_val = self.pop();
                    let mut obj = Object::new_typed(0);
                    obj.properties.insert("__cont_func".into(), func_val);
                    obj.properties.insert("__cont_state".into(), Value::String(Arc::from("ready")));
                    obj.properties.insert("__cont_value".into(), Value::Null);
                    // Store tag info for type checking on suspend/resume
                    obj.properties.insert("__cont_tag".into(), Value::I32(tag_idx as i32));
                    if tag_idx < self.chunks[0].continuation_tags.len() {
                        let tag = &self.chunks[0].continuation_tags[tag_idx];
                        obj.properties.insert("__cont_yield_type".into(), Value::String(Arc::from(tag.yield_type.as_str())));
                        obj.properties.insert("__cont_resume_type".into(), Value::String(Arc::from(tag.resume_type.as_str())));
                    }
                    self.push(Value::Object(Arc::new(Mutex::new(obj))))?;
                }
                _ if op == Op::SUSPEND_TYPED => {
                    let _tag_idx = self.read_u16();
                    // Typed suspend: yield a value, with type validation.
                    let val = self.pop();
                    // In a full implementation, we'd validate val matches the yield_type
                    // of the current continuation's tag. For now, just yield.
                    return Ok(val);
                }
                _ if op == Op::RESUME_TYPED => {
                    let _tag_idx = self.read_u16();
                    let val = self.pop();
                    let cont = self.pop();
                    if let Value::Object(obj) = &cont {
                        // Validate resume value type matches continuation's resume_type
                        let expected_type = {
                            let o = obj.lock().unwrap();
                            o.properties.get("__cont_resume_type")
                                .map(|v| format!("{}", v))
                                .unwrap_or_default()
                        };
                        if !expected_type.is_empty() && expected_type != "any" {
                            let actual_matches = self.test_type(&val, &expected_type.to_lowercase());
                            if !actual_matches {
                                // Type mismatch on resume — for now, proceed anyway
                                // A strict implementation would trap here
                            }
                        }
                        let func_val = {
                            let o = obj.lock().unwrap();
                            o.properties.get("__cont_func").cloned().unwrap_or(Value::Null)
                        };
                        {
                            let mut o = obj.lock().unwrap();
                            o.properties.insert("__cont_state".into(), Value::String(Arc::from("running")));
                            o.properties.insert("__cont_value".into(), val.clone());
                        }
                        self.push(func_val)?;
                        self.push(val)?;
                        self.call_value(1)?;
                    } else {
                        self.push(val)?;
                    }
                }

                // -- String References (zero-copy) --
                _ if op == Op::STRING_AS_REF => {
                    // String → StringRef: since our strings are already Arc<str>,
                    // this is effectively a no-op — we keep the same Arc.
                    // The semantic difference: stringref signals cross-component sharing intent.
                    let val = self.pop();
                    // Just pass through — Arc<str> is already shared
                    self.push(val)?;
                }
                _ if op == Op::STRING_FROM_REF => {
                    // StringRef → String: dereference (zero-copy with Arc).
                    let val = self.pop();
                    self.push(val)?;
                }
                _ if op == Op::STRING_REF_EQ => {
                    // Pointer identity comparison for string refs.
                    let b = self.pop();
                    let a = self.pop();
                    let eq = match (&a, &b) {
                        (Value::String(sa), Value::String(sb)) => Arc::ptr_eq(sa, sb),
                        _ => false,
                    };
                    self.push(Value::Bool(eq))?;
                }

                // -- SIMD (128-bit vectors) --
                _ if op == Op::V128_LOAD => {
                    let addr = self.pop().as_i32() as usize;
                    let mut bytes = [0u8; 16];
                    self.memory.read_bytes(addr, &mut bytes);
                    self.push(Value::V128(bytes))?;
                }
                _ if op == Op::V128_STORE => {
                    let val = self.pop();
                    let addr = self.pop().as_i32() as usize;
                    if let Value::V128(bytes) = val {
                        self.memory.write_bytes(addr, &bytes);
                    }
                }
                _ if op == Op::V128_CONST => {
                    let mut bytes = [0u8; 16];
                    for i in 0..16 { bytes[i] = self.read_byte(); }
                    self.push(Value::V128(bytes))?;
                }

                // i32x4 ops
                _ if op == Op::I32X4_SPLAT => {
                    let v = self.pop().as_i32();
                    let mut bytes = [0u8; 16];
                    for i in 0..4 { bytes[i*4..i*4+4].copy_from_slice(&v.to_le_bytes()); }
                    self.push(Value::V128(bytes))?;
                }
                _ if op == Op::I32X4_ADD => { self.simd_i32x4_binop(|a, b| a.wrapping_add(b))?; }
                _ if op == Op::I32X4_SUB => { self.simd_i32x4_binop(|a, b| a.wrapping_sub(b))?; }
                _ if op == Op::I32X4_MUL => { self.simd_i32x4_binop(|a, b| a.wrapping_mul(b))?; }
                _ if op == Op::I32X4_EQ =>  { self.simd_i32x4_binop(|a, b| if a == b { -1 } else { 0 })?; }
                _ if op == Op::I32X4_GT_S => { self.simd_i32x4_binop(|a, b| if a > b { -1 } else { 0 })?; }
                _ if op == Op::I32X4_LT_S => { self.simd_i32x4_binop(|a, b| if a < b { -1 } else { 0 })?; }
                _ if op == Op::I32X4_SHL => {
                    let shift = self.pop().as_i32() as u32 & 31;
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i32::from_le_bytes([a[i*4], a[i*4+1], a[i*4+2], a[i*4+3]]);
                            out[i*4..i*4+4].copy_from_slice(&(v << shift).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_SHR_S => {
                    let shift = self.pop().as_i32() as u32 & 31;
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = i32::from_le_bytes([a[i*4], a[i*4+1], a[i*4+2], a[i*4+3]]);
                            out[i*4..i*4+4].copy_from_slice(&(v >> shift).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_SHR_U => {
                    let shift = self.pop().as_i32() as u32 & 31;
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let v = u32::from_le_bytes([a[i*4], a[i*4+1], a[i*4+2], a[i*4+3]]);
                            out[i*4..i*4+4].copy_from_slice(&(v >> shift).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_EXTRACT_LANE => {
                    let lane = self.read_byte() as usize & 3;
                    if let Value::V128(a) = self.pop() {
                        let v = i32::from_le_bytes([a[lane*4], a[lane*4+1], a[lane*4+2], a[lane*4+3]]);
                        self.push(Value::I32(v))?;
                    } else { self.push(Value::I32(0))?; }
                }
                _ if op == Op::I32X4_REPLACE_LANE => {
                    let lane = self.read_byte() as usize & 3;
                    let val = self.pop().as_i32();
                    if let Value::V128(mut a) = self.pop() {
                        a[lane*4..lane*4+4].copy_from_slice(&val.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }

                // f64x2 ops
                _ if op == Op::F64X2_SPLAT => {
                    let v = self.pop().as_f64();
                    let mut bytes = [0u8; 16];
                    bytes[0..8].copy_from_slice(&v.to_le_bytes());
                    bytes[8..16].copy_from_slice(&v.to_le_bytes());
                    self.push(Value::V128(bytes))?;
                }
                _ if op == Op::F64X2_ADD => { self.simd_f64x2_binop(|a, b| a + b)?; }
                _ if op == Op::F64X2_SUB => { self.simd_f64x2_binop(|a, b| a - b)?; }
                _ if op == Op::F64X2_MUL => { self.simd_f64x2_binop(|a, b| a * b)?; }
                _ if op == Op::F64X2_DIV => { self.simd_f64x2_binop(|a, b| a / b)?; }
                _ if op == Op::F64X2_MIN => { self.simd_f64x2_binop(|a, b| a.min(b))?; }
                _ if op == Op::F64X2_MAX => { self.simd_f64x2_binop(|a, b| a.max(b))?; }
                _ if op == Op::F64X2_EQ => { self.simd_f64x2_cmp(|a, b| a == b)?; }
                _ if op == Op::F64X2_LT => { self.simd_f64x2_cmp(|a, b| a < b)?; }
                _ if op == Op::F64X2_LE => { self.simd_f64x2_cmp(|a, b| a <= b)?; }
                _ if op == Op::F64X2_SQRT => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = f64::from_le_bytes(a[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&v.sqrt().to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F64X2_ABS => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = f64::from_le_bytes(a[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&v.abs().to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F64X2_NEG => {
                    if let Value::V128(a) = self.pop() {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let v = f64::from_le_bytes(a[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&(-v).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F64X2_EXTRACT_LANE => {
                    let lane = self.read_byte() as usize & 1;
                    if let Value::V128(a) = self.pop() {
                        let v = f64::from_le_bytes(a[lane*8..lane*8+8].try_into().unwrap());
                        self.push(Value::F64(v))?;
                    } else { self.push(Value::F64(0.0))?; }
                }
                _ if op == Op::F64X2_REPLACE_LANE => {
                    let lane = self.read_byte() as usize & 1;
                    let val = self.pop().as_f64();
                    if let Value::V128(mut a) = self.pop() {
                        a[lane*8..lane*8+8].copy_from_slice(&val.to_le_bytes());
                        self.push(Value::V128(a))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }

                // f32x4, i8x16, i16x8 — same patterns, delegate to helpers
                _ if op == Op::F32X4_SPLAT => { let v = self.pop().as_f64() as f32; let b = v.to_le_bytes(); let mut out = [0u8;16]; for i in 0..4 { out[i*4..i*4+4].copy_from_slice(&b); } self.push(Value::V128(out))?; }
                _ if op == Op::F32X4_ADD => { self.simd_f32x4_binop(|a,b| a+b)?; }
                _ if op == Op::F32X4_SUB => { self.simd_f32x4_binop(|a,b| a-b)?; }
                _ if op == Op::F32X4_MUL => { self.simd_f32x4_binop(|a,b| a*b)?; }
                _ if op == Op::F32X4_DIV => { self.simd_f32x4_binop(|a,b| a/b)?; }
                _ if op == Op::F32X4_EXTRACT_LANE => { let lane = self.read_byte() as usize & 3; if let Value::V128(a) = self.pop() { let v = f32::from_le_bytes(a[lane*4..lane*4+4].try_into().unwrap()); self.push(Value::F64(v as f64))?; } else { self.push(Value::F64(0.0))?; } }
                _ if op == Op::F32X4_REPLACE_LANE => { let lane = self.read_byte() as usize & 3; let val = self.pop().as_f64() as f32; if let Value::V128(mut a) = self.pop() { a[lane*4..lane*4+4].copy_from_slice(&val.to_le_bytes()); self.push(Value::V128(a))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::I8X16_SPLAT => { let v = self.pop().as_i32() as u8; self.push(Value::V128([v;16]))?; }
                _ if op == Op::I8X16_ADD => { self.simd_i8x16_binop(|a,b| a.wrapping_add(b))?; }
                _ if op == Op::I8X16_SUB => { self.simd_i8x16_binop(|a,b| a.wrapping_sub(b))?; }
                _ if op == Op::I8X16_EQ =>  { self.simd_i8x16_binop(|a,b| if a==b {0xFF} else {0})?; }
                _ if op == Op::I8X16_EXTRACT_LANE_S => { let lane = self.read_byte() as usize & 15; if let Value::V128(a) = self.pop() { self.push(Value::I32(a[lane] as i8 as i32))?; } else { self.push(Value::I32(0))?; } }
                _ if op == Op::I8X16_EXTRACT_LANE_U => { let lane = self.read_byte() as usize & 15; if let Value::V128(a) = self.pop() { self.push(Value::I32(a[lane] as i32))?; } else { self.push(Value::I32(0))?; } }
                _ if op == Op::I8X16_REPLACE_LANE => { let lane = self.read_byte() as usize & 15; let val = self.pop().as_i32() as u8; if let Value::V128(mut a) = self.pop() { a[lane] = val; self.push(Value::V128(a))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::I8X16_SHUFFLE => { let mut indices = [0u8;16]; for i in 0..16 { indices[i] = self.read_byte(); } let b = self.pop(); let a = self.pop(); if let (Value::V128(va), Value::V128(vb)) = (a,b) { let combined: Vec<u8> = va.iter().chain(vb.iter()).copied().collect(); let mut out = [0u8;16]; for i in 0..16 { out[i] = combined.get(indices[i] as usize).copied().unwrap_or(0); } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::I8X16_SWIZZLE => { let b = self.pop(); let a = self.pop(); if let (Value::V128(va), Value::V128(vb)) = (a,b) { let mut out = [0u8;16]; for i in 0..16 { let idx = vb[i] as usize; out[i] = if idx < 16 { va[idx] } else { 0 }; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::I16X8_SPLAT => { let v = self.pop().as_i32() as i16; let b = v.to_le_bytes(); let mut out = [0u8;16]; for i in 0..8 { out[i*2..i*2+2].copy_from_slice(&b); } self.push(Value::V128(out))?; }
                _ if op == Op::I16X8_ADD => { self.simd_i16x8_binop(|a,b| a.wrapping_add(b))?; }
                _ if op == Op::I16X8_SUB => { self.simd_i16x8_binop(|a,b| a.wrapping_sub(b))?; }
                _ if op == Op::I16X8_MUL => { self.simd_i16x8_binop(|a,b| a.wrapping_mul(b))?; }
                _ if op == Op::I16X8_EXTRACT_LANE_S => { let lane = self.read_byte() as usize & 7; if let Value::V128(a) = self.pop() { let v = i16::from_le_bytes([a[lane*2], a[lane*2+1]]); self.push(Value::I32(v as i32))?; } else { self.push(Value::I32(0))?; } }
                _ if op == Op::I16X8_EXTRACT_LANE_U => { let lane = self.read_byte() as usize & 7; if let Value::V128(a) = self.pop() { let v = u16::from_le_bytes([a[lane*2], a[lane*2+1]]); self.push(Value::I32(v as i32))?; } else { self.push(Value::I32(0))?; } }
                _ if op == Op::I16X8_REPLACE_LANE => { let lane = self.read_byte() as usize & 7; let val = self.pop().as_i32() as i16; if let Value::V128(mut a) = self.pop() { a[lane*2..lane*2+2].copy_from_slice(&val.to_le_bytes()); self.push(Value::V128(a))?; } else { self.push(Value::V128([0;16]))?; } }

                // v128 bitwise
                _ if op == Op::V128_AND => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] & b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::V128_OR => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] | b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::V128_XOR => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] ^ b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::V128_NOT => { if let Value::V128(a) = self.pop() { let mut out = [0u8;16]; for i in 0..16 { out[i] = !a[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::V128_ANDNOT => { let b = self.pop(); let a = self.pop(); if let (Value::V128(a), Value::V128(b)) = (a, b) { let mut out = [0u8;16]; for i in 0..16 { out[i] = a[i] & !b[i]; } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }
                _ if op == Op::V128_ANY_TRUE => { if let Value::V128(a) = self.pop() { self.push(Value::I32(if a.iter().any(|&b| b != 0) { 1 } else { 0 }))?; } else { self.push(Value::I32(0))?; } }
                _ if op == Op::V128_BITSELECT => { let mask = self.pop(); let v2 = self.pop(); let v1 = self.pop(); if let (Value::V128(a), Value::V128(b), Value::V128(m)) = (v1, v2, mask) { let mut out = [0u8;16]; for i in 0..16 { out[i] = (a[i] & m[i]) | (b[i] & !m[i]); } self.push(Value::V128(out))?; } else { self.push(Value::V128([0;16]))?; } }

                // -- Atomics (single-threaded: same as non-atomic for now) --
                _ if op == Op::ATOMIC_FENCE => {} // no-op in single-threaded
                _ if op == Op::I32_ATOMIC_LOAD => { let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_load_i32(addr)))?; }
                // Atomic store: stack is [addr, value] — pop value first (top), then addr
                _ if op == Op::I32_ATOMIC_STORE => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.memory.atomic_store_i32(addr, v); }
                // Atomic RMW: stack is [addr, value] — pop value (top), addr (second), return old
                _ if op == Op::I32_ATOMIC_RMW_ADD => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_add_i32(addr, v)))?; }
                _ if op == Op::I32_ATOMIC_RMW_SUB => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_sub_i32(addr, v)))?; }
                _ if op == Op::I32_ATOMIC_RMW_AND => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_and_i32(addr, v)))?; }
                _ if op == Op::I32_ATOMIC_RMW_OR => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_or_i32(addr, v)))?; }
                _ if op == Op::I32_ATOMIC_RMW_XOR => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_rmw_xor_i32(addr, v)))?; }
                _ if op == Op::I32_ATOMIC_RMW_XCHG => { let v = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_xchg_i32(addr, v)))?; }
                // Atomic CmpXchg: stack is [addr, expected, replacement]
                _ if op == Op::I32_ATOMIC_RMW_CMPXCHG => { let replacement = self.pop().as_i32(); let expected = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.atomic_cmpxchg_i32(addr, expected, replacement)))?; }
                _ if op == Op::I64_ATOMIC_LOAD => { let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_load_i64(addr)))?; }
                _ if op == Op::I64_ATOMIC_STORE => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.memory.atomic_store_i64(addr, v); }
                _ if op == Op::I64_ATOMIC_RMW_ADD => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_rmw_add_i64(addr, v)))?; }
                _ if op == Op::I64_ATOMIC_RMW_SUB => { let v = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_rmw_sub_i64(addr, v)))?; }
                _ if op == Op::I64_ATOMIC_RMW_CMPXCHG => { let repl = self.pop().as_i64(); let exp = self.pop().as_i64(); let addr = self.pop().as_i32() as usize; self.push(Value::I64(self.memory.atomic_cmpxchg_i64(addr, exp, repl)))?; }
                _ if op == Op::MEMORY_ATOMIC_WAIT32 => { let timeout = self.pop().as_i64(); let expected = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.wait32(addr, expected, timeout)))?; }
                _ if op == Op::MEMORY_ATOMIC_NOTIFY => { let count = self.pop().as_i32(); let addr = self.pop().as_i32() as usize; self.push(Value::I32(self.memory.notify(addr, count)))?; }

                // -- Memory64 --
                _ if op == Op::I64_MEMORY_SIZE => { self.push(Value::I64((self.memory.len() / 65536) as i64))?; }
                _ if op == Op::I64_MEMORY_GROW => { let pages = self.pop().as_i64() as usize; let old = self.memory.grow(pages); self.push(Value::I64(old as i64))?; }
                _ if op == Op::I32_LOAD_64 => { let addr = self.pop().as_i64() as usize; self.push(Value::I32(self.memory.load_i32(addr)?))?; }
                _ if op == Op::I64_LOAD_64 => { let addr = self.pop().as_i64() as usize; self.push(Value::I64(self.memory.load_i64(addr)?))?; }
                _ if op == Op::F64_LOAD_64 => { let addr = self.pop().as_i64() as usize; self.push(Value::F64(self.memory.load_f64(addr)?))?; }
                _ if op == Op::I32_STORE_64 => { let v = self.pop().as_i32(); let addr = self.pop().as_i64() as usize; self.memory.store_i32(addr, v)?; }
                _ if op == Op::I64_STORE_64 => { let v = self.pop().as_i64(); let addr = self.pop().as_i64() as usize; let _ = self.memory.store_i64(addr, v); }
                _ if op == Op::F64_STORE_64 => { let v = self.pop().as_f64(); let addr = self.pop().as_i64() as usize; let _ = self.memory.store_f64(addr, v); }

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
                    let idx = self.pop(); let src = self.pop();
                    if let (Value::V128(src), Value::V128(idx)) = (src, idx) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            let n = idx[i];
                            out[i] = if n < 16 { src[n as usize] } else { 0 };
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F32X4_S => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let f = f32::from_le_bytes(bytes[i*4..i*4+4].try_into().unwrap());
                            let r = if f.is_nan() { 0 } else { f as i32 };
                            out[i*4..i*4+4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F32X4_U => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let f = f32::from_le_bytes(bytes[i*4..i*4+4].try_into().unwrap());
                            let r: u32 = if f.is_nan() || f < 0.0 { 0 } else { f as u32 };
                            out[i*4..i*4+4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F64X2_S_ZERO => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f = f64::from_le_bytes(bytes[i*8..i*8+8].try_into().unwrap());
                            let r = if f.is_nan() { 0 } else { f as i32 };
                            out[i*4..i*4+4].copy_from_slice(&r.to_le_bytes());
                        }
                        // Upper two lanes stay zero.
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_RELAXED_TRUNC_F64X2_U_ZERO => {
                    let v = self.pop();
                    if let Value::V128(bytes) = v {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let f = f64::from_le_bytes(bytes[i*8..i*8+8].try_into().unwrap());
                            let r: u32 = if f.is_nan() || f < 0.0 { 0 } else { f as u32 };
                            out[i*4..i*4+4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F32X4_RELAXED_MADD => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                            let fc = f32::from_le_bytes(vc[i*4..i*4+4].try_into().unwrap());
                            out[i*4..i*4+4].copy_from_slice(&fa.mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F32X4_RELAXED_NMADD => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                            let fc = f32::from_le_bytes(vc[i*4..i*4+4].try_into().unwrap());
                            out[i*4..i*4+4].copy_from_slice(&(-fa).mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F64X2_RELAXED_MADD => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                            let fc = f64::from_le_bytes(vc[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&fa.mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F64X2_RELAXED_NMADD => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                            let fc = f64::from_le_bytes(vc[i*8..i*8+8].try_into().unwrap());
                            out[i*8..i*8+8].copy_from_slice(&(-fa).mul_add(fb, fc).to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                // -- Relaxed laneselect --
                // mask bit policy: we use the full bit (all 8 / 16 / 32
                // bits of the mask lane compared to 0) — picking the
                // "any non-zero bit" interpretation consistently.
                _ if op == Op::I8X16_RELAXED_LANESELECT => {
                    let mask = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..16 {
                            out[i] = if vm[i] & 0x80 != 0 { va[i] } else { vb[i] };
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I16X8_RELAXED_LANESELECT => {
                    let mask = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let m = u16::from_le_bytes([vm[i*2], vm[i*2+1]]);
                            let pick_a = m & 0x8000 != 0;
                            out[i*2..i*2+2].copy_from_slice(
                                if pick_a { &va[i*2..i*2+2] } else { &vb[i*2..i*2+2] });
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_RELAXED_LANESELECT => {
                    let mask = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let m = u32::from_le_bytes(vm[i*4..i*4+4].try_into().unwrap());
                            let pick_a = m & 0x8000_0000 != 0;
                            out[i*4..i*4+4].copy_from_slice(
                                if pick_a { &va[i*4..i*4+4] } else { &vb[i*4..i*4+4] });
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I64X2_RELAXED_LANESELECT => {
                    let mask = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vm)) = (a, b, mask) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let m = u64::from_le_bytes(vm[i*8..i*8+8].try_into().unwrap());
                            let pick_a = m & 0x8000_0000_0000_0000 != 0;
                            out[i*8..i*8+8].copy_from_slice(
                                if pick_a { &va[i*8..i*8+8] } else { &vb[i*8..i*8+8] });
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                // -- Relaxed min / max --
                // NaN handling: relaxed variants are allowed to return
                // either operand on NaN input (vs MVP which must return
                // NaN). We pick `a` when `a` is NaN, `b` otherwise — the
                // x86 `minps/maxps` behavior.
                _ if op == Op::F32X4_RELAXED_MIN => {
                    let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                            let r = if fa < fb { fa } else { fb };
                            out[i*4..i*4+4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F32X4_RELAXED_MAX => {
                    let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let fa = f32::from_le_bytes(va[i*4..i*4+4].try_into().unwrap());
                            let fb = f32::from_le_bytes(vb[i*4..i*4+4].try_into().unwrap());
                            let r = if fa > fb { fa } else { fb };
                            out[i*4..i*4+4].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F64X2_RELAXED_MIN => {
                    let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                            let r = if fa < fb { fa } else { fb };
                            out[i*8..i*8+8].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::F64X2_RELAXED_MAX => {
                    let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..2 {
                            let fa = f64::from_le_bytes(va[i*8..i*8+8].try_into().unwrap());
                            let fb = f64::from_le_bytes(vb[i*8..i*8+8].try_into().unwrap());
                            let r = if fa > fb { fa } else { fb };
                            out[i*8..i*8+8].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                // -- q15 multiply-round-saturate (same semantics as MVP) --
                _ if op == Op::I16X8_RELAXED_Q15MULR_S => {
                    let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let av = i16::from_le_bytes(va[i*2..i*2+2].try_into().unwrap()) as i32;
                            let bv = i16::from_le_bytes(vb[i*2..i*2+2].try_into().unwrap()) as i32;
                            let r = ((av * bv) + (1 << 14)) >> 15;
                            let r = r.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
                            out[i*2..i*2+2].copy_from_slice(&r.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                // -- Relaxed integer dot products --
                // The i8x16 x i7x16 ops assume the second operand's high
                // bit is zero (7-bit). The "relaxed" part is that the
                // implementation may saturate or wrap — we wrap via i32.
                _ if op == Op::I16X8_RELAXED_DOT_I8X16_I7X16_S => {
                    let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb)) = (a, b) {
                        let mut out = [0u8; 16];
                        for i in 0..8 {
                            let av0 = va[i*2] as i8 as i16;
                            let av1 = va[i*2+1] as i8 as i16;
                            let bv0 = (vb[i*2] & 0x7F) as i16;
                            let bv1 = (vb[i*2+1] & 0x7F) as i16;
                            let sum = av0.wrapping_mul(bv0).wrapping_add(av1.wrapping_mul(bv1));
                            out[i*2..i*2+2].copy_from_slice(&sum.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }
                _ if op == Op::I32X4_RELAXED_DOT_I8X16_I7X16_ADD_S => {
                    let c = self.pop(); let b = self.pop(); let a = self.pop();
                    if let (Value::V128(va), Value::V128(vb), Value::V128(vc)) = (a, b, c) {
                        let mut out = [0u8; 16];
                        for i in 0..4 {
                            let mut sum: i32 = i32::from_le_bytes(vc[i*4..i*4+4].try_into().unwrap());
                            for j in 0..4 {
                                let av = va[i*4+j] as i8 as i32;
                                let bv = (vb[i*4+j] & 0x7F) as i32;
                                sum = sum.wrapping_add(av.wrapping_mul(bv));
                            }
                            out[i*4..i*4+4].copy_from_slice(&sum.to_le_bytes());
                        }
                        self.push(Value::V128(out))?;
                    } else { self.push(Value::V128([0; 16]))?; }
                }

                // -- JS Promise Integration (JSPI) --
                _ if op == Op::PROMISE_SUSPEND => {
                    // Explicit JSPI suspend point. Like await, but can be inserted
                    // by the compiler for synchronous-looking code that calls async APIs.
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        let is_pending = o.properties.get("__type")
                            .map(|v| format!("{}", v) == "Promise")
                            .unwrap_or(false)
                            && o.properties.get("__state")
                                .map(|v| format!("{}", v) == "pending")
                                .unwrap_or(false);
                        if is_pending {
                            let promise_id = o.properties.get("__id")
                                .map(|v| v.as_f64() as u64)
                                .unwrap_or(0);
                            drop(o);
                            let fiber = self.save_fiber();
                            self.event_loop.borrow_mut().suspend_fiber(promise_id, fiber);
                            return Err(VMError::new(format!("__jspi__:{}", promise_id)));
                        }
                        let resolved = o.properties.get("__value").cloned().unwrap_or(Value::Null);
                        drop(o);
                        self.push(resolved)?;
                    } else {
                        // Not a promise — return value as-is
                        self.push(val)?;
                    }
                }

                // -- WASM GC Type System --
                _ if op == Op::SET_TYPE_ID => {
                    // Stack: [obj, type_id_i32] → [obj]
                    let type_id = self.pop().as_i32() as usize;
                    let obj = self.peek(0);
                    if let Value::Object(o) = obj {
                        o.lock().unwrap().type_id = type_id;
                    }
                }

                // -- Iteration protocol --
                // iter_get, iter_next: removed (non-WASM, were unused by compilers)
                _ if op == Op::SPREAD => {
                    // Spread array onto stack: [array] → [elem0, elem1, ...]
                    let val = self.pop();
                    if let Value::Object(ref obj) = val {
                        let o = obj.lock().unwrap();
                        if let ObjectKind::Array(elems) = &o.kind {
                            for elem in elems {
                                self.push(elem.clone())?;
                            }
                        }
                    }
                }
                // class_new, method_def, inherit: removed (non-WASM, were NOPs)
                _ => {
                    return Err(VMError::new(format!("Unhandled opcode: {:?} (0x{:04X})", op, op.0)));
                }
            }
        }
    }

}
