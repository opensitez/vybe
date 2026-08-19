//! The ONE lowering for [`vybe_ast::AtomicOp`] — the atomic vocabulary's
//! `async_ops.rs`. Every language's spelling (C# `Interlocked`, C
//! `atomic_fetch_*`, Go `sync/atomic`, Java/Kotlin `AtomicInteger`, Pascal
//! `TInterlocked`) normalizes to the node in its walker; this file turns the
//! node into the WASM threads-proposal opcodes and nothing else.
//!
//! The PLACE operand is the load-bearing part. An atomic acts on a word in
//! SHARED linear memory, so the place must resolve to an ADDRESS — which is
//! exactly what the shared-word promotion provides: a binding named as an
//! atomic place is wrapped in `{__ref_kind:"shared", __addr}` at declaration
//! (`collect_atomic_place_idents` → `references::emit_shared_word_new`), and
//! this lowering reads `__addr` off that reference. A place that was never
//! promoted is a COMPILE error here, not a runtime misread: handing the
//! atomics a value where an address belongs is precisely the trap
//! (`addr=100 limit=0`) this module exists to end.

use std::sync::Arc;

use vybe_ast::{AtomicOp, AtomicRmw, ExprKind, Expression, RmwResult};
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

use crate::primitives::pointers::SHARED_ADDR_KEY;
use crate::primitives::{Compiler, threading};

impl Compiler {
    pub(super) fn emit_atomic(&mut self, op: &AtomicOp) -> Result<(), String> {
        let line = self.line;
        match op {
            AtomicOp::Load { place, .. } => {
                self.compile_atomic_place_addr(place)?;
                threading::emit_atomic_load(self.chunk(), line);
            }
            AtomicOp::Store { place, value, .. } => {
                self.compile_atomic_place_addr(place)?;
                self.compile_expr(value)?;
                threading::emit_atomic_store(self.chunk(), line);
                // A store is a statement-shaped operation but this node sits
                // in expression position; its value is the stored value in
                // every language that gives it one (C `atomic_store` is void,
                // Java `set` is void) — null keeps the stack balanced.
                self.emit_null();
            }
            AtomicOp::Rmw {
                op: rmw,
                place,
                operand,
                result,
                ..
            } => {
                // Operand FIRST — its expression may itself read the shared
                // word (`Interlocked.Add(ref x, f(x))`), and the RMW must see
                // a fully evaluated value.
                self.compile_expr(operand)?;
                let operand_slot = self.define_local("__atomic_rmw_operand");
                self.emit_u16(Op::LOCAL_SET, operand_slot);
                self.compile_atomic_place_addr(place)?;
                self.emit_u16(Op::LOCAL_GET, operand_slot);
                match rmw {
                    AtomicRmw::Add => threading::emit_atomic_add(self.chunk(), line),
                    AtomicRmw::Sub => threading::emit_atomic_sub(self.chunk(), line),
                    AtomicRmw::And => threading::emit_atomic_and(self.chunk(), line),
                    AtomicRmw::Or => threading::emit_atomic_or(self.chunk(), line),
                    AtomicRmw::Xor => threading::emit_atomic_xor(self.chunk(), line),
                    AtomicRmw::Xchg => threading::emit_atomic_xchg(self.chunk(), line),
                }
                // WASM's RMW always yields the OLD value; `result` is the
                // walker's declaration of which one the language's function
                // returns. The NEW value is derived from old + operand — never
                // by re-reading the word, which another thread may have
                // changed since.
                if *result == RmwResult::New {
                    match rmw {
                        AtomicRmw::Add => {
                            self.emit_u16(Op::LOCAL_GET, operand_slot);
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                        }
                        AtomicRmw::Sub => {
                            self.emit_u16(Op::LOCAL_GET, operand_slot);
                            crate::primitives::ops::emit_dyn_neg(self.chunk(), line);
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                        }
                        AtomicRmw::Xchg => {
                            // The new value IS the operand.
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, operand_slot);
                        }
                        AtomicRmw::And | AtomicRmw::Or | AtomicRmw::Xor => {
                            // No surveyed language returns the new value for a
                            // bitwise RMW (C and Java both return old). Refuse
                            // rather than guess.
                            return Err(format!(
                                "atomic {rmw:?} with RmwResult::New has no consumer and no \
                                 lowering — the surveyed languages all return the old value"
                            ));
                        }
                    }
                }
            }
            AtomicOp::CompareExchange {
                place,
                expected,
                replacement,
                result,
                ..
            } => {
                if *result == RmwResult::New {
                    return Err(
                        "atomic CompareExchange yields the ORIGINAL value in every surveyed \
                         language (.NET, C, Java); RmwResult::New is unmapped on purpose"
                            .into(),
                    );
                }
                self.compile_expr(expected)?;
                let expected_slot = self.define_local("__atomic_cas_expected");
                self.emit_u16(Op::LOCAL_SET, expected_slot);
                self.compile_expr(replacement)?;
                let replacement_slot = self.define_local("__atomic_cas_replacement");
                self.emit_u16(Op::LOCAL_SET, replacement_slot);
                self.compile_atomic_place_addr(place)?;
                self.emit_u16(Op::LOCAL_GET, expected_slot);
                self.emit_u16(Op::LOCAL_GET, replacement_slot);
                threading::emit_atomic_cmpxchg(self.chunk(), line);
            }
            AtomicOp::Fence { .. } => {
                threading::emit_atomic_fence(self.chunk(), line);
                self.emit_null();
            }
        }
        Ok(())
    }

    /// Push the linear-memory ADDRESS of an atomic place.
    ///
    /// The place is a name whose binding holds a shared-word reference — the
    /// promotion at its declaration guarantees that, and this asks the same
    /// binding facts the deref dispatchers ask. Reading the binding RAW
    /// (no autoderef) and taking `__addr` off it is the one place the
    /// reference object itself is wanted rather than the word's value.
    fn compile_atomic_place_addr(&mut self, place: &Expression) -> Result<(), String> {
        let ExprKind::Ident(name) = &place.kind else {
            return Err(format!(
                "atomic place must be a variable for now (fields and elements \
                 need a member address, not built yet): got {:?}",
                place.kind
            ));
        };
        if let Some(slot) = self.scope().resolve(name) {
            if !self.scope().holds_reference(name).unwrap_or(false) {
                return Err(format!(
                    "atomic place '{name}' was not promoted to shared-word storage — \
                     the pre-scan (`collect_atomic_place_idents`) missed its declaration"
                ));
            }
            self.emit_u16(Op::LOCAL_GET, slot);
        } else {
            if !self.module_atomic_word_globals.contains(name) {
                return Err(format!(
                    "atomic place '{name}' resolves to a global the module scan never \
                     marked — its declaration cannot have been promoted"
                ));
            }
            let key = self.variable_global_binding_key(name);
            self.emit_global_read(&key);
        }
        let addr_key = self
            .chunk()
            .add_constant(Value::String(Arc::from(SHARED_ADDR_KEY)));
        self.emit_struct_field_op(Op::STRUCT_GET, 0, addr_key);
        Ok(())
    }
}
