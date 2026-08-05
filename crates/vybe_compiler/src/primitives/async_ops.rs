//! ONE lowering for the async model — `vybe_ast::AsyncOp` → ECMA §27.2 + JSPI.
//!
//! Walkers normalize seventeen spellings (`Promise.resolve`, `Task.FromResult`,
//! `asyncio.gather`, `Future.wait`, …) into the common vocabulary; this file is
//! the only place that vocabulary is turned into bytecode. The targets are the
//! two layers the architecture allows: the `ecma:promise` host surface
//! (ECMA-262 §27.2 — the SAME objects JS reaches at runtime, so cross-language
//! async values interoperate by construction) and the JSPI suspend import for
//! the sync↔async boundary. Timing is the module's declared
//! nothing here branches on a language.
//!
//! This replaces `promises.rs`, whose emitters had ZERO callers — `.then` in
//! JS only ever worked through runtime member dispatch, and every other
//! language's async API fell through to missing globals ("undefined is not
//! callable", measured 2026-08-04 as csharp_async_await_flow 0/5).

use super::*;
use vybe_ast::{AsyncOp, JoinMode};

impl Compiler {
    pub(super) fn emit_async(&mut self, op: &AsyncOp) -> Result<(), String> {
        let line = self.line;
        match op {
            // §27.2.4.7 PromiseResolve — also ADOPTS thenables, which is what
            // makes `create_task(coro)` / `ensure_future` correct through the
            // same node: an already-started async body is a thenable.
            AsyncOp::Resolved(e) => {
                self.compile_expr(e)?;
                let idx = self.import("ecma:promise", "resolve");
                self.emit_host_call(idx, 1);
            }
            AsyncOp::Rejected(e) => {
                self.compile_expr(e)?;
                let idx = self.import("ecma:promise", "reject");
                self.emit_host_call(idx, 1);
            }
            // §27.2.5.4 — receiver-first host convention: then(p, ok, err).
            AsyncOp::Continue { source, on_fulfilled, on_rejected } => {
                self.compile_expr(source)?;
                match on_fulfilled {
                    Some(f) => self.compile_expr(f)?,
                    None => inst!(self, core_wasm::undefined) }
                match on_rejected {
                    Some(f) => self.compile_expr(f)?,
                    None => inst!(self, core_wasm::undefined) }
                let idx = self.import("ecma:promise", "then");
                self.emit_host_call(idx, 3);
            }
            // §27.2.5.3 finally(p, f).
            AsyncOp::Cleanup { source, on_settled } => {
                self.compile_expr(source)?;
                self.compile_expr(on_settled)?;
                let idx = self.import("ecma:promise", "finally");
                self.emit_host_call(idx, 2);
            }
            // §27.2.4.1-4.5 — one array argument.
            AsyncOp::Join { mode, sources } => {
                for source in sources {
                    self.compile_expr(source)?;
                }
                common::collections::emit_array_new(
                    &mut self.chunks,
                    self.current,
                    sources.len() as u16,
                    line,
                );
                let name = match mode {
                    JoinMode::All => "all",
                    JoinMode::AllSettled => "allSettled",
                    JoinMode::Race => "race",
                    JoinMode::Any => "any" };
                let idx = self.import("ecma:promise", name);
                self.emit_host_call(idx, 1);
            }
            // Scheduled work: resolve(undefined).then(f) — f runs as a JOB
            // under the module's declared discipline, and the result is an
            // async value. No thread is implied; `Spawn` promises ordering,
            // not parallelism.
            AsyncOp::Spawn(f) => {
                inst!(self, core_wasm::undefined);
                let resolve = self.import("ecma:promise", "resolve");
                self.emit_host_call(resolve, 1);
                self.compile_expr(f)?;
                inst!(self, core_wasm::undefined);
                let then = self.import("ecma:promise", "then");
                self.emit_host_call(then, 3);
            }
            // A time-deferred async value: `withResolvers()` mints the
            // promise (§27.2.4.8), the HOST's timer surface fires its
            // `resolve` — the vocabulary says "later", the host owns time.
            AsyncOp::Sleep(duration) => {
                let wr = self.import("ecma:promise", "withResolvers");
                self.emit_host_call(wr, 0);
                let wr_slot = self.define_local("__sleep_resolvers");
                self.emit_u16(Op::LOCAL_TEE, wr_slot);
                let resolve_key = self.str_const("resolve");
                self.emit_struct_field_op(Op::STRUCT_GET, 0, resolve_key);
                self.compile_expr(duration)?;
                let set_timeout = self.import("web:timers", "setTimeout");
                self.emit_host_call(set_timeout, 2);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, wr_slot);
                let promise_key = self.str_const("promise");
                self.emit_struct_field_op(Op::STRUCT_GET, 0, promise_key);
            }
            // Eager-continuation await — a DIFFERENT suspending import from
            // ECMA's `jspi.await`. The semantics were chosen at NORMALIZATION
            // (the C# walker put this node here); the bytecode carries the
            // decision, and the VM implements two spec-shaped operations
            // without consulting any per-module property.
            AsyncOp::AwaitEager(source) => {
                self.compile_expr(source)?;
                let idx = self.import(
                    crate::primitives::functions::JSPI_SUSPEND_MODULE,
                    "await_eager",
                );
                self.emit_host_call(idx, 1);
            }
            // The sync↔async boundary. Same JSPI suspend `await` lowers to —
            // in a fiber, "drive the loop until settled" IS suspending this
            // fiber and letting the scheduler run; the fiber resumes with the
            // value (or the rejection thrown).
            AsyncOp::BlockOn(source) => {
                self.compile_expr(source)?;
                crate::primitives::functions::emit_await(self.chunk(), line);
            }
        }
        Ok(())
    }
}
