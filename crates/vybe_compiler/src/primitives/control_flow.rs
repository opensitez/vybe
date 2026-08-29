//! Generators, for-of iterators, multi-return, pointer-cell promotion, finally/iterator unwinding.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use crate::primitives::class_slots;
use super::*;

impl Compiler {
    /// Terminate the program with `status`, from any depth.
    ///
    /// THE exit primitive — every language's spelling lowers here: PHP
    /// `exit`/`die`, Python `sys.exit`, Ruby `exit`/`exit!`, Lua `os.exit`, Go
    /// `os.Exit`, Java `System.exit`, Pascal `Halt`, COBOL `STOP RUN`, JS
    /// `process.exit`, C `exit`. It belongs beside `Return`/`Break`/`Throw`
    /// because it is the same kind of thing: a non-local transfer that unwinds
    /// every frame. Before this it was expressed three different ways — a
    /// profile builtin bound straight to the host, a hardcoded arm behind a
    /// language-name check, and a walker rewriting it to a plain `Return` —
    /// and only one of the three flushed output buffers.
    ///
    /// WASI does the terminating: `wasi:cli/exit.exit-with-code` ends the guest
    /// instance and hands the status back to the embedder, which is what the
    /// component model requires. The VM never calls `process::exit` itself —
    /// that would tear down a test binary or a server on the first `exit`.
    ///
    /// Two things every language needs and none should reimplement:
    ///   * open output buffers are FLUSHED first. Ending the run skips the
    ///     module epilogue, which is where `emit_ob_flush_all` normally runs, so
    ///     `ob_start(); echo 'x'; exit(0);` printed nothing until this moved
    ///     here.
    ///   * a missing status is 0.
    ///
    /// What DIFFERS per language is only how the argument is spelled — Lua's
    /// `os.exit(true)` means success, PHP's string argument is a message
    /// printed with status 0, Python's goes to stderr with status 1. That is
    /// normalization, and it belongs in each walker, not here.
    ///
    /// Stack: [status?] → [] (never returns at runtime).
    pub(super) fn emit_exit_from_stack(&mut self, argc: u8) -> Result<(), String> {
        let line = self.line;
        if argc == 0 {
            self.emit_const(Value::F64(0.0));
        } else {
            // Trailing arguments are not part of the status (Lua's
            // `os.exit(code, close)`); the status is the FIRST, so drop back
            // down to it.
            for _ in 1..argc {
                self.emit(Op::DROP);
            }
        }
        let status_slot = self.define_local("__exit_status");
        self.emit_u16(Op::LOCAL_SET, status_slot);

        // One syntax, two meanings — see `Directives::exit_argument`. php's
        // `die("bye")` prints a farewell MESSAGE and exits 0, while `exit(3)`
        // exits with STATUS 3 and prints nothing; which one it is depends on
        // the value's TYPE, so it cannot be decided by a lowering. Every other
        // language leaves this `None` and the argument is simply the status.
        //
        // This replaced two `profile.name == "php"` checks that matched the
        // SPELLINGS `exit`/`die` in shared code.
        if argc >= 1
            && self.directives().exit_argument == Some(vybe_ast::ExitArgument::MessageOrStatus)
        {
            // `status_slot` currently holds the ARGUMENT.
            self.emit_u16(Op::LOCAL_GET, status_slot);
            let typeof_idx = self.import("ecma:value", "typeof");
            self.emit_host_call(typeof_idx, 1);
            self.emit_const(Value::String(std::sync::Arc::from("string")));
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            // The message goes through THE write so it lands in the innermost
            // output buffer exactly as `echo` does: `wasi:logging` appends a
            // newline real php never writes, and writing straight to stdout
            // jumps ahead of buffered content.
            self.emit_u16(Op::LOCAL_GET, status_slot);
            self.emit_common("php.echo_stringify", 1, line);
            let current = self.current;
            common::io::emit_write_or_buffer(&mut self.chunks, current, line);
            // A message exits SUCCESSFULLY — the string is output, never a status.
            self.emit_const(Value::F64(0.0));
            self.emit_u16(Op::LOCAL_SET, status_slot);
            self.chunk().emit_end(line);
        }

        let current = self.current;
        common::io::emit_ob_flush_all(&mut self.chunks, current, line);

        let exit_idx = self.import("wasi:cli/exit", "exit-with-code");
        self.emit_u16(Op::LOCAL_GET, status_slot);
        self.emit_host_call(exit_idx, 1);

        // Unreachable once the host call takes effect, but the chunk still has
        // to be well formed for the paths that validate it.
        self.emit_null();
        self.emit_return_through_finally(1)?;
        Ok(())
    }

    /// `StmtKind::Exit` — the statement form. A missing status is 0.
    pub(super) fn compile_exit_stmt(&mut self, status: Option<&Expression>) -> Result<(), String> {
        match status {
            Some(expr) => {
                self.compile_expr(expr)?;
                self.emit_exit_from_stack(1)
            }
            None => self.emit_exit_from_stack(0),
        }
    }

    /// Emit a `for v in gen():` loop that drives the generator via
    ///   block $exit
    ///     loop $loop
    ///       local.get $cont
    ///       gen.next            ;; pushes (value, has_more)
    ///       br_if 0             ;; break out when has_more == 0
    ///       local.set $v        ;; assign yielded value
    ///       <body>
    ///       br $loop
    ///     end
    ///   end
    /// Emit a lazy for-of loop for a custom iterable (one that has an `iterator()`
    /// method per the [Symbol.iterator] protocol). Calls `next()` per iteration.
    /// On entry: `iter_slot` holds the iterable value (not yet advanced).
    /// Emits: BLOCK $exit + LOOP { call iterator(), then loop calling next() }
    pub(super) fn compile_for_of_custom_iterator_lazy(
        &mut self,
        iter_slot: u16,
        var: &str,
        body: &[Statement],
        else_body: Option<&[Statement]>,
    ) -> Result<(), String> {
        let line = self.line;
        let _iterator_key = self.str_const("iterator");
        let _next_key_c = self.str_const("next");
        let done_key_c = self.resolve_slot_interned(&class_slots::ClassSlot::internal("done"));
        let value_key_c = self.resolve_slot_interned(&class_slots::ClassSlot::internal("value"));

        // it = iter_slot.iterator() with __js_this = iter_slot
        let it_slot = self.define_local("__cit_it");
        let next_method_slot = self.define_local("__cit_next");
        let step_slot = self.define_local("__cit_step");
        let done_slot = self.define_local("__cit_done");
        let did_break_slot = self.define_local("__cit_did_break");
        inst!(self, core_wasm::bool_const, false);
        self.emit_u16(Op::LOCAL_SET, did_break_slot);

        // Get iterator method via STRUCT_GET — the TypeRegistry resolves
        // methods registered by crate::primitives::classes::register_type, including
        // "iterator" (the walker-normalized [Symbol.iterator]).
        self.emit_u16(Op::LOCAL_GET, iter_slot);
        self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal("iterator"));
        let iter_fn_slot = self.define_local("__cit_iter_fn");
        self.emit_u16(Op::LOCAL_SET, iter_fn_slot);

        // Call iterator() with __js_this = iter_slot
        self.emit_u16(Op::LOCAL_GET, iter_slot);
        self.emit_global_write("__js_this");
        self.emit_u16(Op::LOCAL_GET, iter_fn_slot);
        self.emit_direct_callable_invoke(0);
        self.emit_u16(Op::LOCAL_SET, it_slot);

        // Emit BLOCK + LOOP
        let block_patch = self.chunk().emit_block(line);
        let (loop_patch, _) = self.chunk().emit_loop_s(line);
        self.label_depth += 2;

        // next_method = it.next via STRUCT_GET
        self.emit_u16(Op::LOCAL_GET, it_slot);
        self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal("next"));
        self.emit_u16(Op::LOCAL_SET, next_method_slot);

        // Call next() with __js_this = it
        self.emit_u16(Op::LOCAL_GET, it_slot);
        self.emit_global_write("__js_this");
        self.emit_u16(Op::LOCAL_GET, next_method_slot);
        self.emit_direct_callable_invoke(0);
        self.emit_u16(Op::LOCAL_SET, step_slot);

        // ECMA-262 IteratorNext: next() must return an Object. A primitive
        // result is a TypeError, not an endless loop over missing done/value.
        self.emit_u16(Op::LOCAL_GET, step_slot);
        let typeof_idx = self.import("ecma:value", "typeof");
        self.emit_host_call(typeof_idx, 1);
        self.emit_const(Value::String(Arc::from("object")));
        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        self.chunk().emit_if(line);
        self.emit_const(Value::String(Arc::from("Iterator result is not an object")));
        self.emit_js_exception_ctor_from_message_value("TypeError")?;
        common::errors::emit_throw(self.chunk(), line);
        self.chunk().emit_end(line);

        // step.done check
        self.emit_u16(Op::LOCAL_GET, step_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &done_key_c);
        self.emit_u16(Op::LOCAL_SET, done_slot);
        self.emit_u16(Op::LOCAL_GET, done_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        self.chunk().emit_br_if(1, line); // done → exit block

        // var = step.value
        self.emit_u16(Op::LOCAL_GET, step_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &value_key_c);
        let var_slot = self.define_source_local(var);
        self.emit_u16(Op::LOCAL_SET, var_slot);

        // Loop body in $body block for break/continue targeting
        let body_block = self.chunk().emit_block(line);
        self.label_depth += 1;
        let break_depth = self.label_depth - 2; // $exit
        let continue_depth = self.label_depth; // $body
        self.loops.push(LoopCtx {
            label: self.pending_label.take(),
            break_label_depth: break_depth,
            continue_label_depth: continue_depth,
            did_break_slot: Some(did_break_slot),
            iterator_close_slot: Some(it_slot),
            is_continuable: true,
            finally_depth: self.frame_cf().active_finally_blocks.len(),
        });
        for s in body {
            self.compile_stmt(s)?;
        }
        self.loops.pop();
        self.chunk().emit_end(line);
        self.chunk().patch_block(body_block);
        self.label_depth -= 1;

        // continue → loop
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(loop_patch);
        self.chunk().emit_end(line);
        self.chunk().patch_block(block_patch);
        self.label_depth -= 2;

        if let Some(else_stmts) = else_body {
            // Python/Ruby else: runs if no break
            let skip = self.chunk().emit_block(line);
            self.label_depth += 1;
            self.emit_u16(Op::LOCAL_GET, did_break_slot);
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            };
            self.chunk().emit_br_if(0, line);
            for s in else_stmts {
                self.compile_stmt(s)?;
            }
            self.chunk().emit_end(line);
            self.chunk().patch_block(skip);
            self.label_depth -= 1;
        }

        Ok(())
    }

    pub(super) fn compile_generator_for_in(
        &mut self,
        var: &str,
        key: Option<&str>,
        iter: &Expression,
        body: &[Statement],
        else_body: Option<&[Statement]>,
    ) -> Result<(), String> {
        use crate::ast::ExprKind;
        // Compile and stash the continuation.
        let (callee, args) = match &iter.kind {
            ExprKind::Call { callee, args, .. } => (callee, args),
            _ => unreachable!("compile_generator_for_in expects Call"),
        };
        self.compile_call(callee, args)?;
        let cont_slot = self.define_local("__gen_cont");
        self.emit_u16(Op::LOCAL_SET, cont_slot);

        self.compile_generator_for_in_cont(var, key, cont_slot, body, else_body)
    }

    pub(super) fn compile_generator_for_in_cont(
        &mut self,
        var: &str,
        key: Option<&str>,
        cont_slot: u16,
        body: &[Statement],
        else_body: Option<&[Statement]>,
    ) -> Result<(), String> {
        let key_index_slot = self.maybe_define_buffered_generator_key_index_slot(key);
        let did_break_slot = self.define_local("__gen_for_did_break");
        inst!(self, core_wasm::bool_const, false);
        self.emit_u16(Op::LOCAL_SET, did_break_slot);

        let line = self.line;
        let block_patch = self.chunk().emit_block(line);
        let (loop_patch, _) = self.chunk().emit_loop_s(line);
        self.label_depth += 2;

        // Advance the generator. GEN_NEXT pops cont and pushes (value, has_more).
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        let line = self.line;
        crate::primitives::generators::emit_next(self.chunk(), line);
        let has_more_slot = self.define_local("__gen_has_more");
        self.emit_u16(Op::LOCAL_SET, has_more_slot);
        let value_slot = self.define_local("__gen_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        if self.profile.buffered_iterator_methods {
            self.emit_buffered_generator_foreach_state(cont_slot, has_more_slot, value_slot);
        } else {
            self.emit_u16(Op::LOCAL_GET, has_more_slot);
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            };
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_not(self.chunk(), line);
            };
            // br_if_label 1 → jump to $exit when has_more was 0.
            self.chunk().emit_br_if(1, line);
        }

        if let Some(key_name) = key {
            let key_slot = self.define_source_local(key_name);
            if self.profile.buffered_iterator_methods {
                self.emit_buffered_generator_key_binding(key_slot, value_slot, key_index_slot);
            } else {
                self.emit_null();
                self.emit_u16(Op::LOCAL_SET, key_slot);
            }
        }

        let var_slot = self.define_source_local(var);
        if self.profile.buffered_iterator_methods {
            self.emit_buffered_generator_value_binding(var_slot, value_slot);
        } else {
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_u16(Op::LOCAL_SET, var_slot);
        }

        if let Some(key_index_slot) = key_index_slot {
            self.emit_u16(Op::LOCAL_GET, key_index_slot);
            self.emit_const(Value::F64(1.0));
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_add(self.chunk(), line);
            };
            self.emit_u16(Op::LOCAL_SET, key_index_slot);
        }

        // Compile loop body inside a `$body` block so `continue` can
        // target it without rerunning the advance.
        let body_block = self.chunk().emit_block(line);
        self.label_depth += 1;
        let break_depth = self.label_depth - 2; // $exit
        let continue_depth = self.label_depth - 0; // $body
        self.loops.push(LoopCtx {
            label: self.pending_label.take(),
            break_label_depth: break_depth,
            continue_label_depth: continue_depth,
            did_break_slot: Some(did_break_slot),
            iterator_close_slot: None,
            is_continuable: true,
            finally_depth: self.frame_cf().active_finally_blocks.len(),
        });
        for s in body {
            self.compile_stmt(s)?;
        }
        self.loops.pop();
        self.chunk().emit_end(line);
        self.chunk().patch_block(body_block);
        self.label_depth -= 1;

        // Continue the loop.
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(loop_patch);
        self.chunk().emit_end(line);
        self.chunk().patch_block(block_patch);
        self.label_depth -= 2;

        let skip_cleanup = self.chunk().emit_block(line);
        self.label_depth += 1;
        self.emit_u16(Op::LOCAL_GET, did_break_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(0, line);

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        let is_done_idx = self.import("ecma:value", "isGeneratorDone");
        self.emit_host_call(is_done_idx, 1);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        }
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        inst!(self, core_wasm::undefined);
        self.emit_generator_control_packet_from_stack("return");
        let line = self.line;
        crate::primitives::generators::emit_resume(self.chunk(), line);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        inst!(self, core_wasm::bool_const, true);
        self.class_set(
            class_slots::ObjSource::Stack,
            &class_slots::ClassSlot::internal("__vybe_gen_returned"),
            class_slots::ValueSource::Stack,
        );
        self.chunk().emit_end(line);

        self.chunk().emit_end(line);
        self.chunk().patch_block(skip_cleanup);
        self.label_depth -= 1;

        if let Some(else_stmts) = else_body {
            let skip_else = self.chunk().emit_block(line);
            self.label_depth += 1;
            self.emit_u16(Op::LOCAL_GET, did_break_slot);
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            };
            self.chunk().emit_br_if(0, line);
            for s in else_stmts {
                self.compile_stmt(s)?;
            }
            self.chunk().emit_end(line);
            self.chunk().patch_block(skip_else);
            self.label_depth -= 1;
        }
        Ok(())
    }

    /// Arity of the currently-compiling function if the pre-scan tagged
    /// it multi-return, else `None`. Driven off `current_func_name` so
    /// it automatically tracks function boundaries without a parallel
    /// stack.
    pub(super) fn current_multi_return_arity(&self) -> Option<u8> {
        let name = self.current_func_name.as_deref()?;
        self.multi_return_functions.get(name).copied()
    }

    pub(super) fn multi_return_arity_for_callee(&self, callee: &Expression) -> Option<u8> {
        match &callee.kind {
            ExprKind::Ident(name) => self.multi_return_functions.get(&self.canon(name)).copied(),
            ExprKind::Member { object, field, .. } => {
                if let ExprKind::Ident(object_name) = &object.kind {
                    let qualified = self.canon(&format!("{}.{}", object_name, field));
                    if let Some(&arity) = self.multi_return_functions.get(&qualified) {
                        return Some(arity);
                    }
                }

                self.multi_return_functions.get(&self.canon(field)).copied()
            }
            _ => None,
        }
    }

    /// Emit the CALL for a multi-value receive context *without* the
    /// trailing repack that `compile_expr` would normally add. The
    /// destructure path consumes the N raw stack values directly.
    pub(super) fn compile_call_raw(&mut self, value: &Expression) -> Result<(), String> {
        if let ExprKind::Call { callee, args, .. } = &value.kind {
            if let ExprKind::Ident(name) = &callee.kind {
                if self.multi_return_arity_for_callee(callee).is_some() {
                    self.emit_var_get(name);
                    for arg in args {
                        self.compile_expr(&arg.value)?;
                    }
                    self.emit_direct_callable_invoke(args.len() as u8);
                    return Ok(());
                }
            }
            self.compile_call(callee, args)
        } else {
            self.compile_expr(value)
        }
    }

    /// Pack the top-N stack values — produced by a multi-value CALL —
    /// into a single array/tuple so downstream uses see the expected
    /// single-value semantics. The last pushed value becomes element
    /// `n-1`; order matches what a destructure would assign.
    pub(super) fn pack_multi_value_result(&mut self, n: u8) {
        let line = self.line;
        // Reserve N consecutive slots via the existing scope helper —
        // `emit_pack_n` stashes each stack value into a slot, then
        // rebuilds the array from those slots in declaration order.
        let mut first = 0u16;
        for i in 0..n {
            let s = self.define_local("__mv_pack");
            if i == 0 {
                first = s;
            }
        }
        common::collections::emit_pack_n(&mut self.chunks, self.current, n as u16, first, line);
        // A multi-value-return function is returning a tuple (the ABI is only
        // taken for functions whose returns are same-arity tuple literals), so
        // the repacked array must carry the tuple tag — otherwise a tuple that
        // flows through a `return` reprs/`type()`s as a plain list, unlike the
        // identical value built inline (which `ExprKind::Tuple` tags). Languages
        // that don't tag tuples (e.g. Lua's multi-value rows) are unaffected.
        if self.profile.tuple_literals_tagged {
            common::tuples::emit_tag(&mut self.chunks, self.current, line);
        }
        if !self.profile.multi_value_row_marker.is_empty() {
            let row_slot = self.define_local("__mv_pack_row");
            self.emit_u16(Op::LOCAL_SET, row_slot);
            self.stamp_multi_value_row_slot(row_slot);
            self.emit_u16(Op::LOCAL_GET, row_slot);
        }
    }

    /// Stamp the profile's multi-value row marker on `row_slot`. A profile that
    /// declares no marker does not distinguish rows from arrays — nothing is
    /// emitted, so every call site can be unconditional.
    pub(super) fn stamp_multi_value_row_slot(&mut self, row_slot: u16) {
        let marker = self.profile.multi_value_row_marker.clone();
        if marker.is_empty() {
            return;
        }
        self.emit_u16(Op::LOCAL_GET, row_slot);
        self.emit_const(Value::Bool(true));
        self.class_set(
            class_slots::ObjSource::Stack,
            &class_slots::ClassSlot::internal(&marker),
            class_slots::ValueSource::Stack,
        );
    }

    /// Return `Some((N, [ident...]))` when `targets`/`value` match the
    /// "multi-value receive" shape:
    ///   * exactly one target, a tuple-destructure of N plain identifiers
    ///   * value is a direct `Ident(name)` call to a function the pre-scan
    ///     tagged multi-return with matching arity N
    /// For any other shape we return `None` and fall through to the
    /// existing heap-tuple destructuring path.
    pub(super) fn detect_multi_value_receive(
        &self,
        targets: &[Expression],
        value: &Expression,
    ) -> Option<(u8, Vec<String>)> {
        if targets.len() != 1 {
            return None;
        }
        let idents = match &targets[0].kind {
            ExprKind::Destructure(DestructurePattern::Array(pats)) => {
                let mut names = Vec::with_capacity(pats.len());
                for p in pats {
                    match p {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(n), _) => {
                            names.push(n.clone());
                        }
                        _ => return None,
                    }
                }
                names
            }
            _ => return None,
        };
        let multi_n = match &value.kind {
            ExprKind::Call { callee, args, .. } => {
                let _ = args;
                self.multi_return_arity_for_callee(callee)?
            }
            _ => return None,
        };
        if multi_n as usize != idents.len() {
            return None;
        }
        Some((multi_n, idents))
    }

    /// Walk top-level function declarations and record every function
    /// whose explicit `Return` statements all carry a tuple literal of
    /// the same arity. Those functions opt into the WASM multi-value
    /// ABI: callee sets `chunk.result_arity = N` and pushes the tuple
    /// elements unpacked; caller destructures directly off the stack.
    pub(super) fn collect_multi_return_functions(&mut self, stmts: &[Statement]) {
        for stmt in stmts {
            match &stmt.kind {
                StmtKind::FunctionDecl { name, body, .. } => {
                    if let Some(arity) = uniform_tuple_return_arity(body) {
                        let cname = self.canon(name);
                        self.multi_return_functions.insert(cname, arity);
                    }
                    self.collect_multi_return_functions(body);
                }
                StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
                    self.collect_multi_return_functions(body);
                }
                StmtKind::If {
                    then_body,
                    elifs,
                    else_body,
                    ..
                } => {
                    self.collect_multi_return_functions(then_body);
                    for (_, body) in elifs {
                        self.collect_multi_return_functions(body);
                    }
                    if let Some(body) = else_body {
                        self.collect_multi_return_functions(body);
                    }
                }
                StmtKind::While {
                    body, else_body, ..
                } => {
                    self.collect_multi_return_functions(body);
                    if let Some(body) = else_body {
                        self.collect_multi_return_functions(body);
                    }
                }
                StmtKind::For { init, body, .. } => {
                    if let Some(init) = init {
                        self.collect_multi_return_functions(std::slice::from_ref(init));
                    }
                    self.collect_multi_return_functions(body);
                }
                StmtKind::ForIn {
                    body, else_body, ..
                } => {
                    self.collect_multi_return_functions(body);
                    if let Some(body) = else_body {
                        self.collect_multi_return_functions(body);
                    }
                }
                StmtKind::DoWhile { body, .. }
                | StmtKind::With { body, .. }
                | StmtKind::Using { body, .. }
                | StmtKind::Lock { body, .. } => {
                    self.collect_multi_return_functions(body);
                }
                StmtKind::Try {
                    body,
                    catches,
                    else_body,
                    finally,
                } => {
                    self.collect_multi_return_functions(body);
                    for catch in catches {
                        self.collect_multi_return_functions(&catch.body);
                    }
                    if let Some(body) = else_body {
                        self.collect_multi_return_functions(body);
                    }
                    if let Some(body) = finally {
                        self.collect_multi_return_functions(body);
                    }
                }
                StmtKind::Switch { cases, default, .. } => {
                    for case in cases {
                        self.collect_multi_return_functions(&case.body);
                    }
                    if let Some(body) = default {
                        self.collect_multi_return_functions(body);
                    }
                }
                _ => {}
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Helpers
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn scope(&self) -> &Scope {
        self.scopes.last().unwrap()
    }
    pub(super) fn scope_mut(&mut self) -> &mut Scope {
        self.scopes.last_mut().unwrap()
    }

    pub(super) fn pointer_binding_key(&self, name: &str) -> String {
        if self.case_sensitive {
            name.to_string()
        } else {
            self.canon(name)
        }
    }

    /// Does `name` denote a reference here?
    ///
    /// Two stores, because there are two questions. A LOCAL or parameter is a
    /// per-binding fact and lives on the binding, so scope resolution answers it
    /// and shadowing is free — a local that shadows an outer reference correctly
    /// reports `false`, and a `global $g` that IS the outer binding correctly
    /// falls through to the global store, because `open_names` routes it there.
    ///
    /// A promoted GLOBAL is a module-wide fact and lives in one module-wide set.
    ///
    /// These used to be one `HashMap<chunk_idx, HashSet<name>>`. That could not
    /// express either question: a name marked in ANY chunk leaked into every
    /// other chunk through a module-wide fallback, and the guard bolted on to
    /// stop that ("this chunk has a local of that name → not a cell") could not
    /// tell a shadowing local from the binding itself.
    pub(super) fn binding_uses_pointer_cell(&self, name: &str) -> bool {
        self.pointer_cell_binding_fact(name, true)
    }

    /// Has a wrap ALREADY happened for this binding?
    ///
    /// What a PROMOTION site must ask. [`Self::binding_uses_pointer_cell`] also
    /// answers `true` for the module-wide pre-pass, which is a "readers must
    /// deref" hint about a wrap that may still be AHEAD in this forward pass —
    /// so a promotion that consults it concludes the work is done, skips the
    /// wrap, and hands out the unwrapped slot. The callee is then aliased to
    /// nothing: its reads see a plain value and its writes land nowhere the
    /// caller can see.
    ///
    /// The same confusion already cost `promote_global_binding_to_pointer_cell`
    /// (§10e.1); the local promotion kept asking the wrong question because the
    /// hint set was, until call sites began declaring by-reference arguments,
    /// almost always empty for php.
    pub(super) fn binding_already_pointer_cell(&self, name: &str) -> bool {
        self.pointer_cell_binding_fact(name, false)
    }

    fn pointer_cell_binding_fact(&self, name: &str, include_pending: bool) -> bool {
        // `global $g` / `nonlocal x` names the OUTER binding. Any local record
        // for it aliases that binding rather than shadowing it, so its
        // properties must be read from where the binding lives — otherwise a
        // promoted global is read RAW inside the one function that declared its
        // intent to use it, and the cell object reaches arithmetic.
        // Both spellings: `ScopeDecl` writes `open_names` through
        // `self.canon(name)` unconditionally, while `pointer_binding_key` canons
        // only when the language folds case — php variables are case-SENSITIVE,
        // so for php the two diverge and asking with either one alone misses.
        let declared_open =
            self.scope().declared_open(name) || self.scope().declared_open(&self.canon(name));
        if !declared_open {
            if let Some(holds) = self.scope().holds_reference(name) {
                return holds;
            }
        }
        let key = self.pointer_binding_key(name);
        // Either a wrap that already happened, or — for a READER only — the
        // module-wide pre-pass saying one WILL happen later in this
        // compilation. The second is what makes
        // `function r(){ global $g; echo $g; }` — emitted before
        // `$g = 1; w($g);` promotes `$g` — deref correctly: whether a global is
        // a cell is a whole-module property being decided in a forward pass.
        self.promoted_global_cells.contains(&key)
            || (include_pending && self.module_addr_taken_globals.contains(&key))
    }

    /// Record that `name` denotes a reference.
    ///
    /// Routed by RESOLUTION, not by the caller: if the name is a local or
    /// parameter of this scope the flag goes on that binding; otherwise it is a
    /// promoted global and goes in the module-wide set. Call sites still pass a
    /// name, so none of them had to change — the store is chosen where the
    /// answer is actually known.
    ///
    /// The Go case the old chunk-keyed map needed a special guard for now falls
    /// out for free: `main`'s local `c` promoted to a cell sets the flag on
    /// `main`'s binding, and the receiver `c` of `func (c Counter) Peek()` is a
    /// DIFFERENT binding whose flag was never set.
    pub(super) fn mark_pointer_cell_binding(&mut self, name: &str) {
        // A binding at MODULE scope is the global every other chunk sees, so the
        // fact is recorded in BOTH stores — it is genuinely both. Only an inner
        // scope's binding is private enough to live on the binding alone.
        //
        // Without this, `$g = 1; w($g);` at module scope promotes `$g` as a
        // script-scope LOCAL (that arm wins), module scope derefs it correctly,
        // and a `global $g;` in some other function — which reaches the same
        // storage as a GLOBAL — finds the global store empty and reads the cell
        // object raw.
        let at_module_scope = self.scopes.len() <= 1;
        let bound_locally = self
            .scopes
            .last_mut()
            .is_some_and(|s| s.set_holds_reference(name));
        if !bound_locally || at_module_scope {
            self.mark_promoted_global_cell(name);
        }
    }

    /// Record that the MODULE-LEVEL binding for `name` holds a cell.
    ///
    /// For callers that already know the storage is global. Both the canon here
    /// and in `binding_uses_pointer_cell` must match what `ScopeDecl` writes
    /// into `open_names`, which is `self.canon(name)`.
    pub(super) fn mark_promoted_global_cell(&mut self, name: &str) {
        let key = self.pointer_binding_key(name);
        self.promoted_global_cells.insert(key);
    }

    pub(super) fn resolve_named_local_slot(&self, name: &str) -> Option<u16> {
        self.scope().resolve(name)
    }

    /// Bind every [`PassBy::Alias`] parameter to the reference it was handed.
    ///
    /// MARK, never promote: the caller already passed a reference
    /// (`compile_ref_aware_call_arg`), so the slot holds one on entry.
    /// `promote_local_binding_to_pointer_cell` would wrap it in a SECOND cell —
    /// a reference to a reference — and every read would then deref one level
    /// short. Marking alone is what makes reads auto-deref and writes store
    /// through to the caller's storage.
    ///
    /// This is the parameter half of aliasing. Without it the argument arrives
    /// as a reference and the body treats it as an ordinary value, which reads
    /// back the cell OBJECT rather than what it points at.
    /// Takes an ITERATOR, not a slice: methods hold their parameters as
    /// `Vec<&Param>` and define their locals in a third prologue of their own
    /// (`classes.rs`), which is how they went without this call entirely.
    pub(super) fn bind_alias_params<'a>(&mut self, params: impl IntoIterator<Item = &'a Param>) {
        for p in params {
            if p.pass_by == PassBy::Alias {
                self.mark_pointer_cell_binding(&p.name);
            }
        }
    }

    pub(super) fn promote_local_binding_to_pointer_cell(&mut self, name: &str) -> Option<u16> {
        let slot = self.resolve_named_local_slot(name)?;
        if !self.binding_already_pointer_cell(name) {
            crate::primitives::references::emit_cell_new_from_local(
                &mut self.chunks,
                self.current,
                slot,
                self.line,
            );
            self.emit_u16(Op::LOCAL_SET, slot);
            self.mark_pointer_cell_binding(name);
        }
        Some(slot)
    }

    /// The module-global key a VARIABLE is stored under.
    ///
    /// Must match `emit_var_get`/`emit_var_set` exactly. A language may mangle
    /// it — php stores `$c` as `__php_var_c` so a global `$foo` cannot collide
    /// with a function `foo` — so reading `canon` here looks up a key that was
    /// never written. That divergence is what `globals.rs` exists to prevent,
    /// and it made `promote_global_binding_to_pointer_cell` refuse every php
    /// global: `$d = &$c` then fell through to a fresh DETACHED cell and
    /// aliasing silently degraded to a copy at module scope.
    pub(super) fn variable_global_binding_key(&self, name: &str) -> String {
        let canon_name = self.canon(name);
        self.variable_global_key(name, &canon_name)
    }

    /// Promote a module global to a pointer cell. Taking the address of a NAME
    /// means the name denotes storage, so this always succeeds.
    ///
    /// It used to refuse unless `profile.globals_may_be_undeclared` (set by c
    /// alone) or the name was already in `defined_globals` — a per-language
    /// switch in shared code whose failure mode was the worst available one:
    /// returning `false` dropped the caller into
    /// `emit_wrap_top_of_stack_in_pointer_cell`, a fresh DETACHED cell, so
    /// aliasing degraded to a silent COPY instead of raising anything. A place
    /// resolves or errors; it never quietly becomes a copy.
    ///
    /// Promoting a name that has no value yet is correct, not a fallback: the
    /// cell is created from whatever the global currently holds, which is what
    /// php's reference auto-vivification (`$r = &$undefined`) already means, and
    /// what c's `globals_may_be_undeclared` was granting itself as a special
    /// case.
    pub(super) fn promote_global_binding_to_pointer_cell(&mut self, name: &str) -> bool {
        let global_key = self.variable_global_binding_key(name);
        // The real-promotion set ONLY, never `binding_uses_pointer_cell`: that
        // also answers `true` for the module-wide address-taken PRE-PASS, which
        // is a "readers must deref" hint, not a record that the wrap happened.
        // Consulting it here made promotion skip itself — readers deref a global
        // that was never wrapped.
        let already_promoted = self
            .promoted_global_cells
            .contains(&self.pointer_binding_key(name));
        if !already_promoted {
            let value_slot = self.define_local("__ref_global_value");
            self.emit_global_read(&global_key);
            self.emit_u16(Op::LOCAL_SET, value_slot);
            crate::primitives::references::emit_cell_new(
                &mut self.chunks,
                self.current,
                value_slot,
                self.line,
            );
            self.emit_global_write(&global_key);
            // The GLOBAL store, not the routing helper: this site just promoted
            // a global and knows it. Routing by resolution would put the flag on
            // whatever local happens to share the name in the CURRENT scope —
            // at module scope that is the script scope's own binding — and then
            // a `global $g;` in some other function finds the global store empty
            // and reads the cell object raw.
            self.mark_promoted_global_cell(name);
        }

        true
    }

    pub(super) fn emit_wrap_top_of_stack_in_pointer_cell(&mut self) {
        let value_slot = self.define_local("__ref_cell_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        crate::primitives::references::emit_cell_new(
            &mut self.chunks,
            self.current,
            value_slot,
            self.line,
        );
    }

    pub(super) fn is_pointer_runtime_field(field: &str) -> bool {
        matches!(field, "__ref_kind" | "__base" | "__idx" | "__value")
    }

    pub(super) fn emit_string_slot_eq_literal(&mut self, slot: u16, literal: &str) {
        self.emit_u16(Op::LOCAL_GET, slot);
        let test = self.import("wasm:js-string", "test");
        self.emit_host_call(test, 1);
        let line = self.line;
        // ⛔ BOTH ARMS PRODUCE A VALUE, so the block is not void. `equals`
        // returns i32 and the else arm is `I32(0)`, which makes this an
        // i32-result block — `emit_if_i32`, the same shape every other
        // comparison chain uses. Declared `(0, 0)` it was only survivable
        // because the VM shares one operand stack across blocks; wasm blocks
        // do not, and the surplus is rejected at the `else`.
        self.chunk().emit_if_i32(line);
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_const(Value::String(Arc::from(literal)));
        let eq = self.import("wasm:js-string", "equals");
        self.emit_host_call(eq, 2);
        self.chunk().emit_else(line);
        self.emit_const(Value::I32(0));
        self.chunk().emit_end(line);
    }

    pub(super) fn emit_string_eq_literal(&mut self, literal: &str) {
        let slot = self.define_local("__string_eq_candidate");
        self.emit_u16(Op::LOCAL_SET, slot);
        self.emit_string_slot_eq_literal(slot, literal);
    }

    pub(super) fn emit_raw_string_slot_eq_literal(&mut self, slot: u16, literal: &str) {
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_const(Value::String(Arc::from(literal)));
        let eq = self.import("wasm:js-string", "equals");
        self.emit_host_call(eq, 2);
    }

    pub(super) fn emit_autoderef_pointer_cell(&mut self) {
        let obj_slot = self.define_local("__ref_autoderef_obj");
        self.emit_u16(Op::LOCAL_SET, obj_slot);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        inst!(self, recipes::is_object);
        let obj_line = self.line;
        // ⛔ EVERY CONDITIONAL IN THIS FUNCTION YIELDS A VALUE — an autoderef
        // leaves the pointed-to value on the stack down every one of the
        // cell / carray / shared / passthrough arms. Declared `emit_if` they
        // were all `(0, 0)`: void blocks whose arms each push. The VM tolerated
        // it because its blocks share one operand stack; wasm blocks do not.
        self.chunk().emit_if_value(obj_line);

        let kind_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("__ref_kind"));

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_string_eq_literal("cell");
        let cell_line = self.line;
        self.chunk().emit_if_value(cell_line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        crate::primitives::references::emit_cell_load(&mut self.chunks, self.current, self.line);
        self.chunk().emit_else(cell_line);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_string_eq_literal("carray");
        let carray_line = self.line;
        self.chunk().emit_if_value(carray_line);

        let base_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("__base"));
        let idx_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("__idx"));
        let base_slot = self.define_local("__ref_carray_base");

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &base_key);
        self.emit_u16(Op::LOCAL_SET, base_slot);

        self.emit_u16(Op::LOCAL_GET, base_slot);
        inst!(self, recipes::is_object);
        let base_obj_line = self.line;
        self.chunk().emit_if_value(base_obj_line);

        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_string_eq_literal("cell");
        let base_cell_line = self.line;
        self.chunk().emit_if_value(base_cell_line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        crate::primitives::references::emit_cell_load(&mut self.chunks, self.current, self.line);
        self.chunk().emit_else(base_cell_line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &idx_key);
        common::collections::emit_get(&mut self.chunks, self.current, self.line);
        self.chunk().emit_end(base_cell_line);

        self.chunk().emit_else(base_obj_line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &idx_key);
        common::collections::emit_get(&mut self.chunks, self.current, self.line);
        self.chunk().emit_end(base_obj_line);

        self.chunk().emit_else(carray_line);
        // Third shape: a word in SHARED linear memory. The load is the WASM
        // atomic one — an ordinary read of an atomically-updated binding must
        // see the other thread's write.
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_string_eq_literal(crate::primitives::pointers::SHARED_KIND);
        let shared_line = self.line;
        self.chunk().emit_if_value(shared_line);
        self.class_get(class_slots::ObjSource::Local(obj_slot), &class_slots::ClassSlot::internal(crate::primitives::pointers::SHARED_ADDR_KEY));
        {
            let line = self.line;
            crate::primitives::threading::emit_atomic_load(self.chunk(), line);
        }
        self.chunk().emit_else(shared_line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.chunk().emit_end(shared_line);
        self.chunk().emit_end(carray_line);

        self.chunk().emit_end(cell_line);

        self.chunk().emit_else(obj_line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.chunk().emit_end(obj_line);
    }

    /// Store `value_slot` THROUGH the reference in `ptr_slot`. Consumes
    /// nothing off the stack and leaves nothing on it.
    ///
    /// The mirror of `emit_autoderef_pointer_cell`: same shapes, same order,
    /// same passthrough for a non-reference. Every write to a name bound to a
    /// reference goes through here, so a `carray` binding stores into its
    /// container instead of growing a dead `__value` field on the reference
    /// object — which is what a cell-only store did, silently.
    pub(super) fn emit_store_through_pointer(&mut self, ptr_slot: u16, value_slot: u16) {
        self.emit_u16(Op::LOCAL_GET, ptr_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        // `ref.test` pushes `Value::I32(0|1)` and `Op::IF` takes an i32 — the
        // ToBoolean ladder here was a no-op. See `operators::emit_to_primitive`.
        self.chunk().emit_if(line);

        let kind_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("__ref_kind"));

        self.emit_u16(Op::LOCAL_GET, ptr_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_const(Value::String(Arc::from("cell")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        }
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, ptr_slot);
        crate::primitives::references::emit_cell_store(
            &mut self.chunks,
            self.current,
            value_slot,
            self.line,
        );
        self.emit(Op::DROP);

        let line = self.line;
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, ptr_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_const(Value::String(Arc::from("carray")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        }
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        let base_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("__base"));
        let idx_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("__idx"));
        let base_slot = self.define_local("__ref_store_carray_base");
        let idx_slot = self.define_local("__ref_store_carray_idx");

        self.emit_u16(Op::LOCAL_GET, ptr_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &base_key);
        self.emit_u16(Op::LOCAL_SET, base_slot);

        self.emit_u16(Op::LOCAL_GET, ptr_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &idx_key);
        self.emit_u16(Op::LOCAL_SET, idx_slot);

        self.emit_u16(Op::LOCAL_GET, base_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        // `ref.test` pushes `Value::I32(0|1)` and `Op::IF` takes an i32 — the
        // ToBoolean ladder here was a no-op. See `operators::emit_to_primitive`.
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_const(Value::String(Arc::from("cell")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        }
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, base_slot);
        crate::primitives::references::emit_cell_store(
            &mut self.chunks,
            self.current,
            value_slot,
            self.line,
        );
        self.emit(Op::DROP);

        let line = self.line;
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::collections::emit_set(&mut self.chunks, self.current, self.line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);

        let line = self.line;
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::collections::emit_set(&mut self.chunks, self.current, self.line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);

        let line = self.line;
        self.chunk().emit_else(line);
        // Third shape: shared word — mirror of the load dispatcher's arm. A
        // plain assignment to a name bound to shared storage IS an atomic
        // store; anything weaker would tear against a concurrent RMW.
        self.emit_u16(Op::LOCAL_GET, ptr_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &kind_key);
        self.emit_string_eq_literal(crate::primitives::pointers::SHARED_KIND);
        let shared_line = self.line;
        self.chunk().emit_if(shared_line);
        self.class_get(class_slots::ObjSource::Local(ptr_slot), &class_slots::ClassSlot::internal(crate::primitives::pointers::SHARED_ADDR_KEY));
        self.emit_u16(Op::LOCAL_GET, value_slot);
        {
            let line = self.line;
            crate::primitives::threading::emit_atomic_store(self.chunk(), line);
        }
        self.chunk().emit_end(shared_line);
        self.chunk().emit_end(line);

        let line = self.line;
        self.chunk().emit_end(line);

        let line = self.line;
        self.chunk().emit_else(line);
        self.chunk().emit_end(line);
    }

    pub(super) fn compile_address_of_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            ExprKind::Ident(name) => {
                // Callable address syntax is language-specific and must lower
                // to `ExprKind::FuncRef` before reaching this storage-address path.
                self.compile_ident_reference(name);
                return Ok(());
            }
            ExprKind::Unary {
                op: UnaryOp::Deref,
                expr,
            } => {
                self.compile_expr(expr)?;
                return Ok(());
            }
            // `&a[i]` denotes the SLOT, so it resolves like a name does. Falling
            // through to the rvalue wrap below would read the element and box the
            // COPY — a detached cell, aliasing degraded to a silent copy.
            ExprKind::Index { object, index, .. } => {
                self.compile_index_reference(object, index)?;
                return Ok(());
            }
            // Same rule as `Index`, and it must live HERE as well as in the
            // `RefOf` arm: fixing only one spelling makes a binding's aliasing
            // depend on which node the walker happened to build (§2). php emits
            // `Unary{AddrOf, Member}` for `&$o->p`.
            ExprKind::Member { object, field, .. } => {
                self.compile_member_reference(object, field)?;
                return Ok(());
            }
            _ => {}
        }

        self.compile_expr(expr)?;
        self.emit_wrap_top_of_stack_in_pointer_cell();
        Ok(())
    }

    /// `&name` where `name` denotes STORAGE — this frame's local, else the
    /// module global.
    ///
    /// A NAME always has storage, so this always resolves. It never falls
    /// through to the rvalue wrap, which would hand back a DETACHED cell —
    /// §4's rule that a place must never reach the rvalue path.
    ///
    /// Shared by BOTH spellings. `RefOf(PlaceExpr::Ident)` used to carry only
    /// the local half and wrapped every global, so go's `var g int; p := &g;
    /// *p = 2` left `g` at its old value while the identical program written
    /// with a local worked (§10d, pre-existing #2). One concept, two spellings,
    /// one resolution — the recurring failure this plan exists to end.
    pub(super) fn compile_ident_reference(&mut self, name: &str) {
        if let Some(slot) = self.promote_local_binding_to_pointer_cell(name) {
            self.emit_u16(Op::LOCAL_GET, slot);
            return;
        }
        self.promote_global_binding_to_pointer_cell(name);
        // Same key the promotion wrote — see `variable_global_binding_key`.
        let global_key = self.variable_global_binding_key(name);
        self.emit_global_read(&global_key);
    }

    /// `&container[key]` → `{__ref_kind:"carray", __base, __idx}`.
    ///
    /// The base is the container VALUE, so the reference aliases whatever the
    /// variable holds rather than a copy of it. Nothing here is per-container:
    /// `emit_autoderef_pointer_cell` and the `RefLoad` store arm both already
    /// route a carray through the VM's polymorphic indexed access, which
    /// dispatches on the base's `ObjectKind` at runtime.
    pub(super) fn compile_index_reference(
        &mut self,
        object: &Expression,
        index: &Expression,
    ) -> Result<(), String> {
        self.compile_expr(object)?;
        let base_slot = self.define_local("__ref_index_base");
        self.emit_u16(Op::LOCAL_SET, base_slot);

        self.compile_expr(index)?;
        let key_slot = self.define_local("__ref_index_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        crate::primitives::references::emit_carray_new(
            &mut self.chunks,
            self.current,
            base_slot,
            key_slot,
            self.line,
        );
        Ok(())
    }

    /// `&obj.field` → the same `{__base, __idx}` as `&obj[key]`, with the field
    /// name as the key.
    ///
    /// A member IS an indexed access with a constant string key: an instance is
    /// a `STRUCT_NEW` object, and both `Op::ARRAY_GET` and `Op::ARRAY_SET` fall
    /// through to that object's property bag for a non-numeric key — the very
    /// bag a name-keyed `STRUCT_GET` reads. So the reference reads and writes
    /// the same storage the plain member access does, and there is no third
    /// pointer kind for members.
    pub(super) fn compile_member_reference(
        &mut self,
        object: &Expression,
        field: &str,
    ) -> Result<(), String> {
        self.compile_index_reference(object, &Expression::string(field))
    }

    pub(super) fn compile_deref_expr(&mut self, expr: &Expression) -> Result<(), String> {
        self.compile_expr(expr)?;
        self.emit_autoderef_pointer_cell();
        Ok(())
    }

    /// `local_count` to the new high-water mark.
    ///
    /// Why this exists: helpers in `emitter/` (`emit_invoke_method`,
    /// `emit_get_range`, `emit_array_pair`, `emit_runtime_helper_call_*`) allocate
    /// scratch slots starting at `chunk.local_count`. If `chunk.local_count`
    /// isn't kept in sync with `scope.next_slot` during compilation, those
    /// scratch slots overlap named locals (params, rest-collection slots,
    /// user `let` bindings) and silently corrupt them.
    ///
    /// This is the historical root cause of the variadic-param-corruption
    /// bug — see `tests/js/test_variadic_bug.rs`. Maintaining the
    /// invariant `chunk.local_count >= scope.next_slot` at all times
    /// makes every helper using `chunk.local_count` for scratch correct
    /// by construction.
    /// Define a binding the SOURCE declared — a parameter, a `let`/`var`/`Dim`,
    /// a catch binding, a loop variable.
    ///
    /// The distinction matters because `_G` / `globalThis` / `$GLOBALS` /
    /// `globals()` expose module-level SOURCE bindings and nothing else. That
    /// used to be decided by `!name.starts_with("__")`, which is a spelling
    /// answering a provenance question: a user variable named `__x` vanished
    /// from the namespace, and a temporary that forgot the prefix leaked in.
    ///
    /// `define_local` is the compiler-temporary path and stays the default:
    /// 856 of the 909 call sites are emitter scratch. Marking the 53 source
    /// sites is the smaller, checkable half.
    pub(crate) fn define_source_local(&mut self, name: &str) -> u16 {
        self.scopes
            .last_mut()
            .unwrap()
            .set_pending_origin(vybe_runtime::chunk::LocalOrigin::Source);
        self.define_local(name)
    }

    /// [`define_source_local`], carrying a declared type.
    pub(crate) fn define_source_local_typed(
        &mut self,
        name: &str,
        type_hint: Option<vybe_ast::TypeHint>,
    ) -> u16 {
        self.scopes
            .last_mut()
            .unwrap()
            .set_pending_origin(vybe_runtime::chunk::LocalOrigin::Source);
        self.define_local_typed(name, type_hint)
    }

    /// Reconcile the SCOPE's slot allocator with the CHUNK's.
    ///
    /// ⛔ There are two allocators over ONE index space: `Scope::next_slot`
    /// (named locals) and `Chunk::alloc_scratch` (compiler temporaries), and
    /// nothing reconciles them except this. `alloc_scratch` has no rewind, so
    /// a scope that allocates without catching up first hands out a slot the
    /// chunk already gave to a temporary — and the two writes silently share
    /// storage.
    ///
    /// **This is why it is a method rather than a block copied into each
    /// caller.** It used to be inline in `define_local` only, so the JS `var`
    /// path — which calls `Scope::define_at_function_scope` directly — skipped
    /// it: `function h() { try { throw 1 } catch (e) { var c = 3; } return c; }`
    /// returned the try/catch's own `caught` flag (`true`) instead of `3`,
    /// because `c` and the flag were the same slot. Same shape as the
    /// `case_sensitive` bug in `directives.md` §1: a precondition every call
    /// site had to remember, and one did not.
    pub(crate) fn sync_scope_allocator(&mut self) {
        let scope = self.scopes.last_mut().unwrap();
        let chunk_locals = self.chunks[self.current].local_count;
        if scope.next_slot < chunk_locals {
            scope.next_slot = chunk_locals;
        }
        if let Some(dup_slot) = self.chunks[self.current].dup_slot {
            if scope.next_slot <= dup_slot {
                scope.next_slot = dup_slot + 1;
            }
        }
    }

    /// The OTHER half: tell the chunk which slots the scope has taken.
    ///
    /// ⛔ `finalize_local_count` takes the `max` of the two allocators, but it
    /// runs at the END of the chunk — it SIZES the frame and does nothing to
    /// stop them overlapping while compiling. `Chunk::alloc_scratch` hands out
    /// `local_count` and bumps it, so without this a temporary is issued the
    /// slot a named local already owns.
    ///
    /// Measured: `function h() { try { throw 1 } catch (e) { var c = 3; } return c; }`
    /// answered `true` — the `dyn_to_bool` temporary in the catch epilogue and
    /// the hoisted `c` were the same slot. Only the FIRST `var` in the arm was
    /// affected, which is the signature of exactly one contested index.
    fn publish_scope_allocator(&mut self) {
        let next = self.scopes.last().unwrap().next_slot;
        let chunk = &mut self.chunks[self.current];
        if chunk.local_count < next {
            chunk.local_count = next;
            if chunk.local_count > chunk.scratch_high_water {
                chunk.scratch_high_water = chunk.local_count;
            }
        }
    }

    /// Define a name in the enclosing FUNCTION's scope rather than the current
    /// block — ECMA-262's VariableEnvironment (§9.1.1.3), which is where a
    /// `VariableStatement` binding lives and why it outlives its block.
    ///
    /// ⛔ Always use this rather than `Scope::define_at_function_scope`: the
    /// allocator reconciliation above is not optional, and the raw method
    /// cannot do it because a `Scope` cannot see the chunk.
    pub(crate) fn define_var_scoped(
        &mut self,
        name: &str,
        type_hint: Option<vybe_ast::TypeHint>,
    ) -> u16 {
        self.sync_scope_allocator();
        let slot = self
            .scopes
            .last_mut()
            .unwrap()
            .define_at_function_scope(name, type_hint);
        self.publish_scope_allocator();
        slot
    }

    pub(crate) fn define_local(&mut self, name: &str) -> u16 {
        self.sync_scope_allocator();
        // Two independent questions, deliberately asked separately: does this
        // language scope variables to the function rather than the block, and
        // is this name a variable at all (rather than a compiler temporary)?
        // They coincide in PHP and need not anywhere else.
        let slot = if self.profile.function_scoped_variables && self.is_variable_name(name) {
            self.scopes
                .last_mut()
                .unwrap()
                .define_at_function_scope(name, None)
        } else {
            self.scopes.last_mut().unwrap().define(name)
        };
        let high = self.scopes.last().unwrap().next_slot;
        let cur = self.current;
        if high > self.chunks[cur].local_count {
            self.chunks[cur].local_count = high;
        }
        self.track_lexical_name(name);
        slot
    }

    /// Record a user binding for sloppy-mode unresolvable-read detection.
    /// No-op unless the profile has `unresolved_reference_throws` (JS only), so
    /// no other language pays for it or changes behavior. Internal temporaries
    /// (`__`-prefixed) are skipped — they are never read as user identifiers.
    pub(super) fn track_lexical_name(&mut self, name: &str) {
        if self.profile.unresolved_reference_throws && !name.starts_with("__") {
            let cname = self.canon(name);
            self.program_lexical_names.insert(cname);
        }
    }

    /// Stack: [coll, idx] → [coll, idx_norm]. For languages where
    /// negative array indices wrap from the end (Python `arr[-1]`,
    /// Ruby, PHP). Maps return length 0 from `ARRAY_LENGTH` so this
    /// is a no-op on dict-style collections (negative integer keys
    /// stay negative). Strings return char count → `s[-1]` works.
    pub(crate) fn emit_negative_index_wrap(&mut self) {
        let line = self.line;
        let arr_slot = self.define_local("__neg_idx_arr");
        let idx_slot = self.define_local("__neg_idx_i");
        // Stash [coll, idx] into locals (LOCAL_SET peeks; DROP pops).
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.emit_u16(Op::LOCAL_SET, arr_slot);
        // A string/Map key (`h["a"]`) must skip the wrap — the `idx < 0` compare
        // below coerces its operand, so a string key would hit `toF64` and trap.
        // Guard on js-string.test (small-int array indices are i32, which
        // js-number.test would reject — that's why we test for string, not number).
        let str_test = self.import("wasm:js-string", "test");
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_host_call(str_test, 1);
        self.chunk().emit_if(line); // string key → skip wrap
        self.chunk().emit_else(line); // non-string → wrap
        // if idx < 0: idx = arr.length + idx
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_const(Value::I32(0));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        };
        let block_p = self.chunk().emit_block(line);
        self.label_depth += 1;
        self.chunk().emit_br_if(0, line); // skip wrap if !(idx < 0)
        self.emit_u16(Op::LOCAL_GET, arr_slot);
        common::collections::emit_array_length(&mut self.chunks[self.current], line);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.chunk().emit_end(line);
        self.chunk().patch_block(block_p);
        self.label_depth -= 1;
        self.chunk().emit_end(line); // close the string-key guard `if`
        // Re-push [arr, idx_norm] for the caller's emit_get.
        self.emit_u16(Op::LOCAL_GET, arr_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
    }

    /// Emit RETURN, draining any function-local BLOCK/LOOP labels first.
    ///
    /// WASM `return` exits the function directly; it does not emit
    /// synthetic `end` instructions for surrounding blocks. The VM
    /// records each call frame's label-stack base and truncates labels
    /// when the frame returns, which keeps bytecode structurally valid
    /// even when a return appears inside an `if`/`else`, loop, or block.
    pub(crate) fn emit_return(&mut self) {
        self.emit(Op::RETURN);
    }

    pub(super) fn emit_active_finally_blocks(&mut self) -> Result<(), String> {
        if self.frame_cf().active_finally_blocks.is_empty() {
            return Ok(());
        }

        let original = self.frame_cf().active_finally_blocks.clone();
        for idx in (0..original.len()).rev() {
            self.frame_cf_mut().active_finally_blocks = original[..idx].to_vec();
            self.emit_finally_action(&original[idx])?;
        }
        self.frame_cf_mut().active_finally_blocks = original;
        Ok(())
    }

    pub(super) fn emit_finally_action(&mut self, action: &FinallyAction) -> Result<(), String> {
        match action {
            FinallyAction::Statements(stmts) => {
                for stmt in stmts {
                    self.compile_stmt(stmt)?;
                }
            }
            FinallyAction::ResourceDispose { slot, method, line } => {
                self.label_depth += 1;
                common::errors::emit_resource_dispose(self.chunk(), *slot, method, *line);
                self.label_depth -= 1;
            }
        }
        Ok(())
    }

    pub(super) fn emit_throw_through_finally(&mut self) -> Result<(), String> {
        // Inline ONLY the finallys whose try handler already fired (their
        // sequenced finally would be skipped by this throw). Finallys with
        // LIVE runtime handlers (enclosing try bodies) are run by the
        // runtime on unwind — inlining them here executed them twice.
        if self.frame_cf().fired_finally_indices.is_empty() {
            let line = self.line;
            common::errors::emit_throw(self.chunk(), line);
            return Ok(());
        }
        // Save the exception, run the fired finallys (innermost first),
        // then re-throw. Mirrors emit_return_through_finally's slicing so
        // control-flow statements INSIDE a finally see the right stack.
        let exc_slot = self.define_local("__throw_finally_exc");
        self.emit_u16(Op::LOCAL_SET, exc_slot);
        let fired = self.frame_cf().fired_finally_indices.clone();
        let original = self.frame_cf().active_finally_blocks.clone();
        for &idx in fired.iter().rev() {
            if idx >= original.len() {
                continue;
            }
            self.frame_cf_mut().active_finally_blocks = original[..idx].to_vec();
            self.emit_finally_action(&original[idx])?;
        }
        self.frame_cf_mut().active_finally_blocks = original;
        self.emit_u16(Op::LOCAL_GET, exc_slot);
        let line = self.line;
        common::errors::emit_throw(self.chunk(), line);
        Ok(())
    }

    pub(super) fn emit_return_through_finally(
        &mut self,
        result_count: usize,
    ) -> Result<(), String> {
        // Preferred path: a single-value return that must cross only
        // `try_table` joins (no `using`/dispose, no ref-out packing) routes
        // through the innermost join — `finally` runs OUTSIDE the handler, then
        // the join dispatch re-issues this return, chaining out to the real
        // frame exit. Same spec-correct lowering as break/continue.
        let ref_out_slots = self.current_ref_out_params.clone().unwrap_or_default();
        if result_count == 1
            && ref_out_slots.is_empty()
            && !self.frame_cf().finally_joins.is_empty()
            && self.frame_cf().finally_joins.len() == self.frame_cf().active_finally_blocks.len()
        {
            let val_slot = self.define_local("__return_val_0");
            self.emit_u16(Op::LOCAL_SET, val_slot);
            self.emit_completion_br(completion::RETURN, Some(val_slot));
            return Ok(());
        }

        let slots: Vec<u16> = (0..result_count)
            .map(|idx| self.define_local(&format!("__return_val_{}", idx)))
            .collect();
        for idx in (0..result_count).rev() {
            self.emit_u16(Op::LOCAL_SET, slots[idx]);
        }

        if !self.frame_cf().active_finally_blocks.is_empty() {
            self.emit_active_finally_blocks()?;
        }

        for slot in &slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }

        let ref_out_slots = self.current_ref_out_params.clone().unwrap_or_default();
        if !ref_out_slots.is_empty() && self.current_multi_return_arity().is_none() {
            for slot in &ref_out_slots {
                self.emit_u16(Op::LOCAL_GET, *slot);
            }
            let pack_count = result_count + ref_out_slots.len();
            let mut first = 0u16;
            for index in 0..pack_count {
                let slot = self.define_local(&format!("__return_pack_{}", index));
                if index == 0 {
                    first = slot;
                }
            }
            common::collections::emit_pack_n(
                &mut self.chunks,
                self.current,
                pack_count as u16,
                first,
                self.line,
            );
        }
        // The JS async-wrapper `try_table` handlers are popped by the VM's
        // `RETURN` (frame-scoped handler cleanup) — no explicit opcode needed.
        self.emit_return();
        Ok(())
    }

    pub(super) fn emit_break_through_finally(&mut self, label: Option<&str>) -> Result<(), String> {
        let target_ctx = if let Some(lbl) = label {
            self.loops
                .iter()
                .rev()
                .find(|c| c.label.as_deref() == Some(lbl))
        } else {
            self.loops.last()
        };

        if let Some(ctx) = target_ctx {
            let target_finally_depth = ctx.finally_depth;
            let nested_finally_count = self.frame_cf()
                .active_finally_blocks
                .len()
                .saturating_sub(target_finally_depth);
            if nested_finally_count > 0 {
                // Preferred path: if EVERY enclosing finally between here and
                // the target loop is a `try_table` join, route to the innermost
                // one. It runs its `finally` OUTSIDE the handler, then re-issues
                // this break — chaining out through the remaining joins. This is
                // the spec-correct lowering (a throwing finally propagates
                // rather than being self-caught). See `finally_joins`.
                if label.is_none()
                    && self.finally_join_route(ctx.break_label_depth, nested_finally_count)
                {
                    self.emit_completion_br(completion::BREAK, None);
                    return Ok(());
                }
                // Fallback (labeled exit, or a `using`/mixed finally without a
                // join): inline the finally bodies. The exited `try_table`
                // handlers are popped structurally by the `br` below.
                let original = self.frame_cf().active_finally_blocks.clone();
                for idx in (target_finally_depth..original.len()).rev() {
                    self.frame_cf_mut().active_finally_blocks = original[..idx].to_vec();
                    self.emit_finally_action(&original[idx])?;
                }
                self.frame_cf_mut().active_finally_blocks = original;
            }
        }

        if let Some(depth) = self.break_depth(label) {
            let line = self.line;
            self.chunk().emit_br(depth.into(), line);
        }
        Ok(())
    }

    /// True if all `nested_finally_count` finallys between the current point
    /// and a loop whose exit is at `break_label_depth` are `try_table` joins
    /// (so the completion-code route can chain through them). False if any is a
    /// `using`/dispose without a registered join — then the caller inlines.
    fn finally_join_route(&self, break_label_depth: u32, nested_finally_count: usize) -> bool {
        let joins_between = self.frame_cf()
            .finally_joins
            .iter()
            .filter(|j| j.join_label_depth > break_label_depth)
            .count();
        joins_between == nested_finally_count && !self.frame_cf().finally_joins.is_empty()
    }

    /// Store `code` in the innermost join's completion slot (plus save the
    /// return value into its `ret_slot` when `ret` is given) and `br` to that
    /// join, where `finally` runs outside the handler and dispatches onward.
    fn emit_completion_br(&mut self, code: f64, ret: Option<u16>) {
        let join = self.frame_cf().finally_joins.last().expect("join present");
        let completion_slot = join.completion_slot;
        let ret_slot = join.ret_slot;
        let join_label_depth = join.join_label_depth;
        if let Some(ret_local) = ret {
            self.emit_u16(Op::LOCAL_GET, ret_local);
            self.emit_u16(Op::LOCAL_SET, ret_slot);
        }
        self.emit_const(Value::F64(code));
        self.emit_u16(Op::LOCAL_SET, completion_slot);
        let depth = self.label_depth.saturating_sub(join_label_depth);
        let line = self.line;
        self.chunk().emit_br(depth, line);
    }

    pub(super) fn emit_continue_through_finally(
        &mut self,
        label: Option<&str>,
    ) -> Result<(), String> {
        let target_ctx = if let Some(lbl) = label {
            self.loops
                .iter()
                .rev()
                .find(|c| c.label.as_deref() == Some(lbl) && c.is_continuable)
        } else {
            self.loops.iter().rev().find(|c| c.is_continuable)
        };

        if let Some(ctx) = target_ctx {
            let target_finally_depth = ctx.finally_depth;
            let nested_finally_count = self.frame_cf()
                .active_finally_blocks
                .len()
                .saturating_sub(target_finally_depth);
            if nested_finally_count > 0 {
                if label.is_none()
                    && self.finally_join_route(ctx.continue_label_depth, nested_finally_count)
                {
                    self.emit_completion_br(completion::CONTINUE, None);
                    return Ok(());
                }
                // Fallback: inline finally bodies (labeled / using / mixed).
                let original = self.frame_cf().active_finally_blocks.clone();
                for idx in (target_finally_depth..original.len()).rev() {
                    self.frame_cf_mut().active_finally_blocks = original[..idx].to_vec();
                    self.emit_finally_action(&original[idx])?;
                }
                self.frame_cf_mut().active_finally_blocks = original;
            }
        }

        if let Some(depth) = self.continue_depth(label) {
            let line = self.line;
            self.chunk().emit_br(depth.into(), line);
        }
        Ok(())
    }

    pub(super) fn current_chunk_is_js_async(&self) -> bool {
        // `is_async` is source truth and includes async GENERATORS, but
        // their bodies must not promise-wrap returns or convert throws to
        // rejections: a generator completes/throws through `resume`, and
        // the §27.6.1.2 promise surface lives in the attached
        // `__vybe_async_generator_next` driver instead.
        self.async_wraps_body_in_try()
            && self.chunks[self.current].is_async
            && !self.chunks[self.current].is_generator
    }

    /// Same as `define_local` but with a type hint — sugar around
    /// `Scope::define_typed`. Keeps the sync invariant.
    pub(crate) fn define_local_typed(
        &mut self,
        name: &str,
        type_hint: Option<vybe_ast::TypeHint>,
    ) -> u16 {
        {
            let scope = self.scopes.last_mut().unwrap();
            let chunk_locals = self.chunks[self.current].local_count;
            if scope.next_slot < chunk_locals {
                scope.next_slot = chunk_locals;
            }
            if let Some(dup_slot) = self.chunks[self.current].dup_slot {
                if scope.next_slot <= dup_slot {
                    scope.next_slot = dup_slot + 1;
                }
            }
        }
        // See `define_local` — same two independent questions.
        let slot = if self.profile.function_scoped_variables && self.is_variable_name(name) {
            self.scopes
                .last_mut()
                .unwrap()
                .define_at_function_scope(name, type_hint)
        } else {
            self.scopes
                .last_mut()
                .unwrap()
                .define_typed(name, type_hint)
        };
        let high = self.scopes.last().unwrap().next_slot;
        let cur = self.current;
        if high > self.chunks[cur].local_count {
            self.chunks[cur].local_count = high;
        }
        self.track_lexical_name(name);
        slot
    }
    pub(crate) fn chunk(&mut self) -> &mut Chunk {
        &mut self.chunks[self.current]
    }

    pub(crate) fn reserve_local_slot(&mut self, slot: u16) {
        self.chunks[self.current].local_count = self.chunks[self.current].local_count.max(slot + 1);
    }

    pub(crate) fn emit(&mut self, op: Op) {
        let l = self.line;
        self.chunks[self.current].emit_op(op, l);
    }
    /// `ref.null extern` — the lenient null every dynamic language's
    /// `null`/`NULL`/`None`/`nil` compiles to. `ref.null` takes a heaptype
    /// immediate per spec; the GC-heap heaptypes give a typed null that traps
    /// on the GC accessors, which is NOT what these sites want.
    pub(crate) fn emit_null(&mut self) {
        let l = self.line;
        self.chunks[self.current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, l);
    }
    pub(crate) fn emit_u16(&mut self, op: Op, v: u16) {
        let l = self.line;
        self.chunks[self.current].emit_op_u16(op, v, l);
    }
    /// `array.new_fixed $t N`. Type index `0` = dynamic-language array literal.
    /// `struct.get`/`get_s`/`get_u`/`set` — `(typeidx, idx)`; typeidx 0 keeps
    /// `idx` a field-NAME constant index.
    pub(crate) fn emit_struct_field_op(
        &mut self,
        op: vybe_runtime::opcode::Op,
        typeidx: u16,
        idx: u16,
    ) {
        let l = self.line;
        self.chunks[self.current].emit_struct_field_op(op, typeidx, idx, l);
    }

    /// `struct.new` — typeidx 0 is the dynamic object-literal form; `count`
    /// is then the number of key/value pairs on the stack.
    pub(crate) fn emit_struct_new(&mut self, typeidx: u16, count: u16) {
        let l = self.line;
        self.chunks[self.current].emit_struct_new(typeidx, count, l);
    }

    pub(crate) fn emit_array_new_fixed(&mut self, typeidx: u16, count: u16) {
        let l = self.line;
        self.chunks[self.current].emit_array_new_fixed(typeidx, count, l);
    }
    #[allow(dead_code)]
    pub(crate) fn emit_u8(&mut self, op: Op, v: u8) {
        let l = self.line;
        self.chunks[self.current].emit_op_u8(op, v, l);
    }
    pub(crate) fn emit_u8_u8(&mut self, op: Op, a: u8, b: u8) {
        let l = self.line;
        self.chunks[self.current].emit_op_u8_u8(op, a, b, l);
    }
    /// Constant encoding lives in ONE place — `primitives::datetime::push_const`
    /// — so this method and the 1432 adapter call sites that use the free
    /// function cannot encode the same literal two different ways.
    pub(crate) fn emit_const(&mut self, val: Value) {
        let l = self.line;
        crate::primitives::datetime::push_const(&mut self.chunks[self.current], val, l);
    }

    /// Compute WASM `br` depth for `break`.
    pub(super) fn break_depth(&self, label: Option<&str>) -> Option<u8> {
        let ctx = if let Some(lbl) = label {
            self.loops
                .iter()
                .rev()
                .find(|c| c.label.as_deref() == Some(lbl))?
        } else {
            self.loops.last()?
        };
        Some((self.label_depth - ctx.break_label_depth) as u8)
    }

    pub(super) fn iterator_close_slot_for_break(&self, label: Option<&str>) -> Option<u16> {
        let ctx = if let Some(lbl) = label {
            self.loops
                .iter()
                .rev()
                .find(|c| c.label.as_deref() == Some(lbl))?
        } else {
            self.loops.last()?
        };
        ctx.iterator_close_slot
    }

    pub(super) fn emit_js_iterator_close(&mut self, iterator_slot: u16) {
        if !self.profile.ecma_iterator_result_shape {
            return;
        }
        let line = self.line;
        let return_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("return"));
        let _function_str = self.str_const("function");
        let return_fn_slot = self.define_local("__iterator_close_return");

        self.emit_u16(Op::LOCAL_GET, iterator_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &return_key);
        self.emit_u16(Op::LOCAL_SET, return_fn_slot);

        self.emit_u16(Op::LOCAL_GET, return_fn_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        inst!(self, core_wasm::string_const, "function");
        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, iterator_slot);
        self.emit_global_write("__js_this");
        self.emit_u16(Op::LOCAL_GET, return_fn_slot);
        self.emit_direct_callable_invoke(0);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);
    }

    pub(super) fn emit_active_js_iterator_closes(&mut self) {
        if !self.profile.ecma_iterator_result_shape {
            return;
        }
        let slots: Vec<u16> = self
            .loops
            .iter()
            .rev()
            .filter_map(|ctx| ctx.iterator_close_slot)
            .collect();
        for slot in slots {
            self.emit_js_iterator_close(slot);
        }
    }

    /// Compute WASM `br` depth for `continue`.
    pub(super) fn continue_depth(&self, label: Option<&str>) -> Option<u8> {
        let ctx = if let Some(lbl) = label {
            self.loops
                .iter()
                .rev()
                .find(|c| c.label.as_deref() == Some(lbl))?
        } else {
            // Skip switch/labeled-block contexts — `continue` targets the
            // nearest actual loop (ECMA-262 §14.8.1).
            self.loops.iter().rev().find(|c| c.is_continuable)?
        };
        Some((self.label_depth - ctx.continue_label_depth) as u8)
    }
}


// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// Linkable chunk builders — the standalone-chunk packaging of what the
// `emit_*` forms splice inline. A language prefix in a name records which
// frontend first needed a linkable chunk, not a language-specific meaning.

// ── iif(c, a, b) → value — VB IIf eager-evaluated ternary ────────
//
// Args are evaluated before call (eager — both branches always run),
// matching .NET `IIf(condition, truePart, falsePart)`. SELECT picks the
// correct one. Note: this is NOT a short-circuiting `If(...)` — VB has
// distinct lazy `If(c, a, b)` operator handled at compile time elsewhere.
pub fn build_iif(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_iif");
    c.arity = 3;
    c.local_count = 3;
    // SELECT pops [a, b, cond]; returns a if cond truthy.
    // Args land in locals in declaration order: cond=0, a=1, b=2.
    // We need stack [a, b, cond] for SELECT.
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // a (true branch)
    c.emit_op_u16(Op::LOCAL_GET, 2, 0); // b (false branch)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // cond
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_op(Op::SELECT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── `goto` / label → structured control flow ──────────────────────────────
//
// WASM has no goto, so every language that supports one lowers HERE. This
// used to live in the C walker, with PHP reaching across into
// `vybe_language_c::walker::lower_gotos` to borrow it — a language crate
// acting as the home for shared machinery.

fn goto_stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}

fn goto_expr(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn goto_ident(name: &str) -> Expression {
    goto_expr(ExprKind::Ident(name.to_string()))
}

fn goto_int_lit(value: i64) -> Expression {
    goto_expr(ExprKind::Lit(Literal::Int(value)))
}

fn goto_assign(target: Expression, value: Expression) -> Expression {
    goto_expr(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })
}

/// Language-agnostic `goto`/label → structured control flow: split the block at
/// each label into numbered sub-blocks, then run a `while(true) { switch(pc) }`
/// state machine (`goto L` becomes `pc = block(L); continue dispatch`). WASM has
/// no goto, so every language that supports it lowers here. `pc_name` is the
/// program-counter variable name in the TARGET language's convention (C uses a
/// bare identifier; PHP passes a `$`-prefixed name); `dispatch_label` names the
/// wrapping loop for the labeled `continue`.
/// `fold_labels` folds label names to lowercase, for the frontends that declare
/// `case_sensitive = false` (pascal, vb, cobol, fortran). It is the same
/// name-kind folding as [`Scope::fold_case`] for locals and
/// `LanguageProfile::fold_callable_names` for callables — a label is simply a
/// third kind of name — so it is REQUIRED rather than defaulted, per the
/// `fold_case` lesson in `documentation/directives.md`: 23 of 33 call sites
/// forgot that guard once and silently broke Go.
pub fn lower_gotos(
    body: Vec<Statement>,
    pc_name_arg: &str,
    dispatch_label_arg: &str,
    fold_labels: bool,
) -> Vec<Statement> {
    let fold = |name: &str| {
        if fold_labels {
            name.to_lowercase()
        } else {
            name.to_string()
        }
    };
    // A label written as `L: stmt` arrives as `StmtKind::Labeled`, not a bare
    // `Label` — Go and PHP both spell goto targets that way. Only a label some
    // `goto` actually NAMES is flattened into a split point; a `Labeled` that
    // exists for `break L` / `continue L` on a loop is left alone, or the loop
    // would lose the target its own jumps resolve against.
    let mut goto_targets = std::collections::HashSet::new();
    for s in &body {
        collect_goto_targets(s, &mut goto_targets);
    }
    let body: Vec<Statement> = body
        .into_iter()
        .flat_map(|s| match s.kind {
            StmtKind::Labeled { label, body: inner } if goto_targets.contains(&label) => {
                let mut out = vec![Statement::new(StmtKind::Label(label))];
                match inner.kind {
                    StmtKind::Block(stmts) => out.extend(stmts),
                    other => out.push(Statement::new(other)),
                }
                out
            }
            other => vec![Statement::new(other)],
        })
        .collect();

    let mut label_to_block: std::collections::HashMap<String, i64> =
        std::collections::HashMap::new();
    let mut blocks: Vec<Vec<Statement>> = vec![Vec::new()];

    for s in body {
        if let StmtKind::Label(name) = s.kind {
            let idx = blocks.len() as i64;
            label_to_block.insert(fold(&name), idx);
            blocks.push(Vec::new());
        } else if let Some(last) = blocks.last_mut() {
            last.push(s);
        }
    }

    if label_to_block.is_empty() {
        return blocks.into_iter().next().unwrap_or_default();
    }

    // A DECLARATION before the first label is not a step in the flow, and must
    // stay visible to every branch — so it is hoisted above the dispatch loop
    // rather than sealed inside block 0. `Statement::is_declaration` answers
    // this from the node, so it is right for every language at once.
    let mut prelude = Vec::new();
    if let Some(first_block) = blocks.first_mut() {
        while first_block
            .first()
            .map(Statement::is_declaration)
            .unwrap_or(false)
        {
            prelude.push(first_block.remove(0));
        }
    }

    let dispatch_label = dispatch_label_arg.to_string();
    let pc_name = pc_name_arg.to_string();

    let mut switch_cases = Vec::new();
    let total_blocks = blocks.len();
    for (idx, block) in blocks.into_iter().enumerate() {
        let next_pc = if idx + 1 < total_blocks {
            goto_int_lit((idx + 1) as i64)
        } else {
            goto_int_lit(-1)
        };
        let mut case_body = vec![goto_stmt(StmtKind::Expr(goto_assign(
            goto_ident(&pc_name),
            next_pc,
        )))];
        case_body.extend(rewrite_gotos_in_stmts(
            block,
            &label_to_block,
            &pc_name,
            &dispatch_label,
            fold_labels,
        ));
        case_body.push(goto_stmt(StmtKind::Break(BreakTarget::Implicit)));
        switch_cases.push(SwitchCase {
            conditions: vec![CaseCondition::Value(goto_int_lit(idx as i64))],
            body: case_body,
        });
    }

    let while_body = vec![
        goto_stmt(StmtKind::Switch {
            expr: goto_ident(&pc_name),
            cases: switch_cases,
            default: Some(vec![goto_stmt(StmtKind::Break(BreakTarget::Implicit))]),
        }),
        goto_stmt(StmtKind::If {
            cond: goto_expr(ExprKind::Binary {
                op: BinOp::Lt,
                left: Box::new(goto_ident(&pc_name)),
                right: Box::new(goto_int_lit(0)),
            }),
            then_body: vec![goto_stmt(StmtKind::Break(BreakTarget::Implicit))],
            elifs: vec![],
            else_body: None,
        }),
    ];

    let mut lowered = prelude;
    lowered.push(goto_stmt(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(pc_name.clone()),
            // The block counter IS an integer, in every language — so the
            // declaration says so once, canonically, and each frontend's own
            // spelling machinery renders it (`Integer` for pascal, nothing for
            // lua). Passing the spelling in would put one language's word for
            // `int` inside a shared pass.
            type_hint: Some("int".to_string().into()),
            init: Some(goto_int_lit(0)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    lowered.push(goto_stmt(StmtKind::Labeled {
        label: dispatch_label,
        body: Box::new(goto_stmt(StmtKind::While {
            cond: goto_expr(ExprKind::Lit(Literal::Bool(true))),
            body: while_body,
            else_body: None,
        })),
    }));
    lowered
}

fn rewrite_gotos_in_stmts(
    stmts: Vec<Statement>,
    label_to_block: &std::collections::HashMap<String, i64>,
    pc_name: &str,
    dispatch_label: &str,
    fold_labels: bool,
) -> Vec<Statement> {
    let mut out = Vec::new();
    for stmt_in in stmts {
        match stmt_in.kind {
            StmtKind::GoTo(target) => {
                // Folded on BOTH sides or not at all — the map was keyed the
                // same way when the labels were collected.
                let key = if fold_labels {
                    target.to_lowercase()
                } else {
                    target.clone()
                };
                if let Some(idx) = label_to_block.get(&key) {
                    out.push(goto_stmt(StmtKind::Expr(goto_assign(
                        goto_ident(pc_name),
                        goto_int_lit(*idx),
                    ))));
                    out.push(goto_stmt(StmtKind::Continue(ContinueTarget::Label(
                        dispatch_label.to_string(),
                    ))));
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                let then_body =
                    rewrite_gotos_in_stmts(then_body, label_to_block, pc_name, dispatch_label, fold_labels);
                let elifs = elifs
                    .into_iter()
                    .map(|(c, b)| {
                        (
                            c,
                            rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label, fold_labels),
                        )
                    })
                    .collect();
                let else_body = else_body
                    .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label, fold_labels));
                out.push(goto_stmt(StmtKind::If {
                    cond,
                    then_body,
                    elifs,
                    else_body,
                }));
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                out.push(goto_stmt(StmtKind::For {
                    init,
                    cond,
                    update,
                    body: rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label, fold_labels),
                }));
            }
            StmtKind::ForIn {
                var,
                key,
                iter,
                body,
                of,
                else_body,
                is_async,
            } => {
                out.push(goto_stmt(StmtKind::ForIn {
                    var,
                    key,
                    iter,
                    body: rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label, fold_labels),
                    of,
                    else_body: else_body.map(|b| {
                        rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label, fold_labels)
                    }),
                    is_async,
                }));
            }
            StmtKind::While {
                cond,
                body,
                else_body,
            } => {
                out.push(goto_stmt(StmtKind::While {
                    cond,
                    body: rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label, fold_labels),
                    else_body: else_body.map(|b| {
                        rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label, fold_labels)
                    }),
                }));
            }
            StmtKind::DoWhile { body, cond, until } => {
                out.push(goto_stmt(StmtKind::DoWhile {
                    body: rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label, fold_labels),
                    cond,
                    until,
                }));
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                let cases = cases
                    .into_iter()
                    .map(|mut c| {
                        c.body =
                            rewrite_gotos_in_stmts(c.body, label_to_block, pc_name, dispatch_label, fold_labels);
                        c
                    })
                    .collect();
                let default = default
                    .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label, fold_labels));
                out.push(goto_stmt(StmtKind::Switch {
                    expr,
                    cases,
                    default,
                }));
            }
            StmtKind::Block(body) => {
                out.push(goto_stmt(StmtKind::Block(rewrite_gotos_in_stmts(
                    body,
                    label_to_block,
                    pc_name,
                    dispatch_label,
                    fold_labels,
                ))));
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                let body = rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label, fold_labels);
                let catches = catches
                    .into_iter()
                    .map(|mut c| {
                        c.body =
                            rewrite_gotos_in_stmts(c.body, label_to_block, pc_name, dispatch_label, fold_labels);
                        c
                    })
                    .collect();
                let else_body = else_body
                    .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label, fold_labels));
                let finally = finally
                    .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label, fold_labels));
                out.push(goto_stmt(StmtKind::Try {
                    body,
                    catches,
                    else_body,
                    finally,
                }));
            }
            StmtKind::Labeled { label, body } => {
                out.push(goto_stmt(StmtKind::Labeled {
                    label,
                    body: Box::new(goto_stmt(match body.kind {
                        StmtKind::Block(inner) => StmtKind::Block(rewrite_gotos_in_stmts(
                            inner,
                            label_to_block,
                            pc_name,
                            dispatch_label,
                            fold_labels,
                        )),
                        other => other,
                    })),
                }));
            }
            StmtKind::Label(_) => {}
            _ => out.push(stmt_in),
        }
    }
    out
}

/// Every label named by a `goto` anywhere inside `s`, however deeply nested.
fn collect_goto_targets(s: &Statement, out: &mut std::collections::HashSet<String>) {
    match &s.kind {
        StmtKind::GoTo(target) => {
            out.insert(target.clone());
        }
        StmtKind::Block(body)
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. } => {
            for inner in body {
                collect_goto_targets(inner, out);
            }
        }
        StmtKind::Labeled { body, .. } => collect_goto_targets(body, out),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for inner in then_body {
                collect_goto_targets(inner, out);
            }
            for (_, b) in elifs {
                for inner in b {
                    collect_goto_targets(inner, out);
                }
            }
            if let Some(b) = else_body {
                for inner in b {
                    collect_goto_targets(inner, out);
                }
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for c in cases {
                for inner in &c.body {
                    collect_goto_targets(inner, out);
                }
            }
            if let Some(b) = default {
                for inner in b {
                    collect_goto_targets(inner, out);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for inner in body {
                collect_goto_targets(inner, out);
            }
            for c in catches {
                for inner in &c.body {
                    collect_goto_targets(inner, out);
                }
            }
            for b in [else_body, finally].into_iter().flatten() {
                for inner in b {
                    collect_goto_targets(inner, out);
                }
            }
        }
        _ => {}
    }
}
