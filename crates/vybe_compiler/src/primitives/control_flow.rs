//! Generators, for-of iterators, multi-return, pointer-cell promotion, finally/iterator unwinding.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

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
            None => self.emit_exit_from_stack(0) }
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
        let js_this = self.str_const("__js_this");
        let _iterator_key = self.str_const("iterator");
        let _next_key_c = self.str_const("next");
        let done_key_c = self.str_const("done");
        let value_key_c = self.str_const("value");

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
        let iterator_key = self.str_const("iterator");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, iterator_key);
        let iter_fn_slot = self.define_local("__cit_iter_fn");
        self.emit_u16(Op::LOCAL_SET, iter_fn_slot);

        // Call iterator() with __js_this = iter_slot
        self.emit_u16(Op::LOCAL_GET, iter_slot);
        self.emit_u16(Op::GLOBAL_SET, js_this);
        self.emit_u16(Op::LOCAL_GET, iter_fn_slot);
        self.emit_u8(Op::CALL_REF, 0);
        self.emit_u16(Op::LOCAL_SET, it_slot);

        // Emit BLOCK + LOOP
        let block_patch = self.chunk().emit_block(line);
        let (loop_patch, _) = self.chunk().emit_loop_s(line);
        self.label_depth += 2;

        // next_method = it.next via STRUCT_GET
        self.emit_u16(Op::LOCAL_GET, it_slot);
        let next_key_c = self.str_const("next");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, next_key_c);
        self.emit_u16(Op::LOCAL_SET, next_method_slot);

        // Call next() with __js_this = it
        self.emit_u16(Op::LOCAL_GET, it_slot);
        self.emit_u16(Op::GLOBAL_SET, js_this);
        self.emit_u16(Op::LOCAL_GET, next_method_slot);
        self.emit_u8(Op::CALL_REF, 0);
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
        self.emit_struct_field_op(Op::STRUCT_GET, 0, done_key_c);
        self.emit_u16(Op::LOCAL_SET, done_slot);
        self.emit_u16(Op::LOCAL_GET, done_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        self.chunk().emit_br_if(1, line); // done → exit block

        // var = step.value
        self.emit_u16(Op::LOCAL_GET, step_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, value_key_c);
        let var_slot = self.define_local(var);
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
            finally_depth: self.active_finally_blocks.len() });
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
            _ => unreachable!("compile_generator_for_in expects Call") };
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
            let key_slot = self.define_local(key_name);
            if self.profile.buffered_iterator_methods {
                self.emit_buffered_generator_key_binding(key_slot, value_slot, key_index_slot);
            } else {
                self.emit_null();
                self.emit_u16(Op::LOCAL_SET, key_slot);
            }
        }

        let var_slot = self.define_local(var);
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
            finally_depth: self.active_finally_blocks.len() });
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
        let returned_key = self.str_const("__vybe_gen_returned");
        self.emit_struct_field_op(Op::STRUCT_SET, 0, returned_key);
        self.emit(Op::DROP);
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
            _ => None }
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
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
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
        let marker_key = self.str_const(&marker);
        self.emit_struct_field_op(Op::STRUCT_SET, 0, marker_key);
        self.emit(Op::DROP);
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
                        _ => return None }
                }
                names
            }
            _ => return None };
        let multi_n = match &value.kind {
            ExprKind::Call { callee, args, .. } => {
                let _ = args;
                self.multi_return_arity_for_callee(callee)?
            }
            _ => return None };
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
                    finally } => {
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

    pub(super) fn binding_uses_pointer_cell(&self, name: &str) -> bool {
        let key = self.pointer_binding_key(name);
        if self
            .pointer_cell_bindings
            .get(&self.current)
            .is_some_and(|bindings| bindings.contains(&key))
        {
            return true;
        }
        // The map is keyed by chunk, and the module-wide fallback below exists
        // for a name this chunk does NOT bind itself — a closure body reading a
        // cell created in its enclosing chunk, or a promoted global.
        //
        // A name this chunk DOES bind (parameter or local) shadows every
        // same-named binding elsewhere, so the fallback must not answer for it.
        // Without this, Go's `c.Bump()` promoting `main`'s local `c` to a cell
        // made every other chunk with a parameter named `c` — such as the
        // receiver of `func (c Counter) Peek() int { return c.n }` — read
        // through a cell that isn't there, yielding `undefined`.
        if self.resolve_named_local_slot(name).is_some() {
            return false;
        }
        self.pointer_cell_bindings
            .values()
            .any(|bindings| bindings.contains(&key))
    }

    pub(super) fn mark_pointer_cell_binding(&mut self, name: &str) {
        let key = self.pointer_binding_key(name);
        self.pointer_cell_bindings
            .entry(self.current)
            .or_default()
            .insert(key);
    }

    pub(super) fn resolve_named_local_slot(&self, name: &str) -> Option<u16> {
        self.scope().resolve(name)
    }

    pub(super) fn promote_local_binding_to_pointer_cell(&mut self, name: &str) -> Option<u16> {
        let slot = self.resolve_named_local_slot(name)?;
        if !self.binding_uses_pointer_cell(name) {
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

    pub(super) fn promote_global_binding_to_pointer_cell(&mut self, name: &str) -> bool {
        let canon_name = self.canon(name);
        if !self.profile.globals_may_be_undeclared && !self.defined_globals.contains(&canon_name) {
            return false;
        }

        if !self.binding_uses_pointer_cell(name) {
            let value_slot = self.define_local("__ref_global_value");
            let idx = self.str_const(&canon_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u16(Op::LOCAL_SET, value_slot);
            crate::primitives::references::emit_cell_new(
                &mut self.chunks,
                self.current,
                value_slot,
                self.line,
            );
            self.emit_u16(Op::GLOBAL_SET, idx);
            self.mark_pointer_cell_binding(name);
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
        self.chunk().emit_if(line);
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
        self.chunk().emit_if(obj_line);

        let kind_key = self.str_const("__ref_kind");

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, kind_key);
        self.emit_string_eq_literal("cell");
        let cell_line = self.line;
        self.chunk().emit_if(cell_line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        crate::primitives::references::emit_cell_load(&mut self.chunks, self.current, self.line);
        self.chunk().emit_else(cell_line);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, kind_key);
        self.emit_string_eq_literal("carray");
        let carray_line = self.line;
        self.chunk().emit_if(carray_line);

        let base_key = self.str_const("__base");
        let idx_key = self.str_const("__idx");
        let base_slot = self.define_local("__ref_carray_base");

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, base_key);
        self.emit_u16(Op::LOCAL_SET, base_slot);

        self.emit_u16(Op::LOCAL_GET, base_slot);
        inst!(self, recipes::is_object);
        let base_obj_line = self.line;
        self.chunk().emit_if(base_obj_line);

        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, kind_key);
        self.emit_string_eq_literal("cell");
        let base_cell_line = self.line;
        self.chunk().emit_if(base_cell_line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        crate::primitives::references::emit_cell_load(&mut self.chunks, self.current, self.line);
        self.chunk().emit_else(base_cell_line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, idx_key);
        common::collections::emit_get(&mut self.chunks, self.current, self.line);
        self.chunk().emit_end(base_cell_line);

        self.chunk().emit_else(base_obj_line);
        self.emit_u16(Op::LOCAL_GET, base_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, idx_key);
        common::collections::emit_get(&mut self.chunks, self.current, self.line);
        self.chunk().emit_end(base_obj_line);

        self.chunk().emit_else(carray_line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.chunk().emit_end(carray_line);

        self.chunk().emit_end(cell_line);

        self.chunk().emit_else(obj_line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.chunk().emit_end(obj_line);
    }

    pub(super) fn compile_address_of_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            ExprKind::Ident(name) => {
                let canon_name = self.canon(name);
                if self.defined_functions.contains(&canon_name) {
                    self.emit_var_get(name);
                    return Ok(());
                }
                if let Some(slot) = self.promote_local_binding_to_pointer_cell(name) {
                    self.emit_u16(Op::LOCAL_GET, slot);
                    return Ok(());
                }
                if self.promote_global_binding_to_pointer_cell(name) {
                    let idx = self.str_const(&canon_name);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    return Ok(());
                }
            }
            ExprKind::Unary {
                op: UnaryOp::Deref,
                expr } => {
                self.compile_expr(expr)?;
                return Ok(());
            }
            _ => {}
        }

        self.compile_expr(expr)?;
        self.emit_wrap_top_of_stack_in_pointer_cell();
        Ok(())
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
    pub(crate) fn define_local(&mut self, name: &str) -> u16 {
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
        if self.active_finally_blocks.is_empty() {
            return Ok(());
        }

        let original = self.active_finally_blocks.clone();
        for idx in (0..original.len()).rev() {
            self.active_finally_blocks = original[..idx].to_vec();
            self.emit_finally_action(&original[idx])?;
        }
        self.active_finally_blocks = original;
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
        if self.fired_finally_indices.is_empty() {
            let line = self.line;
            common::errors::emit_throw(self.chunk(), line);
            return Ok(());
        }
        // Save the exception, run the fired finallys (innermost first),
        // then re-throw. Mirrors emit_return_through_finally's slicing so
        // control-flow statements INSIDE a finally see the right stack.
        let exc_slot = self.define_local("__throw_finally_exc");
        self.emit_u16(Op::LOCAL_SET, exc_slot);
        let fired = self.fired_finally_indices.clone();
        let original = self.active_finally_blocks.clone();
        for &idx in fired.iter().rev() {
            if idx >= original.len() {
                continue;
            }
            self.active_finally_blocks = original[..idx].to_vec();
            self.emit_finally_action(&original[idx])?;
        }
        self.active_finally_blocks = original;
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
            && !self.finally_joins.is_empty()
            && self.finally_joins.len() == self.active_finally_blocks.len()
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

        if !self.active_finally_blocks.is_empty() {
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
            let nested_finally_count = self
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
                let original = self.active_finally_blocks.clone();
                for idx in (target_finally_depth..original.len()).rev() {
                    self.active_finally_blocks = original[..idx].to_vec();
                    self.emit_finally_action(&original[idx])?;
                }
                self.active_finally_blocks = original;
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
        let joins_between = self
            .finally_joins
            .iter()
            .filter(|j| j.join_label_depth > break_label_depth)
            .count();
        joins_between == nested_finally_count && !self.finally_joins.is_empty()
    }

    /// Store `code` in the innermost join's completion slot (plus save the
    /// return value into its `ret_slot` when `ret` is given) and `br` to that
    /// join, where `finally` runs outside the handler and dispatches onward.
    fn emit_completion_br(&mut self, code: f64, ret: Option<u16>) {
        let join = self.finally_joins.last().expect("join present");
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
            let nested_finally_count = self
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
                let original = self.active_finally_blocks.clone();
                for idx in (target_finally_depth..original.len()).rev() {
                    self.active_finally_blocks = original[..idx].to_vec();
                    self.emit_finally_action(&original[idx])?;
                }
                self.active_finally_blocks = original;
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
        self.profile.async_wraps_body_in_try
            && self.chunks[self.current].is_async
            && !self.chunks[self.current].is_generator
    }

    /// Same as `define_local` but with a type hint — sugar around
    /// `Scope::define_typed`. Keeps the sync invariant.
    pub(crate) fn define_local_typed(&mut self, name: &str, type_hint: Option<String>) -> u16 {
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
    pub(crate) fn emit_u8(&mut self, op: Op, v: u8) {
        let l = self.line;
        self.chunks[self.current].emit_op_u8(op, v, l);
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
        let return_key = self.str_const("return");
        let _function_str = self.str_const("function");
        let js_this = self.str_const("__js_this");
        let return_fn_slot = self.define_local("__iterator_close_return");

        self.emit_u16(Op::LOCAL_GET, iterator_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, return_key);
        self.emit_u16(Op::LOCAL_SET, return_fn_slot);

        self.emit_u16(Op::LOCAL_GET, return_fn_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        inst!(self, core_wasm::string_const, "function");
        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, iterator_slot);
        self.emit_u16(Op::GLOBAL_SET, js_this);
        self.emit_u16(Op::LOCAL_GET, return_fn_slot);
        self.emit_u8(Op::CALL_REF, 0);
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
