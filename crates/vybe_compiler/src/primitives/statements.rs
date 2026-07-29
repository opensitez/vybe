//! Statement & declaration compilation: the `compile_stmt` dispatch plus
//! enum/var/pattern/assignment lowering.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — same pattern as
//! `builtins.rs`/`operators.rs`.

use super::*;

impl Compiler {
    pub(super) fn compile_stmt(&mut self, stmt: &Statement) -> Result<(), String> {
        self.line = stmt.span.start_line;
        // Runtime-prelude boundary marker: a frontend that prepends a prelude
        // (e.g. JS) injects a `__vybe_user_code_start__` string-expression right
        // before the user's own code. Record the current bytecode offset on the
        // chunk (the `<script>`) for the debugger's prelude-skip, and emit
        // nothing. Generic — not gated on any language name.
        if let StmtKind::Expr(expr) = &stmt.kind {
            if let ExprKind::Lit(Literal::Str(s)) = &expr.kind {
                if s == "__vybe_user_code_start__" {
                    let off = self.chunks[self.current].code.len() as u32;
                    self.chunks[self.current].user_code_offset = Some(off);
                    return Ok(());
                }
            }
        }
        match &stmt.kind {
            // ── Expression statement ────────────────────────────────────
            StmtKind::Expr(expr) => {
                match &expr.kind {
                    ExprKind::Call { callee, args, .. }
                        if self.profile.name == "go"
                            && matches!(&callee.kind, ExprKind::Ident(name) if name == "__go_named_type")
                            && args.len() == 2 =>
                    {
                        if let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind {
                            let type_name = match &args[1].value.kind {
                                ExprKind::Lit(Literal::Str(type_name)) => Some(type_name.clone()),
                                ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
                                _ => None,
                            };
                            if let Some(type_name) = type_name {
                                self.source_type_aliases.insert(self.canon(name), type_name);
                            }
                        }
                        return Ok(());
                    }
                    ExprKind::Call { callee, args, .. }
                        if self.profile.name == "c"
                            && matches!(&callee.kind, ExprKind::Ident(name) if name == "exit") =>
                    {
                        if let Some(first) = args.first() {
                            self.compile_expr(&first.value)?;
                        } else {
                            self.emit(Op::NULL);
                        }
                        self.emit_return_through_finally(1)?;
                        return Ok(());
                    }
                    ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_lset_stmt")
                            && args.len() == 2 =>
                    {
                        return self.compile_vb_fixed_string_stmt(
                            &args[0].value,
                            &args[1].value,
                            false,
                        );
                    }
                    ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_rset_stmt")
                            && args.len() == 2 =>
                    {
                        return self.compile_vb_fixed_string_stmt(
                            &args[0].value,
                            &args[1].value,
                            true,
                        );
                    }
                    ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_mid_stmt")
                            && args.len() == 4 =>
                    {
                        return self.compile_vb_mid_stmt(
                            &args[0].value,
                            &args[1].value,
                            &args[2].value,
                            &args[3].value,
                        );
                    }
                    ExprKind::Call { callee, args, .. } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_err_raise") =>
                    {
                        return self.compile_vb_err_raise_stmt(args);
                    }
                    ExprKind::Ident(name)
                        if self.is_php_profile()
                            && (name.eq_ignore_ascii_case("exit")
                                || name.eq_ignore_ascii_case("die")) =>
                    {
                        self.emit(Op::NULL);
                        self.emit_return_through_finally(1)?;
                        return Ok(());
                    }
                    // Bare identifier that's a known function → call with 0 args
                    ExprKind::Ident(name) if self.defined_functions.contains(name.as_str()) => {
                        let saved_js_this = self.save_js_this("__js_stmt_prev_this");
                        if self.profile.ambient_this_binding {
                            let line = self.line;
                            common::expressions::emit_undefined(self.chunk(), line);
                            self.set_js_this_from_stack();
                        }
                        self.emit_var_get(name);
                        self.emit_u8(Op::CALL_REF, 0);
                        if saved_js_this.is_some() {
                            let result_slot = self.define_local("__js_stmt_result");
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.restore_js_this(saved_js_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        }
                        self.emit(Op::DROP);
                    }
                    // JS bare member statements evaluate the property access
                    // and discard the result; they are not implicit calls.
                    ExprKind::Member { object, field, .. } => {
                        if self.profile.dynamic_member_access {
                            self.compile_expr(expr)?;
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                        self.compile_expr(object)?;
                        let field_name = self.canon(field);
                        let prop = self.str_const(&field_name);
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.define_local("__fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);
                        let obj_tmp = self.define_local("__obj");
                        self.reserve_local_slot(obj_tmp);
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DROP);
                    }
                    _ => {
                        self.compile_expr(expr)?;
                        self.emit(Op::DROP);
                    }
                }
            }

            // ── Block ───────────────────────────────────────────────────
            StmtKind::Block(stmts) => {
                let all_decls = stmts.iter().all(|s| {
                    matches!(
                        s.kind,
                        StmtKind::VarDecl { .. }
                            | StmtKind::FunctionDecl { .. }
                            | StmtKind::ClassDecl { .. }
                            | StmtKind::EnumDecl { .. }
                    )
                });
                let hoisted_deconstruction = is_hoisted_deconstruction_block(stmts);
                // A block that declares a lexical binding (`let`/`const`/
                // `class`) is its own scope even when it contains *only*
                // declarations — otherwise `{ let x = 42; }` would leak `x` to
                // the enclosing scope. (`var` is function-scoped and correctly
                // skips this.) Driven by the profile capability, not a language
                // name.
                let has_lexical = self.profile.lexical_block_scope
                    && stmts.iter().any(|s| {
                        matches!(
                            &s.kind,
                            StmtKind::VarDecl {
                                kind: VarDeclKind::Let | VarDeclKind::Const,
                                ..
                            } | StmtKind::ClassDecl { .. }
                        )
                    });
                let make_scope = (!all_decls && !hoisted_deconstruction) || has_lexical;
                if make_scope {
                    self.scope_mut().begin_scope();
                }
                let saved_strict = self.in_strict;
                if self.profile.ecma_strict_mode && Self::stmts_have_use_strict_directive(stmts) {
                    self.in_strict = true;
                }
                for s in stmts {
                    self.compile_stmt(s)?;
                }
                self.in_strict = saved_strict;
                if make_scope {
                    self.scope_mut().end_scope();
                }
            }

            // ── Variable declarations ───────────────────────────────────
            StmtKind::VarDecl { declarations, kind } => {
                for decl in declarations {
                    self.compile_var_declarator(decl, kind)?;
                }
            }

            // ── Assignment ──────────────────────────────────────────────
            StmtKind::Assign { targets, value } => {
                if self.profile.name == "fortran" {
                    if let [target] = targets.as_slice() {
                        let is_whole_array_target =
                            matches!(target.kind, ExprKind::Ident(_) | ExprKind::Member { .. })
                                && self.expr_is_array_like(target);
                        if is_whole_array_target {
                            let line = self.line;
                            let value_slot = self.define_local("__fortran_array_fill_value");
                            self.compile_expr(value)?;
                            self.emit_u16(Op::LOCAL_SET, value_slot);

                            if !self.expr_is_array_like(value) {
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                fn_call!(self, "ecma:array", "isArray", 1);
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                self.chunk().emit_if(line);

                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                self.emit_array_clone_from_stack();
                                self.compile_assign_target(target)?;

                                self.chunk().emit_else(line);

                                self.compile_expr(target)?;
                                let array_slot = self.define_local("__fortran_array_fill_target");
                                self.emit_u16(Op::LOCAL_SET, array_slot);

                                self.emit_u16(Op::LOCAL_GET, array_slot);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                self.emit_const(Value::I32(0));
                                self.emit_const(Value::I32(i32::MAX));
                                common::collections::emit_fill(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.compile_assign_target(target)?;

                                self.chunk().emit_end(line);
                                return Ok(());
                            }

                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_array_clone_from_stack();
                            self.compile_assign_target(target)?;
                            return Ok(());
                        }
                    }
                }
                if targets.len() == 1 {
                    if let ExprKind::Ident(name) = &targets[0].kind {
                        let binding_key = self.canon(name);
                        if let Some(binding) = self.resolve_reflection_binding_expr(value) {
                            self.reflection_bindings.insert(binding_key, binding);
                        } else {
                            self.reflection_bindings.remove(&binding_key);
                        }
                    }
                }
                if self.profile.namespaces.use_dotnet && targets.len() == 1 {
                    if let ExprKind::Binary { op, left, right } = &value.kind {
                        if self.assign_target_matches_expr(&targets[0], left)
                            && self.is_csharp_delegate_handler_expr(right)
                        {
                            match op {
                                BinOp::Add => {
                                    self.compile_expr(left)?;
                                    self.compile_expr(right)?;
                                    common::delegates::emit_combine(
                                        &mut self.chunks,
                                        self.current,
                                        self.line,
                                    );
                                    self.compile_assign_target(&targets[0])?;
                                    return Ok(());
                                }
                                BinOp::Sub => {
                                    self.compile_expr(left)?;
                                    self.compile_expr(right)?;
                                    common::delegates::emit_remove(
                                        &mut self.chunks,
                                        self.current,
                                        self.line,
                                    );
                                    self.compile_assign_target(&targets[0])?;
                                    return Ok(());
                                }
                                _ => {}
                            }
                        }
                    }
                }
                // Multi-value receive: `a, b, c = callee(...)` where the
                // callee is a direct identifier call to a function the
                // pre-scan marked multi-return with matching arity. We
                // skip the heap-tuple alloc: compile the call, then let
                // each destructured element LOCAL_SET off the stack.
                if let Some((_arity, idents)) = self.detect_multi_value_receive(targets, value) {
                    // Compile the call inline so the Call-expression path
                    // in `expressions.rs` does NOT re-pack the results —
                    // we want the raw N values on the stack for direct
                    // destructuring.
                    self.compile_call_raw(value)?;
                    // Stack now holds [v0, v1, …, v(N-1)] with v(N-1) at
                    // TOS. Reverse assignment maps v_i to the i-th target.
                    // Inside a function, a fresh ident that doesn't already
                    // resolve should become a new local — C#'s
                    // `var (a, b) = f();` introduces new names, and this
                    // lets the walker emit a single Assign statement
                    // without juggling a Block + VarDecl pair.
                    let in_function = self.scopes.len() > 1;
                    for name in idents.iter().rev() {
                        if in_function
                            && self.scope().resolve(name).is_none()
                            && (self.case_sensitive || self.scope().resolve_ci(name).is_none())
                        {
                            self.define_local(name);
                        }
                        self.emit_var_set(name);
                    }
                    return Ok(());
                } else {
                    let prefer_numeric_add = matches!(targets.as_slice(), [target] if self.expr_prefers_numeric_add(target));
                    self.compile_expr_with_numeric_add_hint(value, prefer_numeric_add)?;
                    if let [target] = targets.as_slice() {
                        self.emit_assignment_type_coercion_for_target(target);
                    }
                    if let [target] = targets.as_slice() {
                        if let ExprKind::Ident(name) = &target.kind {
                            if let Some(type_hint) = self.lookup_var_type_hint(name) {
                                if let Some(target_len) = Self::vb_fixed_string_len(type_hint) {
                                    self.emit_vb_fixed_string_adjust_from_stack(target_len, false);
                                }
                            }
                        }
                    }
                    if let [target] = targets.as_slice() {
                        if let ExprKind::Ident(name) = &target.kind {
                            let type_hint = self.lookup_var_type_hint(name).map(str::to_string);
                            self.maybe_promote_pascal_array_literal_to_set(
                                type_hint.as_deref(),
                                value,
                            );
                        }
                    }
                    // PHP reference assignment: `$b = &$a` — the first
                    // assignment stores the cell itself (GLOBAL_SET/LOCAL_SET),
                    // then mark `$b` as pointer-cell so SUBSEQUENT writes
                    // go through cell_store.
                    let is_ref_assign = self.is_php_profile()
                        && matches!(
                            &value.kind,
                            ExprKind::Unary {
                                op: UnaryOp::AddrOf,
                                ..
                            }
                        );
                    for (i, target) in targets.iter().enumerate() {
                        if i < targets.len() - 1 {
                            inst!(self, core_wasm::dup);
                        }
                        self.compile_assign_target(target)?;
                    }
                    // Mark targets as pointer-cell AFTER the first
                    // assignment so the initial store uses GLOBAL/LOCAL_SET
                    // (writes the cell itself), and subsequent writes
                    // use cell_store (writes through the cell).
                    if is_ref_assign {
                        for target in targets.iter() {
                            if let ExprKind::Ident(name) = &target.kind {
                                self.mark_pointer_cell_binding(name);
                            }
                        }
                    }
                }
            }

            StmtKind::CompoundAssign { target, op, value } => {
                if self.profile.namespaces.use_dotnet
                    && matches!(op, CompoundOp::Add | CompoundOp::Sub)
                    && self.is_csharp_delegate_handler_expr(value)
                {
                    match op {
                        CompoundOp::Add => {
                            self.compile_expr(target)?;
                            self.compile_expr(value)?;
                            common::delegates::emit_combine(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            self.compile_assign_target(target)?;
                            return Ok(());
                        }
                        CompoundOp::Sub => {
                            self.compile_expr(target)?;
                            self.compile_expr(value)?;
                            common::delegates::emit_remove(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            self.compile_assign_target(target)?;
                            return Ok(());
                        }
                        _ => {}
                    }
                }
                if matches!(op, CompoundOp::NullCoalesce) {
                    self.compile_expr(target)?;
                    let current_slot = self.define_local("__null_coalesce_current");
                    self.emit_u16(Op::LOCAL_SET, current_slot);

                    self.emit_u16(Op::LOCAL_GET, current_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.compile_expr(value)?;
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, current_slot);
                    self.chunk().emit_end(line);
                    self.compile_assign_target(target)?;
                    return Ok(());
                }
                // Dynamic-typed languages: desugar `t OP= v` → `t = t OP v`
                // and reuse the full type-aware binary routing so compound
                // assignment dispatches BigInt/number/string identically to
                // the plain operator (e.g. `exp >>= 1n` hits the bigint path).
                if self.profile.dynamic_numeric_dispatch {
                    if let Some(binop) = compound_op_to_binop(op) {
                        let binexpr = Expression::new(ExprKind::Binary {
                            op: binop,
                            left: Box::new(target.clone()),
                            right: Box::new(value.clone()),
                        });
                        self.compile_expr(&binexpr)?;
                        self.compile_assign_target(target)?;
                        return Ok(());
                    }
                }
                // Load current value
                self.compile_expr(target)?;
                let prefer_numeric_add =
                    matches!(op, CompoundOp::Add) && self.expr_prefers_numeric_add(target);
                self.compile_expr_with_numeric_add_hint(value, prefer_numeric_add)?;
                if prefer_numeric_add {
                    self.emit(Op::F64_ADD);
                } else {
                    self.compile_compound_op(op);
                }
                self.compile_assign_target(target)?;
            }

            // ── If / Elif / Else (structured CF with label tracking) ──
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                let line = self.line;
                self.compile_expr(cond)?;
                self.emit_condition_truthiness_from_stack();
                self.chunk().emit_if(line);
                self.label_depth += 1;

                self.scope_mut().begin_scope();
                for s in then_body {
                    self.compile_stmt(s)?;
                }
                self.scope_mut().end_scope();

                if !elifs.is_empty() || else_body.is_some() {
                    let line = self.line;
                    self.chunk().emit_else(line);
                    if let Some((elif_cond, elif_body)) = elifs.first() {
                        let nested = Statement::new(StmtKind::If {
                            cond: elif_cond.clone(),
                            then_body: elif_body.clone(),
                            elifs: elifs.iter().skip(1).cloned().collect(),
                            else_body: else_body.clone(),
                        });
                        self.compile_stmt(&nested)?;
                    } else if let Some(else_stmts) = else_body {
                        self.scope_mut().begin_scope();
                        for s in else_stmts {
                            self.compile_stmt(s)?;
                        }
                        self.scope_mut().end_scope();
                    }
                }

                let line = self.line;
                self.chunk().emit_end(line);
                self.label_depth -= 1;
            }

            // ── While (compiler_common::loops) ─────────────────────────
            StmtKind::While {
                cond,
                body,
                else_body,
            } => {
                let line = self.line;
                let lp = common::loops::emit_loop_start(&mut self.chunks, self.current, line);
                // block + loop = 2 label stack entries
                let break_depth = self.label_depth + 1; // block is first (break target)
                let continue_depth = self.label_depth + 2; // loop is second (continue target)
                self.label_depth += 2;
                self.loop_states.push(lp);
                self.loops.push(LoopCtx {
                    label: self.pending_label.take(),
                    break_label_depth: break_depth,
                    continue_label_depth: continue_depth,
                    did_break_slot: None,
                    iterator_close_slot: None,
                    is_continuable: true,
                    finally_depth: self.active_finally_blocks.len(),
                });
                self.compile_expr(cond)?;
                self.emit_condition_truthiness_from_stack();
                let line = self.line;
                common::loops::emit_loop_cond(&mut self.chunks, self.current, line);
                for s in body {
                    self.compile_stmt(s)?;
                }
                self.loops.pop();
                let lp = self.loop_states.pop().unwrap();
                let line = self.line;
                common::loops::emit_loop_end(&mut self.chunks, self.current, lp, line);
                self.label_depth -= 2; // block + loop closed
                if let Some(else_stmts) = else_body {
                    for s in else_stmts {
                        self.compile_stmt(s)?;
                    }
                }
            }

            // ── For C-style (compiler_common::loops) ────────────────────
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                self.scope_mut().begin_scope();
                if let Some(init_stmt) = init {
                    self.compile_stmt(init_stmt)?;
                }
                let loop_capture_name = if self.profile.for_loop_per_iteration_binding {
                    init.as_ref().and_then(|stmt| match &stmt.kind {
                        StmtKind::VarDecl { declarations, .. } if declarations.len() == 1 => {
                            match &declarations[0].pattern {
                                BindingPattern::Ident(name) => Some(self.canon(name)),
                                _ => None,
                            }
                        }
                        _ => None,
                    })
                } else {
                    None
                };
                let line = self.line;
                // For C-style with update: use block { loop { cond, block $body { body }, update, br loop } }
                let block_patch = self.chunk().emit_block(line);
                self.label_depth += 1; // block
                let (loop_patch, _) = self.chunk().emit_loop_s(line);
                self.label_depth += 1; // loop
                let break_depth = self.label_depth - 1; // the block
                if let Some(c) = cond {
                    self.compile_expr(c)?;
                    self.emit_condition_truthiness_from_stack();
                } else {
                    inst!(self, core_wasm::bool_const, true);
                }
                let line = self.line;
                common::loops::emit_loop_cond(&mut self.chunks, self.current, line);
                // Body block for continue-to-update
                let body_block = if update.is_some() {
                    let bp = self.chunk().emit_block(line);
                    self.label_depth += 1;
                    Some(bp)
                } else {
                    None
                };
                let continue_depth = self.label_depth; // innermost = continue target (body block or loop)
                let lp = common::loops::LoopState {
                    block_patch,
                    loop_patch,
                    body_block_patch: body_block,
                };
                self.loop_states.push(lp);
                self.loops.push(LoopCtx {
                    label: self.pending_label.take(),
                    break_label_depth: break_depth,
                    continue_label_depth: continue_depth,
                    did_break_slot: None,
                    iterator_close_slot: None,
                    is_continuable: true,
                    finally_depth: self.active_finally_blocks.len(),
                });
                if let Some(loop_capture_name) = loop_capture_name.clone() {
                    self.capture_by_value_vars.push(loop_capture_name);
                }
                for s in body {
                    self.compile_stmt(s)?;
                }
                if loop_capture_name.is_some() {
                    self.capture_by_value_vars.pop();
                }
                self.loops.pop();
                let lp = self.loop_states.pop().unwrap();
                // Close body block (continue lands here)
                if let Some(bp) = lp.body_block_patch {
                    self.chunk().emit_end(line);
                    self.chunk().patch_block(bp);
                    self.label_depth -= 1;
                }
                if let Some(u) = update {
                    self.compile_expr(u)?;
                    self.emit(Op::DROP);
                }
                let line = self.line;
                self.chunk().emit_br(0, line); // br loop
                self.chunk().emit_end(line); // end loop
                self.chunk().patch_loop(lp.loop_patch);
                self.label_depth -= 1;
                self.chunk().emit_end(line); // end block
                self.chunk().patch_block(lp.block_patch);
                self.label_depth -= 1;
                self.scope_mut().end_scope();
            }

            // ── ForIn / ForOf ───────────────────────────────────────────
            StmtKind::ForIn {
                var,
                key,
                iter,
                body,
                else_body,
                of,
                is_async,
                ..
            } => {
                // Specialisation: if `iter` is a direct call to a
                // function the pre-pass tagged as a true generator,
                // emit a `GEN_NEXT`-driven loop rather than the
                // array-index loop. This is the only path that makes
                // `for v in @generator_fn()` iterate lazily via the
                // WASM stack-switching coroutine machinery.
                if self.is_direct_generator_call(iter) {
                    self.compile_generator_for_in(
                        var,
                        key.as_deref(),
                        iter,
                        body,
                        else_body.as_deref(),
                    )?;
                } else {
                    let line = self.line;
                    self.compile_expr(iter)?;
                    let iter_slot = self.define_local("__forin_iter");
                    self.emit_u16(Op::LOCAL_SET, iter_slot);

                    // `for await`: resolve [Symbol.asyncIterator] up front
                    // (§7.4.2 GetIterator ASYNC) — an async-generator method
                    // returns a generator continuation, which the runtime-
                    // generator gate below then drives lazily.
                    if *is_async && *of && key.is_none() && self.profile.async_wraps_body_in_try {
                        common::generators::emit_resolve_async_iterator(
                            &mut self.chunks,
                            self.current,
                            iter_slot,
                            line,
                        );
                    }

                    let runtime_generator_done = if *of && key.is_none() {
                        // Large PHP foreach bodies can exceed the i16 reach of
                        // flat BR/BR_IF patching when we emit the runtime
                        // generator fast-path inline. Use structured label
                        // branches here so skipping the generator path does not
                        // depend on relative byte offsets.
                        let done_block = self.chunk().emit_block(line);
                        self.label_depth += 1;

                        let normal_path_gate = self.chunk().emit_block(line);
                        self.label_depth += 1;

                        self.emit_u16(Op::LOCAL_GET, iter_slot);
                        let is_gen_idx = self.import("ecma:value", "isGenerator");
                        self.emit_host_call(is_gen_idx, 1);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                        };
                        self.chunk().emit_br_if(0, line);

                        self.compile_generator_for_in_cont(
                            var,
                            key.as_deref(),
                            iter_slot,
                            body,
                            else_body.as_deref(),
                        )?;
                        self.chunk().emit_br(1, line);

                        self.chunk().emit_end(line);
                        self.chunk().patch_block(normal_path_gate);
                        self.label_depth -= 1;
                        Some(done_block)
                    } else {
                        None
                    };

                    // Gate 2: custom iterable with bytecode [Symbol.iterator].
                    // Uses lazy next() loop so break/return() work on infinite
                    // iterators. Only for for-of (not spread/destructuring).
                    if self.profile.ecma_iterator_result_shape
                        && *of
                        && key.is_none()
                        && runtime_generator_done.is_some()
                    {
                        let line = self.line;
                        let custom_iter_gate = self.chunk().emit_block(line);
                        self.label_depth += 1;

                        self.emit_u16(Op::LOCAL_GET, iter_slot);
                        let iterator_key = self.str_const("iterator");
                        self.emit_u16(Op::STRUCT_GET, iterator_key);
                        // IsCallable(iter.iterator) via the shared reflection
                        // substrate — replaces the retired VM-internal
                        // REF_IS_FUNC opcode.
                        common::reflection::emit_is_callable(&mut self.chunks, self.current, line);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                        };
                        self.chunk().emit_br_if(0, line);

                        self.compile_for_of_custom_iterator_lazy(
                            iter_slot,
                            &var.clone(),
                            body,
                            else_body.as_deref(),
                        )?;
                        self.chunk().emit_br(1, line);

                        self.chunk().emit_end(line);
                        self.chunk().patch_block(custom_iter_gate);
                        self.label_depth -= 1;
                    }

                    // Python/JS-style `for x in obj` yields KEYS for dict-like
                    // objects (Map/Ordinary) and values for sequences — one
                    // shared, type-dispatched primitive. Skips the generic
                    // iterForOf+values path below.
                    let natural_object_iter =
                        self.profile.for_in_object_yields_keys && *of && key.is_none();
                    if natural_object_iter {
                        self.emit_u16(Op::LOCAL_GET, iter_slot);
                        common::collections::emit_iter_natural(
                            &mut self.chunks,
                            self.current,
                            self.line,
                        );
                        self.emit_u16(Op::LOCAL_SET, iter_slot);
                    }

                    // Materialize iterable → array via common emitter.
                    // All languages use iterForOf which handles Array, Map,
                    // Set, String, and custom iterables uniformly.
                    if *of && key.is_none() && !natural_object_iter {
                        self.emit_u16(Op::LOCAL_GET, iter_slot);
                        common::collections::emit_iter_for_of(
                            &mut self.chunks,
                            self.current,
                            self.line,
                        );
                        self.emit_u16(Op::LOCAL_SET, iter_slot);
                    }

                    self.emit_u16(Op::LOCAL_GET, iter_slot);

                    let iter_type_hint = match &iter.kind {
                        ExprKind::Ident(name) => {
                            self.lookup_var_type_hint(name).map(str::to_string)
                        }
                        _ => self.infer_expr_type_hint(iter),
                    };

                    let iterates_dictionary_entries = key.is_none()
                        && *of
                        && iter_type_hint
                            .as_deref()
                            .map(Self::is_dictionary_type_hint)
                            .unwrap_or(false);
                    let iterates_sorted_dictionary_entries = key.is_none()
                        && *of
                        && iter_type_hint
                            .as_deref()
                            .map(Self::is_sorted_dictionary_type_hint)
                            .unwrap_or(false);
                    let iterates_sorted_set_values = key.is_none()
                        && *of
                        && iter_type_hint
                            .as_deref()
                            .map(Self::is_sorted_set_type_hint)
                            .unwrap_or(false);

                    // Pick the polymorphic iteration primitive. All three
                    // dispatch on Array / Map / Ordinary uniformly so PHP
                    // assoc arrays, Python dicts, JS objects, Ruby hashes
                    // iterate correctly without per-language code.
                    //
                    //   for v in X       → values(X)        (Python for)
                    //   for k => v in X  → entries(X)       (PHP foreach, Ruby each_pair, JS for..of of Map/entries)
                    //   for k in X       → keys(X)          (JS for..in, Python dict iter-keys)
                    if key.is_some() || iterates_dictionary_entries {
                        common::collections::emit_iter_entries(
                            &mut self.chunks,
                            self.current,
                            line,
                        );
                    } else if *of {
                        common::collections::emit_iter_values(&mut self.chunks, self.current, line);
                    } else {
                        common::collections::emit_iter_keys(&mut self.chunks, self.current, line);
                    }

                    if iterates_sorted_dictionary_entries {
                        self.emit_common("dotnet.sorted_dictionary_entries", 1, line);
                    } else if iterates_sorted_set_values {
                        common::collections::emit_sorted(&mut self.chunks, self.current, line);
                    }

                    let arr_slot = self.define_local("__forin_arr");
                    self.emit_u16(Op::LOCAL_SET, arr_slot);
                    let idx_slot = self.define_local("__forin_idx");
                    // Allocate did_break slot BEFORE the for-in scaffolding
                    // so the assign-to-false initializer doesn't sit inside
                    // any of the for's blocks. Only when `else` is present
                    // — keeps the cost off the common case.
                    let did_break_slot = if else_body.is_some() {
                        let slot = self.define_local("__for_did_break");
                        inst!(self, core_wasm::bool_const, false);
                        self.emit_u16(Op::LOCAL_SET, slot);
                        Some(slot)
                    } else {
                        None
                    };
                    let lp = common::loops::emit_for_in_start(
                        &mut self.chunks,
                        self.current,
                        arr_slot,
                        idx_slot,
                        line,
                    );
                    // for_in_start emits: block + loop + cond + block $body = 3 labels
                    let break_depth = self.label_depth + 1; // outer block
                    let continue_depth = self.label_depth + 3; // body block (innermost)
                    self.label_depth += 3;

                    if let Some(k_name) = key {
                        // Entries path: TOS is a [k, v] pair. Destructure
                        // into key_var and var, then run body.
                        //
                        // Stack at loop body entry: [pair]
                        //   DUP; index 0 → key_var
                        //   index 1 → value_var
                        let pair_slot = self.define_local("__forin_pair");
                        self.emit_u16(Op::LOCAL_SET, pair_slot);
                        // key = pair[0]
                        self.emit_u16(Op::LOCAL_GET, pair_slot);
                        self.emit_const(Value::I32(0));
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        let key_slot = self.define_local(k_name);
                        self.emit_u16(Op::LOCAL_SET, key_slot);

                        // var = pair[1]
                        self.emit_u16(Op::LOCAL_GET, pair_slot);
                        self.emit_const(Value::I32(1));
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        let var_slot = self.define_local(var);
                        self.emit_u16(Op::LOCAL_SET, var_slot);
                    } else if iterates_dictionary_entries {
                        let var_slot = self.define_local(var);
                        self.emit_u16(Op::LOCAL_SET, var_slot);
                    } else {
                        // Values path: TOS is the value, bind directly.
                        // `for await (let v of …)` per ECMA-262 §13.7.5
                        // performs `Await(value)` between iterator-step
                        // and binding. Emit the WASM JSPI suspend op so
                        // promise values unwrap before the body runs;
                        // non-promises pass through unchanged.
                        if *is_async {
                            crate::primitives::functions::emit_await(self.chunk(), line);
                        }
                        let value_type_hint = iter_type_hint.as_deref().and_then(|type_hint| {
                            type_hint
                                .trim()
                                .trim_end_matches('?')
                                .trim()
                                .strip_suffix("()")
                                .map(str::to_string)
                        });
                        let var_slot = if let Some(type_hint) = value_type_hint {
                            self.define_local_typed(var, Some(type_hint))
                        } else {
                            self.define_local(var)
                        };
                        self.emit_u16(Op::LOCAL_SET, var_slot);
                    }

                    self.loop_states.push(lp);
                    self.loops.push(LoopCtx {
                        label: self.pending_label.take(),
                        break_label_depth: break_depth,
                        continue_label_depth: continue_depth,
                        did_break_slot,
                        iterator_close_slot: None,
                        is_continuable: true,
                        finally_depth: self.active_finally_blocks.len(),
                    });
                    for s in body {
                        self.compile_stmt(s)?;
                    }
                    self.loops.pop();
                    let lp = self.loop_states.pop().unwrap();
                    common::loops::emit_for_in_end(
                        &mut self.chunks,
                        self.current,
                        idx_slot,
                        lp,
                        line,
                    );
                    self.label_depth -= 3;
                    if let Some(else_stmts) = else_body {
                        // Python/Ruby for-else: skip else if any `break` fired.
                        // Wrap in `block { br_if 0 (if did_break); ...else... }`.
                        let dbs = did_break_slot
                            .expect("did_break_slot allocated when else_body present");
                        let skip = self.chunk().emit_block(line);
                        self.label_depth += 1;
                        self.emit_u16(Op::LOCAL_GET, dbs);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_br_if(0, line); // skip else if did_break
                        for s in else_stmts {
                            self.compile_stmt(s)?;
                        }
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(skip);
                        self.label_depth -= 1;
                    }

                    if let Some(done_block) = runtime_generator_done {
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(done_block);
                        self.label_depth -= 1;
                    }
                }
            }

            // ── DoWhile (compiler_common::loops) ────────────────────────
            StmtKind::DoWhile { body, cond, until } => {
                let line = self.line;
                let lp = common::loops::emit_do_loop_start(&mut self.chunks, self.current, line);
                let break_depth = self.label_depth + 1;
                let continue_depth = self.label_depth + 2;
                self.label_depth += 2;
                self.loop_states.push(lp);
                self.loops.push(LoopCtx {
                    label: self.pending_label.take(),
                    break_label_depth: break_depth,
                    continue_label_depth: continue_depth,
                    did_break_slot: None,
                    iterator_close_slot: None,
                    is_continuable: true,
                    finally_depth: self.active_finally_blocks.len(),
                });
                for s in body {
                    self.compile_stmt(s)?;
                }
                self.compile_expr(cond)?;
                self.emit_condition_truthiness_from_stack();
                self.loops.pop();
                let lp = self.loop_states.pop().unwrap();
                let line = self.line;
                common::loops::emit_do_loop_end(&mut self.chunks, self.current, lp, *until, line);
                self.label_depth -= 2;
            }

            // ── Switch / Select Case ────────────────────────────────────
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                // Save switch expression to a local so checks can read it
                // without leaving it on the stack during body execution.
                self.compile_expr(expr)?;
                let sw_slot = self.define_local("__sw_expr");
                self.emit_u16(Op::LOCAL_SET, sw_slot);

                // Switch uses a BLOCK for break — push onto loop stack so break can find it
                let line = self.line;
                let switch_block = self.chunk().emit_block(line);
                self.label_depth += 1;
                let switch_lp = common::loops::LoopState {
                    block_patch: switch_block,
                    loop_patch: 0,
                    body_block_patch: None,
                };
                self.loop_states.push(switch_lp);
                self.loops.push(LoopCtx {
                    label: self.pending_label.take(),
                    break_label_depth: self.label_depth,
                    continue_label_depth: self.label_depth,
                    did_break_slot: None,
                    iterator_close_slot: None,
                    is_continuable: false,
                    finally_depth: self.active_finally_blocks.len(),
                });

                // Merge legacy `default` field into the cases list.
                // New walkers emit default as a case with empty conditions
                // in source order. Old walkers may still use the separate
                // `default` field — append it at the end if present.
                let mut all_cases: Vec<&SwitchCase> = cases.iter().collect();
                let default_case_storage;
                if let Some(def) = default {
                    if !def.is_empty() && !cases.iter().any(|c| c.conditions.is_empty()) {
                        default_case_storage = SwitchCase {
                            conditions: vec![],
                            body: def.clone(),
                        };
                        all_cases.push(&default_case_storage);
                    }
                }

                let dispatch_slot = self.define_local("__switch_dispatch");
                self.emit_const(Value::F64(-1.0));
                self.emit_u16(Op::LOCAL_SET, dispatch_slot);

                for (i, case) in all_cases.iter().enumerate() {
                    if case.conditions.is_empty() {
                        continue;
                    }
                    let case_match_slot = self.define_local("__switch_case_match");
                    self.emit_const(Value::Bool(false));
                    self.emit_u16(Op::LOCAL_SET, case_match_slot);
                    for cond in &case.conditions {
                        match cond {
                            CaseCondition::Value(val) => {
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(val)?;
                                // JS switch uses === (strict equality, no type coercion per §14.12.1).
                                // Other languages use regular equality.
                                if self.profile.ecma_switch_strict_equality {
                                    self.compile_binop(&BinOp::StrictEq);
                                } else {
                                    {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                                    };
                                }
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                };
                                self.chunk().emit_if(line);
                                self.emit_const(Value::Bool(true));
                                self.emit_u16(Op::LOCAL_SET, case_match_slot);
                                self.chunk().emit_end(line);
                            }
                            CaseCondition::Range { from, to } => {
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(from)?;
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                                };
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                };
                                self.chunk().emit_if(line);
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(to)?;
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_le(self.chunk(), line);
                                };
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                };
                                self.chunk().emit_if(line);
                                self.emit_const(Value::Bool(true));
                                self.emit_u16(Op::LOCAL_SET, case_match_slot);
                                self.chunk().emit_end(line);
                                self.chunk().emit_end(line);
                            }
                            CaseCondition::Comparison { op, expr: cmp_expr } => {
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(cmp_expr)?;
                                match op {
                                    ComparisonOp::Eq => {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                                    }
                                    ComparisonOp::NotEq => {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                                    }
                                    ComparisonOp::Lt => {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                                    }
                                    ComparisonOp::LtEq => {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_le(self.chunk(), line);
                                    }
                                    ComparisonOp::Gt => {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                                    }
                                    ComparisonOp::GtEq => {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                                    }
                                }
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                };
                                self.chunk().emit_if(line);
                                self.emit_const(Value::Bool(true));
                                self.emit_u16(Op::LOCAL_SET, case_match_slot);
                                self.chunk().emit_end(line);
                            }
                        }
                    }

                    self.emit_u16(Op::LOCAL_GET, dispatch_slot);
                    self.emit_const(Value::F64(0.0));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, case_match_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_const(Value::F64(i as f64));
                    self.emit_u16(Op::LOCAL_SET, dispatch_slot);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                }

                let default_idx = all_cases.iter().position(|c| c.conditions.is_empty());
                if let Some(default_idx) = default_idx {
                    self.emit_u16(Op::LOCAL_GET, dispatch_slot);
                    self.emit_const(Value::F64(0.0));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_const(Value::F64(default_idx as f64));
                    self.emit_u16(Op::LOCAL_SET, dispatch_slot);
                    self.chunk().emit_end(line);
                }

                for (i, case) in all_cases.iter().enumerate() {
                    self.emit_u16(Op::LOCAL_GET, dispatch_slot);
                    self.emit_const(Value::F64(i as f64));
                    if self.profile.switch_fallthrough {
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_le(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if_value(line);
                        self.emit_u16(Op::LOCAL_GET, dispatch_slot);
                        self.emit_const(Value::F64(0.0));
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        self.chunk().emit_else(line);
                        self.emit_const(Value::Bool(false));
                        self.chunk().emit_end(line);
                    } else {
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                        };
                    }
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.label_depth += 1;
                    for s in &case.body {
                        self.compile_stmt(s)?;
                    }
                    self.chunk().emit_end(line);
                    self.label_depth -= 1;
                }
                self.loops.pop();
                let switch_lp = self.loop_states.pop().unwrap();
                let line = self.line;
                self.chunk().emit_end(line);
                self.chunk().patch_block(switch_lp.block_patch);
                self.label_depth -= 1;
            }

            // ── Try / Catch / Finally ───────────────────────────────────
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                let line = self.line;
                let finally_exc_slot = if catches.is_empty() && finally.is_some() {
                    let slot = self.define_local("__try_finally_exc");
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, slot);
                    Some(slot)
                } else {
                    None
                };
                // For a try-WITH-finally, allocate the completion state so a
                // `break`/`continue`/`return` inside the body can run `finally`
                // OUTSIDE the handler (see `finally_joins`): it stores a code
                // here and `br`s to the join instead of inlining `finally`
                // under the `try_table`. Default NORMAL (fall through).
                let completion = if finally.is_some() {
                    let completion_slot = self.define_local("__try_completion");
                    let ret_slot = self.define_local("__try_ret");
                    self.emit_const(Value::F64(completion::NORMAL));
                    self.emit_u16(Op::LOCAL_SET, completion_slot);
                    Some((completion_slot, ret_slot))
                } else {
                    None
                };
                let after_try_block = self.chunk().emit_block(line);
                self.label_depth += 1;
                // Register the join NOW: its `br` target is `after_try_block`,
                // whose `end` lands exactly on the `finally` emission below,
                // outside every handler. `label_depth - join_label_depth` from
                // any point in the body is the `br` depth to reach it.
                if let Some((completion_slot, ret_slot)) = completion {
                    self.finally_joins.push(FinallyJoin {
                        join_label_depth: self.label_depth,
                        completion_slot,
                        ret_slot,
                    });
                }
                let catch_jump =
                    common::errors::emit_try_start(&mut self.chunks[self.current], line);
                // `try_table` is now a structural block (spec): it pushes its
                // own label at runtime, so the body sits one level deeper.
                // Scope this +1 to the body only — catch arms / else / finally
                // are emitted AFTER the `end` below and must see the original
                // depth.
                self.label_depth += 1;
                if let Some(fin) = finally.clone() {
                    self.active_finally_blocks
                        .push(FinallyAction::Statements(fin));
                }
                let saved_try_strict = self.in_strict;
                if self.profile.ecma_strict_mode && Self::stmts_have_use_strict_directive(body) {
                    self.in_strict = true;
                }
                for s in body {
                    self.compile_stmt(s)?;
                }
                self.in_strict = saved_try_strict;
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                self.label_depth -= 1;
                // Python else: runs if no exception
                if let Some(else_stmts) = else_body {
                    for s in else_stmts {
                        self.compile_stmt(s)?;
                    }
                }
                self.chunk().emit_br(0, line);
                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);
                // Entering this try's catch-arms section: ITS runtime handler
                // has fired, so ITS finally (sequenced after the arms) is the
                // one a `throw` inside an arm must inline — enclosing trys'
                // finallys still have LIVE handlers and must NOT be inlined
                // (the runtime runs them; inlining doubled the finally).
                let fired_finally = if finally.is_some() && !catches.is_empty() {
                    let idx = self.active_finally_blocks.len() - 1;
                    self.fired_finally_indices.push(idx);
                    true
                } else {
                    false
                };
                if catches.is_empty() {
                    if let Some(exc_slot) = finally_exc_slot {
                        self.emit_u16(Op::LOCAL_SET, exc_slot);
                    } else {
                        self.emit(Op::DROP);
                    }
                } else {
                    // Multi-catch dispatch: each arm tests the exception's
                    // canonical __exception_type field. If it matches one of
                    // the arm's types, run the body; otherwise fall through
                    // to the next arm. The exception object is on TOS at
                    // every step. A catch-all arm (empty types or "Exception")
                    // catches everything. After all arms, any unmatched
                    // exception is re-thrown.
                    let exc_slot = self.define_local("__caught_exception");
                    let handled_slot = self.define_local("__catch_handled");
                    self.emit_u16(Op::LOCAL_SET, exc_slot);
                    self.emit_const(Value::Bool(false));
                    self.emit_u16(Op::LOCAL_SET, handled_slot);
                    for c in catches {
                        // When the language models its exceptions as real
                        // classes (`throwable_is_root`), catch types must match
                        // the real class names in the `__type`/`__types` chain —
                        // do NOT canonicalize, which conflates `Error` and
                        // `Exception` (both map to "Exception") and would erase
                        // PHP's two distinct branches.
                        let types: Vec<&str> = c
                            .types
                            .iter()
                            .map(|t| {
                                if self.profile.throwable_is_root {
                                    t.trim()
                                } else {
                                    common::errors::canonical_exception_name(t)
                                }
                            })
                            .collect();
                        // `Throwable` is the universal root everywhere. In
                        // PHP/Java (`throwable_is_root`), `Exception` is only a
                        // branch — the `Error` branch is a sibling — so
                        // `catch (Exception)` matches via the `__types` chain
                        // below, not as a catch-all. In Python/.NET/Ruby,
                        // `Exception` is the root and catches everything.
                        let exception_catches_all = !self.profile.throwable_is_root;
                        let is_catch_all = types.is_empty()
                            || types.iter().any(|t| {
                                *t == "Throwable"
                                    || (exception_catches_all
                                        && (*t == "Exception" || *t == "BaseException"))
                            });

                        let arm_match_slot = self.define_local("__catch_arm_match");
                        self.emit_const(Value::Bool(is_catch_all));
                        self.emit_u16(Op::LOCAL_SET, arm_match_slot);

                        if !is_catch_all {
                            for ty in &types {
                                let mut expected_names = vec![(*ty).to_string()];
                                if !self.case_sensitive {
                                    let canon_ty = self.canon(ty);
                                    if canon_ty != *ty {
                                        expected_names.push(canon_ty);
                                    }
                                }

                                // Single identity test per candidate type:
                                // `REF_TEST` → `test_type` resolves type-registry
                                // subtype (the real class hierarchy — now that the
                                // host registers Error/Exception as distinct roots
                                // per §20.5, PHP's sibling model via its prelude,
                                // etc.), `__type`/`__types` chain, and prototype
                                // identity. One unified mechanism, no per-language
                                // branching, no stamp string-compares.
                                for expected in &expected_names {
                                    self.emit_u16(Op::LOCAL_GET, exc_slot);
                                    let line = self.line;
                                    let idx = self.str_const(expected);
                                    self.chunks[self.current].emit_op_u16(Op::REF_TEST, idx, line);
                                    {
                                        let line = self.line;
                                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                    };
                                    self.chunk().emit_if(line);
                                    inst!(self, core_wasm::bool_const, true);
                                    self.emit_u16(Op::LOCAL_SET, arm_match_slot);
                                    self.chunk().emit_end(line);
                                }
                            }
                        }

                        self.emit_u16(Op::LOCAL_GET, handled_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if_value(line);
                        self.emit_u16(Op::LOCAL_GET, arm_match_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_else(line);
                        self.emit_const(Value::Bool(false));
                        self.chunk().emit_end(line);
                        self.chunk().emit_if(line);
                        // The catch body executes inside this arm-match IF —
                        // a real WASM control frame the VM pushes onto its
                        // label_stack. `break`/`continue` inside the catch body
                        // derive their `br` depth from `label_depth`, so it must
                        // count this open IF or the branch targets the wrong
                        // frame and the enclosing loop never exits (hang).
                        // ECMA-262 §14.2: abrupt completion still exits the loop.
                        self.label_depth += 1;

                        if let Some(ref var) = c.var_name {
                            self.scope_mut().begin_scope();
                            let slot = self.define_local(var);
                            self.emit_u16(Op::LOCAL_GET, exc_slot);
                            self.emit_u16(Op::LOCAL_SET, slot);
                        } else {
                            self.scope_mut().begin_scope();
                        }

                        if let Some(cond) = &c.when_clause {
                            self.compile_expr(cond)?;
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            };
                            self.chunk().emit_if(line);
                            // The when-clause adds a second open IF around the
                            // catch body — count it too.
                            self.label_depth += 1;
                        }

                        self.catch_depth += 1;
                        for s in &c.body {
                            self.compile_stmt(s)?;
                        }
                        self.catch_depth = self.catch_depth.saturating_sub(1);
                        self.emit_const(Value::Bool(true));
                        self.emit_u16(Op::LOCAL_SET, handled_slot);
                        if c.when_clause.is_some() {
                            self.label_depth -= 1;
                            self.chunk().emit_end(line);
                        }
                        self.scope_mut().end_scope();
                        self.label_depth -= 1;
                        self.chunk().emit_end(line);
                    }
                    // Fallthrough = no arm matched. Re-throw (through finally if any).
                    self.emit_u16(Op::LOCAL_GET, handled_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, exc_slot);
                    self.emit_throw_through_finally()?;
                    self.chunk().emit_end(line);
                }
                self.chunk().emit_end(line);
                self.chunk().patch_block(after_try_block);
                self.label_depth -= 1;
                if fired_finally {
                    self.fired_finally_indices.pop();
                }
                if finally.is_some() {
                    self.active_finally_blocks.pop();
                }
                // Pop the join BEFORE emitting the `finally` body: a
                // `break`/`continue`/`return` inside `finally` itself belongs
                // to the ENCLOSING try/loop, not this one (whose finally is
                // now running).
                if finally.is_some() {
                    self.finally_joins.pop();
                }
                if let Some(fin) = finally {
                    for s in fin {
                        self.compile_stmt(s)?;
                    }
                }
                // Completion dispatch — runs AFTER `finally`, OUTSIDE the
                // handler. A non-local exit that jumped here re-issues itself,
                // now chaining to the enclosing join (or the real loop/return
                // target) since this try's join + finally are already popped.
                if let Some((completion_slot, ret_slot)) = completion {
                    let emit_eq_branch = |c: &mut Self, code: f64| {
                        c.emit_u16(Op::LOCAL_GET, completion_slot);
                        c.emit_const(Value::F64(code));
                        let ln = c.line;
                        crate::primitives::ops::emit_dyn_eq(c.chunk(), ln);
                        crate::primitives::ops::emit_dyn_to_bool(c.chunk(), ln);
                        c.chunk().emit_if(ln);
                        c.label_depth += 1;
                    };
                    emit_eq_branch(self, completion::BREAK);
                    self.emit_break_through_finally(None)?;
                    self.label_depth -= 1;
                    self.chunk().emit_end(line);
                    emit_eq_branch(self, completion::CONTINUE);
                    self.emit_continue_through_finally(None)?;
                    self.label_depth -= 1;
                    self.chunk().emit_end(line);
                    emit_eq_branch(self, completion::RETURN);
                    self.emit_u16(Op::LOCAL_GET, ret_slot);
                    self.emit_return_through_finally(1)?;
                    self.label_depth -= 1;
                    self.chunk().emit_end(line);
                }
                if let Some(exc_slot) = finally_exc_slot {
                    if self.catch_depth > 0 {
                        return Ok(());
                    }
                    self.emit_u16(Op::LOCAL_GET, exc_slot);
                    self.emit(Op::REF_IS_NULL);
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, exc_slot);
                    let line = self.line;
                    common::errors::emit_throw(self.chunk(), line);
                    self.chunk().emit_end(line);
                }
            }

            // ── Return ──────────────────────────────────────────────────
            StmtKind::Return(val) => {
                // Multi-value path: `return a, b, c` in a function the
                // pre-scan marked as multi-return. We push each element
                // separately (no heap tuple allocation) and let the VM's
                // `RETURN` pop N values per `chunk.result_arity`.
                let multi_n = self.current_multi_return_arity();
                if let (Some(n), Some(v)) = (multi_n, val) {
                    if let ExprKind::Tuple(elems) = &v.kind {
                        if elems.len() == n as usize {
                            for elem in elems {
                                self.compile_expr(elem)?;
                            }
                            self.emit_return_through_finally(n as usize)?;
                            return Ok(());
                        }
                    }
                }

                if let Some(v) = val {
                    self.compile_expr(v)?;
                    if let Some((ctx_chunk, this_slot)) = self.js_derived_ctor_ctx {
                        if ctx_chunk == self.current && self.profile.name == "js" {
                            let return_slot = self.define_local("__js_derived_return_value");
                            self.emit_u16(Op::LOCAL_SET, return_slot);
                            self.emit_u16(Op::LOCAL_GET, return_slot);
                            let line = self.line;
                            crate::primitives::instructions::recipes::is_object(self.chunk(), line);
                            self.chunk().emit_if(line);
                            self.emit_u16(Op::LOCAL_GET, return_slot);
                            self.chunk().emit_else(line);
                            crate::primitives::classes::emit_this_initialized_guard(
                                self.chunk(),
                                this_slot,
                                line,
                            );
                            self.emit_u16(Op::LOCAL_GET, this_slot);
                            self.chunk().emit_end(line);
                        }
                    }
                } else if let Some(rs) = self.current_result_slot {
                    // ResultSlot return: return the result slot value
                    self.emit_u16(Op::LOCAL_GET, rs);
                } else if self.current_chunk_is_js_async() {
                    inst!(self, core_wasm::undefined);
                } else {
                    self.emit(Op::NULL);
                }

                if self.current_chunk_is_js_async() {
                    let resolve_idx = self.import("ecma:promise", "resolve");
                    self.emit_host_call(resolve_idx, 1);
                }
                self.emit_return_through_finally(1)?;
            }

            // ── Break ───────────────────────────────────────────────────
            StmtKind::Break(target) => {
                match target {
                    // Exit Sub / Exit Function → RETURN (not a loop break)
                    BreakTarget::Kind(ExitKind::Sub) | BreakTarget::Kind(ExitKind::Function) => {
                        // Return with current result slot value, or null
                        if let Some(result_slot) = self.current_result_slot {
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else {
                            self.emit(Op::NULL);
                        }
                        self.emit_return_through_finally(1)?;
                    }
                    BreakTarget::Implicit | BreakTarget::Kind(_) | BreakTarget::Level(_) => {
                        // If the targeted loop has a did_break slot (Python/
                        // Ruby for-else), record that break fired so the
                        // post-loop else clause is skipped.
                        if let Some(slot) = self.loops.last().and_then(|c| c.did_break_slot) {
                            inst!(self, core_wasm::bool_const, true);
                            self.emit_u16(Op::LOCAL_SET, slot);
                        }
                        if let Some(iterator_slot) = self.iterator_close_slot_for_break(None) {
                            self.emit_js_iterator_close(iterator_slot);
                        }
                        self.emit_break_through_finally(None)?;
                    }
                    BreakTarget::Label(label) => {
                        if let Some(slot) = self
                            .loops
                            .iter()
                            .rev()
                            .find(|c| c.label.as_deref() == Some(label))
                            .and_then(|c| c.did_break_slot)
                        {
                            inst!(self, core_wasm::bool_const, true);
                            self.emit_u16(Op::LOCAL_SET, slot);
                        }
                        if let Some(iterator_slot) = self.iterator_close_slot_for_break(Some(label))
                        {
                            self.emit_js_iterator_close(iterator_slot);
                        }
                        self.emit_break_through_finally(Some(label))?;
                    }
                    BreakTarget::Value(expr) => {
                        self.compile_expr(expr)?;
                        self.emit_return_through_finally(1)?;
                    }
                }
            }

            // ── Continue ────────────────────────────────────────────────
            StmtKind::Continue(target) => match target {
                ContinueTarget::Implicit | ContinueTarget::Kind(_) | ContinueTarget::Level(_) => {
                    self.emit_continue_through_finally(None)?;
                }
                ContinueTarget::Label(label) => {
                    self.emit_continue_through_finally(Some(label))?;
                }
            },

            // ── Throw ───────────────────────────────────────────────────
            StmtKind::Throw { expr, cause: _ } => {
                if let Some(v) = expr {
                    self.compile_expr(v)?;
                } else {
                    self.emit(Op::NULL);
                }
                self.emit_active_js_iterator_closes();
                // Inside a catch arm, the VM exception handler is no longer active for
                // this try block, so we must inline the finally block before throwing.
                // In the try body, the VM routes exceptions to the catch handler first.
                if self.catch_depth > 0 {
                    self.emit_throw_through_finally()?;
                } else {
                    let line = self.line;
                    common::errors::emit_throw(self.chunk(), line);
                }
            }

            // ── Function declaration ────────────────────────────────────
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                body,
                modifiers: _,
                handles,
                is_async,
                is_generator,
                is_sub,
            } => {
                self.compile_function_decl(
                    name,
                    params,
                    return_type,
                    body,
                    *is_sub,
                    *is_generator,
                    handles,
                    *is_async,
                )?;
            }

            // ── Class declaration ───────────────────────────────────────
            StmtKind::ClassDecl {
                name,
                parents,
                interfaces,
                members,
                modifiers,
                ..
            } => {
                let cname = self.canon(name);
                self.defined_globals.insert(cname.clone());
                self.defined_classes.insert(cname.clone());
                let inferred_parents;
                let effective_parents: &[String] = if self.should_infer_winforms_form(name, parents)
                {
                    inferred_parents = vec!["Form".to_string()];
                    &inferred_parents
                } else {
                    parents
                };
                // Every language's profile has `uses_normalize_class = true`
                // after Phase 3. ClassDecl always goes through
                // walker → normalize_class → emit_class → compile_class.
                // If a new language is added that hasn't written its
                // normalizer yet, `emit_class_from_ast` returns an error
                // loudly rather than silently picking a legacy path.
                let span = stmt.span.clone();
                crate::primitives::class_normalize::emit::emit_class_from_ast(
                    self,
                    span,
                    &cname,
                    effective_parents,
                    interfaces,
                    members,
                    modifiers,
                    self.profile.name == "fortran",
                )?;
            }

            // ── Interface declaration ───────────────────────────────────
            StmtKind::InterfaceDecl { .. } => {
                // No-op — interfaces are type-level only
            }

            // ── Enum declaration ────────────────────────────────────────
            // Compiles to a namespace object: Color = { Red: 0, Green: 1, Blue: 2 }
            // Bare member references (e.g. Pascal `c := Green`) are resolved at
            // compile time via the enum_members map.
            StmtKind::EnumDecl {
                name,
                members,
                is_flags,
                backing_type: _,
                interfaces,
                body_members,
                ..
            } => {
                let cname = self.canon(name);
                if *is_flags {
                    self.enum_flags.insert(cname.clone());
                } else {
                    self.enum_flags.remove(&cname);
                }

                match self.profile.name.as_str() {
                    "dart" => {
                        self.compile_dart_enum_decl(
                            name,
                            interfaces,
                            body_members,
                            members,
                            stmt.span,
                        )?;
                        return Ok(());
                    }
                    _ => {}
                }

                let mut next_val = 0i64;
                let mut value_names = HashMap::new();
                for m in members {
                    if let Some(ref v) = m.value {
                        if let ExprKind::Lit(Literal::Int(n)) = &v.kind {
                            next_val = *n;
                        }
                    }
                    next_val += 1;
                    let mname = self.canon(&m.name);
                    // Register member → enum type for bare-name resolution
                    self.enum_members.insert(mname, cname.clone());
                    value_names.insert(next_val - 1, m.name.clone());
                }
                self.enum_value_names
                    .insert(cname.clone(), value_names.clone());
                if let Some((_, leaf)) = cname.rsplit_once('.') {
                    self.enum_value_names
                        .entry(leaf.to_string())
                        .or_insert(value_names);
                }
                self.compile_shared_enum_decl(name, interfaces, body_members, members, stmt.span)?;
                self.defined_globals.insert(cname);
            }

            // ── Struct declaration (same as class) ──────────────────────
            // Structs compile through the same pipeline as classes: no
            // parent, no interfaces (struct `interfaces` list is ignored
            // by legacy compile_class anyway), same normalize → emit
            // path. Treated as a parent-less class by the walker's
            // normalize_class for the active language.
            StmtKind::StructDecl { name, members, .. } => {
                let cn = self.canon(name);
                self.defined_globals.insert(cn.clone());
                self.defined_classes.insert(cn.clone());
                let span = stmt.span.clone();
                crate::primitives::class_normalize::emit::emit_class_from_ast(
                    self,
                    span,
                    &cn,
                    &[],
                    &[],
                    members,
                    &crate::ast::ClassModifiers::default(),
                    true,
                )?;
            }

            // ── Module declaration (VB) ─────────────────────────────────
            // Models WASM Component Model: members are exports of the module.
            // - Members compile as globals (so call_ref works)
            // - Bare member names register in enum_members map → resolve to Module.Member
            // - A namespace struct is built so qualified `Module.Member` works too
            StmtKind::ModuleDecl { name, members, .. } => {
                let module_name = self.canon(name);
                self.defined_classes.insert(module_name.clone());
                self.register_module_static_container(&module_name, members);
                let mut member_names: Vec<String> = Vec::new();

                // First pass: compile all members as globals + collect names
                for m in members {
                    match m {
                        ClassMember::Method(stmt) => {
                            if let StmtKind::FunctionDecl { name: mname, .. } = &stmt.kind {
                                let mn = self.canon(mname);
                                let saved_class = self.current_class.clone();
                                let saved_implicit_self = self.current_class_implicit_self;
                                let saved_member_static = self.current_member_is_static;
                                self.current_class = Some(module_name.clone());
                                self.current_class_implicit_self = false;
                                self.current_member_is_static = true;
                                self.compile_stmt(stmt)?;
                                self.current_class = saved_class;
                                self.current_class_implicit_self = saved_implicit_self;
                                self.current_member_is_static = saved_member_static;
                                member_names.push(mn);
                            }
                        }
                        ClassMember::Field {
                            name: fname, init, ..
                        } => {
                            if let Some(init_expr) = init {
                                self.compile_expr(init_expr)?;
                            } else {
                                self.emit(Op::NULL);
                            }
                            let cname = self.canon(fname);
                            let idx = self.str_const(&cname);
                            self.emit_u16(Op::GLOBAL_SET, idx);
                            self.defined_globals.insert(cname.clone());
                            member_names.push(cname);
                        }
                        ClassMember::Const {
                            name: cname, value, ..
                        } => {
                            // Compile value once, install as global
                            // `<Class>.<Const>` (legacy access path)
                            // AND stamp on the class object so PHP
                            // `Class::Const` static access (struct_get
                            // on class) resolves to the value.
                            self.compile_expr(value)?;
                            let val_slot = self.define_local("__class_const_val");
                            self.emit_u16(Op::LOCAL_SET, val_slot);

                            let cn = self.canon(cname);
                            let idx = self.str_const(&cn);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.emit_u16(Op::GLOBAL_SET, idx);
                            self.defined_globals.insert(cn.clone());
                            member_names.push(cn.clone());

                            // Stamp on class object for static access.
                            // `name` here is the enclosing class name; on
                            // module-level Const blocks it's the module
                            // name, but the class object lookup will
                            // miss harmlessly in that case.
                            let class_canon = self.canon(name);
                            if self.defined_globals.contains(&class_canon) {
                                let cg_idx = self.str_const(&class_canon);
                                self.emit_u16(Op::GLOBAL_GET, cg_idx);
                                self.emit_u16(Op::LOCAL_GET, val_slot);
                                let field_idx = self.str_const(cname);
                                self.emit_u16(Op::STRUCT_SET, field_idx);
                                self.emit(Op::DROP);
                            }
                        }
                        ClassMember::NestedType(stmt) => {
                            // Nested types get their own globals; attach them to the
                            // module object so `Module.Type.Member` resolves through the
                            // same shared namespace path used by classes.
                            if let Some(cn) = match &stmt.kind {
                                StmtKind::ClassDecl { name: cname, .. }
                                | StmtKind::StructDecl { name: cname, .. }
                                | StmtKind::EnumDecl { name: cname, .. }
                                | StmtKind::InterfaceDecl { name: cname, .. }
                                | StmtKind::ModuleDecl { name: cname, .. } => {
                                    Some(self.canon(cname))
                                }
                                _ => None,
                            } {
                                member_names.push(cn);
                            }
                            self.compile_stmt(stmt)?;
                        }
                        ClassMember::Constructor { params, body, .. } => {
                            // Module-level constructor — compile as a function named after constructor_name
                            let ctor_stmt = Statement::new(StmtKind::FunctionDecl {
                                name: self.profile.constructor_name.clone(),
                                params: params.clone(),
                                return_type: None,
                                body: body.clone(),
                                modifiers: Modifiers::default(),
                                handles: Vec::new(),
                                is_async: false,
                                is_generator: false,
                                is_sub: true,
                            });
                            let saved_class = self.current_class.clone();
                            let saved_implicit_self = self.current_class_implicit_self;
                            let saved_member_static = self.current_member_is_static;
                            self.current_class = Some(module_name.clone());
                            self.current_class_implicit_self = false;
                            self.current_member_is_static = true;
                            self.compile_stmt(&ctor_stmt)?;
                            self.current_class = saved_class;
                            self.current_class_implicit_self = saved_implicit_self;
                            self.current_member_is_static = saved_member_static;
                            member_names.push(self.canon(&self.profile.constructor_name));
                        }
                        _ => {}
                    }
                }

                if member_names
                    .iter()
                    .any(|mn| mn.eq_ignore_ascii_case("__static_init__"))
                {
                    let init_idx = self.str_const("__static_init__");
                    self.emit_u16(Op::GLOBAL_GET, init_idx);
                    self.emit_u8(Op::CALL_REF, 0);
                    self.emit(Op::DROP);
                }

                // Second pass: build namespace struct { member1: global, member2: global, ... }
                self.emit_u16(Op::STRUCT_NEW, 0);
                for mn in &member_names {
                    inst!(self, core_wasm::dup);
                    let gidx = self.str_const(mn);
                    self.emit_u16(Op::GLOBAL_GET, gidx);
                    let key = self.str_const(mn);
                    self.emit_u16(Op::STRUCT_SET, key);
                    self.emit(Op::DROP);
                    // Register bare member → module name for qualified resolution
                    self.enum_members.insert(mn.clone(), module_name.clone());
                }
                let mod_idx = self.str_const(&module_name);
                self.emit_u16(Op::GLOBAL_SET, mod_idx);
                self.defined_globals.insert(module_name);
            }

            // ── Namespace declaration ───────────────────────────────────
            // C#/VB namespace: container of types. Compiles members as top-level globals
            // (matches .NET behavior — within the same compilation unit, bare type access
            // works without import). Also builds namespace struct for qualified access.
            StmtKind::NamespaceDecl { name, body } => {
                let local_ns_name = self.canon(name).replace('\\', ".");
                let ns_name = match self.current_namespace.as_deref() {
                    Some(prefix) if !prefix.is_empty() => format!("{prefix}.{local_ns_name}"),
                    _ => local_ns_name,
                };
                if ns_name.is_empty() {
                    let prev_namespace = self.current_namespace.clone();
                    self.current_namespace = None;
                    for s in body {
                        self.compile_stmt(s)?;
                    }
                    self.current_namespace = prev_namespace;
                    return Ok(());
                }
                let mut member_names: Vec<(String, String, bool)> = Vec::new();
                let mut qualified_body: Vec<Statement> = Vec::with_capacity(body.len());
                for s in body {
                    let mut qualified = s.clone();
                    match &mut qualified.kind {
                        StmtKind::ClassDecl { name: cn, .. }
                        | StmtKind::StructDecl { name: cn, .. }
                        | StmtKind::EnumDecl { name: cn, .. }
                        | StmtKind::InterfaceDecl { name: cn, .. }
                        | StmtKind::ModuleDecl { name: cn, .. } => {
                            let member_name = self.canon(cn);
                            let qualified_name = if member_name.contains('.') {
                                member_name.clone()
                            } else {
                                format!("{ns_name}.{member_name}")
                            };
                            member_names.push((member_name, qualified_name.clone(), true));
                            *cn = qualified_name;
                        }
                        StmtKind::FunctionDecl { name: cn, .. } => {
                            let member_name = self.canon(cn);
                            let qualified_name = if member_name.contains('.') {
                                member_name.clone()
                            } else {
                                format!("{ns_name}.{member_name}")
                            };
                            member_names.push((member_name, qualified_name.clone(), false));
                            *cn = qualified_name;
                        }
                        _ => {}
                    }
                    qualified_body.push(qualified);
                }
                for (_, qualified_name, is_type_like) in &member_names {
                    self.defined_globals.insert(qualified_name.clone());
                    if *is_type_like {
                        self.defined_classes.insert(qualified_name.clone());
                    } else {
                        self.defined_functions.insert(qualified_name.clone());
                    }
                }
                let prev_namespace = self.current_namespace.clone();
                self.current_namespace = Some(ns_name.clone());
                for s in &qualified_body {
                    self.compile_stmt(s)?;
                }
                self.current_namespace = prev_namespace;

                for (member_name, qualified_name, is_type_like) in &member_names {
                    let qualified_idx = self.str_const(qualified_name);
                    self.defined_globals.insert(qualified_name.clone());
                    if *is_type_like {
                        self.defined_classes.insert(qualified_name.clone());
                    } else {
                        self.defined_functions.insert(qualified_name.clone());
                    }
                    let suffix = format!(".{member_name}");
                    let has_qualified_collision = if *is_type_like {
                        self.defined_classes
                            .iter()
                            .any(|name| name != qualified_name && name.ends_with(&suffix))
                    } else {
                        self.defined_functions
                            .iter()
                            .any(|name| name != qualified_name && name.ends_with(&suffix))
                    };
                    if !has_qualified_collision
                        && !self.defined_globals.contains(member_name)
                        && !self.defined_classes.contains(member_name)
                        && !self.defined_functions.contains(member_name)
                    {
                        let source_idx = self.str_const(member_name);
                        self.emit_u16(Op::GLOBAL_GET, qualified_idx);
                        self.emit_u16(Op::GLOBAL_SET, source_idx);
                    }
                }

                // Build namespace struct
                self.emit_u16(Op::STRUCT_NEW, 0);
                for (member_name, qualified_name, _) in &member_names {
                    inst!(self, core_wasm::dup);
                    let gidx = self.str_const(qualified_name);
                    self.emit_u16(Op::GLOBAL_GET, gidx);
                    let key = self.str_const(member_name);
                    self.emit_u16(Op::STRUCT_SET, key);
                    self.emit(Op::DROP);
                }
                let ns_idx = self.str_const(&ns_name);
                self.emit_u16(Op::GLOBAL_SET, ns_idx);
                self.defined_globals.insert(ns_name.clone());

                let namespace_parts: Vec<&str> = ns_name
                    .split('.')
                    .map(|part| part.trim())
                    .filter(|part| !part.is_empty())
                    .collect();
                if namespace_parts.len() > 1 {
                    for depth in 1..namespace_parts.len() {
                        let parent_name = self.canon(&namespace_parts[..depth].join("."));
                        let child_name = self.canon(&namespace_parts[..=depth].join("."));
                        let child_key = self.canon(namespace_parts[depth]);

                        if self.defined_globals.contains(&parent_name) {
                            let parent_idx = self.str_const(&parent_name);
                            self.emit_u16(Op::GLOBAL_GET, parent_idx);
                        } else {
                            self.emit_u16(Op::STRUCT_NEW, 0);
                        }
                        inst!(self, core_wasm::dup);
                        let child_idx = self.str_const(&child_name);
                        self.emit_u16(Op::GLOBAL_GET, child_idx);
                        let key_idx = self.str_const(&child_key);
                        self.emit_u16(Op::STRUCT_SET, key_idx);
                        self.emit(Op::DROP);
                        let parent_idx = self.str_const(&parent_name);
                        self.emit_u16(Op::GLOBAL_SET, parent_idx);
                        self.defined_globals.insert(parent_name);
                    }
                }
            }

            // ── Delegate declaration ────────────────────────────────────
            StmtKind::DelegateDecl { .. } => {
                // No-op — delegates are type-level
            }

            // ── With ────────────────────────────────────────────────────
            StmtKind::With { items, body, .. } => {
                self.scope_mut().begin_scope();
                if let Some(first) = items.first() {
                    self.compile_expr(&first.expr)?;
                    let slot = if let Some(ref var) = first.var {
                        self.define_local(var)
                    } else {
                        self.define_local("__with_target")
                    };
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.with_targets.push(slot);
                }
                for s in body {
                    self.compile_stmt(s)?;
                }
                if !items.is_empty() {
                    self.with_targets.pop();
                }
                self.scope_mut().end_scope();
            }

            // ── Using ───────────────────────────────────────────────────
            // ECMA-334 §13.14: `using (var r = expr) { body; }` is
            // equivalent to:
            //
            //     var r = expr;
            //     try { body; } finally { r?.Dispose(); }
            //
            // Wrapping in real try/finally bytecode means an exception
            // escaping the body still triggers Dispose — matching the
            // C# semantic exercised by `using_disposes_on_exception`.
            // Cross-language: Python `with`, Java try-with-resources,
            // JS Stage 3 `using` share the same lowering.
            StmtKind::Using {
                var,
                resource,
                body,
            } => {
                let resource_type_hint = self
                    .infer_expr_type_hint(resource)
                    .map(|type_hint| self.resolve_source_type_alias(&type_hint));
                self.compile_expr(resource)?;
                let slot = self.define_local_typed(var, resource_type_hint);
                self.emit_u16(Op::LOCAL_SET, slot);

                let line = self.line;
                let after_using_try = self.chunk().emit_block(line);
                let catch_jump =
                    common::errors::emit_try_start(&mut self.chunks[self.current], line);
                // `try_table` is a structural block — count it while compiling
                // the protected body (see the main `try` site above).
                self.label_depth += 1;
                self.active_finally_blocks
                    .push(FinallyAction::ResourceDispose {
                        slot,
                        method: "Dispose".to_string(),
                        line,
                    });
                for s in body {
                    self.compile_stmt(s)?;
                }
                self.active_finally_blocks.pop();
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                self.label_depth -= 1;
                self.chunk().emit_br(0, line);
                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);
                // Catch arm: dispose, then rethrow the exception
                // (which is on TOS after `patch_catch`).
                let exc_slot = self.define_local("__using_exc");
                self.emit_u16(Op::LOCAL_SET, exc_slot);
                self.label_depth += 1;
                common::errors::emit_resource_dispose(self.chunk(), slot, "Dispose", line);
                self.label_depth -= 1;
                self.emit_u16(Op::LOCAL_GET, exc_slot);
                common::errors::emit_throw(&mut self.chunks[self.current], line);
                // Normal-completion path: dispose, fall through.
                self.chunk().emit_end(line);
                self.chunk().patch_block(after_using_try);
                self.label_depth += 1;
                common::errors::emit_resource_dispose(self.chunk(), slot, "Dispose", line);
                self.label_depth -= 1;
            }

            // ── Lock ────────────────────────────────────────────────────
            StmtKind::Lock { body, .. } => {
                // No real locking in our VM — just compile body
                for s in body {
                    self.compile_stmt(s)?;
                }
            }

            // ── ReDim ───────────────────────────────────────────────────
            // VB `ReDim arr(N)` allocates a fresh array of N+1 elements;
            // `ReDim Preserve arr(N)` allocates a new array AND copies the
            // old elements over (extending with defaults if growing). The
            // upper bound is inclusive (N → N+1 length).
            StmtKind::ReDim {
                array,
                bounds,
                preserve,
            } => {
                if let Some(size_expr) = bounds.first() {
                    let line = self.line;
                    if *preserve {
                        // Allocate new array of N+1, then iterate the OLD
                        // array via compiler_common::loops::emit_for_in_start
                        // and copy each element into new[i] (bounded by
                        // new_len). This reuses the canonical for-in loop
                        // emit pattern that every other iteration site uses.
                        let old_slot = self.define_local("__redim_old");
                        let new_slot = self.define_local("__redim_new");
                        let new_len_slot = self.define_local("__redim_nlen");
                        let idx_slot = self.define_local("__redim_idx");
                        let old_len_slot = self.define_local("__redim_olen");
                        let fill_idx_slot = self.define_local("__redim_fill_idx");
                        let default_slot = self.define_local("__redim_default");

                        // old = arr
                        self.emit_var_get(array);
                        self.emit_u16(Op::LOCAL_SET, old_slot);
                        // new_len = N + 1
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                        };
                        self.emit_u16(Op::LOCAL_SET, new_len_slot);
                        self.emit_u16(Op::LOCAL_GET, old_slot);
                        common::collections::emit_len(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, old_len_slot);
                        // new = newWithLength(new_len) via common::collections
                        self.emit_u16(Op::LOCAL_GET, new_len_slot);
                        common::collections::emit_new_with_length(
                            &mut self.chunks,
                            self.current,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_SET, new_slot);

                        // Iterate old array with the canonical for-in helper.
                        // The helper leaves [element] on the stack each pass
                        // and exposes the index in `idx_slot`.
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks,
                            self.current,
                            old_slot,
                            idx_slot,
                            line,
                        );
                        // Stack: [element]. If idx >= new_len, drop and break
                        // (don't write past the new array). Otherwise
                        // new[idx] = element.
                        let elem_slot = self.define_local("__redim_el");
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_GET, new_len_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                        };
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        // in bounds: new[idx] = element via common::collections::emit_set.
                        self.emit_u16(Op::LOCAL_GET, new_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        // emit_set preserves [val] — drop it.
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);

                        common::loops::emit_for_in_end(
                            &mut self.chunks,
                            self.current,
                            idx_slot,
                            lp,
                            line,
                        );

                        // Fill any grown tail with the array's default value.
                        // Until arrays carry static element metadata, infer the
                        // default from the first existing element's runtime
                        // category: numbers -> 0, bools -> false, refs -> null.
                        self.emit(Op::NULL);
                        self.emit_u16(Op::LOCAL_SET, default_slot);
                        self.emit_u16(Op::LOCAL_GET, old_len_slot);
                        self.emit_const(Value::F64(0.0));
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                        };
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);

                        self.emit_u16(Op::LOCAL_GET, old_slot);
                        inst!(self, core_wasm::i32_const, 0);
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        let seed_slot = self.define_local("__redim_seed");
                        self.emit_u16(Op::LOCAL_SET, seed_slot);

                        self.emit_u16(Op::LOCAL_GET, seed_slot);
                        fn_call!(self, "wasm:js-boolean", "test", 1);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        inst!(self, core_wasm::bool_const, false);
                        self.emit_u16(Op::LOCAL_SET, default_slot);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, seed_slot);
                        fn_call!(self, "wasm:js-number", "test", 1);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        inst!(self, core_wasm::i32_const, 0);
                        self.emit_u16(Op::LOCAL_SET, default_slot);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);

                        self.emit_u16(Op::LOCAL_GET, old_len_slot);
                        self.emit_u16(Op::LOCAL_SET, fill_idx_slot);
                        let fill_block = self.chunk().emit_block(line);
                        let (fill_loop, _) = self.chunk().emit_loop_s(line);
                        self.emit_u16(Op::LOCAL_GET, fill_idx_slot);
                        self.emit_u16(Op::LOCAL_GET, new_len_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_br_if(1, line);

                        self.emit_u16(Op::LOCAL_GET, new_slot);
                        self.emit_u16(Op::LOCAL_GET, fill_idx_slot);
                        self.emit_u16(Op::LOCAL_GET, default_slot);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, fill_idx_slot);
                        self.emit_const(Value::F64(1.0));
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                        };
                        self.emit_u16(Op::LOCAL_SET, fill_idx_slot);
                        self.chunk().emit_br(0, line);
                        self.chunk().emit_end(line);
                        self.chunk().patch_loop(fill_loop);
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(fill_block);

                        // arr = new
                        self.emit_u16(Op::LOCAL_GET, new_slot);
                        self.emit_var_set(array);
                    } else {
                        // ReDim arr(N) — non-preserving. N is the upper
                        // bound; length is N+1. Emit through
                        // `common::collections` (Phase D2).
                        let line = self.line;
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                        };
                        common::collections::emit_new_with_length(
                            &mut self.chunks,
                            self.current,
                            line,
                        );
                        self.emit_var_set(array);
                    }
                }
            }

            // ── Events ──────────────────────────────────────────────────
            // AddHandler/RemoveHandler are language-agnostic statements; the
            // canonical AST holds the control + handler as Expressions, so any
            // frontend (.NET, MAUI, Flutter, …) can produce the same node by
            // mapping its surface syntax (`Handles X.Y`, `obj.Y += h`, etc.).
            //
            // The handler is registered under the SOURCE-CODE-STABLE control
            // identifier (field name, class name for `Me`/`This`, or runtime
            StmtKind::Erase { array } => {
                let line = self.line;
                let Some(binding) = self.lookup_array_binding(array).cloned() else {
                    self.emit(Op::NULL);
                    self.emit_var_set(array);
                    self.emit(Op::DROP);
                    return Ok(());
                };

                if !binding.is_fixed {
                    self.emit(Op::NULL);
                    self.emit_var_set(array);
                    self.emit(Op::DROP);
                    return Ok(());
                }

                let old_slot = self.define_local("__erase_old");
                let len_slot = self.define_local("__erase_len");
                let new_slot = self.define_local("__erase_new");

                self.emit_var_get(array);
                self.emit_u16(Op::LOCAL_SET, old_slot);

                self.emit_u16(Op::LOCAL_GET, old_slot);
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, len_slot);

                self.emit_u16(Op::LOCAL_GET, len_slot);
                common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, new_slot);

                self.emit_u16(Op::LOCAL_GET, new_slot);
                self.emit_default_value_for_type_hint(
                    binding
                        .type_hint
                        .as_deref()
                        .map(|type_hint| type_hint.trim().trim_end_matches("()").trim()),
                );
                self.emit_const(Value::I32(0));
                self.emit_const(Value::I32(i32::MAX));
                common::collections::emit_fill(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, new_slot);
                self.emit_var_set(array);
                self.emit(Op::DROP);
            }
            // `__control_name` for general expressions). This decouples the
            // registry key from the runtime `.Name` property — renaming a
            // control via `btn.Name = "x"` doesn't break wired-up handlers.
            StmtKind::AddHandler {
                control,
                event,
                handler,
            } => {
                self.compile_add_handler_stmt(control, event, handler)?;
            }

            StmtKind::RemoveHandler {
                control,
                event,
                handler,
            } => {
                self.compile_remove_handler_stmt(control, event, handler)?;
            }

            StmtKind::RaiseEvent { event_name, args } => {
                self.compile_raise_event_stmt(event_name, args)?;
            }

            // ── VB legacy error handling ────────────────────────────────
            StmtKind::OnErrorResumeNext => { /* no-op in bytecode VM */ }
            StmtKind::OnErrorGoTo(_) => { /* no-op */ }
            StmtKind::GoTo(_) => { /* no-op — structured bytecode doesn't support arbitrary gotos */
            }
            StmtKind::Label(_) => { /* no-op */ }

            // ── VB legacy file I/O ──────────────────────────────────────
            StmtKind::OpenFile {
                path,
                mode,
                file_number,
            } => {
                let path_slot = self.define_local("__vb_open_path");
                let file_slot = self.define_local("__vb_open_file_number");
                let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");
                let mode_text = match mode {
                    FileMode::Input => "Input",
                    FileMode::Output => "Output",
                    FileMode::Append => "Append",
                    FileMode::Binary => "Binary",
                    FileMode::Random => "Random",
                };

                self.compile_expr(path)?;
                self.emit_u16(Op::LOCAL_SET, path_slot);

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);

                self.emit_u16(Op::LOCAL_GET, path_slot);
                self.emit_const(Value::String(Arc::from(mode_text)));
                self.emit_u16(Op::LOCAL_GET, file_slot);
                let idx = self.import("wasi:filesystem", "openFile");
                self.emit_host_call(idx, 3);
                self.emit(Op::DROP);

                self.emit_ensure_global_map("__vb_file_path_by_handle");
                self.emit_u16(Op::GLOBAL_GET, path_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_u16(Op::LOCAL_GET, path_slot);
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);

                self.emit_ensure_global_map("__vb_file_eof_by_handle");
                self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_const(Value::Bool(false));
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);

                self.emit_global_map_set_null("__vb_record_rows_by_handle", file_slot);
                self.emit_global_map_set_null("__vb_record_next_index_by_handle", file_slot);
                self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
            }
            StmtKind::CloseFile(file_num) => {
                let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");
                if let Some(fnum) = file_num {
                    let file_slot = self.define_local("__vb_close_file_number");
                    self.compile_expr(fnum)?;
                    self.emit_u16(Op::LOCAL_SET, file_slot);

                    self.emit_u16(Op::LOCAL_GET, file_slot);
                    let idx = self.import("wasi:filesystem", "closeFile");
                    self.emit_host_call(idx, 1);
                    self.emit(Op::DROP);

                    self.emit_ensure_global_map("__vb_file_path_by_handle");
                    self.emit_u16(Op::GLOBAL_GET, path_map_key);
                    self.emit_u16(Op::LOCAL_GET, file_slot);
                    self.emit(Op::NULL);
                    self.emit(Op::ARRAY_SET);
                    self.emit(Op::DROP);

                    self.emit_ensure_global_map("__vb_file_eof_by_handle");
                    self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                    self.emit_u16(Op::LOCAL_GET, file_slot);
                    self.emit_const(Value::Bool(false));
                    self.emit(Op::ARRAY_SET);
                    self.emit(Op::DROP);

                    self.emit_global_map_set_null("__vb_record_rows_by_handle", file_slot);
                    self.emit_global_map_set_null("__vb_record_next_index_by_handle", file_slot);
                    self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                } else {
                    self.emit(Op::NULL);
                    let idx = self.import("wasi:filesystem", "closeFile");
                    self.emit_host_call(idx, 1);
                    self.emit(Op::DROP);
                }
            }
            StmtKind::PrintFile { file_number, items } => {
                self.compile_expr(file_number)?;
                for item in items {
                    self.compile_expr(item)?;
                }
                let idx = self.import("wasi:filesystem", "printFile");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::DROP);
            }
            StmtKind::WriteFile { file_number, items } => {
                self.compile_expr(file_number)?;
                for item in items {
                    self.compile_expr(item)?;
                }
                let idx = self.import("wasi:filesystem", "writeFile_handle");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::DROP);
            }
            StmtKind::InputFile {
                file_number,
                variables,
            } => {
                let file_slot = self.define_local("__vb_input_file_number");
                let values_slot = self.define_local("__vb_input_values");
                let rows_slot = self.define_local("__vb_input_rows");
                let len_slot = self.define_local("__vb_input_len");
                let idx_slot = self.define_local("__vb_input_idx");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);

                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);
                self.emit_global_map_get_into_local(
                    "__vb_record_next_index_by_handle",
                    file_slot,
                    idx_slot,
                );
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if(line);
                inst!(self, core_wasm::i32_const, 0);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_GET, file_slot);
                let idx = self.import("wasi:filesystem", "inputFile");
                self.emit_host_call(idx, 1);
                self.emit_u16(Op::LOCAL_SET, values_slot);

                for (index, variable) in variables.iter().enumerate() {
                    self.emit_u16(Op::LOCAL_GET, values_slot);
                    self.emit_const(Value::F64(index as f64));
                    self.emit(Op::ARRAY_GET);
                    self.emit_assignment_type_coercion_for_target(variable);
                    self.compile_assign_target(variable)?;
                }

                if variables.is_empty() {
                    self.emit_u16(Op::LOCAL_GET, values_slot);
                    self.emit(Op::DROP);
                }

                self.emit_ensure_global_map("__vb_file_eof_by_handle");
                self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                inst!(self, core_wasm::i32_const, 1);
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);

                self.emit_global_map_set_from_local(
                    "__vb_record_next_index_by_handle",
                    file_slot,
                    idx_slot,
                );

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, len_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                };
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);
            }
            StmtKind::LineInput {
                file_number,
                variable,
            } => {
                let file_slot = self.define_local("__vb_line_input_file_number");
                let rows_slot = self.define_local("__vb_line_input_rows");
                let len_slot = self.define_local("__vb_line_input_len");
                let idx_slot = self.define_local("__vb_line_input_idx");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);

                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);
                self.emit_global_map_get_into_local(
                    "__vb_record_next_index_by_handle",
                    file_slot,
                    idx_slot,
                );
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if(line);
                inst!(self, core_wasm::i32_const, 0);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_GET, file_slot);
                let idx = self.import("wasi:filesystem", "lineInput");
                self.emit_host_call(idx, 1);
                self.emit_var_set(variable);

                self.emit_ensure_global_map("__vb_file_eof_by_handle");
                self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                inst!(self, core_wasm::i32_const, 1);
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);

                self.emit_global_map_set_from_local(
                    "__vb_record_next_index_by_handle",
                    file_slot,
                    idx_slot,
                );

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, len_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                };
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);
            }
            StmtKind::StartFile {
                file_number,
                key_index,
                key_value,
                relation,
            } => {
                let line = self.line;
                let file_slot = self.define_local("__vb_start_file_number");
                let rows_slot = self.define_local("__vb_start_rows");
                let len_slot = self.define_local("__vb_start_len");
                let key_slot = self.define_local("__vb_start_key");
                let found_slot = self.define_local("__vb_start_found");
                let idx_slot = self.define_local("__vb_start_idx");
                let row_slot = self.define_local("__vb_start_row");
                let values_slot = self.define_local("__vb_start_values");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);

                self.compile_expr(key_value)?;
                self.emit_u16(Op::LOCAL_SET, key_slot);
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, found_slot);

                let state = common::loops::emit_for_in_start(
                    &mut self.chunks,
                    self.current,
                    rows_slot,
                    idx_slot,
                    line,
                );
                self.emit_u16(Op::LOCAL_SET, row_slot);
                self.emit_u16(Op::LOCAL_GET, row_slot);
                self.emit_const(Value::String(Arc::from(",")));
                fn_call!(self, "ecma:string", "split", 2);
                self.emit_u16(Op::LOCAL_SET, values_slot);
                self.emit_u16(Op::LOCAL_GET, values_slot);
                self.emit_const(Value::F64(*key_index as f64));
                self.emit(Op::ARRAY_GET);
                self.emit_u16(Op::LOCAL_GET, key_slot);
                self.emit_file_key_compare(*relation);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_SET, found_slot);
                self.chunks[self.current].emit_br(state.break_depth(0).into(), line);
                self.chunk().emit_end(line);
                common::loops::emit_for_in_end(
                    &mut self.chunks,
                    self.current,
                    idx_slot,
                    state,
                    line,
                );

                self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                self.emit_u16(Op::LOCAL_GET, found_slot);
                self.emit(Op::REF_IS_NULL);
                self.chunk().emit_if(line);
                self.emit_global_map_set_from_local(
                    "__vb_record_next_index_by_handle",
                    file_slot,
                    len_slot,
                );
                self.emit_global_map_set_const(
                    "__vb_file_eof_by_handle",
                    file_slot,
                    Value::Bool(true),
                );
                self.chunk().emit_else(line);
                self.emit_global_map_set_from_local(
                    "__vb_record_next_index_by_handle",
                    file_slot,
                    found_slot,
                );
                self.emit_global_map_set_const(
                    "__vb_file_eof_by_handle",
                    file_slot,
                    Value::Bool(false),
                );
                self.chunk().emit_end(line);
            }
            StmtKind::InputRecordFile {
                file_number,
                variables,
                key_index,
                key_value,
            } => {
                let line = self.line;
                let file_slot = self.define_local("__vb_record_file_number");
                let rows_slot = self.define_local("__vb_record_rows");
                let len_slot = self.define_local("__vb_record_len");
                let idx_slot = self.define_local("__vb_record_idx");
                let row_slot = self.define_local("__vb_record_row");
                let values_slot = self.define_local("__vb_record_values");
                let found_slot = self.define_local("__vb_record_found");
                let key_slot = key_value
                    .as_ref()
                    .map(|_| self.define_local("__vb_record_key"));

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);

                if let Some(key_expr) = key_value {
                    let key_slot = key_slot.expect("key slot allocated when key_value exists");
                    let key_index = key_index.unwrap_or(0);

                    self.compile_expr(key_expr)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, found_slot);

                    let state = common::loops::emit_for_in_start(
                        &mut self.chunks,
                        self.current,
                        rows_slot,
                        idx_slot,
                        line,
                    );
                    self.emit_u16(Op::LOCAL_SET, row_slot);
                    self.emit_u16(Op::LOCAL_GET, row_slot);
                    self.emit_const(Value::String(Arc::from(",")));
                    fn_call!(self, "ecma:string", "split", 2);
                    self.emit_u16(Op::LOCAL_SET, values_slot);
                    self.emit_u16(Op::LOCAL_GET, values_slot);
                    self.emit_const(Value::F64(key_index as f64));
                    self.emit(Op::ARRAY_GET);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    self.emit_file_key_compare(FileKeyRelation::Equal);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_SET, found_slot);
                    self.chunks[self.current].emit_br(state.break_depth(0).into(), line);
                    self.chunk().emit_end(line);
                    common::loops::emit_for_in_end(
                        &mut self.chunks,
                        self.current,
                        idx_slot,
                        state,
                        line,
                    );

                    self.emit_u16(Op::LOCAL_GET, found_slot);
                    self.emit(Op::REF_IS_NULL);
                    self.chunk().emit_if(line);
                    self.emit_record_assign_nulls(variables);
                    self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                    self.emit_global_map_set_const(
                        "__vb_file_eof_by_handle",
                        file_slot,
                        Value::Bool(true),
                    );
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, rows_slot);
                    self.emit_u16(Op::LOCAL_GET, found_slot);
                    self.emit(Op::ARRAY_GET);
                    self.emit_u16(Op::LOCAL_SET, row_slot);
                    self.emit_u16(Op::LOCAL_GET, row_slot);
                    self.emit_const(Value::String(Arc::from(",")));
                    fn_call!(self, "ecma:string", "split", 2);
                    self.emit_u16(Op::LOCAL_SET, values_slot);
                    self.emit_record_assign_values_from_local(values_slot, variables);
                    self.emit_global_map_set_from_local(
                        "__vb_record_current_index_by_handle",
                        file_slot,
                        found_slot,
                    );
                    self.emit_u16(Op::LOCAL_GET, found_slot);
                    inst!(self, core_wasm::i32_const, 1);
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit_global_map_set_from_local(
                        "__vb_record_next_index_by_handle",
                        file_slot,
                        idx_slot,
                    );
                    self.emit_global_map_set_const(
                        "__vb_file_eof_by_handle",
                        file_slot,
                        Value::Bool(false),
                    );
                    self.chunk().emit_end(line);
                } else {
                    self.emit_global_map_get_into_local(
                        "__vb_record_next_index_by_handle",
                        file_slot,
                        idx_slot,
                    );
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::REF_IS_NULL);
                    self.chunk().emit_if(line);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, rows_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::ARRAY_GET);
                    self.emit_u16(Op::LOCAL_SET, row_slot);
                    self.emit_u16(Op::LOCAL_GET, row_slot);
                    self.emit_const(Value::String(Arc::from(",")));
                    fn_call!(self, "ecma:string", "split", 2);
                    self.emit_u16(Op::LOCAL_SET, values_slot);
                    self.emit_record_assign_values_from_local(values_slot, variables);
                    self.emit_global_map_set_from_local(
                        "__vb_record_current_index_by_handle",
                        file_slot,
                        idx_slot,
                    );
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    inst!(self, core_wasm::i32_const, 1);
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit_global_map_set_from_local(
                        "__vb_record_next_index_by_handle",
                        file_slot,
                        idx_slot,
                    );
                    self.emit_global_map_set_const(
                        "__vb_file_eof_by_handle",
                        file_slot,
                        Value::Bool(false),
                    );
                    self.chunk().emit_else(line);
                    self.emit_record_assign_nulls(variables);
                    self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                    self.emit_global_map_set_const(
                        "__vb_file_eof_by_handle",
                        file_slot,
                        Value::Bool(true),
                    );
                    self.chunk().emit_end(line);
                }
            }
            StmtKind::RewriteRecordFile {
                file_number,
                items,
                field_formats,
            } => {
                let line = self.line;
                let file_slot = self.define_local("__vb_rewrite_file_number");
                let rows_slot = self.define_local("__vb_rewrite_rows");
                let len_slot = self.define_local("__vb_rewrite_len");
                let current_slot = self.define_local("__vb_rewrite_current");
                let line_slot = self.define_local("__vb_rewrite_line");
                let items_slot = self.define_local("__vb_rewrite_items");
                let path_slot = self.define_local("__vb_rewrite_path");
                let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);
                self.emit_global_map_get_into_local(
                    "__vb_record_current_index_by_handle",
                    file_slot,
                    current_slot,
                );
                self.emit_u16(Op::LOCAL_GET, current_slot);
                self.emit(Op::REF_IS_NULL);
                self.emit(Op::I32_EQZ);
                self.chunk().emit_if(line);

                for (index, item) in items.iter().enumerate() {
                    self.compile_expr(item)?;
                    self.emit_record_rewrite_field_format(
                        field_formats.get(index).and_then(|format| format.as_ref()),
                    );
                }
                common::collections::emit_array_new(
                    &mut self.chunks,
                    self.current,
                    items.len() as u16,
                    line,
                );
                self.emit_u16(Op::LOCAL_SET, items_slot);
                self.emit_u16(Op::LOCAL_GET, items_slot);
                self.emit_const(Value::String(Arc::from(",")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, line_slot);

                self.emit_u16(Op::LOCAL_GET, rows_slot);
                self.emit_u16(Op::LOCAL_GET, current_slot);
                self.emit_u16(Op::LOCAL_GET, line_slot);
                common::collections::emit_set(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);

                self.emit_ensure_global_map("__vb_file_path_by_handle");
                self.emit_u16(Op::GLOBAL_GET, path_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit(Op::ARRAY_GET);
                self.emit_u16(Op::LOCAL_SET, path_slot);

                self.emit_u16(Op::LOCAL_GET, path_slot);
                self.emit_u16(Op::LOCAL_GET, rows_slot);
                self.emit_const(Value::String(Arc::from("\n")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
                let write_file_idx = self.import("wasi:filesystem", "writeFile");
                self.emit_host_call(write_file_idx, 2);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);
            }

            // ── Export ──────────────────────────────────────────────────
            StmtKind::Export {
                declaration,
                default,
                ..
            } => {
                if let Some(decl) = declaration {
                    self.compile_stmt(decl)?;
                }
                if let Some(expr) = default {
                    self.compile_expr(expr)?;
                    let idx = self.str_const("default");
                    self.emit_u16(Op::GLOBAL_SET, idx);
                }
            }

            // ── Labeled statement ───────────────────────────────────────
            StmtKind::Labeled { label, body } => {
                // Store label so the next loop/switch push picks it up.
                self.pending_label = Some(label.clone());
                // Check if this is a non-loop body (plain block etc.). If so,
                // we need to emit a WASM block + LoopCtx so that `break label`
                // can find the label (ECMA-262 §14.13: labeled block statements
                // accept `break <label>`).
                let is_loop_body = matches!(
                    &body.kind,
                    StmtKind::While { .. }
                        | StmtKind::DoWhile { .. }
                        | StmtKind::For { .. }
                        | StmtKind::ForIn { .. }
                        | StmtKind::Switch { .. }
                );
                let block_patch = if !is_loop_body {
                    let line = self.line;
                    let bp = self.chunk().emit_block(line);
                    self.label_depth += 1;
                    let lp = common::loops::LoopState {
                        block_patch: bp,
                        loop_patch: 0,
                        body_block_patch: None,
                    };
                    self.loop_states.push(lp);
                    self.loops.push(LoopCtx {
                        label: self.pending_label.take(),
                        break_label_depth: self.label_depth,
                        continue_label_depth: self.label_depth,
                        did_break_slot: None,
                        iterator_close_slot: None,
                        is_continuable: false,
                        finally_depth: self.active_finally_blocks.len(),
                    });
                    Some(bp)
                } else {
                    None
                };
                self.compile_stmt(body)?;
                self.pending_label = None;
                if let Some(_) = block_patch {
                    let line = self.line;
                    self.chunk().emit_end(line);
                    let lp = self.loop_states.pop().unwrap();
                    self.chunk().patch_block(lp.block_patch);
                    self.loops.pop();
                    self.label_depth -= 1;
                }
            }

            // ── Echo (PHP/debug print) ──────────────────────────────────
            StmtKind::Echo(exprs) => {
                let line = self.line;
                let log_idx = self.import("wasi:logging/logging", "log");
                let php_echo = self.profile.name == "php";
                if self.profile.name == "cobol" {
                    if exprs.is_empty() {
                        self.emit_const(Value::String(Arc::from("")));
                    } else {
                        self.compile_expr(&exprs[0])?;
                        let line = self.line;
                        common::strings::emit_to_string(self.chunk(), line);
                        for expr in exprs.iter().skip(1) {
                            self.compile_expr(expr)?;
                            let line = self.line;
                            common::strings::emit_to_string(self.chunk(), line);
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                            };
                        }
                    }
                    common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);
                } else {
                    // PHP echo writes raw bytes to stdout (no newline) —
                    // the WASI 0.3 stream surface, not wasi:logging.log.
                    let php_write_idx = if php_echo {
                        Some(self.import("wasi:cli/stdout", "write-via-stream"))
                    } else {
                        None
                    };
                    for expr in exprs {
                        self.compile_expr(expr)?;
                        if php_echo {
                            // PHP: when echoing an object with `__toString`,
                            // call the method and print its result. Other
                            // values pass through. The check is a runtime
                            // struct_get on the value's `__toString` slot;
                            // if non-null, invoke as a method.
                            //
                            // Also: PHP `echo null;` writes no bytes (vs.
                            // Vybe's normal flow which would log ""); skip
                            // the write call when the expression is null so
                            // test-runner output entries match PHP-stdout
                            // bytes.
                            let v_slot = self.define_local("__echo_v");
                            self.emit_u16(Op::LOCAL_SET, v_slot);
                            self.emit_u16(Op::LOCAL_GET, v_slot);
                            self.emit(Op::REF_IS_NULL);
                            self.emit(Op::I32_EQZ);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            // Probe the ToString SLOT, not the PHP spelling. A
                            // value echoed here may have been declared in any
                            // language — the slot is filled by Python's
                            // `__str__`, Ruby's `to_s` and C#'s `ToString`
                            // just as much as by PHP's `__toString`, so this
                            // stringifies a foreign object correctly instead of
                            // falling through to the default rendering.
                            self.emit_u16(Op::LOCAL_GET, v_slot);
                            let ts_key = self.str_const(&vybe_ast::protocol_slot_key(
                                vybe_ast::ProtocolSlot::ToString,
                            ));
                            self.emit_u16(Op::STRUCT_GET, ts_key);
                            let fn_slot = self.define_local("__echo_ts_fn");
                            self.emit_u16(Op::LOCAL_SET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit(Op::REF_IS_NULL);
                            self.emit(Op::I32_EQZ);
                            let line = self.line;
                            self.chunk().emit_if_value(line);
                            // Has __toString — invoke (fn, this).
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, v_slot);
                            self.emit_u8(Op::CALL_REF, 1);
                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, v_slot);
                            self.chunk().emit_end(line);
                            self.emit_common("php.echo_stringify", 1, line);
                            let out_slot = self.define_local("__echo_out");
                            self.emit_u16(Op::LOCAL_SET, out_slot);
                            let rd_slot = self.define_local("__echo_rd");
                            let wr_slot = self.define_local("__echo_wr");
                            common::io::emit_write_stdout_with_imports(
                                self.chunk(),
                                php_write_idx.unwrap(),
                                rd_slot,
                                wr_slot,
                                line,
                                |c| c.emit_op_u16(Op::LOCAL_GET, out_slot, line),
                            );
                            self.chunk().emit_end(line);
                        } else {
                            common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);
                        }
                    }
                }
            }

            // ── Delete ──────────────────────────────────────────────────
            StmtKind::Delete(exprs) => {
                for expr in exprs {
                    match &expr.kind {
                        ExprKind::Member { object, field, .. } => {
                            self.compile_expr(object)?;
                            self.emit(Op::NULL);
                            let field_name = self.canon(field);
                            let idx = self.str_const(&field_name);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ExprKind::Index { object, index, .. } => {
                            let line = self.line;
                            if self.is_python_profile() {
                                if let ExprKind::Slice { lower, upper, step } = &index.kind {
                                    if step.is_none() {
                                        self.compile_expr(object)?;
                                        let obj_tmp = self.define_local("__delete_slice_obj");
                                        self.emit_u16(Op::LOCAL_SET, obj_tmp);

                                        if let Some(lower) = lower {
                                            self.compile_expr(lower)?;
                                        } else {
                                            inst!(self, core_wasm::i32_const, 0);
                                        }
                                        let start_tmp = self.define_local("__delete_slice_start");
                                        self.emit_u16(Op::LOCAL_SET, start_tmp);

                                        if let Some(upper) = upper {
                                            self.compile_expr(upper)?;
                                        } else {
                                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                            common::collections::emit_len(
                                                &mut self.chunks,
                                                self.current,
                                                line,
                                            );
                                        }
                                        let end_tmp = self.define_local("__delete_slice_end");
                                        self.emit_u16(Op::LOCAL_SET, end_tmp);

                                        self.emit_u16(Op::LOCAL_GET, end_tmp);
                                        self.emit_u16(Op::LOCAL_GET, start_tmp);
                                        self.emit(Op::I32_SUB);
                                        let count_tmp = self.define_local("__delete_slice_count");
                                        self.emit_u16(Op::LOCAL_SET, count_tmp);

                                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                        self.emit_u16(Op::LOCAL_GET, start_tmp);
                                        self.emit_u16(Op::LOCAL_GET, count_tmp);
                                        common::collections::emit_remove_range(
                                            &mut self.chunks,
                                            self.current,
                                            line,
                                        );
                                        self.emit(Op::DROP);
                                        continue;
                                    } else {
                                        // `del a[i:j:k]` — strided deletion via
                                        // the shared slices emitter.
                                        self.compile_expr(object)?;
                                        if let Some(lower) = lower {
                                            self.compile_expr(lower)?;
                                        } else {
                                            self.emit(Op::NULL);
                                        }
                                        if let Some(upper) = upper {
                                            self.compile_expr(upper)?;
                                        } else {
                                            self.emit(Op::NULL);
                                        }
                                        if let Some(step) = step {
                                            self.compile_expr(step)?;
                                        } else {
                                            self.emit(Op::NULL);
                                        }
                                        let opts = crate::primitives::slices::Options::new(
                                            self.profile.slice_step_zero_raises,
                                        );
                                        crate::primitives::slices::emit_strided_del(
                                            &mut self.chunks,
                                            self.current,
                                            line,
                                            opts,
                                        );
                                        continue;
                                    }
                                }
                            }

                            self.compile_expr(object)?;
                            let obj_tmp = self.define_local("__delete_obj");
                            self.emit_u16(Op::LOCAL_SET, obj_tmp);
                            self.compile_expr(index)?;
                            let key_tmp = self.define_local("__delete_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            let is_array_idx = self.import("ecma:array", "isArray");
                            self.chunk().emit_call(is_array_idx, 1, line);
                            inst!(self, core_wasm::i32_const, 0);
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                            };
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            self.chunk().emit_if(line);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            common::collections::emit_remove_at(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            self.emit(Op::DROP);

                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            common::dict::emit_method_delete(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            self.chunk().emit_end(line);
                        }
                        _ => {
                            // Delete on non-member is a no-op
                        }
                    }
                }
            }

            // ── Assert ──────────────────────────────────────────────────
            StmtKind::Assert { test, msg } => {
                self.compile_expr(test)?;
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);
                // Raise a CANONICAL AssertionError (same construction as
                // `raise AssertionError(msg)`), not a bare string — typed
                // `except AssertionError:` matches through the shared shape.
                self.emit_u16(Op::STRUCT_NEW, 0);
                inst!(self, core_wasm::dup);
                if let Some(m) = msg {
                    self.compile_expr(m)?;
                } else {
                    self.emit_const(Value::String(Arc::from("Assertion failed")));
                }
                common::errors::emit_exception_new_finalize(
                    &mut self.chunks[self.current],
                    "AssertionError",
                    line,
                );
                if !self.profile.throwable_is_root {
                    common::errors::emit_stamp_exception_ancestors(
                        &mut self.chunks[self.current],
                        "AssertionError",
                        line,
                    );
                }
                common::errors::emit_throw(&mut self.chunks[self.current], line);
                self.chunk().emit_end(line);
            }

            // ── Scope declarations (Python global/nonlocal, PHP global) ─
            StmtKind::ScopeDecl { kind, names } => {
                if self.profile.name == "php" && matches!(kind, ScopeDeclKind::Global) {
                    let globals: Vec<String> = names.iter().map(|name| self.canon(name)).collect();
                    if let Some(frame) = self.php_function_globals.last_mut() {
                        for name in globals {
                            frame.insert(name);
                        }
                    }
                }
            }

            // ── Match statement (Python) ────────────────────────────────
            StmtKind::MatchStatement { subject, cases } => {
                let line = self.line;
                self.compile_expr(subject)?;
                let subject_slot = self.define_local("__match_subject");
                self.emit_u16(Op::LOCAL_SET, subject_slot);
                let matched_slot = self.define_local("__match_done");
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                for case in cases {
                    self.emit_u16(Op::LOCAL_GET, matched_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if(line);
                    let pattern_match_slot =
                        self.emit_match_pattern_match_slot(&case.pattern, subject_slot)?;
                    self.emit_u16(Op::LOCAL_GET, pattern_match_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    if let Some(guard) = &case.guard {
                        self.compile_expr(guard)?;
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if(line);
                        self.emit_match_pattern_bindings(&case.pattern, subject_slot)?;
                        for s in &case.body {
                            self.compile_stmt(s)?;
                        }
                        self.emit_const(Value::Bool(true));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.chunk().emit_end(line);
                    } else {
                        self.emit_match_pattern_bindings(&case.pattern, subject_slot)?;
                        for s in &case.body {
                            self.compile_stmt(s)?;
                        }
                        self.emit_const(Value::Bool(true));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                    }
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                }
            }

            // ── WASM linear memory / data segments ──────────────────────
            // Module-global: these belong to the script chunk (chunk 0), which
            // the VM instantiates from — it allocates the declared pages and
            // writes active data into linear memory before `_start` runs.
            StmtKind::MemoryDecl {
                min_pages,
                max_pages,
            } => {
                self.chunks[0].memory_min_pages.push(*min_pages);
                self.chunks[0].memory_max_pages.push(*max_pages);
            }
            StmtKind::TableDecl { min_size, max_size } => {
                self.chunks[0].table_min_sizes.push(*min_size);
                self.chunks[0].table_max_sizes.push(*max_size);
            }
            StmtKind::DataSegment {
                memory_index,
                offset,
                bytes,
            } => {
                let data_index = self.chunks[0].data_segments.len() as u32;
                self.chunks[0].data_segments.push(bytes.clone());
                // Active segment (has a constant offset) → recorded for the VM's
                // instantiation-time copy. Passive segments (offset None) stay in
                // `data_segments` for `memory.init`.
                if let Some(off) = offset {
                    let offset_val =
                        const_eval_u64(off, &self.global_const_values).ok_or_else(|| {
                            "data segment offset must be a constant integer expression".to_string()
                        })?;
                    self.chunks[0].active_data_segments.push(
                        vybe_bytecode::chunk::ActiveDataSegment {
                            memory_index: *memory_index,
                            offset: offset_val,
                            data_index,
                        },
                    );
                }
            }

            // ── WASM exception handling (canonical `try_table`) ─────────
            StmtKind::WasmTagDecl { name, arity } => {
                // A tag is a load-time entity resolved by name; importing it
                // here (and at every `throw`/`catch` site) coalesces to one
                // entity. Nothing is emitted into the instruction stream.
                self.chunks[self.current].import_exception_tag(format!("wast:tag:{name}"), *arity);
            }
            StmtKind::WasmThrow { tag, args } => {
                let line = self.line;
                for a in args {
                    self.compile_expr(a)?;
                }
                let tag_idx = self.chunks[self.current]
                    .import_exception_tag(format!("wast:tag:{tag}"), args.len() as u8);
                let chunk = &mut self.chunks[self.current];
                chunk.emit_op(Op::THROW, line);
                chunk.emit((tag_idx >> 8) as u8, line);
                chunk.emit((tag_idx & 0xff) as u8, line);
            }
            StmtKind::WasmRethrow { exnref_local } => {
                // Legacy `rethrow N` → `throw_ref` of the exnref captured by
                // the target catch handler (bound to `exnref_local`).
                self.emit_var_get(exnref_local);
                self.emit(Op::THROW_REF);
            }
            StmtKind::WasmTryTable { body, catches } => {
                let line = self.line;
                // Resolve each clause's kind + tag index up front.
                let mut clause_kinds: Vec<u8> = Vec::with_capacity(catches.len());
                let mut clause_tags: Vec<u16> = Vec::with_capacity(catches.len());
                for c in catches {
                    let (kind, tag_idx) = match (&c.tag, c.capture_ref) {
                        (Some(name), false) => (
                            common::errors::CATCH_KIND_CATCH,
                            self.chunks[self.current].import_exception_tag(
                                format!("wast:tag:{name}"),
                                c.payload_binds.len() as u8,
                            ),
                        ),
                        (Some(name), true) => (
                            common::errors::CATCH_KIND_CATCH_REF,
                            self.chunks[self.current].import_exception_tag(
                                format!("wast:tag:{name}"),
                                c.payload_binds.len() as u8,
                            ),
                        ),
                        (None, false) => (common::errors::CATCH_KIND_CATCH_ALL, 0u16),
                        (None, true) => (common::errors::CATCH_KIND_CATCH_ALL_REF, 0u16),
                    };
                    clause_kinds.push(kind);
                    clause_tags.push(tag_idx);
                }

                // Join block: normal completion and every handler branch here.
                let after = self.chunk().emit_block(line);
                self.label_depth += 1;

                // try_table with one clause per catch (offsets patched later),
                // via the shared single-source-of-truth primitive.
                let clauses: Vec<common::errors::TryTableClause> = (0..catches.len())
                    .map(|i| common::errors::TryTableClause {
                        kind: clause_kinds[i],
                        tag: clause_tags[i],
                    })
                    .collect();
                let offset_positions =
                    common::errors::emit_try_table(&mut self.chunks[self.current], &clauses, line);

                // Body sits one label level deeper (try_table is a block).
                self.label_depth += 1;
                for s in body {
                    self.compile_stmt(s)?;
                }
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                self.label_depth -= 1;

                // Normal completion: skip every handler → join block's `end`.
                self.chunk().emit_br(0, line);

                // Handlers in clause order. Each patches its clause offset to
                // its start, binds the delivered exnref/payload, runs, then
                // branches to the join (the last falls through to the `end`).
                for (i, c) in catches.iter().enumerate() {
                    common::errors::patch_catch(
                        &mut self.chunks[self.current],
                        offset_positions[i],
                    );
                    if c.capture_ref {
                        if let Some(exnref) = &c.exnref_bind {
                            let slot = self.define_local(exnref);
                            self.emit_u16(Op::LOCAL_SET, slot);
                        }
                    }
                    for bind in c.payload_binds.iter().rev() {
                        let slot = self.define_local(bind);
                        self.emit_u16(Op::LOCAL_SET, slot);
                    }
                    for s in &c.body {
                        self.compile_stmt(s)?;
                    }
                    if i + 1 < catches.len() {
                        self.chunk().emit_br(0, line);
                    }
                }

                // Close the join block.
                self.chunk().emit_end(line);
                self.chunk().patch_block(after);
                self.label_depth -= 1;
            }

            // ── Empty ───────────────────────────────────────────────────
            StmtKind::Empty => {}
        }
        Ok(())
    }

    pub(super) fn compile_enum_decl_as_class(
        &mut self,
        name: &str,
        parent: Option<&str>,
        interfaces: &[String],
        members: Vec<ClassMember>,
        span: Span,
    ) -> Result<(), String> {
        let parents: Vec<String> = parent.into_iter().map(|value| value.to_string()).collect();
        let cname = self.canon(name);
        self.defined_globals.insert(cname.clone());
        self.defined_classes.insert(cname.clone());
        crate::primitives::class_normalize::emit::emit_class_from_ast(
            self,
            span,
            &cname,
            &parents,
            interfaces,
            &members,
            &ClassModifiers::default(),
            false,
        )
    }

    pub(super) fn compile_shared_enum_decl(
        &mut self,
        name: &str,
        interfaces: &[String],
        body_members: &[ClassMember],
        members: &[EnumMember],
        span: Span,
    ) -> Result<(), String> {
        let static_modifiers = {
            let mut modifiers = Modifiers::default();
            modifiers.is_static = true;
            modifiers
        };
        let mut synthetic_members = body_members.to_vec();
        let mut next_val = 0i64;

        for member in members {
            let (value_expr, numeric_value) = if let Some(value) = &member.value {
                if let ExprKind::Lit(Literal::Int(n)) = &value.kind {
                    next_val = *n;
                    (value.clone(), Some(*n))
                } else {
                    // Non-literal member value (e.g. `1 << 0`): the forward
                    // field still works, but we can't key a compile-time
                    // reverse entry off it, so skip its reverse map entry.
                    (value.clone(), None)
                }
            } else {
                (
                    Expression::new(ExprKind::Lit(Literal::Int(next_val))),
                    Some(next_val),
                )
            };
            // Forward entry: `Color.Red = 0`.
            synthetic_members.push(ClassMember::Field {
                name: member.name.clone(),
                type_hint: Some(name.to_string()),
                init: Some(value_expr),
                modifiers: static_modifiers.clone(),
                with_events: false,
                array_bounds: None,
            });
            // Reverse entry: `Color[0] = "Red"` — the TypeScript numeric-enum
            // shape, so `value → name` lookups (ToString/GetName) can read the
            // enum object at runtime (`EnumType[value]`) instead of relying on
            // compile-time ordinal tables. Keyed by the value's string form,
            // which an integer index resolves to.
            if let Some(nv) = numeric_value {
                synthetic_members.push(ClassMember::Field {
                    name: nv.to_string(),
                    type_hint: None,
                    init: Some(Expression::string(&member.name)),
                    modifiers: static_modifiers.clone(),
                    with_events: false,
                    array_bounds: None,
                });
            }
            next_val += 1;
        }

        self.compile_enum_decl_as_class(name, None, interfaces, synthetic_members, span)
    }

    pub(super) fn compile_dart_enum_decl(
        &mut self,
        name: &str,
        interfaces: &[String],
        body_members: &[ClassMember],
        members: &[EnumMember],
        span: Span,
    ) -> Result<(), String> {
        let mut synthetic_members = body_members.to_vec();
        let static_modifiers = {
            let mut modifiers = Modifiers::default();
            modifiers.is_static = true;
            modifiers
        };
        let mut values_array = Vec::new();

        for (index, member) in members.iter().enumerate() {
            let obj_expr = Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("index"),
                    value: Expression::new(ExprKind::Lit(Literal::Int(index as i64))),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("name"),
                    value: Expression::string(&member.name),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("__type"),
                    value: Expression::string(name),
                },
            ]));
            synthetic_members.push(ClassMember::Field {
                name: member.name.clone(),
                type_hint: None,
                init: Some(obj_expr.clone()),
                modifiers: static_modifiers.clone(),
                with_events: false,
                array_bounds: None,
            });
            values_array.push(ArrayElement {
                key: None,
                value: obj_expr,
                spread: false,
                by_ref: false,
            });
        }

        synthetic_members.push(ClassMember::Field {
            name: "values".into(),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Array(values_array))),
            modifiers: static_modifiers,
            with_events: false,
            array_bounds: None,
        });

        self.compile_enum_decl_as_class(name, Some("Enum"), interfaces, synthetic_members, span)
    }

    /// Member names a class pattern's POSITIONAL sub-patterns test, resolved at
    /// compile time — `None` means "read `__match_args__` at runtime instead".
    ///
    /// PEP 634 keys positional sub-patterns off the class's `__match_args__`.
    /// A class that declares none has nothing to read, so its CONSTRUCTOR's
    /// parameter order is the mapping — the same default a dataclass-generated
    /// `__match_args__` would carry, and the only ordering available for a
    /// plain class. Without it `case P(x, y)` silently binds `undefined`.
    /// A class that DOES declare `__match_args__` is left to the runtime
    /// lookup: only its value, not the fact of it, decides the mapping.
    fn class_pattern_positional_names(&self, cls: &Expression) -> Option<Vec<String>> {
        let ExprKind::Ident(class_name) = &cls.kind else {
            return None;
        };
        let pending = self.pending_classes.get(&self.canon(class_name))?;
        if pending
            .static_fields
            .iter()
            .any(|field| field == "__match_args__")
        {
            return None;
        }
        let ctor_key = self.canon(&self.profile.constructor_name);
        let names = pending
            .instance_method_overloads
            .get(&ctor_key)
            .and_then(|overloads| overloads.first())
            .map(|overload| overload.signature.param_names.clone())
            .or_else(|| Some(pending.fields.clone()))?;
        // Empty means "nothing resolved" — fall back to the runtime lookup
        // rather than emit a name-less read.
        (!names.is_empty()).then_some(names)
    }

    pub(super) fn emit_match_pattern_match_slot(
        &mut self,
        pattern: &Pattern,
        value_slot: u16,
    ) -> Result<u16, String> {
        let matched_slot = self.define_local("__match_pattern_ok");
        self.emit_const(Value::Bool(true));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        match pattern {
            Pattern::Value(expr) | Pattern::Singleton(expr) => {
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.compile_expr(expr)?;
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.chunk().emit_end(line);
            }
            Pattern::Sequence(items) => {
                let star_index = items
                    .iter()
                    .position(|item| matches!(item, Pattern::Star(_)));
                let suffix_count = star_index.map(|index| items.len() - index - 1).unwrap_or(0);
                let required_len = if star_index.is_some() {
                    items.len().saturating_sub(1)
                } else {
                    items.len()
                };
                let len_slot = self.define_local("__match_seq_len");
                self.emit_u16(Op::LOCAL_GET, value_slot);
                common::collections::emit_len(&mut self.chunks, self.current, self.line);
                self.emit_u16(Op::LOCAL_SET, len_slot);

                self.emit_u16(Op::LOCAL_GET, len_slot);
                self.emit_const(Value::F64(required_len as f64));
                {
                    let line = self.line;
                    if star_index.is_some() {
                        crate::primitives::ops::emit_dyn_ge(self.chunk(), line)
                    } else {
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line)
                    }
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.chunk().emit_end(line);

                if suffix_count == 0 {
                    let prefix_len = star_index.unwrap_or(items.len());
                    for (index, item) in items.iter().take(prefix_len).enumerate() {
                        let elem_slot = self.define_local("__match_seq_item");
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit_const(Value::F64(index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        let item_match_slot =
                            self.emit_match_pattern_match_slot(item, elem_slot)?;
                        self.emit_u16(Op::LOCAL_GET, item_match_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.emit(Op::I32_EQZ);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_const(Value::Bool(false));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.chunk().emit_end(line);
                    }
                }
            }
            Pattern::As {
                pattern: Some(sub_pattern),
                ..
            } => {
                let sub_match_slot = self.emit_match_pattern_match_slot(sub_pattern, value_slot)?;
                self.emit_u16(Op::LOCAL_GET, sub_match_slot);
                self.emit_u16(Op::LOCAL_SET, matched_slot);
            }
            Pattern::Or(patterns) => {
                if let Some(first) = patterns.first() {
                    let first_match_slot = self.emit_match_pattern_match_slot(first, value_slot)?;
                    self.emit_u16(Op::LOCAL_GET, first_match_slot);
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                }
            }
            // `case Point(x=1, y=2)` / `case Point(1, 2)` (PEP 634) and the C#
            // property-pattern shape `o is Point { X: 1 }` — same three tests:
            // the subject's type, then each sub-pattern against a member.
            Pattern::Class {
                cls,
                patterns,
                kw_patterns,
            } => {
                // Type test. `emit_instanceof` matches a class NAME against the
                // `__type` / `__types` stamps `emit_new_typed_object` writes, so
                // a named class goes in as its canonical name rather than as the
                // constructor value.
                self.emit_u16(Op::LOCAL_GET, value_slot);
                match &cls.kind {
                    ExprKind::Ident(class_name) => {
                        let canon = self.canon(class_name);
                        self.emit_const(Value::String(Arc::from(canon.as_str())));
                    }
                    _ => self.compile_expr(cls)?,
                }
                common::reflection::emit_instanceof(&mut self.chunks, self.current, self.line);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.chunk().emit_end(line);

                // Positional sub-patterns name their member through the class's
                // `__match_args__` (PEP 634 §3.3): the i-th pattern tests the
                // attribute `__match_args__[i]`. Languages whose patterns are
                // always keyed (C# property patterns) carry none, so this is a
                // no-op for them — the empty vec is the gate.
                if !patterns.is_empty() {
                    let declared_names = self.class_pattern_positional_names(cls);
                    let margs_slot = self.define_local("__match_args");
                    if declared_names.is_none() {
                        self.compile_expr(cls)?;
                        self.emit_const(Value::String(Arc::from("__match_args__")));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, margs_slot);
                    }
                    for (index, sub_pattern) in patterns.iter().enumerate() {
                        let attr_slot = self.define_local("__match_pos_attr");
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        match declared_names.as_ref().and_then(|names| names.get(index)) {
                            Some(field) => {
                                self.emit_const(Value::String(Arc::from(field.as_str())));
                            }
                            None => {
                                self.emit_u16(Op::LOCAL_GET, margs_slot);
                                self.emit_const(Value::F64(index as f64));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                            }
                        }
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, attr_slot);

                        let sub_match_slot =
                            self.emit_match_pattern_match_slot(sub_pattern, attr_slot)?;
                        self.emit_u16(Op::LOCAL_GET, sub_match_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.emit(Op::I32_EQZ);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_const(Value::Bool(false));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.chunk().emit_end(line);
                    }
                }

                // Keyed sub-patterns: `Point(x=1)` / `Point { X: 1 }`.
                for (member_name, sub_pattern) in kw_patterns {
                    let attr_slot = self.define_local("__match_class_attr");
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit_const(Value::String(Arc::from(member_name.as_str())));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    self.emit_u16(Op::LOCAL_SET, attr_slot);

                    let sub_match_slot =
                        self.emit_match_pattern_match_slot(sub_pattern, attr_slot)?;
                    self.emit_u16(Op::LOCAL_GET, sub_match_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    self.emit_const(Value::Bool(false));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.chunk().emit_end(line);
                }
            }
            Pattern::Wildcard
            | Pattern::Star(_)
            | Pattern::As { pattern: None, .. }
            | Pattern::Mapping(_) => {}
        }
        Ok(matched_slot)
    }

    pub(super) fn emit_match_pattern_bindings(
        &mut self,
        pattern: &Pattern,
        value_slot: u16,
    ) -> Result<(), String> {
        match pattern {
            Pattern::As { pattern, name } => {
                if let Some(sub_pattern) = pattern {
                    self.emit_match_pattern_bindings(sub_pattern, value_slot)?;
                }
                if let Some(name) = name {
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let slot = self
                        .scope()
                        .resolve(name)
                        .unwrap_or_else(|| self.define_local(name));
                    self.emit_u16(Op::LOCAL_SET, slot);
                }
            }
            Pattern::Sequence(items) => {
                let star_index = items
                    .iter()
                    .position(|item| matches!(item, Pattern::Star(_)));
                let suffix_count = star_index.map(|index| items.len() - index - 1).unwrap_or(0);

                if suffix_count == 0 {
                    let prefix_len = star_index.unwrap_or(items.len());
                    for (index, item) in items.iter().take(prefix_len).enumerate() {
                        let elem_slot = self.define_local("__match_bind_item");
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit_const(Value::F64(index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        self.emit_match_pattern_bindings(item, elem_slot)?;
                    }

                    if let Some(star_pos) = star_index {
                        if let Pattern::Star(Some(name)) = &items[star_pos] {
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_const(Value::F64(star_pos as f64));
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            common::collections::emit_len(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            common::collections::emit_slice(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            let slot = self
                                .scope()
                                .resolve(name)
                                .unwrap_or_else(|| self.define_local(name));
                            self.emit_u16(Op::LOCAL_SET, slot);
                        }
                    }
                }
            }
            Pattern::Or(patterns) => {
                if let Some(first) = patterns.first() {
                    self.emit_match_pattern_bindings(first, value_slot)?;
                }
            }
            // Mirrors the tests in `emit_match_pattern_match_slot`: a captured
            // name inside a class pattern (`case Point(x=n)`) binds from the
            // member the test read, so the member lookups repeat here.
            Pattern::Class {
                cls,
                patterns,
                kw_patterns,
            } => {
                if !patterns.is_empty() {
                    let declared_names = self.class_pattern_positional_names(cls);
                    let margs_slot = self.define_local("__match_args_bind");
                    if declared_names.is_none() {
                        self.compile_expr(cls)?;
                        self.emit_const(Value::String(Arc::from("__match_args__")));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, margs_slot);
                    }
                    for (index, sub_pattern) in patterns.iter().enumerate() {
                        let attr_slot = self.define_local("__match_pos_attr_bind");
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        match declared_names.as_ref().and_then(|names| names.get(index)) {
                            Some(field) => {
                                self.emit_const(Value::String(Arc::from(field.as_str())));
                            }
                            None => {
                                self.emit_u16(Op::LOCAL_GET, margs_slot);
                                self.emit_const(Value::F64(index as f64));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                            }
                        }
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, attr_slot);
                        self.emit_match_pattern_bindings(sub_pattern, attr_slot)?;
                    }
                }
                for (member_name, sub_pattern) in kw_patterns {
                    let attr_slot = self.define_local("__match_class_attr_bind");
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit_const(Value::String(Arc::from(member_name.as_str())));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    self.emit_u16(Op::LOCAL_SET, attr_slot);
                    self.emit_match_pattern_bindings(sub_pattern, attr_slot)?;
                }
            }
            Pattern::Value(_)
            | Pattern::Singleton(_)
            | Pattern::Wildcard
            | Pattern::Star(_)
            | Pattern::Mapping(_) => {}
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Variable declarator compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_var_declarator(
        &mut self,
        decl: &VarDeclarator,
        kind: &VarDeclKind,
    ) -> Result<(), String> {
        match &decl.pattern {
            BindingPattern::Ident(name) => {
                let reflection_binding = decl
                    .init
                    .as_ref()
                    .and_then(|expr| self.resolve_reflection_binding_expr(expr));
                let init_type_hint = decl
                    .init
                    .as_ref()
                    .and_then(|expr| self.infer_expr_type_hint(expr));
                let declared_type_hint = decl.type_hint.clone();
                let mut inferred_type_hint = declared_type_hint
                    .clone()
                    .or_else(|| init_type_hint.clone());

                // VB often spells dynamically-created controls as `As Object`
                // even though the initializer is a concrete dotnet wrapper such
                // as `Window.Forms.Button()`. Keep that concrete wrapper type so
                // later lowering (`AddHandler`, instance method dispatch, etc.)
                // stays on the same WinForms adapter path as designer forms.
                if self.profile.namespaces.use_dotnet {
                    let declared_is_object = declared_type_hint
                        .as_deref()
                        .map(|type_hint| self.resolve_source_type_alias(type_hint))
                        .map(|type_hint| {
                            matches!(
                                Self::normalize_type_hint(&type_hint).as_str(),
                                "object" | "system.object"
                            )
                        })
                        .unwrap_or(false);
                    if declared_is_object {
                        if let Some(init_type_hint) = init_type_hint.as_deref() {
                            let resolved_init = self.resolve_source_type_alias(init_type_hint);
                            if self
                                .resolve_pending_class_name_for_type_hint(&resolved_init)
                                .is_some()
                            {
                                inferred_type_hint = Some(resolved_init);
                            }
                        }
                    }
                }
                let resolved_type_hint = inferred_type_hint
                    .as_deref()
                    .map(|type_hint| self.resolve_source_type_alias(type_hint));
                if decl.array_bounds.is_some() {
                    if let Some(type_hint) = inferred_type_hint.as_mut() {
                        if !type_hint.trim_end().ends_with("()") {
                            type_hint.push_str("()");
                        }
                    }
                }
                let is_pascal_type_alias_decl = self.profile.name == "pascal"
                    && *kind == VarDeclKind::Const
                    && decl.init.is_none()
                    && decl.array_bounds.is_none()
                    && self.scopes.len() == 1
                    && self.scope().depth == 0;
                if is_pascal_type_alias_decl {
                    if let Some(type_hint) =
                        resolved_type_hint.as_deref().or(decl.type_hint.as_deref())
                    {
                        self.source_type_aliases
                            .insert(self.canon(name), type_hint.to_string());
                    }
                    return Ok(());
                }
                if inferred_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim().ends_with("()"))
                    || decl.array_bounds.is_some()
                    || resolved_type_hint.as_deref().is_some_and(|type_hint| {
                        self.pascal_array_type_hint_metadata(type_hint).is_some()
                    })
                {
                    let array_type_hint = resolved_type_hint
                        .clone()
                        .or_else(|| inferred_type_hint.clone());
                    let pascal_bounds = array_type_hint
                        .as_deref()
                        .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint));
                    let is_fixed = decl
                        .array_bounds
                        .as_ref()
                        .is_some_and(|bounds| !bounds.is_empty())
                        || resolved_type_hint
                            .as_deref()
                            .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint))
                            .is_some_and(|metadata| metadata.is_fixed);
                    self.record_array_binding(
                        name,
                        ArrayBindingMetadata {
                            is_fixed,
                            type_hint: array_type_hint,
                            pascal_bounds,
                        },
                    );
                }
                // Top-level vars → globals.
                // `let`/`const` inside a block scope (depth > 0) are locals
                // even at the top level — they respect block scoping.
                // ECMA-262 §10.2.11: `var` inside a function is function-
                // scoped (a local), only script-level `var` is global.
                let is_toplevel = self.scopes.len() == 1 && self.scope().depth == 0;
                let is_hoisted =
                    *kind == VarDeclKind::Var && self.profile.hoist_var && self.scopes.len() == 1;

                if *kind == VarDeclKind::Static {
                    let binding =
                        self.ensure_static_local_binding(name, inferred_type_hint.clone())?;
                    let flag_idx = self.str_const(&binding.init_flag_name);
                    self.emit_u16(Op::GLOBAL_GET, flag_idx);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_var_decl_initializer_value(decl, resolved_type_hint.as_deref())?;
                    let value_idx = self.str_const(&binding.global_name);
                    self.emit_u16(Op::GLOBAL_SET, value_idx);
                    inst!(self, core_wasm::bool_const, true);
                    self.emit_u16(Op::GLOBAL_SET, flag_idx);
                    self.chunk().emit_end(line);

                    let binding_key = self.canon(name);
                    if let Some(binding) = reflection_binding {
                        self.reflection_bindings.insert(binding_key, binding);
                    } else {
                        self.reflection_bindings.remove(&binding_key);
                    }
                    return Ok(());
                }

                // Recursive local lambdas need their binding slot defined
                // before compiling the initializer so captures resolve to the
                // enclosing local rather than an unresolved global.
                let mut predeclared_local_slot: Option<u16> = None;
                if !is_toplevel && !is_hoisted {
                    if let Some(init_expr) = decl.init.as_ref() {
                        let recursive_lambda_init = matches!(
                            init_expr.kind,
                            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_)
                        );
                        if recursive_lambda_init {
                            let slot = if *kind == VarDeclKind::Var && self.profile.hoist_var {
                                self.scope_mut()
                                    .define_at_function_scope(name, inferred_type_hint.clone())
                            } else {
                                self.define_local_typed(name, inferred_type_hint.clone())
                            };
                            self.emit(Op::NULL);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            predeclared_local_slot = Some(slot);
                        }
                    }
                }

                if let Some(ref init_expr) = decl.init {
                    self.compile_expr_with_value_copy(init_expr)?;
                    let effective_type_hint =
                        resolved_type_hint.as_deref().or(decl.type_hint.as_deref());
                    let skip_c_coerce = if self.profile.name == "c" {
                        let is_array_type = effective_type_hint
                            .map(|hint| hint.contains('['))
                            .unwrap_or(false)
                            || decl.array_bounds.is_some();
                        let is_char_string_init =
                            matches!(init_expr.kind, ExprKind::Lit(Literal::Str(_)))
                                && effective_type_hint
                                    .map(|hint| {
                                        let lower = hint.to_ascii_lowercase();
                                        lower.contains("char")
                                    })
                                    .unwrap_or(false);
                        is_array_type || is_char_string_init
                    } else {
                        false
                    };
                    if !skip_c_coerce {
                        self.coerce_c_value_for_type_hint(effective_type_hint)?;
                    }
                    self.maybe_promote_pascal_array_literal_to_set(
                        decl.type_hint.as_deref(),
                        init_expr,
                    );
                    // ECMA-262 §10.2.9 SetFunctionName — anonymous
                    // function expressions assigned to a binding take
                    // the binding name as their `name` property.
                    // Covers `const f = () => x` / `const f = function() {}`.
                    if self.profile.ecma_lexical_declarations {
                        let should_infer_name = match &init_expr.kind {
                            ExprKind::Lambda { .. } => true,
                            ExprKind::FunctionExpr(stmt) => {
                                matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name.is_empty())
                            }
                            ExprKind::ClassExpr { name, .. } => name.is_none(),
                            _ => false,
                        };
                        if should_infer_name {
                            let line = self.line;
                            inst!(self, core_wasm::dup);
                            self.emit_const(Value::String(Arc::from(name.as_str())));
                            let name_key = self.str_const("name");
                            self.chunk().emit_op_u16(Op::STRUCT_SET, name_key, line);
                            self.emit(Op::DROP);
                        }
                    }
                } else if decl.array_bounds.is_some() || decl.type_hint.is_some() {
                    self.emit_var_decl_initializer_value(decl, resolved_type_hint.as_deref())?;
                } else {
                    self.emit(Op::NULL);
                }

                if is_toplevel || is_hoisted {
                    let cn = self.canon(name);
                    let idx = self.str_const(&cn);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    if self.profile.ecma_lexical_declarations
                        && *kind == VarDeclKind::Var
                        && is_toplevel
                    {
                        let global_this_key = self.str_const("globalThis");
                        let field_key = self.str_const(&cn);
                        self.emit_u16(Op::GLOBAL_GET, global_this_key);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        self.emit_u16(Op::STRUCT_SET, field_key);
                        self.emit(Op::DROP);
                    }
                    if let Some(type_hint) = inferred_type_hint.as_deref() {
                        self.global_type_hints
                            .insert(cn.clone(), Self::normalize_type_hint(type_hint));
                    }
                    if *kind == VarDeclKind::Const && self.profile.ecma_lexical_declarations {
                        self.const_globals.insert(cn.clone());
                    }
                    // Record the global's constant init value (if it is a
                    // constant expression) so extended-const data/elem offsets
                    // can resolve `global.get $g` at compile time (WASM). Keyed
                    // by the declared name to match the offset's `Ident`.
                    if let Some(init) = &decl.init {
                        if let Some(v) = const_eval_i128(init, &self.global_const_values) {
                            self.global_const_values.insert(name.to_string(), v as i64);
                        }
                    }
                    self.defined_globals.insert(cn);
                } else {
                    // ECMA-262 §10.2.11: `var` is function-scoped (must
                    // survive enclosing-block exits). `let` / `const`
                    // are block-scoped. The scope helper picks the right
                    // depth based on the kind.
                    let slot = if let Some(slot) = predeclared_local_slot {
                        slot
                    } else if *kind == VarDeclKind::Var && self.profile.hoist_var {
                        self.scope_mut()
                            .define_at_function_scope(name, inferred_type_hint.clone())
                    } else {
                        self.define_local_typed(name, inferred_type_hint.clone())
                    };
                    if *kind == VarDeclKind::Const && self.profile.ecma_lexical_declarations {
                        self.scope_mut().mark_const(slot);
                    }
                    self.emit_u16(Op::LOCAL_SET, slot);
                    // If this local is captured by inner closures, also store
                    // the initial value in the shared env array so closures
                    // see the same value.
                    if let (Some(env_slot), Some(idx)) =
                        (self.shared_env_slot, self.shared_env_index(name))
                    {
                        let l = self.line;
                        self.emit_u16(Op::LOCAL_GET, slot);
                        crate::primitives::closures::emit_env_set(self.chunk(), env_slot, idx, l);
                    }
                    // If this local's address is taken anywhere in the function,
                    // box it in a pointer cell now (once), so a `&name` inside a
                    // loop reuses this cell rather than re-wrapping every
                    // iteration. Reads/writes become cell-aware via the mark.
                    if self.current_addr_taken_locals.contains(name) {
                        self.promote_local_binding_to_pointer_cell(name);
                    }
                }

                let binding_key = self.canon(name);
                if let Some(binding) = reflection_binding {
                    self.reflection_bindings.insert(binding_key, binding);
                } else {
                    self.reflection_bindings.remove(&binding_key);
                }
            }
            BindingPattern::Object(_) | BindingPattern::Array(_) => {
                // Destructuring `let { a, b } = expr` / `let [a, b] = expr`.
                // Compile RHS, then recursively bind via the helper so
                // arbitrary nesting (`{ a: { b: { c } } }`) works.
                if let Some(ref init_expr) = decl.init {
                    self.compile_expr(init_expr)?;
                    self.compile_destructure_bind(&decl.pattern)?;
                }
            }
        }
        Ok(())
    }

    /// Recursively bind a `BindingPattern` from a value on TOS. Consumes
    /// the value. Used by `let { a: { b: { c } } } = ...` and friends.
    /// Defines locals at every leaf ident — call sites must already be
    /// in the right scope.
    pub(super) fn compile_destructure_bind(
        &mut self,
        pattern: &BindingPattern,
    ) -> Result<(), String> {
        match pattern {
            BindingPattern::Ident(name) => {
                let slot = self.define_local(name);
                self.emit_u16(Op::LOCAL_SET, slot);
                if let (Some(env_slot), Some(idx)) =
                    (self.shared_env_slot, self.shared_env_index(name))
                {
                    let l = self.line;
                    self.emit_u16(Op::LOCAL_GET, slot);
                    crate::primitives::closures::emit_env_set(self.chunk(), env_slot, idx, l);
                }
            }
            BindingPattern::Object(props) => {
                let obj_slot = self.define_local("__destruct_obj");
                self.emit_u16(Op::LOCAL_SET, obj_slot);
                // Collect non-rest named keys for rest exclusion.
                let named_keys: Vec<String> = props
                    .iter()
                    .filter(|p| !p.is_rest)
                    .map(|p| p.key.clone())
                    .collect();
                for prop in props {
                    if prop.is_rest {
                        // Build rest = Object.assign({}, src) then delete named keys.
                        let new_idx = self.import("ecma:object", "new");
                        self.emit_host_call(new_idx, 0);
                        let rest_slot = self.define_local("__rest_obj");
                        self.emit_u16(Op::LOCAL_SET, rest_slot);
                        self.emit_u16(Op::LOCAL_GET, rest_slot);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        let assign_idx = self.import("ecma:object", "assign");
                        self.emit_host_call(assign_idx, 2);
                        self.emit(Op::DROP); // drop assign's return (target already in slot)
                        for named in &named_keys {
                            self.emit_u16(Op::LOCAL_GET, rest_slot);
                            self.emit_const(Value::String(Arc::from(named.as_str())));
                            let del_idx = self.import("ecma:object", "delete");
                            self.emit_host_call(del_idx, 2);
                            self.emit(Op::DROP); // drop bool result
                        }
                        self.emit_u16(Op::LOCAL_GET, rest_slot);
                        let rest_var_slot = self.define_local(&prop.key);
                        self.emit_u16(Op::LOCAL_SET, rest_var_slot);
                        continue;
                    }
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_const(Value::String(Arc::from(prop.key.as_str())));
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    // JS: if own-property lookup returned Undefined, the key may live
                    // on the prototype chain — fall back to ecma:object.get which walks
                    // __proto__. We don't use ecma:object.get unconditionally because
                    // CALL_IMPORT triggers the JSPI auto-check: if the own value happens
                    // to be a pending Promise (e.g. Promise.withResolvers destructuring),
                    // the fiber would be suspended before the await expression even runs.
                    if self.profile.async_wraps_body_in_try {
                        let value_slot = self.define_local("__destruct_prop_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if_value(line);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_const(Value::String(Arc::from(prop.key.as_str())));
                        let get_idx = self.import("ecma:object", "get");
                        self.emit_host_call(get_idx, 2);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.chunk().emit_end(line);
                    }
                    if let Some(ref default) = prop.default {
                        let value_slot = self.define_local("__destruct_default_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.compile_expr(default)?;
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.chunk().emit_end(line);
                    }
                    let target = match &prop.value {
                        Some(p) => p.clone(),
                        None => BindingPattern::Ident(prop.key.clone()),
                    };
                    self.compile_destructure_bind(&target)?;
                }
            }
            BindingPattern::Array(elems) => {
                // JS profile: if the value is a generator (ObjectKind::
                // Continuation, e.g. `let [a,b] = gen()`), drain it via
                // the WASM stack-switching `__stdlib_drain_generator`
                // helper into a real Array first. ARRAY_GET on a
                // Continuation returns undefined otherwise.
                if self.profile.supports_spread_arguments {
                    common::collections::emit_spread_iterable(
                        &mut self.chunks,
                        self.current,
                        self.line,
                    );
                }
                let arr_slot = self.define_local("__destruct_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                for (i, elem) in elems.iter().enumerate() {
                    match elem {
                        ArrayPatternElem::Pattern(pat, default) => {
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_const(Value::F64(i as f64));
                            {
                                let l = self.line;
                                common::collections::emit_get(&mut self.chunks, self.current, l);
                            }
                            if let Some(def) = default {
                                let value_slot = self.define_local("__destruct_default_value");
                                self.emit_u16(Op::LOCAL_SET, value_slot);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                fn_call!(self, "wasm:js-undefined", "test", 1);
                                let line = self.line;
                                self.chunk().emit_if_value(line);
                                self.compile_expr(def)?;
                                self.chunk().emit_else(line);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                self.chunk().emit_end(line);
                            }
                            self.compile_destructure_bind(pat)?;
                        }
                        ArrayPatternElem::Rest(name) => {
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_const(Value::F64(i as f64));
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            {
                                let l = self.line;
                                common::collections::emit_len(&mut self.chunks, self.current, l);
                            }
                            let line = self.line;
                            common::collections::emit_slice(&mut self.chunks, self.current, line);
                            let slot = self.define_local(name);
                            self.emit_u16(Op::LOCAL_SET, slot);
                        }
                        ArrayPatternElem::Hole => {}
                    }
                }
            }
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Assignment target
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_array_pattern_assignment_from_slot(
        &mut self,
        arr_slot: u16,
        elems: &[ArrayPatternElem],
    ) -> Result<(), String> {
        for (i, elem) in elems.iter().enumerate() {
            match elem {
                ArrayPatternElem::Pattern(BindingPattern::Ident(name), _) => {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    self.emit_var_set(name);
                }
                ArrayPatternElem::Pattern(BindingPattern::Array(items), _) => {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    let nested_slot = self.define_local("__destruct_nested_arr");
                    self.emit_u16(Op::LOCAL_SET, nested_slot);
                    self.compile_array_pattern_assignment_from_slot(nested_slot, items)?;
                }
                ArrayPatternElem::Pattern(BindingPattern::Object(_), _) => {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    if let ArrayPatternElem::Pattern(pattern, _) = elem {
                        self.compile_destructure_bind(pattern)?;
                    }
                }
                ArrayPatternElem::Rest(name) => {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    {
                        let l = self.line;
                        common::collections::emit_len(&mut self.chunks, self.current, l);
                    }
                    let line = self.line;
                    common::collections::emit_slice(&mut self.chunks, self.current, line);
                    self.emit_var_set(name);
                }
                ArrayPatternElem::Hole => {}
            }
        }
        Ok(())
    }

    pub(super) fn compile_assign_target(&mut self, target: &Expression) -> Result<(), String> {
        match &target.kind {
            ExprKind::Ident(name) => {
                // FuncName := value assigns to Result slot (Pascal/VB)
                if let Some(ref fn_name) = self.current_func_name.clone() {
                    let matches = if self.case_sensitive {
                        name == fn_name
                    } else {
                        name.eq_ignore_ascii_case(fn_name)
                    };
                    if matches {
                        if let Some(rs) = self.current_result_slot {
                            self.emit_u16(Op::LOCAL_SET, rs);
                            return Ok(());
                        }
                    }
                }
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                    || self.has_static_local_binding(name);

                // Implicit self field write (only if NOT a local)
                if !is_local && self.is_class_field(name) {
                    let tmp = self.define_local("__field_tmp");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    if self.emit_self_ref() {
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        let field_name = self
                            .current_class
                            .as_deref()
                            .and_then(|class_name| {
                                self.visible_instance_field_storage_name_for_class(class_name, name)
                            })
                            .unwrap_or_else(|| self.canon(name));
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);
                        return Ok(());
                    }
                    self.emit_u16(Op::LOCAL_GET, tmp);
                }
                let stored_type_hint = self.lookup_var_type_hint(name).map(str::to_string);
                self.coerce_c_value_for_type_hint(stored_type_hint.as_deref())?;
                self.emit_var_set(name);
            }
            ExprKind::StaticAccess { class, member } => {
                let value_tmp = self.define_local("__static_access_value");
                self.emit_u16(Op::LOCAL_SET, value_tmp);

                self.compile_expr(class)?;
                let class_tmp = self.define_local("__static_access_class");
                self.emit_u16(Op::LOCAL_SET, class_tmp);

                if let ExprKind::Ident(name) = &member.kind {
                    if self.private_member_access_forbidden(name) {
                        self.emit_private_access_denied(name)?;
                        return Ok(());
                    }
                    let field_name = match &class.kind {
                        ExprKind::Ident(class_name) => {
                            self.js_member_storage_name_for_class(class_name, name)
                        }
                        _ => self.canon(name),
                    };
                    if self.profile.supports_private_fields && name.starts_with('#') {
                        let setter_name = format!("__set_{}", field_name);
                        self.emit_u16(Op::LOCAL_GET, class_tmp);
                        self.emit_const(Value::String(Arc::from(setter_name.as_str())));
                        // `has` (proto-walk, raw key) not `hasOwn`: the private
                        // accessor key is `__get_/__set___js_private_*` — a `__`
                        // key that `hasOwn` hides, and under prototype dispatch the
                        // accessor lives on the class prototype, not the instance.
                        let has_own_idx = self.import("ecma:object", "has");
                        let line = self.line;
                        self.emit_host_call(has_own_idx, 2);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);

                        self.emit_u16(Op::LOCAL_GET, class_tmp);
                        let setter_key = self.str_const(&setter_name);
                        self.emit_u16(Op::STRUCT_GET, setter_key);
                        self.emit_u16(Op::LOCAL_GET, class_tmp);
                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                        self.emit_u8(Op::CALL_REF, 2);
                        self.emit(Op::DROP);

                        self.chunk().emit_else(line);
                        self.emit_js_private_brand_check(class_tmp, &field_name)?;
                        self.emit_u16(Op::LOCAL_GET, class_tmp);
                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, class_tmp);
                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);
                    }
                } else {
                    self.emit_u16(Op::LOCAL_GET, class_tmp);
                    self.compile_expr(member)?;
                    self.emit_u16(Op::LOCAL_GET, value_tmp);
                    let line = self.line;
                    common::collections::emit_set(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                }
            }
            ExprKind::Member { object, field, .. } => {
                if self.private_member_access_forbidden(field) {
                    self.emit_private_access_denied(field)?;
                    return Ok(());
                }
                if self.class_prototype_dispatch()
                    && self.profile.ecma_object_literals
                    && matches!(&object.kind, ExprKind::Super)
                {
                    let value_tmp = self.define_local("__js_super_set_value");
                    self.emit_u16(Op::LOCAL_SET, value_tmp);

                    let receiver_tmp = self.define_local("__js_super_set_receiver");
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self
                        .scope()
                        .resolve(&self_kw)
                        .or_else(|| self.scope().resolve_ci(&self_kw))
                    {
                        self.emit_u16(Op::LOCAL_GET, slot);
                    } else {
                        let js_this = self.str_const("__js_this");
                        self.emit_u16(Op::GLOBAL_GET, js_this);
                    }
                    self.emit_u16(Op::LOCAL_SET, receiver_tmp);

                    self.emit_js_super_home_base();
                    let setter_key = self.str_const(&format!("__set_{}", field));
                    self.emit_u16(Op::STRUCT_GET, setter_key);
                    let setter_tmp = self.define_local("__js_super_setter");
                    self.emit_u16(Op::LOCAL_SET, setter_tmp);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, receiver_tmp);
                    self.emit_u16(Op::LOCAL_GET, value_tmp);
                    let field_idx = self.str_const(field);
                    self.emit_u16(Op::STRUCT_SET, field_idx);
                    self.emit(Op::DROP);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit_u16(Op::LOCAL_GET, receiver_tmp);
                    self.emit_u16(Op::LOCAL_GET, value_tmp);
                    self.emit_u8(Op::CALL_REF, 2);
                    self.emit(Op::DROP);
                    self.chunk().emit_end(line);
                    return Ok(());
                }
                // .NET control property write resolves through the component
                // descriptor to a direct `vybe:gui` host call — no emitted
                // accessor. Stack on entry is [value]. The generic property
                // setter takes (this, "Key", value); dedicated setters (this,
                // value).
                if self.profile.namespaces.use_dotnet
                    && !self.expr_user_value_type_name(object).is_some()
                {
                    if let Some(type_hint) = self.infer_expr_type_hint(object) {
                        let class_name = Self::normalize_type_hint(&type_hint);
                        if let Some(target) =
                            vybe_bytecode::namespaces::lookup_type_property_setter_target(
                                &self.profile.namespaces.type_scopes,
                                &class_name,
                                field,
                            )
                        {
                            match target {
                                vybe_bytecode::component_model::InstancePropertyTarget::Host {
                                    module,
                                    func,
                                    key,
                                } => {
                                    let value_tmp = self.define_local("__dotnet_prop_value");
                                    self.emit_u16(Op::LOCAL_SET, value_tmp);
                                    self.compile_expr(object)?;
                                    let idx = self.import(&module, &func);
                                    if let Some(key) = key {
                                        self.emit_const(Value::String(Arc::from(key.as_str())));
                                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                                        self.emit_host_call(idx, 3);
                                    } else {
                                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                                        self.emit_host_call(idx, 2);
                                    }
                                    self.emit(Op::DROP);
                                    return Ok(());
                                }
                                vybe_bytecode::component_model::InstancePropertyTarget::Common { emit } => {
                                    let value_tmp = self.define_local("__dotnet_prop_value");
                                    self.emit_u16(Op::LOCAL_SET, value_tmp);
                                    self.compile_expr(object)?;
                                    self.emit_u16(Op::LOCAL_GET, value_tmp);
                                    self.emit_common(&emit, 2, self.line);
                                    self.emit(Op::DROP);
                                    return Ok(());
                                }
                            }
                        }
                    }
                }
                if let ExprKind::Ident(obj_name) = &object.kind {
                    if let Some(key) = self.generic_static_member_key(obj_name, field) {
                        let tmp = self.define_local("__tmp");
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        let idx = self.str_const(&key);
                        self.emit_u16(Op::GLOBAL_SET, idx);
                        return Ok(());
                    }

                    let needs_value_type_writeback =
                        self.expr_user_value_type_name(object).is_some()
                            || (self.profile.name == "fortran"
                                && self
                                    .lookup_var_type_hint(obj_name)
                                    .and_then(|type_hint| {
                                        self.resolve_pending_class_name_for_type_hint(type_hint)
                                    })
                                    .is_some());
                    if needs_value_type_writeback {
                        let value_tmp = self.define_local("__tmp");
                        let obj_tmp = self.define_local("__value_type_member_obj");
                        self.emit_u16(Op::LOCAL_SET, value_tmp);

                        self.compile_expr(object)?;
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);

                        let field_name = self
                            .field_storage_name_for_receiver(object, field)
                            .unwrap_or_else(|| self.js_member_storage_name(field));
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_var_set(obj_name);
                        return Ok(());
                    }
                }

                // Proxy set-trap dispatch (JS profile, only when the
                // module references `Proxy`). Stack on entry is [value]
                // (caller pushed it); the dispatcher needs [obj, key,
                // value] so we re-stash, push obj + key string, reload
                // value, then call.
                if self.uses_proxy {
                    let tmp = self.define_local("__proxy_set_v");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let line = self.line;
                    if self.in_strict {
                        vybe_bytecode::registry::hooks(&self.profile.name)
                            .proxy_set_bool
                            .unwrap()(&mut self.chunks, self.current, line);
                        self.emit_strict_set_failure_check()?;
                    } else {
                        vybe_bytecode::registry::hooks(&self.profile.name)
                            .proxy_set
                            .unwrap()(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP); // adapter leaves [value] on stack
                    }
                    return Ok(());
                }
                let tmp = self.define_local("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                let field_name = self
                    .field_storage_name_for_receiver(object, field)
                    .unwrap_or_else(|| self.js_member_storage_name(field));
                if self.profile.name == "fortran" {
                    if let ExprKind::Index {
                        object: collection_owner,
                        index,
                        ..
                    } = &object.kind
                    {
                        let line = self.line;
                        let coll_tmp = self.define_local("__fortran_index_member_coll");
                        let key_tmp = self.define_local("__fortran_index_member_key");
                        let elem_tmp = self.define_local("__fortran_index_member_elem");
                        let field_idx = self.str_const(&field_name);

                        self.compile_expr(collection_owner)?;
                        self.emit_u16(Op::LOCAL_SET, coll_tmp);

                        self.compile_array_index_operand_for_owner(collection_owner, index)?;
                        self.emit_u16(Op::LOCAL_SET, key_tmp);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, elem_tmp);

                        self.emit_u16(Op::LOCAL_GET, elem_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        self.emit_u16(Op::STRUCT_SET, field_idx);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, elem_tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.compile_assign_target(collection_owner)?;
                        return Ok(());
                    }
                }
                if self.profile.supports_private_fields && field.starts_with('#') {
                    self.compile_expr(object)?;
                    let obj_tmp = self.define_local("__js_private_set_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_tmp);

                    let setter_name = format!("__set_{}", field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(setter_name.as_str())));
                    // `has` (proto-walk, raw key) not `hasOwn`: the private
                    // accessor key is `__get_/__set___js_private_*` — a `__`
                    // key that `hasOwn` hides, and under prototype dispatch the
                    // accessor lives on the class prototype, not the instance.
                    let has_own_idx = self.import("ecma:object", "has");
                    let line = self.line;
                    self.emit_host_call(has_own_idx, 2);
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let setter_key = self.str_const(&setter_name);
                    self.emit_u16(Op::STRUCT_GET, setter_key);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u8(Op::CALL_REF, 2);
                    self.emit(Op::DROP);

                    self.chunk().emit_else(line);

                    let getter_name = format!("__get_{}", field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(getter_name.as_str())));
                    // `has` (proto-walk, raw key) not `hasOwn`: the private
                    // accessor key is `__get_/__set___js_private_*` — a `__`
                    // key that `hasOwn` hides, and under prototype dispatch the
                    // accessor lives on the class prototype, not the instance.
                    let has_own_idx = self.import("ecma:object", "has");
                    self.emit_host_call(has_own_idx, 2);
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::String(Arc::from(
                        "Cannot set private accessor without a setter",
                    )));
                    self.emit_js_exception_ctor_from_message_value("TypeError")?;
                    common::errors::emit_throw(self.chunk(), line);
                    self.chunk().emit_else(line);

                    self.emit_js_private_brand_check(obj_tmp, &field_name)?;
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_SET, idx);
                    self.emit(Op::DROP);

                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                    return Ok(());
                }
                if self.profile.namespaces.use_dotnet {
                    self.compile_expr(object)?;
                    if !Self::is_pointer_runtime_field(field) {
                        self.emit_autoderef_pointer_cell();
                    }
                    let obj_tmp = self.define_local("__member_set_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_tmp);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let setter_key = self.str_const(&format!("__set_{}", field_name));
                    self.emit_u16(Op::STRUCT_GET, setter_key);
                    let setter_tmp = self.define_local("__member_setter");
                    self.emit_u16(Op::LOCAL_SET, setter_tmp);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_SET, idx);
                    self.emit(Op::DROP);

                    self.chunk().emit_else(line);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u8(Op::CALL_REF, 2);
                    self.emit(Op::DROP);

                    self.chunk().emit_end(line);
                    return Ok(());
                }

                self.compile_expr(object)?;
                if !Self::is_pointer_runtime_field(field) {
                    self.emit_autoderef_pointer_cell();
                }
                // JS `Object.keys` / `Object.entries` need insertion order
                // (ECMA-262 §7.3.22). The HashMap backing properties is
                // non-deterministic, so we mirror each direct write into
                // `__keys` via the host trackKey helper. Only fires for
                // JS — other languages don't promise insertion order or
                // pay the host-call overhead.
                if self.profile.ecma_object_literals && !field_name.starts_with("__") {
                    let line = self.line;
                    inst!(self, core_wasm::dup);
                    self.emit_const(Value::String(Arc::from(field_name.as_str())));
                    let track_idx = self.import("ecma:object", "trackKey");
                    self.chunk().emit_call(track_idx, 2, line);
                    self.emit(Op::DROP);
                }
                // JS profile member writes route through `ecma:object.set`
                // for ECMA-262 §10.1.5 OrdinarySet enforcement: frozen /
                // sealed / preventExtensions gates + `__set_<key>`
                // accessor dispatch in one place. Internal `__*` keys
                // bypass — VM bookkeeping (proxy, prototype, type stamps)
                // that the gates would block.
                if self.profile.ecma_object_literals && !field_name.starts_with("__") {
                    // Bind `__js_this = obj` so a setter installed by
                    // `Object.defineProperty` (arity-1 `set(val)`) sees
                    // the receiver via the JS method-call protocol.
                    // Stack on entry: [obj]. Stash, set __js_this,
                    // re-push, call, restore.
                    let line = self.line;
                    let obj_slot = self.define_local("__js_set_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    let saved_this = self.save_js_this("__js_prev_this_set");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_const(Value::String(Arc::from(field_name.as_str())));
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let set_idx = self.import("ecma:object", "set");
                    if self.in_strict {
                        // §13.15.2: strict assignment failures throw — the
                        // host's OrdinarySet gates read this 4th arg.
                        inst!(self, core_wasm::bool_const, true);
                        self.chunk().emit_call(set_idx, 4, line);
                    } else {
                        self.chunk().emit_call(set_idx, 3, line);
                    }
                    self.emit(Op::DROP);
                    self.restore_js_this(saved_this);
                } else {
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_SET, idx);
                    self.emit(Op::DROP);
                }
                // globalThis.X = val also sets X in module global scope
                // so bare `X` references resolve (§19.3 global object semantics).
                if matches!(&object.kind, ExprKind::Ident(n) if n == "globalThis") {
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let g_idx = self.str_const(&field_name);
                    self.emit_u16(Op::GLOBAL_SET, g_idx);
                }
            }
            ExprKind::Unary {
                op: UnaryOp::Deref,
                expr,
            }
            | ExprKind::RefLoad(expr) => {
                let value_slot = self.define_local("__ref_store_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.compile_expr(expr)?;
                let ptr_slot = self.define_local("__ref_store_ptr");
                self.emit_u16(Op::LOCAL_SET, ptr_slot);

                self.emit_u16(Op::LOCAL_GET, ptr_slot);
                inst!(self, recipes::is_object);
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                let kind_key = self.str_const("__ref_kind");

                self.emit_u16(Op::LOCAL_GET, ptr_slot);
                self.emit_u16(Op::STRUCT_GET, kind_key);
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
                self.emit_u16(Op::STRUCT_GET, kind_key);
                self.emit_const(Value::String(Arc::from("carray")));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                }
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                let base_key = self.str_const("__base");
                let idx_key = self.str_const("__idx");
                let base_slot = self.define_local("__ref_store_carray_base");
                let idx_slot = self.define_local("__ref_store_carray_idx");

                self.emit_u16(Op::LOCAL_GET, ptr_slot);
                self.emit_u16(Op::STRUCT_GET, base_key);
                self.emit_u16(Op::LOCAL_SET, base_slot);

                self.emit_u16(Op::LOCAL_GET, ptr_slot);
                self.emit_u16(Op::STRUCT_GET, idx_key);
                self.emit_u16(Op::LOCAL_SET, idx_slot);

                self.emit_u16(Op::LOCAL_GET, base_slot);
                inst!(self, recipes::is_object);
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, base_slot);
                self.emit_u16(Op::STRUCT_GET, kind_key);
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
                self.chunk().emit_end(line);

                let line = self.line;
                self.chunk().emit_end(line);

                let line = self.line;
                self.chunk().emit_else(line);
                self.chunk().emit_end(line);
            }
            ExprKind::Index { object, index, .. } => {
                // A class declaring `operator []=` / `__setitem__` makes
                // `x[i] = v` a call to it. Resolved from the receiver's
                // declared type, so array/dict/string stores are untouched.
                // Stack on entry: `[value]`.
                if !matches!(index.kind, ExprKind::Range { .. } | ExprKind::Slice { .. })
                    && self.expr_has_user_index_setter(object)
                {
                    let line = self.line;
                    let value_slot = self.define_local("__idx_set_value");
                    let obj_slot = self.define_local("__idx_set_recv");
                    let key_slot = self.define_local("__idx_set_key");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.compile_expr(index)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                    let setter = self.str_const(&vybe_ast::protocol_slot_key(
                        vybe_ast::ProtocolSlot::SetItem,
                    ));
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::STRUCT_GET, setter);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_op_u8(Op::CALL_REF, 3, line);
                    self.emit(Op::DROP);
                    return Ok(());
                }
                if self.profile.name == "fortran" {
                    if let ExprKind::Slice { lower, upper, step } = &index.kind {
                        if step.is_none() {
                            let line = self.line;
                            let value_tmp = self.define_local("__fortran_slice_value");
                            let obj_tmp = self.define_local("__fortran_slice_obj");
                            let start_tmp = self.define_local("__fortran_slice_start");
                            let end_tmp = self.define_local("__fortran_slice_end");
                            let count_tmp = self.define_local("__fortran_slice_count");
                            let replacement_tmp = self.define_local("__fortran_slice_replacement");
                            let string_value_tmp =
                                self.define_local("__fortran_slice_string_value");

                            self.emit_u16(Op::LOCAL_SET, value_tmp);

                            self.compile_expr(object)?;
                            self.emit_u16(Op::LOCAL_SET, obj_tmp);

                            if let Some(lower) = lower {
                                self.compile_expr(lower)?;
                            } else {
                                inst!(self, core_wasm::i32_const, 0);
                            }
                            self.emit_u16(Op::LOCAL_SET, start_tmp);

                            if let Some(upper) = upper {
                                self.compile_expr(upper)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            self.emit_u16(Op::LOCAL_SET, end_tmp);

                            self.emit_u16(Op::LOCAL_GET, end_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit(Op::I32_SUB);
                            self.emit_u16(Op::LOCAL_SET, count_tmp);

                            let known_string_object = self
                                .infer_expr_type_hint(object)
                                .as_deref()
                                .is_some_and(Self::is_string_type_hint);
                            if known_string_object {
                                inst!(self, core_wasm::bool_const, true);
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                fn_call!(self, "wasm:js-string", "test", 1);
                            }
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            };
                            self.chunk().emit_if(line);

                            let to_string = self.import("ecma:string", "String");
                            let pad_end = self.import("ecma:string", "padEnd");

                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            self.emit_host_call(to_string, 1);
                            self.emit_u16(Op::LOCAL_SET, string_value_tmp);

                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            common::strings::emit_length(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                            };
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            };
                            self.chunk().emit_if(line);

                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            inst!(self, core_wasm::i32_const, 0);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::strings::emit_substring(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_SET, string_value_tmp);

                            self.chunk().emit_end(line);

                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            self.emit_const(Value::String(Arc::from(" ")));
                            self.emit_host_call(pad_end, 3);
                            self.emit_u16(Op::LOCAL_SET, string_value_tmp);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            inst!(self, core_wasm::i32_const, 0);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            common::strings::emit_substring(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            common::strings::emit_str_concat(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, end_tmp);
                            self.emit_const(Value::I32(i32::MAX));
                            common::strings::emit_substring(self.chunk(), line);
                            common::strings::emit_str_concat(self.chunk(), line);
                            self.compile_assign_target(object)?;

                            self.chunk().emit_else(line);

                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            fn_call!(self, "ecma:array", "isArray", 1);
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            };
                            self.chunk().emit_if(line);

                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            self.emit_u16(Op::LOCAL_SET, replacement_tmp);

                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::collections::emit_new_with_length(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            inst!(self, core_wasm::dup);
                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            self.emit_const(Value::I32(0));
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::collections::emit_fill(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_SET, replacement_tmp);
                            self.chunk().emit_end(line);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::collections::emit_remove_range(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit_u16(Op::LOCAL_GET, replacement_tmp);
                            common::collections::emit_insert_range(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            self.emit(Op::DROP);
                            self.chunk().emit_end(line);
                            return Ok(());
                        }
                    }

                    if matches!(&object.kind, ExprKind::Index { .. }) {
                        let line = self.line;
                        let value_tmp = self.define_local("__fortran_nested_index_value");
                        let coll_tmp = self.define_local("__fortran_nested_index_coll");
                        let key_tmp = self.define_local("__fortran_nested_index_key");

                        self.emit_u16(Op::LOCAL_SET, value_tmp);

                        self.compile_expr(object)?;
                        self.emit_u16(Op::LOCAL_SET, coll_tmp);

                        self.compile_array_index_operand_for_owner(object, index)?;
                        self.emit_u16(Op::LOCAL_SET, key_tmp);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.compile_assign_target(object)?;
                        return Ok(());
                    }
                }
                if self.is_python_profile() {
                    if let ExprKind::Slice { lower, upper, step } = &index.kind {
                        // Slice assignment via the shared slices emitter: no-step
                        // splices (variable length, negative-aware), step does a
                        // positional (equal-length) assignment. Value is on TOS.
                        let line = self.line;
                        let value_tmp = self.define_local("__py_slice_value");
                        self.emit_u16(Op::LOCAL_SET, value_tmp);
                        self.compile_expr(object)?;
                        if let Some(lower) = lower {
                            self.compile_expr(lower)?;
                        } else {
                            self.emit(Op::NULL);
                        }
                        if let Some(upper) = upper {
                            self.compile_expr(upper)?;
                        } else {
                            self.emit(Op::NULL);
                        }
                        if step.is_none() {
                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            crate::primitives::slices::emit_splice_assign(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                        } else {
                            if let Some(step) = step {
                                self.compile_expr(step)?;
                            } else {
                                self.emit(Op::NULL);
                            }
                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            let opts = crate::primitives::slices::Options::new(
                                self.profile.slice_step_zero_raises,
                            );
                            crate::primitives::slices::emit_strided_assign(
                                &mut self.chunks,
                                self.current,
                                line,
                                opts,
                            );
                        }
                        return Ok(());
                    }
                }

                // Proxy set-trap dispatch — same shape as Member assign
                // but the key is a runtime expression.
                if self.uses_proxy {
                    let tmp = self.define_local("__proxy_idx_set_v");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let line = self.line;
                    if self.in_strict {
                        vybe_bytecode::registry::hooks(&self.profile.name)
                            .proxy_set_bool
                            .unwrap()(&mut self.chunks, self.current, line);
                        self.emit_strict_set_failure_check()?;
                    } else {
                        vybe_bytecode::registry::hooks(&self.profile.name)
                            .proxy_set
                            .unwrap()(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                    }
                    return Ok(());
                }
                // PHP `$arr[] = v` — empty bracket with null index is the
                // auto-append form; route through collections::emit_push.
                let is_append = matches!(&index.kind, ExprKind::Lit(crate::ast::Literal::Null));
                let line = self.line;
                let tmp = self.define_local("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                if let Some((args_slot, param_slot, alias_index)) =
                    self.js_arguments_alias_for_index_target(object, index)
                {
                    self.emit_u16(Op::LOCAL_GET, args_slot);
                    self.emit_const(Value::F64(alias_index as f64));
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_set(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u16(Op::LOCAL_SET, param_slot);
                    return Ok(());
                }
                if self.is_php_profile() {
                    if let ExprKind::Member {
                        object: recv,
                        field,
                        null_safe,
                    } = &object.kind
                    {
                        if !*null_safe {
                            let recv_tmp = self.define_local("__php_index_member_recv");
                            let coll_tmp = self.define_local("__php_index_member_coll");
                            let field_name = self
                                .field_storage_name_for_receiver(recv, field)
                                .unwrap_or_else(|| self.canon(field));

                            self.compile_expr(recv)?;
                            self.emit_u16(Op::LOCAL_SET, recv_tmp);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            let field_idx = self.str_const(&field_name);
                            self.emit_u16(Op::STRUCT_GET, field_idx);
                            self.emit_u16(Op::LOCAL_SET, coll_tmp);

                            if is_append {
                                self.emit_u16(Op::LOCAL_GET, coll_tmp);
                                self.emit_u16(Op::LOCAL_GET, tmp);
                                common::collections::emit_push(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.emit(Op::DROP);
                            } else {
                                self.compile_expr(index)?;
                                let key_tmp = self.define_local("__php_index_member_key");
                                self.emit_u16(Op::LOCAL_SET, key_tmp);
                                // Promote to an ordered Map on first string key;
                                // native Map order, no `__keys`/CSV side-band.
                                self.emit_php_promote_empty_array_for_string_key(
                                    coll_tmp, key_tmp, line,
                                );
                                self.emit_u16(Op::LOCAL_GET, coll_tmp);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                self.emit_u16(Op::LOCAL_GET, tmp);
                                common::collections::emit_set(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);

                                self.emit_u16(Op::LOCAL_GET, recv_tmp);
                                self.emit_u16(Op::LOCAL_GET, coll_tmp);
                                self.emit_u16(Op::STRUCT_SET, field_idx);
                                self.emit(Op::DROP);
                            }

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.emit_u16(Op::LOCAL_GET, coll_tmp);
                            self.emit_u16(Op::STRUCT_SET, field_idx);
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                    }
                }
                if self.profile.name == "fortran" {
                    if let ExprKind::Member {
                        object: recv,
                        field,
                        null_safe,
                    } = &object.kind
                    {
                        if !*null_safe {
                            let recv_tmp = self.define_local("__fortran_index_member_recv");
                            let coll_tmp = self.define_local("__fortran_index_member_coll");
                            let key_tmp = self.define_local("__fortran_index_member_key");
                            let field_name = self.canon(field);
                            let field_idx = self.str_const(&field_name);

                            self.compile_expr(recv)?;
                            self.emit_u16(Op::LOCAL_SET, recv_tmp);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.emit_u16(Op::STRUCT_GET, field_idx);
                            self.emit_u16(Op::LOCAL_SET, coll_tmp);

                            self.compile_array_index_operand_for_owner(object, index)?;
                            self.emit_u16(Op::LOCAL_SET, key_tmp);

                            self.emit_u16(Op::LOCAL_GET, coll_tmp);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            self.emit_u16(Op::LOCAL_GET, tmp);
                            common::collections::emit_set(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.emit_u16(Op::LOCAL_GET, coll_tmp);
                            self.emit_u16(Op::STRUCT_SET, field_idx);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.compile_assign_target(recv)?;

                            return Ok(());
                        }
                    }
                }
                // PHP auto-vivification: $x[$k][] = $v → ensure $x[$k]
                // is an array before pushing. If undefined, create [].
                if is_append && self.is_php_profile() {
                    if let ExprKind::Index {
                        object: parent,
                        index: key,
                        ..
                    } = &object.kind
                    {
                        let parent_tmp = self.define_local("__vivify_parent");
                        let key_tmp = self.define_local("__vivify_key");
                        let sub_tmp = self.define_local("__vivify_sub");
                        self.compile_expr(parent)?;
                        self.emit_u16(Op::LOCAL_SET, parent_tmp);
                        self.compile_expr(key)?;
                        self.emit_u16(Op::LOCAL_SET, key_tmp);
                        // sub = parent[key]
                        self.emit_u16(Op::LOCAL_GET, parent_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit(Op::ARRAY_GET);
                        self.emit_u16(Op::LOCAL_SET, sub_tmp);
                        // if sub is null/undefined → create [] and set
                        self.emit_u16(Op::LOCAL_GET, sub_tmp);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::ARRAY_NEW_FIXED, 0);
                        self.emit_u16(Op::LOCAL_SET, sub_tmp);
                        self.emit_u16(Op::LOCAL_GET, parent_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, sub_tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);
                        // push value
                        self.emit_u16(Op::LOCAL_GET, sub_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        return Ok(());
                    }
                }
                if is_append {
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    // ecma:array.push leaves [new_length]; drop it.
                    self.emit(Op::DROP);
                } else if self.profile.ecma_object_literals {
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    self.compile_array_index_operand_for_owner(object, index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let set_idx = self.import("ecma:object", "set");
                    if self.in_strict {
                        inst!(self, core_wasm::bool_const, true);
                        self.chunk().emit_call(set_idx, 4, line);
                    } else {
                        self.chunk().emit_call(set_idx, 3, line);
                    }
                    self.emit(Op::DROP);
                } else if self.profile.namespaces.use_dotnet {
                    if self
                        .infer_expr_type_hint(object)
                        .as_deref()
                        .map(Self::normalize_type_hint)
                        .is_some_and(|type_hint| {
                            type_hint.rsplit('.').next().is_some_and(|name| {
                                name.eq_ignore_ascii_case("ObservableCollection")
                            })
                        })
                    {
                        self.compile_expr(object)?;
                        self.compile_collection_key(object, index)?;
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        self.emit_common("dotnet.observable_collection_set_index", 3, line);
                        self.emit(Op::DROP);
                        return Ok(());
                    }
                    if self.profile.namespaces.use_dotnet
                        && self
                            .infer_expr_type_hint(object)
                            .as_deref()
                            .map(Self::normalize_type_hint)
                            .is_some_and(|type_hint| {
                                type_hint
                                    .rsplit('.')
                                    .next()
                                    .is_some_and(|name| name.eq_ignore_ascii_case("StringBuilder"))
                            })
                    {
                        self.compile_expr(object)?;
                        self.compile_collection_key(object, index)?;
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        self.emit_common("dotnet.sb_index_set", 3, line);
                        self.emit(Op::DROP);
                        return Ok(());
                    }

                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    let obj_tmp = self.define_local("__index_set_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_tmp);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let setter_key = self.str_const("__set___index__");
                    self.emit_u16(Op::STRUCT_GET, setter_key);
                    let setter_tmp = self.define_local("__index_setter");
                    self.emit_u16(Op::LOCAL_SET, setter_tmp);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.compile_collection_key(object, index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_set(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);

                    self.chunk().emit_else(line);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.compile_collection_key(object, index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u8(Op::CALL_REF, 3);
                    self.emit(Op::DROP);

                    self.chunk().emit_end(line);
                } else {
                    let is_c_pointer_base_index = if self.profile.name == "c" {
                        match (&object.kind, &index.kind) {
                            (
                                ExprKind::Member {
                                    object: base_owner,
                                    field: base_field,
                                    ..
                                },
                                ExprKind::Member {
                                    object: idx_owner,
                                    field: idx_field,
                                    ..
                                },
                            ) => {
                                matches!(
                                    (&base_owner.kind, &idx_owner.kind, base_field.as_str(), idx_field.as_str()),
                                    (ExprKind::Ident(a), ExprKind::Ident(b), "__base", "__idx") if a == b
                                )
                            }
                            _ => false,
                        }
                    } else {
                        false
                    };

                    if is_c_pointer_base_index {
                        self.compile_expr(object)?;
                        let obj_tmp = self.define_local("__pointer_index_set_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);

                        self.compile_array_index_operand_for_owner(object, index)?;
                        let key_tmp = self.define_local("__pointer_index_set_key");
                        self.emit_u16(Op::LOCAL_SET, key_tmp);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        inst!(self, recipes::is_object);
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);

                        let kind_key = self.str_const("__ref_kind");
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, kind_key);
                        self.emit_const(Value::String(Arc::from("cell")));
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                        }
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        crate::primitives::references::emit_cell_store(
                            &mut self.chunks,
                            self.current,
                            tmp,
                            self.line,
                        );
                        self.emit(Op::DROP);

                        let line = self.line;
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);

                        let line = self.line;
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);
                        return Ok(());
                    }

                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    self.compile_array_index_operand_for_owner(object, index)?;
                    if self.is_php_profile() {
                        let key_tmp = self.define_local("__php_idx_key");
                        let obj_tmp = self.define_local("__php_idx_obj");
                        self.emit_u16(Op::LOCAL_SET, key_tmp);
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);
                        // Promote a sequential array to an ordered Map on first
                        // string key so associative writes are `ObjectKind::Map`
                        // (= Python dict / JS Map — identity-equal pass-around).
                        // Insertion order is then native to the Map (IndexMap);
                        // we deliberately do NOT maintain any `__keys`/CSV side
                        // band — that stamps an extra property onto the Map and
                        // makes `foreach`/`Object.keys` read the stale tracker
                        // instead of the Map's real order.
                        self.emit_php_promote_empty_array_for_string_key(obj_tmp, key_tmp, line);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        if let ExprKind::Ident(name) = &object.kind {
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_var_set(name);
                        }
                        return Ok(());
                    }
                    if self.is_python_profile() {
                        let key_tmp = self.define_local("__py_idx_key");
                        let obj_tmp = self.define_local("__py_idx_obj");
                        self.emit_u16(Op::LOCAL_SET, key_tmp);
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let is_array_idx = self.import("ecma:array", "isArray");
                        self.chunk().emit_call(is_array_idx, 1, line);
                        inst!(self, core_wasm::i32_const, 0);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                        };
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if(line);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let keys_key = self.str_const("__keys");
                        self.emit_u16(Op::STRUCT_GET, keys_key);
                        let keys_tmp = self.define_local("__py_idx_keys");
                        self.emit_u16(Op::LOCAL_SET, keys_tmp);
                        self.emit_u16(Op::LOCAL_GET, keys_tmp);
                        self.emit(Op::REF_IS_NULL);
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, keys_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_index_of(&mut self.chunks, self.current, line);
                        inst!(self, core_wasm::i32_const, 0);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                        };
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, keys_key);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        return Ok(());
                    } else {
                        if self.profile.name == "go" {
                            let go_map_type = match &object.kind {
                                ExprKind::Ident(name) => {
                                    self.lookup_var_type_hint(name).map(str::to_string)
                                }
                                _ => self.infer_expr_type_hint(object),
                            };
                            if go_map_type
                                .as_deref()
                                .is_some_and(|type_hint| type_hint.trim().starts_with("map["))
                            {
                                let key_tmp = self.define_local("__go_idx_key");
                                let obj_tmp = self.define_local("__go_idx_obj");
                                self.emit_u16(Op::LOCAL_SET, key_tmp);
                                self.emit_u16(Op::LOCAL_SET, obj_tmp);

                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                let keys_tmp = self.define_local("__go_idx_keys");
                                self.emit_u16(Op::LOCAL_SET, keys_tmp);
                                self.emit_u16(Op::LOCAL_GET, keys_tmp);
                                self.emit(Op::REF_IS_NULL);
                                self.emit(Op::I32_EQZ);
                                self.chunk().emit_if(line);
                                self.emit_u16(Op::LOCAL_GET, keys_tmp);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                common::collections::emit_index_of(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                inst!(self, core_wasm::i32_const, 0);
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                                };
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                self.chunk().emit_if(line);
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                common::collections::emit_push(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.emit(Op::DROP);
                                self.chunk().emit_end(line);
                                self.chunk().emit_end(line);

                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                self.emit_u16(Op::LOCAL_GET, tmp);
                                common::collections::emit_set(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);
                                return Ok(());
                            }
                        }
                        // JS profile: track insertion order via the
                        // `__keys` side channel so `Object.keys` /
                        // `Object.entries` / `Object.values` see the
                        // correct order. The HashMap backing Ordinary
                        // PHP polyfills that build assoc results
                        // (`array_flip`, `array_diff_assoc`, etc.) and
                        // any JS code that relies on §7.3.22 ordering.
                        if self.profile.ecma_object_literals {
                            let key_tmp = self.define_local("__idx_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp);
                            inst!(self, core_wasm::dup);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            let track_idx = self.import("ecma:object", "trackKey");
                            self.chunk().emit_call(track_idx, 2, line);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                        }
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        // ecma:array.set leaves [null]; drop it.
                        self.emit(Op::DROP);
                    }
                }
            }
            // VB/Pascal: arr(idx) = val — Call used as index because () can
            // represent indexed access in those frontends.
            ExprKind::Call { callee, args, .. } if args.len() == 1 => {
                if self.profile.has_ecma_globals
                    && matches!(&callee.kind, ExprKind::Ident(name) if name == "__len__")
                {
                    let tmp = self.define_local("__tmp");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    self.compile_expr(&args[0].value)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let idx = self.import("ecma:array", "setLength");
                    self.emit_host_call(idx, 2);
                    self.emit(Op::DROP);
                    return Ok(());
                }
                // Route the subscript through the owner-aware normalization
                // path so Pascal char-bound arrays and other declaration-
                // relative indices match the read path.
                let tmp = self.define_local("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.compile_expr(callee)?;
                self.compile_array_index_operand_for_owner(callee, &args[0].value)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                let l = self.line;
                common::collections::emit_set(&mut self.chunks, self.current, l);
                self.emit(Op::DROP); // drop returned null
            }
            ExprKind::Destructure(pattern) => {
                // Destructuring assignment
                match pattern {
                    DestructurePattern::Object(props) => {
                        let obj_slot = self.define_local("__destruct_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        let named_keys: Vec<String> = props
                            .iter()
                            .filter(|p| !p.is_rest)
                            .map(|p| p.key.clone())
                            .collect();
                        for prop in props {
                            if prop.is_rest {
                                let new_idx = self.import("ecma:object", "new");
                                self.emit_host_call(new_idx, 0);
                                let rest_slot = self.define_local("__rest_obj");
                                self.emit_u16(Op::LOCAL_SET, rest_slot);
                                self.emit_u16(Op::LOCAL_GET, rest_slot);
                                self.emit_u16(Op::LOCAL_GET, obj_slot);
                                let assign_idx = self.import("ecma:object", "assign");
                                self.emit_host_call(assign_idx, 2);
                                self.emit(Op::DROP);
                                for named in &named_keys {
                                    self.emit_u16(Op::LOCAL_GET, rest_slot);
                                    self.emit_const(Value::String(Arc::from(named.as_str())));
                                    let del_idx = self.import("ecma:object", "delete");
                                    self.emit_host_call(del_idx, 2);
                                    self.emit(Op::DROP);
                                }
                                self.emit_u16(Op::LOCAL_GET, rest_slot);
                                self.emit_var_set(&prop.key);
                                continue;
                            }
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_const(Value::String(Arc::from(prop.key.as_str())));
                            {
                                let l = self.line;
                                common::collections::emit_get(&mut self.chunks, self.current, l);
                            }
                            let target = match &prop.value {
                                Some(p) => p.clone(),
                                None => BindingPattern::Ident(prop.key.clone()),
                            };
                            self.compile_destructure_bind(&target)?;
                        }
                    }
                    DestructurePattern::Array(elems) => {
                        let arr_slot = self.define_local("__destruct_arr");
                        self.emit_u16(Op::LOCAL_SET, arr_slot);
                        self.compile_array_pattern_assignment_from_slot(arr_slot, elems)?;
                    }
                }
            }
            // JS destructuring assignment shorthand `[a, b] = [b, a]` /
            // `({ x } = obj)` — the walker produces an Array/Object
            // literal for the LHS, but the assignment target re-uses
            // the same shape. Treat each element as a separate
            // assignment to mirror the desugar `let _t = rhs; a = _t[0]; b = _t[1]`.
            ExprKind::Array(elems) => {
                let arr_slot = self.define_local("__assign_destruct_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                for (i, elem) in elems.iter().enumerate() {
                    if elem.spread {
                        continue;
                    }
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    let target = elem.value.clone();
                    self.compile_assign_target(&target)?;
                }
            }
            ExprKind::Tuple(elems) => {
                let arr_slot = self.define_local("__assign_tuple_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                for (i, elem) in elems.iter().enumerate() {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    self.compile_assign_target(elem)?;
                }
            }
            ExprKind::Object(props) => {
                let obj_slot = self.define_local("__assign_destruct_obj");
                self.emit_u16(Op::LOCAL_SET, obj_slot);
                for prop in props {
                    if let crate::ast::ObjectProperty::Shorthand(name) = prop {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        let key = self.str_const(name);
                        self.emit_u16(Op::STRUCT_GET, key);
                        self.emit_var_set(name);
                    } else if let crate::ast::ObjectProperty::KeyValue { key, value } = prop {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        if let ExprKind::Lit(crate::ast::Literal::Str(ref s)) = key.kind {
                            let k = self.str_const(s);
                            self.emit_u16(Op::STRUCT_GET, k);
                        } else {
                            self.compile_expr(key)?;
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        let target = value.clone();
                        self.compile_assign_target(&target)?;
                    }
                }
            }
            _ => {}
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Binary operator emission
    // ════════════════════════════════════════════════════════════════════════
}

/// Evaluate a constant integer expression (a WASM data/elem-segment offset).
/// Plain `i32.const N` → `Lit(Int)`; the Extended Const Expressions proposal
/// also permits `i32.add`/`sub`/`mul` and `global.get` of an immutable global
/// whose initializer is itself a constant expression. Returns `None` for
/// anything non-constant so the caller can report it rather than guess.
fn const_eval_i128(
    expr: &Expression,
    globals: &std::collections::HashMap<String, i64>,
) -> Option<i128> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(n)) | ExprKind::Lit(Literal::BigInt(n)) => Some(*n as i128),
        ExprKind::Lit(Literal::Float(f)) => Some(*f as i128),
        // `global.get $g` lowers to an identifier; resolve it to the global's
        // recorded constant init value (extended-const allows this).
        ExprKind::Ident(name) => globals.get(name).map(|v| *v as i128),
        ExprKind::Binary { op, left, right } => {
            let l = const_eval_i128(left, globals)?;
            let r = const_eval_i128(right, globals)?;
            match op {
                BinOp::Add => Some(l.wrapping_add(r)),
                BinOp::Sub => Some(l.wrapping_sub(r)),
                BinOp::Mul => Some(l.wrapping_mul(r)),
                _ => None,
            }
        }
        // The wast walker lowers a folded `(i32.add …)`/`(i64.mul …)` const
        // expression to `Call(Ident("i32_add"), [a, b])` (not a `Binary`);
        // fold those arithmetic builtins too.
        ExprKind::Call { callee, args, .. } => {
            let ExprKind::Ident(fname) = &callee.kind else {
                return None;
            };
            if args.len() != 2 {
                return None;
            }
            let l = const_eval_i128(&args[0].value, globals)?;
            let r = const_eval_i128(&args[1].value, globals)?;
            match fname.as_str() {
                "i32_add" | "i64_add" => Some(l.wrapping_add(r)),
                "i32_sub" | "i64_sub" => Some(l.wrapping_sub(r)),
                "i32_mul" | "i64_mul" => Some(l.wrapping_mul(r)),
                _ => None,
            }
        }
        _ => None,
    }
}

/// Extended-const evaluation to a `u64` byte offset (wrapping to 32/64-bit is
/// handled by the caller's memory model; a WASM i32 offset wraps mod 2^32).
fn const_eval_u64(
    expr: &Expression,
    globals: &std::collections::HashMap<String, i64>,
) -> Option<u64> {
    const_eval_i128(expr, globals).map(|v| v as u64)
}
