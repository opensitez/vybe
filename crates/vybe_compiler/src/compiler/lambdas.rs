//! Lambda compilation (explicit captures, direct) + component-call.
//!
//! Extracted from `compiler/calls.rs` (`impl Compiler`).

use super::*;

impl Compiler {
    pub(super) fn split_explicit_capture(capture: &str) -> (bool, &str) {
        if let Some(name) = capture.strip_prefix('&') {
            (true, name)
        } else {
            (false, capture)
        }
    }

    pub(super) fn normalize_explicit_capture(&self, capture: &str) -> String {
        let (by_ref, raw_name) = Self::split_explicit_capture(capture);
        let normalized_name = if self.is_php_profile() && !raw_name.starts_with('$') {
            format!("${raw_name}")
        } else {
            raw_name.to_string()
        };

        if by_ref {
            format!("&{normalized_name}")
        } else {
            normalized_name
        }
    }

    pub(super) fn compile_lambda(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        captures: &[String],
    ) -> Result<(), String> {
        self.compile_lambda_with_flags(params, body, captures, false, false, false)
    }

    pub(super) fn compile_lambda_with_flags(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        captures: &[String],
        is_async: bool,
        is_generator: bool,
        is_arrow: bool,
    ) -> Result<(), String> {
        let normalized_captures: Vec<String> = captures
            .iter()
            .map(|capture| self.normalize_explicit_capture(capture))
            .collect();

        if normalized_captures
            .iter()
            .any(|capture| !Self::split_explicit_capture(capture).0)
        {
            return self.compile_lambda_with_explicit_captures(
                params,
                body,
                &normalized_captures,
                is_async,
                is_generator,
            );
        }

        self.compile_lambda_direct(params, body, is_async, is_generator, is_arrow)
    }

    pub(super) fn compile_lambda_with_explicit_captures(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        captures: &[String],
        is_async: bool,
        is_generator: bool,
    ) -> Result<(), String> {
        let capture_bindings: Vec<(String, Option<String>)> = captures
            .iter()
            .filter_map(|capture| {
                let (by_ref, capture_name) = Self::split_explicit_capture(capture);
                if by_ref {
                    None
                } else {
                    Some((
                        capture_name.to_string(),
                        self.lookup_var_type_hint(capture_name).map(str::to_string),
                    ))
                }
            })
            .collect();

        let factory_idx = self.chunks.len();
        let factory = common::functions::create_function_chunk(
            "<lambda_factory>",
            capture_bindings.len() as u8,
        );
        self.chunks.push(factory);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = factory_idx;

        for (capture_name, capture_type) in &capture_bindings {
            self.define_local_typed(capture_name, capture_type.clone());
        }

        // Compile the actual lambda body inside the factory. The inner lambda
        // upvalue-captures the factory's locals (the by-value captures, including
        // __js_this). compile_lambda_direct emits REF_FUNC into the factory chunk,
        // leaving the function reference on the factory's operand stack.
        // PHP `use` closures — never arrows.
        self.compile_lambda_direct(params, body, is_async, is_generator, false)?;

        // Emit RETURN so the factory returns the function reference it just built.
        let line = self.line;
        self.chunks[factory_idx].emit_op(Op::RETURN, line);

        // Collect upvalues AFTER body compilation — the body may have referenced
        // outer-scope variables, registering them as factory upvalues.
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        let inner_scope_idx = self.scopes.len() - 1;
        let uv_names: Vec<Option<String>> = (0..uvs.len())
            .map(|i| self.captured_name_for_upvalue(inner_scope_idx, i as u8))
            .collect();
        self.scopes.pop();
        self.current = saved;

        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], factory_idx, 0, line);
        } else {
            let mut env_slots: Vec<u16> = Vec::new();
            for (i, uv) in uvs.iter().enumerate() {
                if let Some(name) = uv_names[i].clone() {
                    let slot = if uv.is_local {
                        uv.index
                    } else {
                        let parent_env = self.closure_env_slot();
                        let parent_idx = self.closure_env_index(&name);
                        let tmp = self.define_local(&format!("__nested_cap_{}", name));
                        crate::emitter::closures::emit_env_get(
                            self.chunk(),
                            parent_env,
                            parent_idx,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        tmp
                    };
                    env_slots.push(slot);
                }
            }
            crate::emitter::closures::emit_env_new(self.chunk(), &env_slots, line);
            let env_slot = self.define_local(&format!("__closure_env_factory_{}", factory_idx));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], factory_idx, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                env_slot,
                line,
            );
        }
        for capture in captures {
            let (by_ref, capture_name) = Self::split_explicit_capture(capture);
            if !by_ref {
                if self.profile.ambient_this_binding && capture_name == "__js_this" {
                    self.compile_expr(&Expression::new(ExprKind::This))?;
                } else {
                    self.emit_var_get(capture_name);
                }
            }
        }
        self.emit_u8(Op::CALL_REF, capture_bindings.len() as u8);
        Ok(())
    }

    pub(super) fn compile_lambda_direct(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        is_async: bool,
        is_generator: bool,
        is_arrow: bool,
    ) -> Result<(), String> {
        // Walker-lowered generator/async-generator expressions arrive as
        // Lambdas holding the `__gen_fn` contract — they are NOT arrows.
        let is_arrow = is_arrow
            && match body {
                LambdaBody::Block(stmts) => Self::wrapped_generator_kind(stmts).is_none(),
                _ => true,
            };
        let has_rest = params.last().map_or(false, |p| p.is_rest);
        if has_rest {
            self.rest_fixed_arities
                .insert(params.len().saturating_sub(1) as u8);
        }
        // §10.2.11 / §15.3: arrows bind `this` and `new.target` LEXICALLY —
        // captured at CREATION, never read from the ambient call-time
        // globals (member calls set __js_this to the receiver, plain calls
        // null __js_new_target; both are [[Call]]-time bindings arrows must
        // not observe). When the enclosing scope already provides a lexical
        // `this` (method/ctor local, or an outer arrow's capture) the
        // existing upvalue resolution is the capture; otherwise snapshot
        // the current globals into enclosing locals the arrow body's
        // upvalue resolution will find by name.
        if self.profile.ambient_this_binding && is_arrow {
            let self_kw = self.profile.self_keyword.clone();
            let scope_idx = self.scopes.len() - 1;
            let this_reachable = self.scope().resolve(&self_kw).is_some()
                || self.scope().resolve("__js_this").is_some()
                || (scope_idx > 0
                    && (self.resolve_upvalue(scope_idx, &self_kw).is_some()
                        || self.resolve_upvalue(scope_idx, "__js_this").is_some()));
            if !this_reachable {
                let slot = self.define_local("__js_this");
                let js_this = self.str_const("__js_this");
                self.emit_u16(Op::GLOBAL_GET, js_this);
                self.emit_u16(Op::LOCAL_SET, slot);
            }
            let nt_reachable = self.scope().resolve("__js_new_target").is_some()
                || (scope_idx > 0 && self.resolve_upvalue(scope_idx, "__js_new_target").is_some());
            if !nt_reachable {
                let slot = self.define_local("__js_new_target");
                let js_nt = self.str_const("__js_new_target");
                self.emit_u16(Op::GLOBAL_GET, js_nt);
                self.emit_u16(Op::LOCAL_SET, slot);
            }
        }
        // Capture parent's shared env info before switching scope
        let parent_shared_env_slot = self.shared_env_slot;
        let parent_shared_env_names = self.shared_env_names.clone();
        let arity = params.len() as u8;
        let ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk("<lambda>", arity);
        self.chunks.push(chunk);
        self.chunks[ci].is_async = is_async;
        self.chunks[ci].is_generator = is_generator;
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = ci;
        // Runtime TRY_END counts are per-FRAME: a nested chunk must not
        // inherit the enclosing async body's try depth, or its returns pop the
        // CALLER's handlers off the shared runtime handler stack (a lambda
        // compiled inline inside an async fn emitted TRY_END × 2, silently
        // removing the user's enclosing try/catch).
        let saved_async_try_depth = std::mem::take(&mut self.active_async_try_depth);
        let saved_fn = self.current_func_name.replace("<lambda>".into());
        let saved_env_names = std::mem::take(&mut self.closure_env_names);
        let saved_capture_locals = std::mem::take(&mut self.capture_locals);
        let saved_shared_env_slot = self.shared_env_slot.take();
        let saved_shared_env_names = std::mem::take(&mut self.shared_env_names);
        // If parent has a shared env, pre-seed the inner function's
        // closure_env_names so upvalue indices match the shared env layout.
        if !parent_shared_env_names.is_empty() {
            self.closure_env_names = parent_shared_env_names.clone();
        }
        // ECMA-262 §11.2.2: strict mode is inherited by nested functions and
        // additionally turned on by a `"use strict"` directive prologue in
        // this function's own block body. Arrow expression bodies cannot carry
        // a prologue, so they only inherit.
        let saved_strict = self.in_strict;
        let saved_closure_captured = std::mem::take(&mut self.current_closure_captured_locals);
        match body {
            LambdaBody::Block(stmts) => {
                if Self::stmts_have_use_strict_directive(stmts) {
                    self.in_strict = true;
                }
                crate::compiler::collect_closure_captured_idents(
                    stmts,
                    &mut self.current_closure_captured_locals,
                );
            }
            LambdaBody::Expr(expr) => {
                crate::compiler::collect_closure_captured_in_expr(
                    expr,
                    &mut self.current_closure_captured_locals,
                );
            }
        }
        for p in params {
            self.define_local_typed(&p.name, p.type_hint.clone());
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if(line);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot);
                self.chunk().emit_end(line);
            }
        }
        // Snapshot __js_this as a local BEFORE shared env creation so inner
        // arrows can capture it via the shared env / upvalue chain.
        if self.profile.ambient_this_binding && self.scopes.len() > 1 {
            let parent_has_this = self.scopes.len() > 2
                && self.scopes[self.scopes.len() - 2]
                    .resolve("__js_this")
                    .is_some();
            if !parent_has_this {
                let body_has_this = match body {
                    LambdaBody::Block(stmts) => crate::compiler::body_contains_this(stmts),
                    LambdaBody::Expr(expr) => crate::compiler::expr_contains_this(expr),
                };
                if body_has_this {
                    let this_idx = self.str_const("__js_this");
                    self.emit_u16(Op::GLOBAL_GET, this_idx);
                    let this_local = self.define_local("__js_this");
                    self.emit_u16(Op::LOCAL_SET, this_local);
                    self.current_closure_captured_locals
                        .insert("__js_this".to_string());
                }
            }
        }

        if !self.current_closure_captured_locals.is_empty() {
            let mut captured_names: Vec<String> = self
                .current_closure_captured_locals
                .iter()
                .filter(|name| !self.defined_globals.contains(name.as_str()))
                .cloned()
                .collect();
            captured_names.sort();

            {
                let env_size = captured_names.len() as u16;
                let line = self.line;
                for _ in 0..env_size {
                    self.emit(Op::NULL);
                }
                self.chunks[self.current].emit_op_u16(Op::ARRAY_NEW_FIXED, env_size, line);
                let env_slot = self.define_local("__shared_env");
                self.emit_u16(Op::LOCAL_SET, env_slot);
                self.shared_env_slot = Some(env_slot);
                self.shared_env_names = captured_names.clone();
                let mut local_decls: HashSet<String> =
                    params.iter().map(|p| p.name.clone()).collect();
                local_decls.insert("__js_this".to_string());
                if let LambdaBody::Block(stmts) = body {
                    crate::compiler::collect_declared_names(stmts, &mut local_decls);
                }
                for (idx, cap_name) in captured_names.iter().enumerate() {
                    if let Some(param_slot) = self.scope().resolve(cap_name) {
                        self.emit_u16(Op::LOCAL_GET, param_slot);
                        crate::emitter::closures::emit_env_set(
                            self.chunk(),
                            env_slot,
                            idx as u16,
                            line,
                        );
                    } else if !local_decls.contains(cap_name) && parent_shared_env_slot.is_some() {
                        if let Some(parent_idx) =
                            parent_shared_env_names.iter().position(|n| n == cap_name)
                        {
                            let closure_env = self.closure_env_slot();
                            crate::emitter::closures::emit_env_get(
                                self.chunk(),
                                closure_env,
                                parent_idx as u16,
                                line,
                            );
                            crate::emitter::closures::emit_env_set(
                                self.chunk(),
                                env_slot,
                                idx as u16,
                                line,
                            );
                        }
                    }
                }
            }
        }

        let result_slot = if self.profile.function_return == ReturnStyle::ResultSlot {
            let rs = self.define_local("Result");
            self.emit(Op::NULL);
            self.emit_u16(Op::LOCAL_SET, rs);
            let saved_rs = self.current_result_slot.take();
            self.current_result_slot = Some(rs);
            Some((rs, saved_rs))
        } else {
            None
        };
        let saved_result_slot = result_slot.as_ref().map(|(_, saved_rs)| *saved_rs);

        // Same `!is_generator` gate as the declaration/method sites: an
        // async GENERATOR expression completes/throws through `resume`;
        // its promise surface is the attached `.next()` driver.
        let async_try = if is_async && !is_generator && self.profile.async_wraps_body_in_try {
            let line = self.line;
            Some(common::functions::emit_async_body_start(
                &mut self.chunks[self.current],
                line,
            ))
        } else {
            None
        };
        if async_try.is_some() {
            self.active_async_try_depth += 1;
        }

        match body {
            LambdaBody::Expr(expr) => {
                self.compile_expr(expr)?;
                if self.current_chunk_is_js_async() {
                    let resolve_idx = self.import("ecma:promise", "resolve");
                    self.emit_host_call(resolve_idx, 1);
                    self.emit_return_through_finally(1)?;
                } else {
                    self.emit(Op::RETURN);
                }
            }
            LambdaBody::Block(stmts) => {
                for s in stmts {
                    self.compile_stmt(s)?;
                }
            }
        }

        if async_try.is_some() {
            self.active_async_try_depth = self.active_async_try_depth.saturating_sub(1);
        }

        if let Some(catch_jump) = async_try {
            let line = self.line;
            let chunk = &mut self.chunks[self.current];
            common::functions::emit_async_body_fallthrough(chunk, catch_jump, line);
            let resolve_idx = self.import("ecma:promise", "resolve");
            self.emit_host_call(resolve_idx, 1);
            self.emit(Op::RETURN);
            let chunk = &mut self.chunks[self.current];
            common::functions::patch_async_body_catch(chunk, catch_jump);
            let reject_idx = self.import("ecma:promise", "reject");
            self.emit_host_call(reject_idx, 1);
            self.emit(Op::RETURN);
        } else if let Some((rs, saved_rs)) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
            self.current_result_slot = saved_rs;
        } else if matches!(body, LambdaBody::Block(_)) {
            if self.profile.has_undefined_value {
                inst!(self, core_wasm::undefined);
                self.emit(Op::RETURN);
            } else {
                let line = self.line;
                common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
            }
        }
        if let Some(saved_rs) = saved_result_slot {
            self.current_result_slot = saved_rs;
        }

        self.current_func_name = saved_fn;
        self.in_strict = saved_strict;

        let ns = self.scope().next_slot;
        self.chunks[ci].finalize_local_count(ns);
        self.chunks[ci].local_names = self.scope().defined_names.clone();
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        // Resolve upvalue names BEFORE popping the inner scope
        let inner_scope_idx = self.scopes.len() - 1;
        let uv_names: Vec<Option<String>> = (0..uvs.len())
            .map(|i| self.captured_name_for_upvalue(inner_scope_idx, i as u8))
            .collect();
        self.scopes.pop();
        self.current = saved;
        self.active_async_try_depth = saved_async_try_depth;
        self.current_closure_captured_locals = saved_closure_captured;
        self.closure_env_names = saved_env_names;
        self.capture_locals = saved_capture_locals;
        self.shared_env_slot = saved_shared_env_slot;
        self.shared_env_names = saved_shared_env_names;
        let parent_locals = self.scope().locals.clone();
        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 0, line);
        } else if let Some(shared_slot) = parent_shared_env_slot {
            // Parent has a shared env — pass it directly as the upvalue.
            // The inner function's closure_env_names was pre-seeded from
            // parent_shared_env_names, so indices match.
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                shared_slot,
                line,
            );
        } else {
            // No shared env — build a per-closure env (original path).
            let mut env_slots: Vec<u16> = Vec::new();
            let mut env_names: Vec<String> = Vec::new();
            for (i, uv) in uvs.iter().enumerate() {
                if let Some(name) = uv_names[i].clone() {
                    let slot = if uv.is_local {
                        let by_value = parent_locals
                            .iter()
                            .find(|l| l.slot == uv.index)
                            .map(|l| {
                                self.capture_by_value_vars
                                    .iter()
                                    .any(|n| *n == self.canon(&l.name))
                            })
                            .unwrap_or(false);
                        if by_value {
                            let orig_slot = uv.index;
                            self.emit_u16(Op::LOCAL_GET, orig_slot);
                            let snap = self.define_local(&format!("__snap_{}_{}", name, ci));
                            self.emit_u16(Op::LOCAL_SET, snap);
                            snap
                        } else {
                            uv.index
                        }
                    } else {
                        let parent_env = self.closure_env_slot();
                        let parent_idx = self.closure_env_index(&name);
                        let tmp = self.define_local(&format!("__nested_cap_{}", name));
                        crate::emitter::closures::emit_env_get(
                            self.chunk(),
                            parent_env,
                            parent_idx,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        tmp
                    };
                    env_names.push(name);
                    env_slots.push(slot);
                }
            }
            crate::emitter::closures::emit_env_new(self.chunk(), &env_slots, line);
            let env_slot = self.define_local(&format!("__closure_env_{}", ci));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                env_slot,
                line,
            );
        }
        if self.profile.has_function_prototype_bind {
            let length = params
                .iter()
                .take_while(|p| p.default.is_none() && !p.is_rest)
                .count();

            inst!(self, core_wasm::dup);
            self.emit_const(Value::F64(length as f64));
            let length_key = self.str_const("length");
            self.emit_u16(Op::STRUCT_SET, length_key);
            self.emit(Op::DROP);

            inst!(self, core_wasm::dup);
            {
                // Recover the source kind when the walker lowered a generator
                // into a plain wrapper holding `__gen_fn` (obj-literal
                // `*m(){}` methods and generator expressions).
                let (eff_async, eff_generator) = match body {
                    LambdaBody::Block(stmts) => {
                        Self::wrapped_generator_kind(stmts).unwrap_or((is_async, is_generator))
                    }
                    _ => (is_async, is_generator),
                };
                let line = self.line;
                crate::emitter::prototypes::emit_stamp_function_kind_proto(
                    self.chunk(),
                    eff_async,
                    eff_generator,
                    line,
                );
            }

            // §7.2.4 IsConstructor: arrows (and the other Lambda-lowered
            // forms — shorthand methods, generator expressions) have no
            // [[Construct]]. `new` on them must TypeError; the host
            // construct path checks this marker.
            inst!(self, core_wasm::dup);
            self.emit_const(Value::Bool(true));
            let non_ctor_key = self.str_const("__vybe_non_ctor");
            self.emit_u16(Op::STRUCT_SET, non_ctor_key);
            self.emit(Op::DROP);

            // §10.2.9/§10.2.10: name/length are non-enumerable.
            inst!(self, core_wasm::dup);
            {
                let line = self.line;
                crate::emitter::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);
            }

            // §10.2.11: arrows carry a marker — the host uses it for
            // lexical-this (call/apply ignore thisArg) and toString's
            // `=>` form. Threaded from the ExprKind::Lambda compile arm;
            // object-literal shorthand methods pass false.
            if is_arrow {
                inst!(self, core_wasm::dup);
                self.emit_const(Value::Bool(true));
                let arrow_key = self.str_const("__fn_arrow");
                self.emit_u16(Op::STRUCT_SET, arrow_key);
                self.emit(Op::DROP);
            }
        }
        if has_rest {
            self.emit_stamp_rest_metadata_on_stack(params.len().saturating_sub(1));
        }
        Ok(())
    }

    /// ES2024 `Object.groupBy(arr, fn)` — inline loop emitter.
    ///
    /// Stack on entry: [arr, fn]. Result: new object whose keys are the
    /// string results of `fn(item)` and whose values are arrays of matching
    /// items. Uses only already-registered host fns (ecma:object, ecma:array);
    /// no new imports needed.
    pub(super) fn emit_object_group_by(&mut self, line: u32) -> Result<(), String> {
        let fn_slot = self.define_local("__groupby_fn");
        self.emit_u16(Op::LOCAL_SET, fn_slot);
        let arr_slot = self.define_local("__groupby_arr");
        self.emit_u16(Op::LOCAL_SET, arr_slot);

        let new_idx = self.import("ecma:object", "new");
        self.emit_host_call(new_idx, 0);
        let result_slot = self.define_local("__groupby_result");
        self.emit_u16(Op::LOCAL_SET, result_slot);

        self.emit_u16(Op::LOCAL_GET, arr_slot);
        common::collections::emit_len(&mut self.chunks, self.current, line);
        let len_slot = self.define_local("__groupby_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);

        self.emit_const(Value::F64(0.0));
        let i_slot = self.define_local("__groupby_i");
        self.emit_u16(Op::LOCAL_SET, i_slot);

        let loop_state = common::loops::emit_loop_start(&mut self.chunks, self.current, line);

        self.emit_u16(Op::LOCAL_GET, i_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_lt(self.chunk(), line);
        };
        common::loops::emit_loop_cond(&mut self.chunks, self.current, line);

        self.emit_u16(Op::LOCAL_GET, arr_slot);
        self.emit_u16(Op::LOCAL_GET, i_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        let item_slot = self.define_local("__groupby_item");
        self.emit_u16(Op::LOCAL_SET, item_slot);

        // key = fn(item)
        self.emit_u16(Op::LOCAL_GET, fn_slot);
        self.emit_u16(Op::LOCAL_GET, item_slot);
        self.emit_u8(Op::CALL_REF, 1);
        let key_slot = self.define_local("__groupby_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        // if result[key] === undefined: result[key] = []
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);

        // result[key].push(item)
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit_u16(Op::LOCAL_GET, item_slot);
        common::collections::emit_push(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, i_slot);
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_SET, i_slot);

        common::loops::emit_loop_end(&mut self.chunks, self.current, loop_state, line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(())
    }

    pub(super) fn try_compile_dotnet_component_call(
        &mut self,
        parts: &[String],
        args: &[&Expression],
    ) -> Result<bool, String> {
        // namespaceplan.md: platform surfaces are data in the shared tree;
        // the common resolver handles the mounted chain.
        match self.resolve_profile_namespace_chain(parts) {
            Some(super::resolver::Resolution::Tree(
                crate::emitter::namespaces::ResolutionTarget::CommonEmit(emit),
            )) => {
                for a in args {
                    self.compile_expr(a)?;
                }
                let line = self.line;
                self.emit_common(&emit, args.len() as u8, line);
                Ok(true)
            }
            Some(
                super::resolver::Resolution::HostImport { module, func }
                | super::resolver::Resolution::Tree(
                    crate::emitter::namespaces::ResolutionTarget::HostCall { module, func, .. },
                ),
            ) => {
                for a in args {
                    self.compile_expr(a)?;
                }
                let idx = self.import(&module, &func);
                self.emit_host_call(idx, args.len() as u8);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}
