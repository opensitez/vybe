//! Class, constructor, and free-function compilation.
//!
//! Extracted from `compiler.rs` to keep that file navigable. The
//! methods on this `impl Compiler { ... }` block are private by
//! convention (they're only called from other compiler methods) and
//! crate-private for the `dotnet_register` bridge.

use super::*;

impl Compiler {
    // ════════════════════════════════════════════════════════════════════════
    // Function declaration compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_function_decl(
        &mut self, name: &str, params: &[Param], return_type: &Option<String>,
        body: &[Statement], _is_sub: bool, is_generator: bool, handles: &[String],
        is_async: bool,
    ) -> Result<(), String> {
        let cname = self.canon(name);
        self.defined_globals.insert(cname.clone());
        self.defined_functions.insert(cname.clone());
        let name = &cname;

        let has_rest = params.last().map_or(false, |p| p.is_rest);
        let arity: u8 = if has_rest { 255 } else { params.len() as u8 };
        let func_idx = self.chunks.len();
        let mut chunk = common::functions::create_function_chunk(name, arity);
        // Mark the chunk so the WASM emitter can list it in the
        // `vybe.jspi` custom section — JS hosts wrap promising exports
        // via `WebAssembly.promising(...)` at load time.
        chunk.is_async = is_async;
        // Generators: when the source marked the function as a
        // generator (Python `yield`, JS `function*`, C# `yield return`),
        // stamp the chunk so the VM wraps invocations in a
        // `Continuation` instead of executing the body inline. The
        // body itself was compiled with `SUSPEND` at each yield site.
        chunk.is_generator = is_generator;
        if is_generator {
            self.generator_functions.insert(cname.clone());
        }
        // Multi-value tuple returns: if the pre-scan marked this function
        // as a same-arity multi-tuple-return, stamp its result_arity here
        // so the WASM type section emits `(externref^N) -> (externref^N)`
        // and the VM's RETURN knows to pop N values off the stack.
        if let Some(&n) = self.multi_return_functions.get(&cname) {
            chunk.result_arity = n;
        }
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = func_idx;

        // Define params
        for p in params {
            self.scope_mut().define(&p.name);
            // Default parameters
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let has_val = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                self.patch_jump(has_val);
            }
        }

        // Rest param preamble: collect excess args into an array.
        // With arity=255 the VM doesn't truncate excess args. They land in
        // sequential slots after the non-rest params. We scan those slots
        // with unrolled local_get + null-check (local_get is static u16).
        // Caps at 16 rest args which covers all realistic use cases.
        if has_rest {
            let rest_name = &params.last().unwrap().name;
            let rest_slot = self.scope().resolve(rest_name).unwrap();
            // Build array from slots rest_slot..rest_slot+16, stopping at null.
            // Pattern per slot: if local[N] is null → jump to done; else arr.push(local[N])
            // Build rest array via `common::collections` so the provider
            // is swappable in one place. `vybe:js-array.push` returns
            // new_length (ECMA-262), not arr, so we stash arr in a
            // scope-local and reload each iteration.
            let line = self.line;
            // Reserve the 16 rest-arg slots before allocating `__rest_arr` so
            // the accumulator doesn't overwrite an incoming rest argument.
            // (The VM parks overflow args in slots rest_slot..rest_slot+argc-arity;
            // without this reservation, `__rest_arr` landed on the second rest arg,
            // triggering a self-referential push loop.)
            let max_rest = 16u16;
            for i in 1..max_rest {
                self.scope_mut().define(&format!("__rest_reserved_{}", i));
            }
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            let rest_arr = self.scope_mut().define("__rest_arr");
            self.emit_u16(Op::LOCAL_SET, rest_arr); self.emit(Op::DROP);
            let mut done_patches: Vec<usize> = Vec::new();
            for i in 0..max_rest {
                let slot = rest_slot + i;
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                done_patches.push(self.emit_jump(Op::BR_IF_TRUE)); // null → done
                self.emit_u16(Op::LOCAL_GET, rest_arr);
                self.emit_u16(Op::LOCAL_GET, slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP); // drop new_length
            }
            for p in done_patches { self.patch_jump(p); }
            // Store rest array back into the rest_slot param position.
            self.emit_u16(Op::LOCAL_GET, rest_arr);
            self.emit_u16(Op::LOCAL_SET, rest_slot);
            self.emit(Op::DROP);
        }

        // Result slot for functions with return type (Pascal/VB Function).
        // The slot name is profile-driven so VB can keep it internal
        // (`__result__`) and avoid shadowing user classes named `Result`,
        // while Pascal keeps it as `Result` (user-visible per Pascal idiom).
        let result_slot = if return_type.is_some() && self.profile.function_return == ReturnStyle::ResultSlot {
            let slot_name = self.profile.result_slot_name.clone();
            let rs = self.scope_mut().define(&slot_name);
            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
            Some(rs)
        } else {
            None
        };

        let saved_fn = self.current_func_name.take();
        let saved_rs = self.current_result_slot.take();
        self.current_func_name = Some(name.to_string());
        self.current_result_slot = result_slot;

        for s in body { self.compile_stmt(s)?; }

        self.current_func_name = saved_fn;
        self.current_result_slot = saved_rs;

        if let Some(rs) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
        } else {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
        }

        let locals = self.scope().next_slot.max(self.chunks[func_idx].local_count);
        self.chunks[func_idx].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;

        let line = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, uvs.len() as u8, line);
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current].emit(uv.index, line);
        }
        let idx = self.str_const(name);
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);

        // VB `Handles ctrl.Event` clause on a top-level Sub: register the
        // event handler with the canonical GUI binding. The same canonical
        // emit path serves C# `+=`, JS `addEventListener`, etc.
        for handle in handles {
            let parts: Vec<&str> = handle.splitn(2, '.').collect();
            if parts.len() == 2 {
                let line = self.line;
                let bind_idx = self.import("vybe:gui", common::gui::HOST_FN_BIND_EVENT);
                self.emit_var_get(parts[0]);
                common::gui::emit_get_control_name(self.chunk(), line);
                let ev = parts[1].to_lowercase();
                self.emit_const(Value::String(Arc::from(ev.as_str())));
                self.emit_var_get(name);
                common::gui::emit_bind_event(self.chunk(), bind_idx, line);
                self.emit(Op::DROP); // statement: discard host call result
            }
        }

        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Class compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_class(&mut self, name: &str, parent: &Option<String>, members: &[ClassMember]) -> Result<(), String> {
        let self_kw = self.profile.self_keyword.clone();
        let ctor_name = self.profile.constructor_name.clone();
        let result_style = self.profile.function_return.clone();

        // Collect fields and initializers (separate instance vs static)
        // Auto-properties are treated as plain fields (matches old C# compiler).
        let mut fields = Vec::new();
        let mut field_inits: Vec<(String, Option<Expression>)> = Vec::new();
        let mut static_field_inits: Vec<(String, Option<Expression>)> = Vec::new();
        for m in members {
            if let ClassMember::Field { name: fname, init, modifiers, .. } = m {
                let fname = self.canon(fname);
                if modifiers.is_static {
                    static_field_inits.push((fname, init.clone()));
                } else {
                    fields.push(fname.clone());
                    field_inits.push((fname, init.clone()));
                }
            }
            if let ClassMember::Property { name: pname, is_auto, modifiers, .. } = m {
                if *is_auto {
                    let pname_canon = self.canon(pname);
                    if modifiers.is_static {
                        if !static_field_inits.iter().any(|(n, _)| n == &pname_canon) {
                            static_field_inits.push((pname_canon, None));
                        }
                    } else if !fields.contains(&pname_canon) {
                        fields.push(pname_canon.clone());
                        field_inits.push((pname_canon, None));
                    }
                }
            }
        }

        // Store field list for implicit self resolution
        self.pending_classes.insert(name.to_string(), PendingClass {
            parent: parent.clone(),
            fields: fields.clone(),
            statics: Vec::new(), // filled after methods are compiled
        });

        // Compile methods (including constructor body)
        // (name, chunk_idx, is_ctor, is_static)
        let mut method_chunks: Vec<(String, usize, bool, bool)> = Vec::new();
        let saved_class = self.current_class.take();
        self.current_class = Some(name.to_string());

        // Pre-register all method names to avoid value-method hijacking
        for m in members {
            if let ClassMember::Method(stmt) = m {
                if let StmtKind::FunctionDecl { name: mname, .. } = &stmt.kind {
                    self.defined_class_methods.insert(self.canon(mname));
                }
            }
            if let ClassMember::Property { name: pname, .. } = m {
                self.defined_class_methods.insert(self.canon(pname));
            }
        }


        for m in members {
            match m {
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name: mname, params, return_type, body, modifiers, is_sub: _, is_generator, .. } = &stmt.kind {
                        // NOTE: do NOT skip empty-body methods. They still need
                        // a chunk + binding so that callers (e.g. an explicit
                        // constructor calling `InitializeComponent()`) can
                        // dispatch through `me.<method>`. Skipping here is what
                        // caused VB Forms tests to fail with "null is not
                        // callable" — the empty `Sub InitializeComponent` was
                        // never bound on `me`.

                        let is_ctor = if self.case_sensitive {
                            mname == &ctor_name || (modifiers.is_static && mname == "new")
                        } else {
                            mname.eq_ignore_ascii_case(&ctor_name)
                            || modifiers.is_static && mname.eq_ignore_ascii_case("new")
                        };

                        let user_params: Vec<&Param> = if self.profile.explicit_self_param {
                            params.iter().skip(1).collect()
                        } else {
                            params.iter().collect()
                        };
                        let arity = (user_params.len() + 1) as u8; // +1 for self

                        let ci = self.chunks.len();
                        let mut chunk = common::functions::create_function_chunk(mname, arity);
                        // C# `yield return` / JS generator methods on
                        // classes: propagate the generator flag from the
                        // FunctionDecl so the VM wraps invocations in a
                        // `Continuation` instead of executing the body.
                        chunk.is_generator = *is_generator;
                        if *is_generator {
                            let cname = self.canon(mname);
                            self.generator_functions.insert(cname);
                        }
                        self.chunks.push(chunk);
                        self.scopes.push(Scope::new_function());
                        let saved = self.current;
                        self.current = ci;

                        self.scope_mut().define(&self_kw);
                        for p in &user_params { self.scope_mut().define(&p.name); }

                        if is_ctor {
                            for s in body { self.compile_stmt(s)?; }
                            if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                                self.emit_u16(Op::LOCAL_GET, slot);
                                self.emit(Op::RETURN);
                            }
                        } else if return_type.is_some() && result_style == ReturnStyle::ResultSlot {
                            let slot_name = self.profile.result_slot_name.clone();
                            let rs = self.scope_mut().define(&slot_name);
                            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
                            let saved_fn = self.current_func_name.take();
                            let saved_rs = self.current_result_slot.take();
                            self.current_func_name = Some(mname.clone());
                            self.current_result_slot = Some(rs);
                            for s in body { self.compile_stmt(s)?; }
                            self.current_func_name = saved_fn;
                            self.current_result_slot = saved_rs;
                            self.emit_u16(Op::LOCAL_GET, rs);
                            self.emit(Op::RETURN);
                        } else {
                            for s in body { self.compile_stmt(s)?; }
                            let line = self.line;
                            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                        }

                        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
                        self.chunks[ci].local_count = locals;
                        self.scopes.pop();
                        self.current = saved;

                        let bound_name = self.canon(mname);
                        method_chunks.push((bound_name, ci, is_ctor, modifiers.is_static));
                    }
                }
                ClassMember::Constructor { .. } => {
                    // Constructor body is handled by the main constructor flow below
                    // (extracted via ctor_body). No separate chunk needed.
                }
                ClassMember::Property { name: pname, getter, setter, is_auto, .. } => {
                    // Auto-properties are handled as plain fields above — skip getter/setter compilation
                    if *is_auto { continue; }
                    let pname_canon = self.canon(pname);

                    // Getter → __get_<prop>
                    if let Some(getter_body) = getter {
                        let get_name = format!("__get_{}", pname_canon);
                        let ci = self.chunks.len();
                        let chunk = common::functions::create_function_chunk(&get_name, 1); // self
                        self.chunks.push(chunk);
                        self.scopes.push(Scope::new_function());
                        let saved = self.current;
                        self.current = ci;
                        self.scope_mut().define(&self_kw);

                        if getter_body.is_empty() {
                            // Auto-property getter: return backing field
                            if let Some(slot) = self.scope().resolve(&self_kw) {
                                self.emit_u16(Op::LOCAL_GET, slot);
                                let backing = self.str_const(&format!("__{}", pname_canon));
                                self.emit_u16(Op::STRUCT_GET, backing);
                                self.emit(Op::RETURN);
                            }
                        } else {
                            let slot_name = self.profile.result_slot_name.clone();
                            let rs = self.scope_mut().define(&slot_name);
                            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
                            let saved_fn = self.current_func_name.take();
                            let saved_rs = self.current_result_slot.take();
                            self.current_func_name = Some(pname.clone());
                            self.current_result_slot = Some(rs);
                            for s in getter_body { self.compile_stmt(s)?; }
                            self.current_func_name = saved_fn;
                            self.current_result_slot = saved_rs;
                            self.emit_u16(Op::LOCAL_GET, rs);
                            self.emit(Op::RETURN);
                        }

                        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
                        self.chunks[ci].local_count = locals;
                        self.scopes.pop();
                        self.current = saved;
                        method_chunks.push((get_name, ci, false, false));
                    }

                    // Setter → __set_<prop>
                    if let Some(setter_info) = setter {
                        let set_name = format!("__set_{}", pname_canon);
                        let ci = self.chunks.len();
                        let chunk = common::functions::create_function_chunk(&set_name, 2); // self, value
                        self.chunks.push(chunk);
                        self.scopes.push(Scope::new_function());
                        let saved = self.current;
                        self.current = ci;
                        self.scope_mut().define(&self_kw);
                        self.scope_mut().define(&setter_info.param.name);

                        if setter_info.body.is_empty() {
                            // Auto-property setter: set backing field
                            if let Some(self_slot) = self.scope().resolve(&self_kw) {
                                self.emit_u16(Op::LOCAL_GET, self_slot);
                                if let Some(val_slot) = self.scope().resolve(&setter_info.param.name) {
                                    self.emit_u16(Op::LOCAL_GET, val_slot);
                                }
                                let backing = self.str_const(&format!("__{}", pname_canon));
                                self.emit_u16(Op::STRUCT_SET, backing);
                                self.emit(Op::DROP);
                            }
                        } else {
                            for s in &setter_info.body { self.compile_stmt(s)?; }
                        }

                        let line = self.line;
                        common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
                        self.chunks[ci].local_count = locals;
                        self.scopes.pop();
                        self.current = saved;
                        method_chunks.push((set_name, ci, false, false));
                    }
                }
                ClassMember::Const { name: cname, value, .. } => {
                    // Class-level constant → global
                    self.compile_expr(value)?;
                    let global_name = self.canon(&format!("{}.{}", name, cname));
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.emit(Op::DROP);
                    self.defined_globals.insert(global_name);
                }
                ClassMember::Event { .. } => { /* type-level only */ }
                ClassMember::NestedType(stmt) => { self.compile_stmt(stmt)?; }
                _ => {}
            }
        }

        self.current_class = saved_class;

        // Find constructor body and its user arity
        let _ctor = method_chunks.iter().find(|(_, _, is_ctor, _)| *is_ctor);
        let ctor_body: Option<(&Vec<Statement>, &Vec<Param>, Option<&Vec<Expression>>)> = members.iter().find_map(|m| {
            match m {
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name: mname, params, body, modifiers, .. } = &stmt.kind {
                        let is_ctor = if self.case_sensitive {
                            mname == &ctor_name || (modifiers.is_static && mname == "new")
                        } else {
                            mname.eq_ignore_ascii_case(&ctor_name)
                            || modifiers.is_static && mname.eq_ignore_ascii_case("new")
                        };
                        if is_ctor && !body.is_empty() { return Some((body, params, None)); }
                    }
                    None
                }
                ClassMember::Constructor { params, body, base_args, .. } => Some((body, params, base_args.as_ref())),
                _ => None,
            }
        });

        let user_params: Vec<String> = ctor_body.map(|(_, params, _)| {
            if self.profile.explicit_self_param {
                params.iter().skip(1).map(|p| p.name.clone()).collect()
            } else {
                params.iter().map(|p| p.name.clone()).collect()
            }
        }).unwrap_or_default();
        let user_arity = user_params.len() as u8;

        // ── Single constructor function (not split wrapper + body) ──────
        // This is the ONLY function that `new ClassName(args)` calls.
        // It creates the object, initializes fields, binds methods, runs
        // user constructor body, and returns this.
        let ctor_idx = self.chunks.len();
        let ctor_chunk = common::functions::create_function_chunk(name, user_arity);
        self.chunks.push(ctor_chunk);
        self.scopes.push(Scope::new_function());
        let saved_cur = self.current;
        let saved_class2 = self.current_class.take();
        self.current = ctor_idx;
        self.current_class = Some(name.to_string());

        // Define user params (slot 1..N), then this (slot N+1)
        // Also handle default parameter values from the Param structs.
        let ctor_param_defaults: Vec<Option<Expression>> = ctor_body.map(|(_, params, _)| {
            let skip = if self.profile.explicit_self_param { 1 } else { 0 };
            params.iter().skip(skip).map(|p| p.default.clone()).collect()
        }).unwrap_or_default();
        for (i, p) in user_params.iter().enumerate() {
            self.scope_mut().define(p);
            if let Some(Some(default)) = ctor_param_defaults.get(i) {
                let slot = self.scope().resolve(p).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let has_val = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                self.patch_jump(has_val);
            }
        }
        self.scope_mut().define(&self_kw); // this_slot = user_arity
        let this_slot = user_arity as u16;

        let is_child = parent.is_some();
        let line = self.line;

        // Separate instance methods from static methods
        let instance_methods: Vec<&(String, usize, bool, bool)> = method_chunks.iter()
            .filter(|(_, _, ic, is_static)| !*ic && !*is_static)
            .collect();
        let static_methods: Vec<&(String, usize, bool, bool)> = method_chunks.iter()
            .filter(|(_, _, ic, is_static)| !*ic && *is_static)
            .collect();
        let instance_method_names: Vec<String> = instance_methods.iter()
            .map(|(n, _, _, _)| n.clone())
            .collect();

        if is_child {
            // ── Child class ─────────────────────────────────────────────
            // For child classes, the constructor body calls super() which
            // creates the object. We run the body FIRST, then bind methods.
            // This works for both explicit super (JS) and implicit (VB/C#)
            // because super() stores the result in this_slot.

            // ── Step 1: Call parent constructor to get the object ────────
            self.emit(Op::NULL);
            self.emit_u16(Op::LOCAL_SET, this_slot);
            self.emit(Op::DROP);

            // Determine if auto_base_call should kick in:
            // ctor exists + base_args None + profile says auto + parent exists + body has no super
            let has_explicit_base = ctor_body.as_ref().map_or(false, |(_, _, ba)| ba.is_some());
            let auto_base_needed = !has_explicit_base
                && ctor_body.is_some()
                && self.profile.auto_base_call
                && parent.is_some()
                && {
                    let stmts = ctor_body.as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
                    !body_has_super_call(stmts)
                };

            if let Some((_, _, base_args)) = &ctor_body {
                if let Some(bargs) = base_args {
                    // Explicit base_args provided (C#-style `: base(args)`)
                    if let Some(parent_name) = parent {
                        let pname = self.canon(parent_name);
                        let pidx = self.str_const(&pname);
                        self.emit_u16(Op::GLOBAL_GET, pidx);
                        for a in *bargs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL, bargs.len() as u8);
                        self.emit_u16(Op::LOCAL_SET, this_slot);
                        self.emit(Op::DROP);
                    }
                } else if auto_base_needed {
                    // Profile-driven auto base call (VB/C#/Pascal):
                    // body has no super() → auto-call parent() with 0 args.
                    if let Some(parent_name) = parent {
                        let pname = self.canon(parent_name);
                        let pidx = self.str_const(&pname);
                        self.emit_u16(Op::GLOBAL_GET, pidx);
                        self.emit_u8(Op::CALL, 0);
                        self.emit_u16(Op::LOCAL_SET, this_slot);
                        self.emit(Op::DROP);
                    }
                }
                // else: JS pattern — body calls super() itself, sets this_slot
            } else {
                // No explicit constructor — auto-call parent with user args
                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    let pidx = self.str_const(&pname);
                    self.emit_u16(Op::GLOBAL_GET, pidx);
                    for i in 0..user_arity {
                        self.emit_u16(Op::LOCAL_GET, (i as u16) + 1);
                    }
                    self.emit_u8(Op::CALL, user_arity);
                    self.emit_u16(Op::LOCAL_SET, this_slot);
                    self.emit(Op::DROP);
                }
            }

            if has_explicit_base || auto_base_needed || ctor_body.is_none() {
                // C#-style: base call already done above, or no-ctor auto-call done above.
                // Order: re-stamp __type → fields → save base → bind methods → body
                //
                // The parent ctor stamped __type with the parent name. Re-stamp with
                // the child name so `obj is ChildType` returns true.
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_const(Value::String(Arc::from(name)));
                let type_key = self.str_const("__type");
                self.emit_u16(Op::STRUCT_SET, type_key);
                self.emit(Op::DROP);

                for (fname, init) in &field_inits {
                    if let Some(init_expr) = init {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.compile_expr(init_expr)?;
                        common::classes::emit_init_field_end(self.chunk(), fname, line);
                    } else {
                        common::classes::emit_init_field_null(self.chunk(), this_slot, fname, line);
                    }
                }

                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    for method_name in &instance_method_names {
                        common::classes::emit_save_base_method(self.chunk(), this_slot, method_name, line);
                    }
                    common::classes::emit_store_super(self.chunk(), this_slot, &pname, line);
                }

                for (mname, mci, _, _) in &instance_methods {
                    if mname.starts_with("__get_") {
                        let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                        common::classes::emit_bind_getter(self.chunk(), this_slot, prop, *mci, line);
                    } else if mname.starts_with("__set_") {
                        let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                        common::classes::emit_bind_setter(self.chunk(), this_slot, prop, *mci, line);
                    } else {
                        common::classes::emit_bind_method_with_aliases(self.chunk(), this_slot, mname, *mci, line);
                    }
                }

                // Auto-init methods from profile (e.g. InitializeComponent for .NET forms).
                // Emitted after method binding but before user body.
                {
                    let ctor_stmts: &[Statement] = ctor_body
                        .as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
                    let auto_inits = self.profile.auto_init_methods.clone();
                    for aim in &auto_inits {
                        let has_method = instance_methods.iter()
                            .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                        if has_method && !body_calls_method(ctor_stmts, aim) {
                            common::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
                        }
                    }
                }

                if let Some((body, _, _)) = ctor_body {
                    for s in body { self.compile_stmt(s)?; }
                }
            } else {
                // JS/VB/Pascal-style: constructor body contains the super() call which
                // sets this_slot.
                //
                // Order matters for VB.NET / Pascal correctness: real .NET binds the
                // class's instance methods on `this` BEFORE the user-visible body
                // runs, so that user body code can call its own methods (e.g.
                // `InitializeComponent()` inside `New()`). The catch is that
                // method binding can only happen AFTER `this` exists, which in
                // this branch means after the super/inherited call.
                //
                // We split the body at the super call:
                //   body[..=super_idx]   — "preamble": runs to set up `this`
                //   bind methods + save base + field inits + re-stamp __type
                //   body[super_idx+1..]  — "main": user code that can now call
                //                          methods on `this`
                //
                // The walker normalization for each language is responsible for
                // putting a super call in the body for `Inherits` classes:
                //   - VB: walker injects `MyBase.New()` (and an
                //     `Me.__control_name = "<lower class name>"` stamp) at the
                //     top of every ctor body for `Inherits` classes — real
                //     VB.NET semantics where the runtime implicitly calls the
                //     parameterless parent ctor.
                //   - C#: walker sets `base_args = Some(_)` (handled by the
                //     C#-style branch above, not this one).
                //   - Pascal: user writes `inherited Create(...)`.
                //   - JS: user writes `super(...)`.
                //
                // We skip null-init for fields with no explicit initializer because
                // the body may have already assigned them (Pascal pattern:
                // `inherited Create(X); FY := Y;`) — and a no-op null-init would
                // clobber that assignment. Fields default to null on dynamic
                // structs anyway, so this is safe.

                let body_stmts: &[Statement] = ctor_body
                    .as_ref()
                    .map(|(b, _, _)| b.as_slice())
                    .unwrap_or(&[]);

                // Find the index of the first super-call statement in the body.
                // This is the "this exists" boundary. Different walkers emit
                // the super call as different node shapes:
                //   - VB walker: `Expr(SuperCall { method: Some("New"), args })`
                //   - C# walker (when not using : base): same shape
                //   - JS walker: `Expr(Call { callee: Super, args })`
                //   - Pascal walker: `Expr(SuperCall { method: Some("Create"), args })`
                //     OR an `inherited Create(...)` that lowers to a Super call.
                // We match all of them to keep this branch language-agnostic.
                //
                // The walker normalization for VB also injects
                // `Me.__control_name = "..."` immediately after; we include it
                // in the preamble so methods bind onto a fully-stamped `this`.
                let is_super_call = |s: &Statement| -> bool {
                    if let StmtKind::Expr(e) = &s.kind {
                        match &e.kind {
                            ExprKind::SuperCall { .. } => true,
                            ExprKind::Call { callee, .. } => matches!(callee.kind, ExprKind::Super),
                            _ => false,
                        }
                    } else {
                        false
                    }
                };
                let super_idx = body_stmts.iter().position(is_super_call);
                let preamble_end = match super_idx {
                    Some(i) => {
                        // Extend through any immediately-following identity stamps
                        // (Me.__control_name = ..., Me.__type = ..., etc.) so the
                        // method binding sees the canonical control name.
                        let mut end = i + 1;
                        while end < body_stmts.len() && is_identity_stamp(&body_stmts[end]) {
                            end += 1;
                        }
                        end
                    }
                    None => 0,
                };

                // Compile preamble (super call + any identity stamps).
                for s in &body_stmts[..preamble_end] {
                    self.compile_stmt(s)?;
                }

                // Re-stamp __type with the child name (the body's super call
                // stamped it with the parent name).
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_const(Value::String(Arc::from(name)));
                let type_key2 = self.str_const("__type");
                self.emit_u16(Op::STRUCT_SET, type_key2);
                self.emit(Op::DROP);

                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    for method_name in &instance_method_names {
                        common::classes::emit_save_base_method(self.chunk(), this_slot, method_name, line);
                    }
                    common::classes::emit_store_super(self.chunk(), this_slot, &pname, line);
                }

                for (fname, init) in &field_inits {
                    if let Some(init_expr) = init {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.compile_expr(init_expr)?;
                        common::classes::emit_init_field_end(self.chunk(), fname, line);
                    }
                }

                for (mname, mci, _, _) in &instance_methods {
                    if mname.starts_with("__get_") {
                        let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                        common::classes::emit_bind_getter(self.chunk(), this_slot, prop, *mci, line);
                    } else if mname.starts_with("__set_") {
                        let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                        common::classes::emit_bind_setter(self.chunk(), this_slot, prop, *mci, line);
                    } else {
                        common::classes::emit_bind_method_with_aliases(self.chunk(), this_slot, mname, *mci, line);
                    }
                }

                // Auto-init methods from profile (e.g. InitializeComponent for .NET forms).
                // Emitted after method binding so struct_get finds the method,
                // but before user body so controls exist for AddHandler etc.
                let user_body = &body_stmts[preamble_end..];
                let auto_inits = self.profile.auto_init_methods.clone();
                for aim in &auto_inits {
                    let has_method = instance_methods.iter()
                        .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                    if has_method && !body_calls_method(user_body, aim) {
                        common::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
                    }
                }

                // Compile the main body (everything after the preamble).
                for s in user_body {
                    self.compile_stmt(s)?;
                }
            }
        } else {
            // ── Base class ──────────────────────────────────────────────
            common::classes::emit_new_typed_object(self.chunk(), this_slot, name, line);

            // Initialize fields
            for (fname, init) in &field_inits {
                if let Some(init_expr) = init {
                    common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                    self.compile_expr(init_expr)?;
                    common::classes::emit_init_field_end(self.chunk(), fname, line);
                } else {
                    common::classes::emit_init_field_null(self.chunk(), this_slot, fname, line);
                }
            }

            // Bind instance methods
            for (mname, mci, _, _) in &instance_methods {
                if mname.starts_with("__get_") {
                    let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                    common::classes::emit_bind_getter(self.chunk(), this_slot, prop, *mci, line);
                } else if mname.starts_with("__set_") {
                    let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                    common::classes::emit_bind_setter(self.chunk(), this_slot, prop, *mci, line);
                } else {
                    common::classes::emit_bind_method_with_aliases(self.chunk(), this_slot, mname, *mci, line);
                }
            }

            // Auto-init methods from profile (e.g. InitializeComponent for .NET forms).
            let ctor_stmts: &[Statement] = ctor_body
                .as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
            let auto_inits = self.profile.auto_init_methods.clone();
            for aim in &auto_inits {
                let has_method = instance_methods.iter()
                    .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                if has_method && !body_calls_method(ctor_stmts, aim) {
                    common::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
                }
            }

            // Run user constructor body
            if let Some((body, _, _)) = ctor_body {
                for s in body { self.compile_stmt(s)?; }
            }
        }

        // Finalize: instanceof chain
        common::classes::emit_instanceof_chain(&mut self.chunks, self.current, this_slot, name, line);
        common::classes::emit_constructor_return(self.chunk(), this_slot, line);

        let locals = self.scope().next_slot.max(self.chunks[ctor_idx].local_count);
        self.chunks[ctor_idx].local_count = locals;
        self.scopes.pop();
        self.current = saved_cur;
        self.current_class = saved_class2;

        // Store constructor globally and register type
        let ctor_local = self.scope_mut().define(&format!("__{}_ctor", name));
        common::classes::emit_store_constructor(self.chunk(), name, ctor_idx, ctor_local, line);

        // Initialize static fields on the constructor object
        for (fname, init) in &static_field_inits {
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            if let Some(init_expr) = init {
                self.compile_expr(init_expr)?;
            } else {
                self.emit(Op::NULL);
            }
            let fk = self.str_const(fname);
            self.emit_u16(Op::STRUCT_SET, fk);
            self.emit(Op::DROP);
        }

        // Attach static methods to the constructor object
        let mut all_statics: Vec<(String, usize)> = Vec::new();
        for (mname, mci, _, _) in &static_methods {
            common::classes::emit_attach_static_method(self.chunk(), ctor_local, mname, *mci, line);
            all_statics.push((mname.clone(), *mci));
        }

        // Inherit parent's static methods — walk up the chain via PendingClass
        if let Some(parent_name) = parent {
            let mut current_parent = Some(self.canon(parent_name));
            while let Some(ref pname) = current_parent {
                let parent_statics = self.pending_classes.get(pname.as_str())
                    .map(|pc| pc.statics.clone())
                    .unwrap_or_default();
                let next_parent = self.pending_classes.get(pname.as_str())
                    .and_then(|pc| pc.parent.clone());
                for (sname, sci) in &parent_statics {
                    // Only inherit if child doesn't already define it
                    if !all_statics.iter().any(|(n, _)| n == sname) {
                        common::classes::emit_attach_static_method(self.chunk(), ctor_local, sname, *sci, line);
                        all_statics.push((sname.clone(), *sci));
                    }
                }
                current_parent = next_parent;
            }
        }

        // Store statics in PendingClass for grandchildren to inherit
        if let Some(pc) = self.pending_classes.get_mut(name) {
            pc.statics = all_statics;
        }

        let all_methods: Vec<(String, usize)> = method_chunks.iter().map(|(n, c, _, _)| (n.clone(), *c)).collect();
        let parent_str = parent.clone().unwrap_or_default();
        common::classes::register_type(&mut self.chunks, name, &parent_str, fields, all_methods, false, Vec::new(), Some(ctor_idx));

        Ok(())
    }

}
