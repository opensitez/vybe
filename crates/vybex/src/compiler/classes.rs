//! Class, constructor, and free-function compilation.
//!
//! Extracted from `compiler.rs` to keep that file navigable. The
//! methods on this `impl Compiler { ... }` block are private by
//! convention (they're only called from other compiler methods) and
//! crate-private for the `dotnet_register` bridge.

use super::*;
use crate::common::classes::{BaseCall, NormalMethod};

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
        self.function_param_modes.insert(
            cname.clone(),
            params.iter().map(|param| param.pass_by).collect(),
        );
        self.function_min_arity.insert(
            cname.clone(),
            params.iter().take_while(|param| param.default.is_none() && !param.is_rest).count(),
        );
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
        // Function body opens fresh wrt the runtime label_stack —
        // emit_return drains back to this base. Save+restore so nested
        // function decls compose.
        let saved_label_base = self.function_label_base;
        self.function_label_base = self.label_depth;

        // Define params
        for p in params {
            self.define_local_typed(&p.name, p.type_hint.clone());
            // Default parameters: ECMA-262 §15.2.3 — only `undefined`
            // triggers the default (not `null`). The VM now pads
            // missing positional args with `Undefined`, distinct from
            // an explicitly-passed `Null`, so `REF_IS_UNDEFINED` is
            // the correct discriminant for JS. Other languages don't
            // distinguish missing/null and use `REF_IS_NULL` (matches
            // either tag).
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                if self.is_js_profile() {
                    self.emit(Op::REF_IS_UNDEFINED);
                } else {
                    self.emit(Op::REF_IS_NULL);
                }
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
            // is swappable in one place. `ecma:array.push` returns
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
                self.define_local(&format!("__rest_reserved_{}", i));
            }
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            let rest_arr = self.define_local("__rest_arr");
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
            let rs = self.define_local(&slot_name);
            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
            Some(rs)
        } else {
            None
        };
        let ref_out_slots: Vec<u16> = params.iter()
            .filter(|param| matches!(param.pass_by, PassBy::Ref | PassBy::Out))
            .filter_map(|param| self.scope().resolve(&param.name))
            .collect();

        let saved_fn = self.current_func_name.take();
        let saved_rs = self.current_result_slot.take();
        let saved_ref_out = self.current_ref_out_params.take();
        self.current_func_name = Some(name.to_string());
        self.current_result_slot = result_slot;
        self.current_ref_out_params = (!ref_out_slots.is_empty()).then_some(ref_out_slots);

        // ECMA-262 async function semantics: throws inside the body
        // become rejected Promises, normal returns become fulfilled
        // Promises (§27.7.5.3). Wrap the body in TRY_START/TRY_END so
        // uncaught exceptions short-circuit to the Promise.reject
        // path. The `await` opcode handles per-await suspension via
        // JSPI; this wrap just covers terminal throw / return.
        let async_try = if is_async && self.is_js_profile() {
            let line = self.line;
            Some(crate::emitter::errors::emit_try_start(&mut self.chunks[self.current], line))
        } else {
            None
        };

        for s in body { self.compile_stmt(s)?; }

        if let Some(catch_jump) = async_try {
            let line = self.line;
            // Normal exit: wrap return in Promise.resolve(value).
            // The body's compile_stmt may have already emitted RETURNs
            // (early returns); we still need the fall-through path
            // to leave a fulfilled Promise on the stack.
            let chunk = &mut self.chunks[self.current];
            crate::emitter::errors::emit_try_end(chunk, line);
            // After try_end, the body completed normally. Wrap with
            // Promise.resolve(undefined) since no value was left.
            chunk.emit_op(Op::UNDEFINED, line);
            let resolve_idx = self.import("ecma:promise", "resolve");
            self.emit_host_call(resolve_idx, 1);
            self.emit_return();
            // Catch handler — exception value on TOS.
            let chunk = &mut self.chunks[self.current];
            crate::emitter::errors::patch_catch(chunk, catch_jump);
            let reject_idx = self.import("ecma:promise", "reject");
            self.emit_host_call(reject_idx, 1);
            self.emit_return();
        } else if let Some(rs) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit_return_through_finally(1)?;
        } else if self.current_ref_out_params.is_some() {
            self.emit(Op::NULL);
            self.emit_return_through_finally(1)?;
        } else {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
        }

        self.current_func_name = saved_fn;
        self.current_result_slot = saved_rs;
        self.current_ref_out_params = saved_ref_out;

        let locals = self.scope().next_slot.max(self.chunks[func_idx].local_count);
        self.chunks[func_idx].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;
        self.function_label_base = saved_label_base;

        let line = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, uvs.len() as u8, line);
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current].emit(uv.index, line);
        }
        let idx = self.str_const(name);
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);

        if self.is_js_profile() {
            let line = self.line;
            self.emit_common("object.new", 0, line);
            let proto_slot = self.define_local("__js_fn_proto");
            self.emit_u16(Op::LOCAL_SET, proto_slot); self.emit(Op::DROP);

            self.emit_var_get(name);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);

            // ECMA-262 §10.2.4 `length`: number of params before the
            // first one with a default value or rest. Skip rest entirely.
            let length = params.iter().take_while(|p| p.default.is_none() && !p.is_rest).count();
            self.emit_var_get(name);
            self.emit_const(Value::F64(length as f64));
            let length_key = self.str_const("length");
            self.emit_u16(Op::STRUCT_SET, length_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, proto_slot);
            self.emit_var_get(name);
            let ctor_key = self.str_const("constructor");
            self.emit_u16(Op::STRUCT_SET, ctor_key);
            self.emit(Op::DROP);

            self.emit_var_get(name);
            self.emit_u16(Op::LOCAL_GET, proto_slot);
            let proto_key = self.str_const("prototype");
            self.emit_u16(Op::STRUCT_SET, proto_key);
            self.emit(Op::DROP);
        }

        // VB `Handles ctrl.Event` clause on a top-level Sub: register the
        // event handler with the canonical GUI binding. The same canonical
        // emit path serves C# `+=`, JS `addEventListener`, etc.
        for handle in handles {
            let parts: Vec<&str> = handle.splitn(2, '.').collect();
            if parts.len() == 2 {
                let line = self.line;
                let bind_idx = self.import("vybe:gui", common::gui::HOST_FN_BIND_EVENT);
                let ctrl_raw = parts[0].trim();
                let ctrl_canon = self.canon(ctrl_raw);
                let ctrl_key = if ctrl_canon == self.profile.self_keyword
                    || ctrl_canon == "me"
                    || ctrl_canon == "this"
                    || ctrl_canon == "mybase"
                {
                    self.current_class.clone().map(|c| self.canon(&c)).unwrap_or(ctrl_canon)
                } else {
                    ctrl_canon
                };
                self.emit_const(Value::String(Arc::from(ctrl_key.as_str())));
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

    pub(crate) fn compile_class(&mut self, class: &crate::common::classes::NormalClass) -> Result<(), String> {
        // Extract the canonicalised names the orchestration below needs.
        // Canonicalisation happens once here rather than at every caller.
        let cname = self.canon(&class.name);
        let name: &str = &cname;
        let parent_canonical = class.parent.as_ref().map(|p| self.canon(p));
        let parent: &Option<String> = &parent_canonical;

        // Phase 2b.2 complete: passes 1-4 all read NormalClass fields
        // directly. No more ClassMember reconstruction inside
        // compile_class.
        let self_kw = self.profile.self_keyword.clone();
        let ctor_name = self.profile.constructor_name.clone();
        let result_style = self.profile.function_return.clone();

        // Pass 1 (ported to NormalClass): collect fields + initialisers
        // from instance_fields / static_fields, then add backing fields
        // for auto-properties. Reads NormalClass directly; no longer
        // iterates the reconstructed member list.
        let mut fields: Vec<String> = Vec::new();
        let mut field_inits: Vec<(String, Option<String>, Option<Expression>)> = Vec::new();
        let mut static_field_inits: Vec<(String, Option<Expression>)> = Vec::new();
        for f in &class.instance_fields {
            let fname = self.canon(&f.name);
            fields.push(fname.clone());
            field_inits.push((fname, f.type_hint.clone(), f.init.clone()));
        }
        for f in &class.static_fields {
            let fname = self.canon(&f.name);
            static_field_inits.push((fname, f.init.clone()));
        }
        for p in &class.properties {
            // Auto-properties get a backing field named like the property;
            // the runtime reads/writes through auto-emitted __get_/__set_
            // chunks bound later.
            if let Some(auto_field_name) = &p.auto_field {
                let pname_canon = self.canon(auto_field_name);
                if p.is_static {
                    if !static_field_inits.iter().any(|(n, _)| n == &pname_canon) {
                        static_field_inits.push((pname_canon, None));
                    }
                } else if !fields.contains(&pname_canon) {
                    fields.push(pname_canon.clone());
                    field_inits.push((pname_canon, None, None));
                }
            }
        }

        // Events need backing storage on instances so method bodies can
        // read/invoke them via implicit-self resolution (`if (Click != null)
        // Click();`) and subscriptions (`obj.Click += handler`) persist.
        for m in &class.raw_extra_members {
            if let ClassMember::Event { name: ename, .. } = m {
                let fname = self.canon(ename);
                if !fields.contains(&fname) {
                    fields.push(fname.clone());
                    field_inits.push((fname, None, None));
                }
            }
        }

        // Store field list for implicit self resolution
        let static_field_names: Vec<String> = static_field_inits
            .iter()
            .map(|(n, _)| n.clone())
            .collect();
        self.pending_classes.insert(name.to_string(), PendingClass {
            parent: parent.clone(),
            fields: fields.clone(),
            is_value_type: class.is_value_type,
            instance_member_names: class
                .instance_methods
                .iter()
                .map(|method| method.canonical_name.clone())
                .collect(),
            static_fields: static_field_names,
            static_field_types: class.static_fields.iter().filter_map(|f| {
                f.type_hint.as_ref().map(|t| (self.canon(&f.name), Self::normalize_type_hint(t)))
            }).collect(),
            static_method_names: class
                .static_methods
                .iter()
                .map(|m| m.source_name.clone())
                .collect(),
            nested_types: class.raw_extra_members.iter().filter_map(|m| {
                if let ClassMember::NestedType(stmt) = m {
                    match &stmt.kind {
                        StmtKind::ClassDecl { name, .. }
                        | StmtKind::StructDecl { name, .. }
                        | StmtKind::InterfaceDecl { name, .. }
                        | StmtKind::EnumDecl { name, .. } => Some(name.clone()),
                        _ => None,
                    }
                } else {
                    None
                }
            }).collect(),
            statics: Vec::new(), // filled after methods are compiled
        });

        // Compile methods (including constructor body)
        // (name, chunk_idx, is_ctor, is_static)
        let mut method_chunks: Vec<(String, usize, bool, bool)> = Vec::new();
        let saved_class = self.current_class.take();
        let saved_implicit = self.current_class_implicit_self;
        self.current_class = Some(name.to_string());
        self.current_class_implicit_self = class.implicit_self_fields;

        // Pass 2 (ported to NormalClass): pre-register method + property
        // names in `defined_class_methods` so expression-compilation
        // doesn't hijack a method call via the value-method dispatch
        // table. Walks instance_methods + static_methods + properties
        // directly; no reconstructed member iteration.
        for m in class.instance_methods.iter().chain(class.static_methods.iter()) {
            // Use `source_name` so existing compile paths that look up
            // `self.defined_class_methods.contains("ToString")` (from VB
            // call-site compilation) still hit. Canonical-name-only
            // lookups are a Phase 2b.3 concern.
            self.defined_class_methods.insert(self.canon(&m.source_name));
        }
        for p in &class.properties {
            self.defined_class_methods.insert(self.canon(&p.source_name));
        }


        // Pass 3 (ported to NormalClass): compile method chunks,
        // property getter/setter chunks, class-level constants, and
        // nested types. Order matches the former reconstructed-
        // `members` layout so chunk indices stay byte-identical:
        //   instance_methods → static_methods → raw_extra_members
        //   → properties. Constructor body is handled in pass-4 below.

        // --- Instance + static methods ---
        // Each NormalMethod carries the walker's raw modifiers, source
        // name, params, body, return_type, is_generator, is_static
        // flag (implied by which vec the method lives in). That's all
        // the old Method arm needed.
        let mut compile_normal_method = |cc: &mut Compiler, m: &NormalMethod, is_static: bool| -> Result<(), String> {
            let mname = &m.source_name;
            let is_static_init = is_static && mname == "__static_init__";
            let is_ctor = if cc.case_sensitive {
                mname == &ctor_name || (is_static && mname == "new")
            } else {
                mname.eq_ignore_ascii_case(&ctor_name)
                || is_static && mname.eq_ignore_ascii_case("new")
            };

            let user_params: Vec<&Param> = if class.explicit_self_param {
                m.params.iter().skip(1).collect()
            } else {
                m.params.iter().collect()
            };
            let uses_js_this = cc.is_js_profile();
            let has_rest = user_params.last().map_or(false, |p| p.is_rest);
            let arity = if has_rest {
                255u8
            } else if is_static_init {
                user_params.len() as u8
            } else if uses_js_this {
                user_params.len() as u8
            } else {
                (user_params.len() + 1) as u8
            };

            let ci = cc.chunks.len();
            let mut chunk = common::functions::create_function_chunk(mname, arity);
            chunk.is_generator = m.is_generator;
            if m.is_generator {
                let cname_canon = cc.canon(mname);
                cc.generator_functions.insert(cname_canon);
            }
            cc.chunks.push(chunk);
            cc.scopes.push(Scope::new_function());
            let saved = cc.current;
            cc.current = ci;

            if !uses_js_this && !is_static_init {
                cc.define_local(&self_kw);
            }
            for p in &user_params { cc.define_local(&p.name); }

            // Rest param preamble: collect excess args into an array.
            // Mirrors `compile_function_decl` so methods with `params T[] xs`
            // get the same variadic semantics as free functions.
            if has_rest {
                let rest_name = &user_params.last().unwrap().name;
                let rest_slot = cc.scope().resolve(rest_name).unwrap();
                let line = cc.line;
                let max_rest = 16u16;
                for i in 1..max_rest {
                    cc.define_local(&format!("__rest_reserved_{}", i));
                }
                common::collections::emit_array_new(&mut cc.chunks, cc.current, 0, line);
                let rest_arr = cc.define_local("__rest_arr");
                cc.emit_u16(Op::LOCAL_SET, rest_arr); cc.emit(Op::DROP);
                let mut done_patches: Vec<usize> = Vec::new();
                for i in 0..max_rest {
                    let slot = rest_slot + i;
                    cc.emit_u16(Op::LOCAL_GET, slot);
                    cc.emit(Op::REF_IS_NULL);
                    done_patches.push(cc.emit_jump(Op::BR_IF_TRUE));
                    cc.emit_u16(Op::LOCAL_GET, rest_arr);
                    cc.emit_u16(Op::LOCAL_GET, slot);
                    common::collections::emit_push(&mut cc.chunks, cc.current, line);
                    cc.emit(Op::DROP);
                }
                for p in done_patches { cc.patch_jump(p); }
                cc.emit_u16(Op::LOCAL_GET, rest_arr);
                cc.emit_u16(Op::LOCAL_SET, rest_slot);
                cc.emit(Op::DROP);
            }

            // Default parameters (C# `string greeting = "Hello"`): if
            // the slot is null/undefined when the method runs, install
            // the default. JS profile uses `REF_IS_UNDEFINED` (only
            // explicit `undefined` triggers); other languages use
            // `REF_IS_NULL` which matches either tag.
            for p in &user_params {
                if let Some(ref default) = p.default {
                    let slot = cc.scope().resolve(&p.name).unwrap();
                    cc.emit_u16(Op::LOCAL_GET, slot);
                    if uses_js_this {
                        cc.emit(Op::REF_IS_UNDEFINED);
                    } else {
                        cc.emit(Op::REF_IS_NULL);
                    }
                    let has_val = cc.emit_jump(Op::BR_IF_FALSE);
                    cc.compile_expr(default)?;
                    cc.emit_u16(Op::LOCAL_SET, slot); cc.emit(Op::DROP);
                    cc.patch_jump(has_val);
                }
            }

            if is_ctor {
                for s in &m.body { cc.compile_stmt(s)?; }
                if let Some(slot) = cc.scope().resolve(&self_kw).or_else(|| cc.scope().resolve_ci(&self_kw)) {
                    cc.emit_u16(Op::LOCAL_GET, slot);
                    cc.emit(Op::RETURN);
                }
            } else if m.return_type.is_some() && result_style == ReturnStyle::ResultSlot {
                let slot_name = cc.profile.result_slot_name.clone();
                let rs = cc.define_local(&slot_name);
                let returns_self_type = m.return_type.as_deref()
                    .is_some_and(|rt| rt.eq_ignore_ascii_case(&class.name));
                if returns_self_type && body_has_result_member_assign(&m.body) {
                    cc.emit_var_get(&class.name);
                    cc.emit_u8(Op::CALL_REF, 0);
                } else {
                    cc.emit(Op::NULL);
                }
                cc.emit_u16(Op::LOCAL_SET, rs); cc.emit(Op::DROP);
                let saved_fn = cc.current_func_name.take();
                let saved_rs = cc.current_result_slot.take();
                cc.current_func_name = Some(mname.clone());
                cc.current_result_slot = Some(rs);
                for s in &m.body { cc.compile_stmt(s)?; }
                cc.current_func_name = saved_fn;
                cc.current_result_slot = saved_rs;
                cc.emit_u16(Op::LOCAL_GET, rs);
                cc.emit(Op::RETURN);
            } else {
                for s in &m.body { cc.compile_stmt(s)?; }
                let line = cc.line;
                common::functions::emit_function_epilogue(&mut cc.chunks[ci], line);
            }

            let locals = cc.scope().next_slot.max(cc.chunks[ci].local_count);
            cc.chunks[ci].local_count = locals;
            cc.scopes.pop();
            cc.current = saved;

            // `[Symbol.iterator]()` lands as source_name="Symbol.iterator"
            // but must be reachable under canonical `iterator`. For
            // ordinary names, keep source-derived binding so language
            // dunders (PHP `__toString`, Python `__str__`) bind under
            // their literal name + cross-lang aliases via
            // `emit_bind_method_with_aliases`. Only the Symbol.* pseudo-
            // names need the canonical rewrite.
            let bound_name = if mname.starts_with("Symbol.") && !m.canonical_name.is_empty() {
                m.canonical_name.clone()
            } else {
                cc.canon(mname)
            };
            method_chunks.push((bound_name, ci, is_ctor, is_static));
            Ok(())
        };

        for m in &class.instance_methods {
            compile_normal_method(self, m, false)?;
        }
        for m in &class.static_methods {
            compile_normal_method(self, m, true)?;
        }

        // --- Events / Consts / NestedTypes (from raw_extra_members) ---
        for m in &class.raw_extra_members {
            match m {
                ClassMember::Const { name: cname, value, .. } => {
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

        // --- Properties: getter → __get_<prop>, setter → __set_<prop> ---
        for p in &class.properties {
            // Auto-properties are handled as plain fields in pass-1.
            if p.auto_field.is_some() { continue; }
            let pname_canon = self.canon(&p.source_name);
            let prop_is_static = p.is_static;

            if let Some(getter) = &p.getter {
                let get_name = format!("__get_{}", pname_canon);
                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk(
                    &get_name,
                    if prop_is_static { 0 } else { 1 },
                );
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;
                if !prop_is_static {
                    self.define_local(&self_kw);
                }

                if getter.body.is_empty() {
                    // Auto-property getter: return backing field
                    if let Some(slot) = self.scope().resolve(&self_kw) {
                        self.emit_u16(Op::LOCAL_GET, slot);
                        let backing = self.str_const(&format!("__{}", pname_canon));
                        self.emit_u16(Op::STRUCT_GET, backing);
                        self.emit(Op::RETURN);
                    }
                } else {
                    let slot_name = self.profile.result_slot_name.clone();
                    let rs = self.define_local(&slot_name);
                    self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
                    let saved_fn = self.current_func_name.take();
                    let saved_rs = self.current_result_slot.take();
                    self.current_func_name = Some(p.source_name.clone());
                    self.current_result_slot = Some(rs);
                    for s in &getter.body { self.compile_stmt(s)?; }
                    self.current_func_name = saved_fn;
                    self.current_result_slot = saved_rs;
                    self.emit_u16(Op::LOCAL_GET, rs);
                    self.emit(Op::RETURN);
                }

                let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
                self.chunks[ci].local_count = locals;
                self.scopes.pop();
                self.current = saved;
                method_chunks.push((get_name, ci, false, prop_is_static));
            }

            if let Some(setter) = &p.setter {
                let set_name = format!("__set_{}", pname_canon);
                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk(
                    &set_name,
                    if prop_is_static { 1 } else { 2 },
                );
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;
                if !prop_is_static {
                    self.define_local(&self_kw);
                }
                let value_param_name = setter.params.first()
                    .map(|p| p.name.clone())
                    .unwrap_or_else(|| "value".to_string());
                self.define_local(&value_param_name);

                if setter.body.is_empty() {
                    // Auto-property setter: set backing field
                    if let Some(self_slot) = self.scope().resolve(&self_kw) {
                        self.emit_u16(Op::LOCAL_GET, self_slot);
                        if let Some(val_slot) = self.scope().resolve(&value_param_name) {
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                        }
                        let backing = self.str_const(&format!("__{}", pname_canon));
                        self.emit_u16(Op::STRUCT_SET, backing);
                        self.emit(Op::DROP);
                    }
                } else {
                    for s in &setter.body { self.compile_stmt(s)?; }
                }

                let line = self.line;
                common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
                self.chunks[ci].local_count = locals;
                self.scopes.pop();
                self.current = saved;
                method_chunks.push((set_name, ci, false, prop_is_static));
            }
        }

        self.current_class = saved_class;
        self.current_class_implicit_self = saved_implicit;

        // Pass 4 (ported to NormalClass): find the constructor's body,
        // params, and base-call args for the orchestration below.
        //
        // Priority order matches the former `members` iteration:
        // Constructor (reconstruction placed it first) wins over any
        // Method named `New` / `constructor` / etc. Languages that use
        // a conventional method-name constructor (VB `Sub New`, C#
        // `ClassName(...)`) already land in `NormalClass.constructor`
        // through the walker's normalize_class, so the secondary
        // method-scan is only a safety net for languages that haven't
        // categorized a ctor-shaped method into the Constructor slot.
        let _ctor = method_chunks.iter().find(|(_, _, is_ctor, _)| *is_ctor);

        // Convert NormalConstructor.base_call → Option<Vec<Expression>>
        // for the downstream orchestration (which still expects the AST
        // shape). Only `BaseCall::Explicit(args)` carries arguments;
        // `BaseCall::Auto` signals "auto-insert a zero-arg parent call"
        // (C#/Pascal style) and is handled separately from base_args.
        // Walker already turned explicit super() / MyBase.New(...) into
        // `BaseCall::Explicit(args)`.
        let ctor_base_args_from_nc: Option<Vec<Expression>> = class.constructor.as_ref().and_then(|c| {
            if let BaseCall::Explicit(args) = &c.base_call {
                Some(args.iter().map(|a| a.value.clone()).collect())
            } else {
                None
            }
        });
        let ctor_auto_base = class.constructor.as_ref()
            .map(|c| matches!(c.base_call, BaseCall::Auto))
            .unwrap_or(false);

        let ctor_body: Option<(&Vec<Statement>, &Vec<Param>, Option<&Vec<Expression>>)> = {
            if let Some(c) = &class.constructor {
                Some((&c.body, &c.params, ctor_base_args_from_nc.as_ref()))
            } else {
                // Safety-net scan of instance + static methods for a
                // method-name constructor (e.g. a language walker that
                // left a `New` method in the methods vec instead of
                // promoting it to NormalClass.constructor).
                class.instance_methods.iter()
                    .chain(class.static_methods.iter())
                    .find_map(|m| {
                        let is_static = class.static_methods.iter()
                            .any(|sm| sm.source_name == m.source_name);
                        let is_ctor = if self.case_sensitive {
                            m.source_name == ctor_name || (is_static && m.source_name == "new")
                        } else {
                            m.source_name.eq_ignore_ascii_case(&ctor_name)
                            || is_static && m.source_name.eq_ignore_ascii_case("new")
                        };
                        if is_ctor && !m.body.is_empty() {
                            Some((&m.body, &m.params, None))
                        } else {
                            None
                        }
                    })
            }
        };

        let user_params: Vec<String> = ctor_body.map(|(_, params, _)| {
            if class.explicit_self_param {
                params.iter().skip(1).map(|p| p.name.clone()).collect()
            } else {
                params.iter().map(|p| p.name.clone()).collect()
            }
        }).unwrap_or_default();
        // ECMA-262 §15.7.10: a derived class without an explicit constructor
        // gets the implicit `constructor(...args) { super(...args); }`. We
        // model that by synthesizing 16 forwarding slots — the same
        // MAX_VARIADIC convention the spread-call path uses. The parent
        // ctor truncates to its own arity (Undefined-padded for the rest),
        // so the chain composes correctly through any depth of mixins.
        const IMPLICIT_CTOR_FORWARD_ARGS: u8 = 16;
        let synthesized_forward_args = ctor_body.is_none() && parent.is_some();
        let user_arity = if synthesized_forward_args {
            IMPLICIT_CTOR_FORWARD_ARGS
        } else {
            user_params.len() as u8
        };

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
        let saved_implicit2 = self.current_class_implicit_self;
        self.current = ctor_idx;
        self.current_class = Some(name.to_string());
        self.current_class_implicit_self = class.implicit_self_fields;

        // Define user params (slot 1..N), then this (slot N+1)
        // Also handle default parameter values from the Param structs.
        let ctor_param_defaults: Vec<Option<Expression>> = ctor_body.map(|(_, params, _)| {
            let skip = if class.explicit_self_param { 1 } else { 0 };
            params.iter().skip(skip).map(|p| p.default.clone()).collect()
        }).unwrap_or_default();
        for (i, p) in user_params.iter().enumerate() {
            self.define_local(p);
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
        // Synthesized implicit ctor — reserve the forwarding slots so
        // `this_slot` lands above them, matching `user_arity`.
        if synthesized_forward_args {
            for i in 0..IMPLICIT_CTOR_FORWARD_ARGS {
                self.define_local(&format!("__implicit_arg_{}", i));
            }
        }
        self.define_local(&self_kw); // this_slot = user_arity
        let this_slot = user_arity as u16;
        if self.is_js_profile() {
            let js_this = self.str_const("__js_this");
            self.emit_u16(Op::GLOBAL_GET, js_this);
            self.emit_u16(Op::LOCAL_SET, this_slot);
            self.emit(Op::DROP);
        }

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

            // Determine if auto base call should kick in: walker flagged
            // `BaseCall::Auto` on the ctor (C#/Pascal) + parent exists +
            // body doesn't already contain a super call (VB's walker
            // injects `MyBase.New()` into the body, so this check short-
            // circuits there without a profile flag).
            let has_explicit_base = ctor_body.as_ref().map_or(false, |(_, _, ba)| ba.is_some());
            let auto_base_needed = !has_explicit_base
                && ctor_body.is_some()
                && ctor_auto_base
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
                        // Walk local/upvalue scope first so parent
                        // resolved from a closure capture (e.g.
                        // `(Base) => class extends Base`) works.
                        // Falls through to global_get for normal
                        // top-level classes.
                        self.emit_var_get(&pname);
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
                        self.emit_var_get(&pname);
                        self.emit_u8(Op::CALL, 0);
                        self.emit_u16(Op::LOCAL_SET, this_slot);
                        self.emit(Op::DROP);
                    }
                }
                // else: JS pattern — body calls super() itself, sets this_slot
            } else {
                // No explicit constructor — auto-call parent with all
                // received args (per ECMA-262 §15.7.10 implicit ctor:
                // `constructor(...args) { super(...args); }`). The
                // synthesized arity is IMPLICIT_CTOR_FORWARD_ARGS; the
                // parent ctor truncates to its own arity.
                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    self.emit_var_get(&pname);
                    for i in 0..user_arity {
                        self.emit_u16(Op::LOCAL_GET, i as u16);
                    }
                    self.emit_u8(Op::CALL_REF, user_arity);
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

                // JS prototype link: `this.__proto__ = ClassName.prototype`.
                // Lets `Object.getPrototypeOf(d) === Dog.prototype` and
                // prototype-chain walks (`hasIn`, `to_primitive`'s
                // proto traversal) reach methods declared on the class.
                // Skip if the global isn't bound yet (some component
                // types stamp __type without going through class decl).
                if self.is_js_profile() {
                    let class_global = self.str_const(name);
                    let prototype_key = self.str_const("prototype");
                    let proto_link_key = self.str_const("__proto__");
                    let proto_local = self.define_local(&format!("__{}_link_proto", name));
                    self.emit_u16(Op::GLOBAL_GET, class_global);
                    self.emit_u16(Op::STRUCT_GET, prototype_key);
                    self.emit_u16(Op::LOCAL_SET, proto_local);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, proto_local);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit_u16(Op::LOCAL_GET, this_slot);
                    self.emit_u16(Op::LOCAL_GET, proto_local);
                    self.emit_u16(Op::STRUCT_SET, proto_link_key);
                    self.emit(Op::DROP);
                    self.patch_jump(skip);
                }

                for (fname, type_hint, init) in &field_inits {
                    if let Some(init_expr) = init {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.compile_expr(init_expr)?;
                        common::classes::emit_init_field_end(self.chunk(), fname, line);
                    } else if class.is_value_type {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.emit_default_value_for_type_hint(type_hint.as_deref());
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

                // Auto-init methods declared on the class (e.g. InitializeComponent
                // for .NET forms). Walker-populated via normalize_class.
                // Emitted after method binding but before user body.
                {
                    let ctor_stmts: &[Statement] = ctor_body
                        .as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
                    for aim in &class.auto_init_methods {
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

                // Re-stamp WASM GC type_id with the child's id. Each
                // ancestor's super-call stamped its own type_id, so by
                // the time we land here `this.type_id` reflects the
                // direct parent — overwrite with this class's id so
                // multi-level chains (`C extends B extends A`) end up
                // with `c.type_id == C_id`. `is_subtype` then walks the
                // parent pointers to confirm `C` is a subtype of both
                // `B` and `A`. Use `self.canon(name)` so the __tid_
                // global key matches what the VM populated at type-
                // table load time.
                let tid_key = self.str_const(&format!("__tid_{}", self.canon(name)));
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_u16(Op::GLOBAL_GET, tid_key);
                self.emit(Op::SET_TYPE_ID);
                self.emit(Op::DROP);

                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    for method_name in &instance_method_names {
                        common::classes::emit_save_base_method(self.chunk(), this_slot, method_name, line);
                    }
                    common::classes::emit_store_super(self.chunk(), this_slot, &pname, line);
                }

                for (fname, type_hint, init) in &field_inits {
                    if let Some(init_expr) = init {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.compile_expr(init_expr)?;
                        common::classes::emit_init_field_end(self.chunk(), fname, line);
                    } else if class.is_value_type {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.emit_default_value_for_type_hint(type_hint.as_deref());
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

                // Auto-init methods declared on the class (e.g. InitializeComponent
                // for .NET forms). Walker-populated via normalize_class.
                // Emitted after method binding so struct_get finds the method,
                // but before user body so controls exist for AddHandler etc.
                let user_body = &body_stmts[preamble_end..];
                for aim in &class.auto_init_methods {
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
            // Pass canon'd name so __type stamping and the __tid_<canon>
            // global lookup match what `register_type` stored above.
            let canon_name = self.canon(name);
            common::classes::emit_new_typed_object(self.chunk(), this_slot, &canon_name, line);

            // Initialize fields
            for (fname, type_hint, init) in &field_inits {
                if let Some(init_expr) = init {
                    common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                    self.compile_expr(init_expr)?;
                    common::classes::emit_init_field_end(self.chunk(), fname, line);
                } else if class.is_value_type {
                    common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                    self.emit_default_value_for_type_hint(type_hint.as_deref());
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

            // JS prototype link: `this.__proto__ = ClassName.prototype`.
            // Mirrors the wiring in the child-class branch (line ~740);
            // base classes need the same link so prototype-chain walks
            // (`hasIn`, `to_primitive`'s ToPrimitive proto traversal,
            // `Object.getPrototypeOf`) reach methods declared on the
            // class. Skip if the global isn't bound yet.
            if self.is_js_profile() {
                let class_global = self.str_const(name);
                let prototype_key = self.str_const("prototype");
                let proto_link_key = self.str_const("__proto__");
                let proto_local = self.define_local(&format!("__{}_link_proto_base", name));
                self.emit_u16(Op::GLOBAL_GET, class_global);
                self.emit_u16(Op::STRUCT_GET, prototype_key);
                self.emit_u16(Op::LOCAL_SET, proto_local);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, proto_local);
                self.emit(Op::REF_IS_NULL);
                let skip = self.emit_jump(Op::BR_IF_TRUE);
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_u16(Op::LOCAL_GET, proto_local);
                self.emit_u16(Op::STRUCT_SET, proto_link_key);
                self.emit(Op::DROP);
                self.patch_jump(skip);
            }

            // Auto-init methods declared on the class (e.g. InitializeComponent
            // for .NET forms). Walker-populated via normalize_class.
            let ctor_stmts: &[Statement] = ctor_body
                .as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
            for aim in &class.auto_init_methods {
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
        // Capture the ctor's upvalues before popping the scope —
        // mixin / closure-captured parents (`(Base) => class extends
        // Base`) materialize as upvalue references on the ctor body
        // and must be threaded into the runtime function ref.
        let ctor_upvalues = self.scope().upvalues.clone();
        self.scopes.pop();
        self.current = saved_cur;
        self.current_class = saved_class2;
        self.current_class_implicit_self = saved_implicit2;

        // Store constructor globally and register type. Emit upvalue
        // captures inline with REF_FUNC so closure-bound parents
        // resolve at runtime (mirrors `compile_function_decl`'s
        // `emit_ref_func` pattern).
        let ctor_local = self.define_local(&format!("__{}_ctor", name));
        let uv_pairs: Vec<(bool, u8)> = ctor_upvalues
            .iter()
            .map(|uv| (uv.is_local, uv.index))
            .collect();
        common::classes::emit_store_constructor_with_upvalues(
            self.chunk(), name, ctor_idx, ctor_local, &uv_pairs, line,
        );

        if self.is_js_profile() {
            self.emit_common("object.new", 0, line);
            let proto_local = self.define_local(&format!("__{}_prototype", name));
            self.emit_u16(Op::LOCAL_SET, proto_local); self.emit(Op::DROP);

            if let Some(parent_name) = parent {
                let pname = self.canon(parent_name);
                self.emit_var_get(&pname);
                let parent_proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_GET, parent_proto_key);
                let parent_proto_local = self.define_local(&format!("__{}_parent_prototype", name));
                self.emit_u16(Op::LOCAL_SET, parent_proto_local); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, parent_proto_local);
                self.emit(Op::REF_IS_NULL);
                let skip_parent_proto = self.emit_jump(Op::BR_IF_TRUE);
                self.emit_u16(Op::LOCAL_GET, proto_local);
                self.emit_u16(Op::LOCAL_GET, parent_proto_local);
                let proto_link_key = self.str_const("__proto__");
                self.emit_u16(Op::STRUCT_SET, proto_link_key);
                self.emit(Op::DROP);
                self.patch_jump(skip_parent_proto);
            }

            self.emit_u16(Op::LOCAL_GET, proto_local);
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            let ctor_key = self.str_const("constructor");
            self.emit_u16(Op::STRUCT_SET, ctor_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_u16(Op::LOCAL_GET, proto_local);
            let proto_key = self.str_const("prototype");
            self.emit_u16(Op::STRUCT_SET, proto_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_const(Value::String(Arc::from(name)));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);
        }

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

        // Attach nested types onto the constructor object so `Outer.Inner`
        // resolves to the nested class constructor through the same shared
        // class-object path as static methods.
        let nested_types = self.pending_classes.get(name)
            .map(|pc| pc.nested_types.clone())
            .unwrap_or_default();
        for nested in nested_types {
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            let nested_canon = self.canon(&nested);
            let nested_idx = self.str_const(&nested_canon);
            self.emit_u16(Op::GLOBAL_GET, nested_idx);
            let key = self.str_const(&nested_canon);
            self.emit_u16(Op::STRUCT_SET, key);
            self.emit(Op::DROP);
        }

        // Attach static methods to the constructor object
        let mut all_statics: Vec<(String, usize)> = Vec::new();
        for (mname, mci, _, _) in &static_methods {
            common::classes::emit_attach_static_method(self.chunk(), ctor_local, mname, *mci, line);
            all_statics.push((mname.clone(), *mci));
        }

        // Synthetic static constructor hook from language walkers.
        if let Some((_, static_init_ci, _, _)) = static_methods
            .iter()
            .find(|(mname, _, _, _)| mname.eq_ignore_ascii_case("__static_init__"))
        {
            let line = self.line;
            self.emit_u16(Op::REF_FUNC, *static_init_ci as u16);
            self.chunk().emit(0, line);
            self.emit_u8(Op::CALL_REF, 0);
            self.emit(Op::DROP);
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

        // Attach instance methods to the class object so static
        // `super.method()` dispatch can reach them. ECMA-262
        // §13.3.7.4 / §10.2.4 / §10.2.10.2: `super` resolves via
        // [[HomeObject]].[[Prototype]] (the parent class's prototype),
        // NOT the instance prototype chain. Multi-level inheritance
        // (C → B → A) needs B.method when called from C, A.method
        // when called from B — both at compile time. We mirror the
        // method bindings on the class constructor so
        // `GLOBAL_GET(ParentClass) ~ STRUCT_GET(method)` returns the
        // class-level method ref. Instance bindings are unchanged
        // (still per-instance for `this.method()` and override).
        for (mname, mci, _, _) in &instance_methods {
            // Skip getter/setter wrappers — they're bound differently.
            if mname.starts_with("__get_") || mname.starts_with("__set_") {
                continue;
            }
            common::classes::emit_attach_static_method(self.chunk(), ctor_local, mname, *mci, line);
        }

        let all_methods: Vec<(String, usize)> = method_chunks.iter().map(|(n, c, _, _)| (n.clone(), *c)).collect();
        // Canonicalise per language case-sensitivity: case-insensitive
        // languages (VB/Pascal/COBOL/PHP) lowercase here, case-sensitive
        // (JS/TS/Python/C#) preserve. Registry stores whatever the walker
        // produced; runtime `Op::REF_TEST` looks up by the same canon.
        let canon_name = self.canon(name);
        let canon_parent = parent.as_ref().map(|p| self.canon(p)).unwrap_or_default();
        common::classes::register_type(&mut self.chunks, &canon_name, &canon_parent, fields, all_methods, false, Vec::new(), Some(ctor_idx));

        Ok(())
    }

}

fn body_has_result_member_assign(body: &[Statement]) -> bool {
    body.iter().any(stmt_has_result_member_assign)
}

fn stmt_has_result_member_assign(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Assign { targets, .. } => targets.iter().any(expr_is_result_member),
        StmtKind::Block(body)
        | StmtKind::FunctionDecl { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => body_has_result_member_assign(body),
        StmtKind::If { then_body, elifs, else_body, .. } => {
            body_has_result_member_assign(then_body)
                || elifs.iter().any(|(_, body)| body_has_result_member_assign(body))
                || else_body.as_ref().is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::For { init, body, .. } => {
            init.as_ref().is_some_and(|stmt| stmt_has_result_member_assign(stmt))
                || body_has_result_member_assign(body)
        }
        StmtKind::ForIn { body, else_body, .. }
        | StmtKind::While { body, else_body, .. } => {
            body_has_result_member_assign(body)
                || else_body.as_ref().is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::DoWhile { body, .. } => body_has_result_member_assign(body),
        StmtKind::Switch { cases, default, .. } => {
            cases.iter().any(|case| body_has_result_member_assign(&case.body))
                || default.as_ref().is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::Try { body, catches, else_body, finally, .. } => {
            body_has_result_member_assign(body)
                || catches.iter().any(|catch| body_has_result_member_assign(&catch.body))
                || else_body.as_ref().is_some_and(|body| body_has_result_member_assign(body))
                || finally.as_ref().is_some_and(|body| body_has_result_member_assign(body))
        }
        _ => false,
    }
}

fn expr_is_result_member(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Member { object, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Result"))
    )
}
