//! `compile_expr` — the `ExprKind` dispatch. The single largest
//! method in the compiler, split out so edits to other concerns
//! (calls, classes, statements) don't churn a multi-thousand-line
//! file.

use super::*;

impl Compiler {
    /// Push a reference to the class constructor global `ctor_global`.
    /// When the profile opts into autoloading (`supports_autoload`), emit
    /// the autoload-fallback sequence keyed on `autoload_name`; otherwise a
    /// plain `GLOBAL_GET`. Keeps the constructor-resolution sites
    /// language-agnostic — the PHP autoload logic lives in the php emitter.
    fn emit_constructor_global_ref(&mut self, ctor_global: &str, autoload_name: &str) {
        if self.profile.supports_autoload {
            let line = self.line;
            crate::emitter::php::autoload_adapter::emit_constructor_ref_with_autoload(
                self.chunk(),
                ctor_global,
                autoload_name,
                line,
            );
        } else {
            let idx = self.str_const(ctor_global);
            self.emit_u16(Op::GLOBAL_GET, idx);
        }
    }

    /// Like [`Self::emit_constructor_global_ref`] but resolves a primary
    /// constructor global then an optional fallback before autoloading.
    fn emit_dynamic_constructor_global_ref(
        &mut self,
        primary_ctor_global: &str,
        fallback_ctor_global: Option<&str>,
        autoload_name: &str,
    ) {
        if self.profile.supports_autoload {
            let line = self.line;
            crate::emitter::php::autoload_adapter::emit_dynamic_constructor_ref_with_autoload(
                self.chunk(),
                primary_ctor_global,
                fallback_ctor_global,
                autoload_name,
                line,
            );
        } else {
            let idx = self.str_const(primary_ctor_global);
            self.emit_u16(Op::GLOBAL_GET, idx);
        }
    }

    fn emit_js_member_fallback_get(&mut self, obj_slot: u16, field_name: &str) {
        let lookup = self.str_const("__vybe_js_get_method");
        let getter_slot = self.define_local("__js_member_getter");
        let accessor_name = format!("__get_{}", field_name);

        self.emit_u16(Op::GLOBAL_GET, lookup);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(accessor_name.as_str())));
        self.emit_u8(Op::CALL_REF, 2);
        self.emit_u16(Op::LOCAL_SET, getter_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit(Op::REF_IS_UNDEFINED);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::GLOBAL_GET, lookup);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(field_name)));
        self.emit_u8(Op::CALL_REF, 2);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u8(Op::CALL_REF, 1);
        self.chunk().emit_end(line);
    }

    fn emit_js_import_meta_object(&mut self) {
        let global_name = "__js_import_meta";
        let global_idx = self.str_const(global_name);
        let meta_slot = self.define_local("__js_import_meta_value");

        self.emit_u16(Op::GLOBAL_GET, global_idx);
        self.emit_u16(Op::LOCAL_SET, meta_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, meta_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);
        common::dict::emit_new(&mut self.chunks, self.current, line);
        let init_slot = self.define_local("__js_import_meta_init");
        self.emit_u16(Op::LOCAL_SET, init_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_const(Value::String(Arc::from("")));
        let url_key = self.str_const("url");
        self.emit_u16(Op::STRUCT_SET, url_key);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_u16(Op::GLOBAL_SET, global_idx);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_u16(Op::LOCAL_SET, meta_slot);
        self.emit(Op::DROP);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, meta_slot);
        self.emit(Op::REF_IS_UNDEFINED);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);
        common::dict::emit_new(&mut self.chunks, self.current, line);
        let init_slot = self.define_local("__js_import_meta_init");
        self.emit_u16(Op::LOCAL_SET, init_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_const(Value::String(Arc::from("")));
        let url_key = self.str_const("url");
        self.emit_u16(Op::STRUCT_SET, url_key);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_u16(Op::GLOBAL_SET, global_idx);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_u16(Op::LOCAL_SET, meta_slot);
        self.emit(Op::DROP);

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, meta_slot);
    }

    fn try_compile_fortran_array_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if self.profile.name != "fortran" {
            return Ok(false);
        }
        if !matches!(
            op,
            BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow
        ) {
            return Ok(false);
        }

        let left_is_array = self.expr_is_array_like(left);
        let right_is_array = self.expr_is_array_like(right);
        if !left_is_array && !right_is_array {
            return Ok(false);
        }

        let line = self.line;
        let left_slot = self.define_local("__fortran_array_binop_left");
        self.compile_expr(left)?;
        self.emit_u16(Op::LOCAL_SET, left_slot);
        self.emit(Op::DROP);

        let right_slot = self.define_local("__fortran_array_binop_right");
        self.compile_expr(right)?;
        self.emit_u16(Op::LOCAL_SET, right_slot);
        self.emit(Op::DROP);

        let iter_slot = if left_is_array { left_slot } else { right_slot };
        let result_slot = self.define_local("__fortran_array_binop_result");
        let idx_slot = self.define_local("__fortran_array_binop_idx");

        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);

        let state = common::loops::emit_for_in_start(
            &mut self.chunks,
            self.current,
            iter_slot,
            idx_slot,
            line,
        );
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, result_slot);
        if left_is_array {
            self.emit_u16(Op::LOCAL_GET, left_slot);
            self.emit_u16(Op::LOCAL_GET, idx_slot);
            common::collections::emit_get(&mut self.chunks, self.current, line);
        } else {
            self.emit_u16(Op::LOCAL_GET, left_slot);
        }
        if right_is_array {
            self.emit_u16(Op::LOCAL_GET, right_slot);
            self.emit_u16(Op::LOCAL_GET, idx_slot);
            common::collections::emit_get(&mut self.chunks, self.current, line);
        } else {
            self.emit_u16(Op::LOCAL_GET, right_slot);
        }
        self.compile_binop(op);
        common::collections::emit_push(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, state, line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(true)
    }

    // ════════════════════════════════════════════════════════════════════════
    // Expression compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn emit_generator_entry_control(&mut self, control_slot: u16) -> Result<(), String> {
        self.emit_u16(Op::LOCAL_GET, control_slot);
        self.emit(Op::REF_IS_OBJECT);
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        let marker_key = self.str_const("__vybe_generator_control");
        self.emit_u16(Op::STRUCT_GET, marker_key);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        let op_key = self.str_const("op");
        self.emit_u16(Op::STRUCT_GET, op_key);
        self.emit_const(Value::String(Arc::from("throw")));
        self.emit(Op::STR_EQUALS);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        let value_key = self.str_const("value");
        self.emit_u16(Op::STRUCT_GET, value_key);
        self.emit(Op::THROW);

        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, control_slot);
        self.emit_u16(Op::STRUCT_GET, op_key);
        self.emit_const(Value::String(Arc::from("return")));
        self.emit(Op::STR_EQUALS);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        self.emit_u16(Op::STRUCT_GET, value_key);
        self.emit_return_through_finally(1)?;

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        Ok(())
    }

    fn emit_generator_return_from_resume_slot(&mut self, resume_slot: u16) -> Result<(), String> {
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        let value_key = self.str_const("value");
        self.emit_u16(Op::STRUCT_GET, value_key);
        self.emit_return_through_finally(1)
    }

    pub(super) fn emit_generator_resume_value(&mut self) -> Result<(), String> {
        let resume_slot = self.define_local("__yield_resume");
        self.emit_u16(Op::LOCAL_SET, resume_slot);
        self.emit(Op::DROP);

        let result_slot = self.define_local("__yield_resume_value");
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);

        let line = self.line;
        let marker_slot = self.define_local("__yield_resume_is_control");
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        let marker_key = self.str_const("__vybe_generator_control");
        self.emit_u16(Op::STRUCT_GET, marker_key);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, marker_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, marker_slot);
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        let op_key = self.str_const("op");
        self.emit_u16(Op::STRUCT_GET, op_key);
        self.emit_const(Value::String(Arc::from("throw")));
        self.emit(Op::STR_EQUALS);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit(Op::I32_AND);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, resume_slot);
        let value_key = self.str_const("value");
        self.emit_u16(Op::STRUCT_GET, value_key);
        self.emit(Op::THROW);

        self.chunk().emit_else(line);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, marker_slot);
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        self.emit_u16(Op::STRUCT_GET, op_key);
        self.emit_const(Value::String(Arc::from("return")));
        self.emit(Op::STR_EQUALS);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit(Op::I32_AND);
        self.chunk().emit_if(line);

        self.emit_generator_return_from_resume_slot(resume_slot)?;

        self.chunk().emit_else(line);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(())
    }

    fn generator_keyed_yield_parts<'a>(
        &self,
        expr: &'a Expression,
    ) -> Option<(&'a Expression, &'a Expression)> {
        let ExprKind::Object(props) = &expr.kind else {
            return None;
        };

        let mut has_marker = false;
        let mut key_expr = None;
        let mut value_expr = None;
        for prop in props {
            if let ObjectProperty::KeyValue { key, value } = prop {
                match &key.kind {
                    ExprKind::Lit(Literal::Str(name)) if name == "__vybe_generator_yield" => {
                        has_marker = matches!(value.kind, ExprKind::Lit(Literal::Bool(true)));
                    }
                    ExprKind::Lit(Literal::Str(name)) if name == "key" => {
                        key_expr = Some(value);
                    }
                    ExprKind::Lit(Literal::Str(name)) if name == "value" => {
                        value_expr = Some(value);
                    }
                    _ => {}
                }
            }
        }

        if has_marker {
            key_expr.zip(value_expr)
        } else {
            None
        }
    }

    fn emit_generator_payload_store(&mut self) {
        let store_global = self.str_const("__vybe_generator_payloads");
        self.emit_u16(Op::GLOBAL_GET, store_global);
        let store_slot = self.define_local("__gen_payload_store");
        self.emit_u16(Op::LOCAL_SET, store_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, store_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);
        common::dict::emit_new(&mut self.chunks, self.current, line);
        self.emit(Op::DUP);
        self.emit_u16(Op::LOCAL_SET, store_slot);
        self.emit(Op::DROP);
        self.emit_u16(Op::GLOBAL_SET, store_global);
        self.emit(Op::DROP);

        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, store_slot);
    }

    fn emit_next_generator_payload_id(&mut self) {
        let next_global = self.str_const("__vybe_generator_payload_next_id");
        self.emit_u16(Op::GLOBAL_GET, next_global);
        let id_slot = self.define_local("__gen_payload_id_current");
        self.emit_u16(Op::LOCAL_SET, id_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, id_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_SET, id_slot);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, id_slot);
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::GLOBAL_SET, next_global);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, id_slot);
    }

    pub(crate) fn emit_generator_yield_value(&mut self, yielded_slot: u16) {
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        self.emit(Op::REF_IS_OBJECT);
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let marker_key = self.str_const("__vybe_generator_yield");
        self.emit_u16(Op::STRUCT_GET, marker_key);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let payload_id_key = self.str_const("payload_id");
        self.emit_u16(Op::STRUCT_GET, payload_id_key);
        let payload_id_slot = self.define_local("__yield_payload_id");
        self.emit_u16(Op::LOCAL_SET, payload_id_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
        self.emit(Op::REF_IS_NULL);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let value_key = self.str_const("value");
        self.emit_u16(Op::STRUCT_GET, value_key);
        self.chunk().emit_else(line);

        self.emit_generator_payload_store();
        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);

        self.chunk().emit_end(line);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        self.chunk().emit_end(line);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        self.chunk().emit_end(line);
    }

    pub(crate) fn emit_generator_yield_key_or_fallback(
        &mut self,
        yielded_slot: u16,
        fallback_slot: Option<u16>,
    ) {
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        self.emit(Op::REF_IS_OBJECT);
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let marker_key = self.str_const("__vybe_generator_yield");
        self.emit_u16(Op::STRUCT_GET, marker_key);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let key_key = self.str_const("key");
        self.emit_u16(Op::STRUCT_GET, key_key);

        self.chunk().emit_else(line);
        if let Some(slot) = fallback_slot {
            self.emit_u16(Op::LOCAL_GET, slot);
        } else {
            self.emit(Op::NULL);
        }
        self.chunk().emit_end(line);
        self.chunk().emit_else(line);
        if let Some(slot) = fallback_slot {
            self.emit_u16(Op::LOCAL_GET, slot);
        } else {
            self.emit(Op::NULL);
        }
        self.chunk().emit_end(line);
    }

    pub(crate) fn compile_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            // ── Literals ────────────────────────────────────────────────
            ExprKind::Lit(lit) => match lit {
                Literal::Int(n) => self.emit_const(Value::F64(*n as f64)),
                Literal::Float(n) => self.emit_const(Value::F64(*n)),
                Literal::BigInt(n) => self.emit_const(Value::BigInt(*n)),
                Literal::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                Literal::Char(c) => {
                    self.emit_const(Value::String(Arc::from(c.to_string().as_str())))
                }
                Literal::Bool(b) => {
                    if *b {
                        self.emit(Op::TRUE)
                    } else {
                        self.emit(Op::FALSE)
                    }
                }
                Literal::Null => self.emit(Op::NULL),
                Literal::Undefined => {
                    let l = self.line;
                    common::expressions::emit_undefined(self.chunk(), l);
                }
                Literal::Ellipsis => self.emit(Op::NULL),
            },

            // ── Identifier ──────────────────────────────────────────────
            ExprKind::Ident(name) => {
                // JS global constants that aren't variables
                match name.as_str() {
                    "NaN" => {
                        self.emit_const(Value::F64(f64::NAN));
                        return Ok(());
                    }
                    "Infinity" => {
                        self.emit_const(Value::F64(f64::INFINITY));
                        return Ok(());
                    }
                    "__js_import_meta" if self.is_js_profile() => {
                        self.emit_js_import_meta_object();
                        return Ok(());
                    }
                    "undefined" if self.case_sensitive => {
                        let l = self.line;
                        common::expressions::emit_undefined(self.chunk(), l);
                        return Ok(());
                    }
                    _ => {}
                }
                if self.is_python_profile() {
                    match name.as_str() {
                        "__debug__" => {
                            self.emit(Op::TRUE);
                            return Ok(());
                        }
                        "__name__" => {
                            self.emit_const(Value::String(Arc::from("__main__")));
                            return Ok(());
                        }
                        _ => {}
                    }
                }
                if let (Some(function_name), Some(result_slot)) =
                    (self.current_func_name.as_ref(), self.current_result_slot)
                {
                    let matches = if self.case_sensitive {
                        name == function_name
                    } else {
                        name.eq_ignore_ascii_case(function_name)
                    };
                    if matches {
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }
                }
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                    || self.has_static_local_binding(name);

                // Implicit self field access (only if NOT a local)
                if !is_local && self.is_class_field(name) {
                    if self.emit_self_ref() {
                        let field_name = self.canon(name);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        return Ok(());
                    }
                }

                // Bare enum member: `Green` → `TColor.Green`
                if !is_local {
                    let canon_name = self.canon(name);
                    if let Some(enum_type) = self.enum_members.get(&canon_name).cloned() {
                        let type_idx = self.str_const(&enum_type);
                        self.emit_u16(Op::GLOBAL_GET, type_idx);
                        let mem_idx = self.str_const(&canon_name);
                        self.emit_u16(Op::STRUCT_GET, mem_idx);
                        return Ok(());
                    }
                }

                // Bare profile namespace constant (e.g. Pascal `MaxInt`, `Pi`)
                if !is_local && !self.defined_globals.contains(&self.canon(name)) {
                    if let Some(cv) = self.profile.lookup_constant(name) {
                        match cv {
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => {
                                self.emit_const(Value::String(Arc::from(s.as_str())))
                            }
                        }
                        return Ok(());
                    }
                }

                if self.profile.name == "pascal" && !is_local {
                    let canon_name = self.canon(name);
                    if self.function_min_arity.get(&canon_name).copied() == Some(0) {
                        self.emit_var_get(name);
                        self.emit_u8(Op::CALL_REF, 0);
                        return Ok(());
                    }
                }

                if self.profile.name == "vb"
                    && !is_local
                    && self.defined_functions.contains(name.as_str())
                {
                    self.emit_var_get(name);
                    self.emit_u8(Op::CALL_REF, 0);
                    return Ok(());
                }

                // ESM named import of a constant Value (e.g. `ecma:math::PI`).
                // These are NOT callable — they're inlined as compile-time
                // constants from `host_const_bindings` populated during
                // `collect_host_imports`. Function imports stay in
                // `host_import_bindings` and route through CALL_IMPORT.
                if !is_local {
                    let cname = self.canon(name);
                    if let Some(val) = self.host_const_bindings.get(&cname).cloned() {
                        self.emit_const(val);
                        return Ok(());
                    }
                }

                self.emit_var_get(name);
            }

            // ── This / Super ────────────────────────────────────────────
            ExprKind::This => {
                if self.is_js_profile()
                    && self.current_class.is_some()
                    && self.current_func_name.as_deref() != Some("<lambda>")
                    && self.current_func_name.as_deref().is_some_and(|name| {
                        !name.eq_ignore_ascii_case(&self.profile.constructor_name)
                    })
                {
                    let idx = self.str_const("__js_this");
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    return Ok(());
                }

                let self_kw = &self.profile.self_keyword;
                if let Some(slot) = self
                    .scope()
                    .resolve(self_kw)
                    .or_else(|| self.scope().resolve_ci(self_kw))
                    .or_else(|| self.scope().resolve("Self"))
                    .or_else(|| self.scope().resolve("self"))
                    .or_else(|| self.scope().resolve("this"))
                {
                    self.emit_u16(Op::LOCAL_GET, slot);
                } else if self.scopes.len() > 1 {
                    // Arrow function: capture `this` from enclosing scope via upvalue
                    let kw = self.profile.self_keyword.clone();
                    if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, &kw) {
                        self.emit_u8(Op::UPVALUE_GET, uv);
                    } else if self.is_js_profile() {
                        if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, "__js_this") {
                            self.emit_u8(Op::UPVALUE_GET, uv);
                        } else {
                            let idx = self.str_const("__js_this");
                            self.emit_u16(Op::GLOBAL_GET, idx);
                        }
                    } else if self.is_js_profile() {
                        let idx = self.str_const("__js_this");
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    } else {
                        self.emit(Op::NULL);
                    }
                } else if self.is_js_profile() {
                    let idx = self.str_const("__js_this");
                    self.emit_u16(Op::GLOBAL_GET, idx);
                } else {
                    self.emit(Op::NULL);
                }
            }

            ExprKind::Super => {
                // super refers to the parent class constructor.
                // Look up the parent from the current class's PendingClass info.
                if let Some(ref class_name) = self.current_class.clone() {
                    if let Some(parent_name) = self
                        .pending_classes
                        .get(class_name.as_str())
                        .and_then(|pc| pc.parent.clone())
                    {
                        let pname = self.canon(&parent_name);
                        let idx = self.str_const(&pname);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    } else {
                        self.emit(Op::NULL);
                    }
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── Binary ──────────────────────────────────────────────────
            ExprKind::Binary { op, left, right } => {
                let is_csharp_integral_type = |type_hint: &str| {
                    matches!(
                        Self::normalize_type_hint(type_hint).as_str(),
                        "int" | "uint" | "long" | "ulong" | "short" | "ushort" | "byte" | "sbyte"
                    )
                };
                let is_c_integral_type = |type_hint: &str| {
                    let hint = Self::normalize_type_hint(type_hint);
                    hint.contains("int")
                        || hint.contains("long")
                        || hint.contains("short")
                        || hint.contains("char")
                        || hint == "uint8"
                        || hint == "uint32"
                        || hint == "bool"
                        || hint == "_bool"
                };
                let expr_is_csharp_integral = |compiler: &Compiler, expr: &Expression| {
                    matches!(expr.kind, ExprKind::Lit(Literal::Int(_)))
                        || compiler
                            .infer_expr_type_hint(expr)
                            .as_deref()
                            .is_some_and(is_csharp_integral_type)
                };
                let expr_is_c_integral = |compiler: &Compiler, expr: &Expression| {
                    matches!(expr.kind, ExprKind::Lit(Literal::Int(_)))
                        || compiler
                            .infer_expr_type_hint(expr)
                            .as_deref()
                            .is_some_and(is_c_integral_type)
                };

                // Short-circuit for And/Or — generic path for all languages.
                // PHP used a custom truthiness path here that referenced the
                // removed __keys/vybe$assoc_keys_csv side-band and left the
                // stack unbalanced. emit_dyn_to_bool is correct: empty()/isset()
                // already handle PHP array truthiness at the call site.
                if *op == BinOp::And {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip =
                        common::expressions::emit_and_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(
                        &mut self.chunks[self.current],
                        skip,
                    );
                    return Ok(());
                }
                if *op == BinOp::Or {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip =
                        common::expressions::emit_or_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(
                        &mut self.chunks[self.current],
                        skip,
                    );
                    return Ok(());
                }
                // NullCoalesce as binary op
                if *op == BinOp::NullCoalesce {
                    self.compile_expr(left)?;
                    let value_slot = self.define_local("__null_coalesce_left");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.compile_expr(right)?;
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_end(line);
                    return Ok(());
                }
                // Pow → BigInt fast path, then canonical stdlib path.
                if *op == BinOp::Pow {
                    let line = self.line;
                    if self.is_js_profile()
                        && self.infer_expr_type_hint(left).as_deref() == Some("bigint")
                        && self.infer_expr_type_hint(right).as_deref() == Some("bigint")
                    {
                        let idx = self.import("ecma:bigint", "pow");
                        self.compile_expr(left)?;
                        self.compile_expr(right)?;
                        self.emit_host_call(idx, 2);
                        return Ok(());
                    }
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    common::math::emit_pow(self.chunk(), line);
                    return Ok(());
                }
                // InstanceOf → WASM GC `ref.test` opcode with the type name
                // from the const pool. The static (Ident RHS) form covers
                // every real-world `a instanceof TypeName` usage and matches
                // both GC `ref.test ht` and Component Model resource handle
                // typing — both of which require a compile-time-known type.
                // Dynamic RHS (the rare `a instanceof someVariable` JS form)
                // falls through to the host-less polyfill in compile_binop.
                if *op == BinOp::InstanceOf {
                    if self.is_js_profile() {
                        self.compile_expr(left)?;
                        match &right.kind {
                            // For built-in constructors that stamp `__type` on their instances,
                            // pass the name as a string — `js_instanceof` matches it against
                            // `__type` / `__types`, avoiding namespace-alias objects that have
                            // no `.name` property and would always yield `false`.
                            ExprKind::Ident(name)
                                if matches!(
                                    name.as_str(),
                                    "Error"
                                        | "EvalError"
                                        | "RangeError"
                                        | "ReferenceError"
                                        | "SyntaxError"
                                        | "TypeError"
                                        | "URIError"
                                        | "AggregateError"
                                        | "RegExp"
                                        | "Map"
                                        | "Set"
                                        | "WeakMap"
                                        | "WeakSet"
                                        | "WeakRef"
                                        | "FinalizationRegistry"
                                        | "ArrayBuffer"
                                        | "DataView"
                                        | "Date"
                                        | "Promise"
                                        | "SharedArrayBuffer"
                                        | "URL"
                                        | "URLSearchParams"
                                        | "TextEncoder"
                                        | "TextDecoder"
                                ) =>
                            {
                                self.emit_const(Value::String(Arc::from(name.as_str())));
                            }
                            _ => self.compile_expr(right)?,
                        }
                        self.compile_binop(op);
                        return Ok(());
                    }
                    if let crate::ast::ExprKind::Ident(type_name) = &right.kind {
                        self.compile_expr(left)?;
                        let line = self.line;
                        let name_canon = self.canon(type_name);
                        let idx = self.chunk().add_constant(vybe_bytecode::Value::String(
                            std::sync::Arc::from(name_canon.as_str()),
                        ));
                        self.chunk()
                            .emit_op_u16(vybe_bytecode::Op::REF_TEST, idx, line);
                        return Ok(());
                    }
                }
                if self.is_js_profile() && *op == BinOp::In {
                    if let ExprKind::Ident(field) = &left.kind {
                        if field.starts_with('#') {
                            let storage_name = self.js_member_storage_name(field);
                            self.emit_const(Value::String(Arc::from(storage_name.as_str())));
                            self.compile_expr(right)?;
                            self.compile_binop(op);
                            return Ok(());
                        }
                    }
                }
                if self.profile.name == "pascal" && (*op == BinOp::In || *op == BinOp::NotIn) {
                    if !self.expr_is_pascal_set(right) {
                        let line = self.line;
                        self.compile_expr(right)?;
                        self.compile_expr(left)?;
                        common::collections::emit_contains(&mut self.chunks, self.current, line);
                        if *op == BinOp::NotIn {
                            {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                            };
                        }
                        return Ok(());
                    }
                }
                if self.try_compile_pascal_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_python_set_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_csharp_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_dotnet_datetime_timespan_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_fortran_interface_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_fortran_array_binary_operator(op, left, right)? {
                    return Ok(());
                }

                if self.profile.name == "csharp"
                    && *op == BinOp::Div
                    && expr_is_csharp_integral(self, left)
                    && expr_is_csharp_integral(self, right)
                {
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    self.compile_binop(&BinOp::IDiv);
                    return Ok(());
                }
                if self.profile.name == "c"
                    && *op == BinOp::Div
                    && expr_is_c_integral(self, left)
                    && expr_is_c_integral(self, right)
                {
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    self.compile_binop(&BinOp::IDiv);
                    return Ok(());
                }

                if self.is_php_profile() && *op == BinOp::Concat {
                    let line = self.line;
                    self.compile_expr(left)?;
                    self.emit_common("php.echo_stringify", 1, line);
                    self.compile_expr(right)?;
                    self.emit_common("php.echo_stringify", 1, line);
                    self.compile_binop(op);
                    return Ok(());
                }

                // BigInt arithmetic and comparisons via ecma:bigint host fns.
                // These already exist and return Value::BigInt / Value::Bool.
                // `infer_expr_type_hint` returns "bigint" for BigInt literals
                // and for variables initialised with BigInt values.
                if self.is_js_profile() {
                    let left_hint = self.infer_expr_type_hint(left);
                    let right_hint = self.infer_expr_type_hint(right);
                    let left_is_bigint = left_hint.as_deref() == Some("bigint");
                    let right_is_bigint = right_hint.as_deref() == Some("bigint");
                    // The other operand is "known non-BigInt" only when its
                    // type was inferred to something concrete that isn't
                    // bigint. An *unknown* hint (e.g. a reassigned parameter
                    // holding a BigInt at runtime) is NOT a compile-time mix
                    // error — §21.2.1.1 mixing is a RUNTIME TypeError, so we
                    // route to the bigint op and let it coerce/throw.
                    let other_known_non_bigint = if left_is_bigint {
                        right_hint.is_some()
                    } else {
                        left_hint.is_some()
                    };
                    if left_is_bigint && right_is_bigint {
                        let fn_name: Option<&str> = match op {
                            BinOp::Add => Some("add"),
                            BinOp::Sub => Some("sub"),
                            BinOp::Mul => Some("mul"),
                            BinOp::Div => Some("div"),
                            BinOp::Mod => Some("rem"),
                            BinOp::BitAnd => Some("and"),
                            BinOp::BitOr => Some("or"),
                            BinOp::BitXor => Some("xor"),
                            BinOp::Pow => Some("pow"),
                            BinOp::Shl => Some("shl"),
                            BinOp::Shr => Some("shr"),
                            BinOp::Eq | BinOp::StrictEq => Some("eq"),
                            BinOp::NotEq | BinOp::StrictNotEq => Some("ne"),
                            BinOp::Lt => Some("lt"),
                            BinOp::LtEq => Some("le"),
                            BinOp::Gt => Some("gt"),
                            BinOp::GtEq => Some("ge"),
                            _ => None,
                        };
                        if let Some(name) = fn_name {
                            let idx = self.import("ecma:bigint", name);
                            self.compile_expr(left)?;
                            self.compile_expr(right)?;
                            self.emit_host_call(idx, 2);
                            return Ok(());
                        }
                    } else if left_is_bigint || right_is_bigint {
                        let arith_fn: Option<&str> = match op {
                            BinOp::Add => Some("add"),
                            BinOp::Sub => Some("sub"),
                            BinOp::Mul => Some("mul"),
                            BinOp::Div => Some("div"),
                            BinOp::Mod => Some("rem"),
                            BinOp::BitAnd => Some("and"),
                            BinOp::BitOr => Some("or"),
                            BinOp::BitXor => Some("xor"),
                            BinOp::Shl => Some("shl"),
                            BinOp::Shr => Some("shr"),
                            BinOp::Pow => Some("pow"),
                            _ => None,
                        };
                        if let Some(name) = arith_fn {
                            if other_known_non_bigint {
                                // §21.2.1.1: arithmetic between BigInt and a
                                // KNOWN non-BigInt throws TypeError.
                                self.emit_const(Value::String(Arc::from(
                                    "Cannot mix BigInt and other types, use explicit conversions",
                                )));
                                let line = self.line;
                                self.emit_js_exception_ctor_from_message_value("TypeError")?;
                                common::errors::emit_throw(self.chunk(), line);
                                return Ok(());
                            }
                            // Other operand is unknown — route to the bigint
                            // op (its `to_bigint` coerces an actual BigInt and
                            // throws at runtime on a real Number).
                            let idx = self.import("ecma:bigint", name);
                            self.compile_expr(left)?;
                            self.compile_expr(right)?;
                            self.emit_host_call(idx, 2);
                            return Ok(());
                        } else if matches!(
                            op,
                            BinOp::Eq
                                | BinOp::NotEq
                                | BinOp::Lt
                                | BinOp::LtEq
                                | BinOp::Gt
                                | BinOp::GtEq
                        ) {
                            // §7.2.13: BigInt == Number compares numerically.
                            // ecma:bigint comparisons use to_bigint() on both sides, which handles
                            // all numeric types (F64, I32, BigInt) correctly.
                            let fn_name = match op {
                                BinOp::Eq => "eq",
                                BinOp::NotEq => "ne",
                                BinOp::Lt => "lt",
                                BinOp::LtEq => "le",
                                BinOp::Gt => "gt",
                                BinOp::GtEq => "ge",
                                _ => unreachable!(),
                            };
                            let idx = self.import("ecma:bigint", fn_name);
                            self.compile_expr(left)?;
                            self.compile_expr(right)?;
                            self.emit_host_call(idx, 2);
                            return Ok(());
                        }
                        // StrictEq / StrictNotEq with mixed types: fall through to compile_binop
                        // (strict equality of different types returns false/true without coercion).
                    }
                }

                self.compile_expr(left)?;
                self.compile_expr(right)?;
                self.compile_binop(op);
            }

            // ── Unary ───────────────────────────────────────────────────
            ExprKind::Unary { op, expr: inner } => {
                match op {
                    UnaryOp::PreInc | UnaryOp::PostInc => {
                        // ++x / x++ : load, add 1, store
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostInc {
                            self.emit(Op::DUP);
                        }
                        self.emit_const(Value::F64(1.0));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                        };
                        if *op == UnaryOp::PreInc {
                            self.emit(Op::DUP);
                        }
                        self.compile_assign_target(inner)?;
                    }
                    UnaryOp::PreDec | UnaryOp::PostDec => {
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostDec {
                            self.emit(Op::DUP);
                        }
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::F64_SUB);
                        if *op == UnaryOp::PreDec {
                            self.emit(Op::DUP);
                        }
                        self.compile_assign_target(inner)?;
                    }
                    UnaryOp::AddrOf => {
                        self.compile_address_of_expr(inner)?;
                    }
                    UnaryOp::Deref => {
                        self.compile_deref_expr(inner)?;
                    }
                    _ => {
                        if self.is_js_profile() && matches!(op, UnaryOp::Typeof) {
                            if let ExprKind::Ident(name) = &inner.kind {
                                let is_callable_global =
                                    matches!(
                                        name.as_str(),
                                        "eval"
                                            | "parseInt"
                                            | "parseFloat"
                                            | "Function"
                                            | "Object"
                                            | "Boolean"
                                            | "Number"
                                            | "String"
                                            | "Array"
                                            | "Symbol"
                                            | "BigInt"
                                            | "Date"
                                            | "RegExp"
                                            | "Promise"
                                            | "Proxy"
                                            | "Map"
                                            | "Set"
                                            | "WeakMap"
                                            | "WeakSet"
                                            | "ArrayBuffer"
                                            | "SharedArrayBuffer"
                                            | "DataView"
                                            | "Int8Array"
                                            | "Uint8Array"
                                            | "Uint8ClampedArray"
                                            | "Int16Array"
                                            | "Uint16Array"
                                            | "Int32Array"
                                            | "Uint32Array"
                                            | "Float32Array"
                                            | "Float64Array"
                                            | "BigInt64Array"
                                            | "BigUint64Array"
                                    ) || self.profile.lookup_builtin(name).is_some();
                                if is_callable_global {
                                    self.emit_const(Value::String(Arc::from("function")));
                                    return Ok(());
                                }
                            }
                        }
                        self.compile_expr(inner)?;
                        match op {
                            UnaryOp::Neg => {
                                let l = self.line;
                                if self.is_js_profile()
                                    && self.infer_expr_type_hint(inner).as_deref() == Some("bigint")
                                {
                                    let idx = self.import("ecma:bigint", "neg");
                                    self.emit_host_call(idx, 1);
                                } else {
                                    common::math::emit_neg(self.chunk(), l);
                                }
                            }
                            UnaryOp::Pos => {
                                // JS `+v` coerces to number — ECMA-262 §7.1.4 ToNumber.
                                // BigInt is the one primitive exception: unary
                                // plus throws, while explicit Number(1n) still
                                // converts.
                                if self.is_js_profile()
                                    && self.infer_expr_type_hint(inner).as_deref() == Some("bigint")
                                {
                                    self.emit_const(Value::String(Arc::from(
                                        "Cannot convert a BigInt value to a number",
                                    )));
                                    let line = self.line;
                                    self.emit_js_exception_ctor_from_message_value("TypeError")?;
                                    common::errors::emit_throw(self.chunk(), line);
                                    return Ok(());
                                }
                                // For Object operands, ToPrimitive(hint=number)
                                // first (so Symbol.toPrimitive / valueOf
                                // overrides fire with the JS method-call
                                // protocol intact), then `Number(...)` to
                                // finalise the coercion.
                                if self.is_js_profile() {
                                    self.emit_to_primitive("number");
                                }
                                let idx = self.import("ecma:number", "Number");
                                self.emit_host_call(idx, 1);
                            }
                            UnaryOp::Not => {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                                if self.is_js_profile() {
                                    crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                                }
                            }
                            UnaryOp::BitNot => {
                                let l = self.line;
                                if self.is_js_profile()
                                    && self.infer_expr_type_hint(inner).as_deref() == Some("bigint")
                                {
                                    let idx = self.import("ecma:bigint", "not");
                                    self.emit_host_call(idx, 1);
                                } else {
                                    common::expressions::emit_i32_not(self.chunk(), l);
                                }
                            }
                            UnaryOp::Typeof => self.emit(Op::REF_TYPEOF),
                            UnaryOp::Void => {
                                self.emit(Op::DROP);
                                self.emit(Op::UNDEFINED);
                            }
                            UnaryOp::Delete => {
                                self.emit(Op::DROP);
                                self.emit(Op::TRUE);
                            }
                            UnaryOp::Await => {} // handled below in ExprKind::Await
                            _ => {}              // PreInc etc handled above
                        }
                    }
                }
            }

            ExprKind::RefOf(place) => match place.as_ref() {
                PlaceExpr::Ident(name) => {
                    if let Some(slot) = self.promote_local_binding_to_pointer_cell(name) {
                        self.emit_u16(Op::LOCAL_GET, slot);
                    } else {
                        self.compile_expr(&Expression::ident(name))?;
                        self.emit_wrap_top_of_stack_in_pointer_cell();
                    }
                }
                PlaceExpr::Deref(expr) => {
                    self.compile_expr(expr)?;
                }
                PlaceExpr::Member {
                    object,
                    field,
                    null_safe,
                } => {
                    self.compile_expr(&Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: field.clone(),
                        null_safe: *null_safe,
                    }))?;
                    self.emit_wrap_top_of_stack_in_pointer_cell();
                }
                PlaceExpr::Index {
                    object,
                    index,
                    null_safe,
                } => {
                    self.compile_expr(&Expression::new(ExprKind::Index {
                        object: object.clone(),
                        index: index.clone(),
                        null_safe: *null_safe,
                    }))?;
                    self.emit_wrap_top_of_stack_in_pointer_cell();
                }
            },

            ExprKind::RefLoad(expr) => {
                self.compile_deref_expr(expr)?;
            }

            // ── Ternary ─────────────────────────────────────────────────
            ExprKind::Ternary { cond, then, else_ } => {
                self.compile_expr(cond)?;
                self.emit_condition_truthiness_from_stack();
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.compile_expr(then)?;
                self.chunk().emit_else(line);
                self.compile_expr(else_)?;
                self.chunk().emit_end(line);
            }

            // ── Call ────────────────────────────────────────────────────
            ExprKind::Call {
                callee,
                args,
                optional,
            } => {
                if *optional {
                    // Optional call: callee?.() — short-circuit to undefined if callee is null/undefined.
                    // Per ECMA-262 §13.5.9: the result is `undefined` (not null) when short-circuiting.
                    self.compile_expr(callee)?;
                    let tmp = self.define_local("__optional_callee");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit(Op::UNDEFINED);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit(Op::REF_IS_UNDEFINED);
                    let undef_line = self.line;
                    self.chunk().emit_if_value(undef_line);
                    self.emit(Op::UNDEFINED);
                    self.chunk().emit_else(undef_line);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    for a in args {
                        self.compile_expr(&a.value)?;
                    }
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                    self.chunk().emit_end(undef_line);
                    self.chunk().emit_end(line);
                } else {
                    if self.profile.parens_for_index
                        && !args.is_empty()
                        && matches!(&callee.kind,
                            ExprKind::Ident(name) if self.lookup_array_binding(name).is_some()
                        )
                    {
                        self.compile_expr(callee)?;
                        for arg in args {
                            self.compile_array_index_operand_for_owner(callee, &arg.value)?;
                            let line = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, line);
                        }
                        return Ok(());
                    }

                    self.compile_call(callee, args)?;
                    // Multi-value result repack: when the callee is one of
                    // the pre-scanned multi-return functions, CALL leaves
                    // N values on the stack. A destructure-assign consumes
                    // them directly (see `detect_multi_value_receive` in
                    // compiler/mod.rs, which bypasses this branch); every
                    // other use site — `r = f()`, `print(f())`, `f() + g()`
                    // — expects a single value, so we re-pack here.
                    if let Some(n) = self.multi_return_arity_for_callee(callee) {
                        self.pack_multi_value_result(n);
                    }
                }
            }

            // ── Member access ───────────────────────────────────────────
            ExprKind::Member {
                object,
                field,
                null_safe,
            } => {
                if self.js_private_member_access_forbidden(field) {
                    self.emit_js_private_access_denied(field)?;
                    return Ok(());
                }
                let reflection_field = field.split('<').next().unwrap_or(field.as_str()).trim();
                if let Some(binding) = self.resolve_reflection_binding_expr(object) {
                    match (binding, reflection_field) {
                        (ReflectionBinding::Type(type_name), "Name") => {
                            let short_name = self.reflection_type_short_name(&type_name);
                            self.compile_expr(&Expression::string(&short_name))?;
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "FullName") => {
                            let full_name = self.reflection_type_full_name(&type_name);
                            self.compile_expr(&Expression::string(&full_name))?;
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "IsEnum") => {
                            self.emit(if self.reflection_is_enum_type(&type_name) {
                                Op::TRUE
                            } else {
                                Op::FALSE
                            });
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "IsValueType") => {
                            self.emit(if self.reflection_is_value_type(&type_name) {
                                Op::TRUE
                            } else {
                                Op::FALSE
                            });
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "BaseType") => {
                            if let Some(parent_type) = self.reflection_base_type_name(&type_name) {
                                self.compile_reflection_type_value(&parent_type)?;
                            } else {
                                self.emit(Op::NULL);
                            }
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Method {
                                type_name,
                                method_name,
                            },
                            "IsStatic",
                        ) => {
                            let is_static = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.methods.get(&method_name))
                                .map(|meta| meta.is_static)
                                .unwrap_or(false);
                            self.emit(if is_static { Op::TRUE } else { Op::FALSE });
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Property {
                                type_name,
                                property_name,
                            },
                            "CanWrite",
                        ) => {
                            let can_write = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.properties.get(&property_name))
                                .map(|meta| meta.can_write)
                                .unwrap_or(false);
                            self.emit(if can_write { Op::TRUE } else { Op::FALSE });
                            return Ok(());
                        }
                        _ => {}
                    }
                }

                if field.eq_ignore_ascii_case("IsEnum") {
                    if let ExprKind::Lit(Literal::Str(type_name)) = &object.kind {
                        let full_name = type_name
                            .strip_prefix("System.")
                            .unwrap_or(type_name.as_str())
                            .trim();
                        let short_name = full_name.rsplit('.').next().unwrap_or(full_name).trim();
                        let is_enum = [full_name, short_name].into_iter().any(|candidate| {
                            let canon = self.canon(candidate);
                            self.enum_value_names.contains_key(&canon)
                                || self
                                    .enum_value_names
                                    .keys()
                                    .any(|known| known.eq_ignore_ascii_case(candidate))
                        });
                        self.emit(if is_enum { Op::TRUE } else { Op::FALSE });
                        return Ok(());
                    }
                }

                // Namespace constant check (Math.PI, etc.)
                if let ExprKind::Ident(obj_name) = &object.kind {
                    let prefers_type_lookup =
                        self.prefers_type_qualified_member_lookup(obj_name, field);
                    let obj_is_local = self.scope().resolve(obj_name).is_some()
                        || self.has_static_local_binding(obj_name)
                        || (!self.case_sensitive
                            && self.scope().resolve_ci(obj_name).is_some()
                            && !prefers_type_lookup);
                    if !obj_is_local {
                        if let Some(value) = self.enum_member_ordinal(obj_name, field) {
                            self.emit_const(Value::F64(value as f64));
                            return Ok(());
                        }
                    }

                    if let Some(key) = self.generic_static_member_key(obj_name, field) {
                        let idx = self.str_const(&key);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        return Ok(());
                    }

                    let compound = format!("{}.{}", obj_name, field);
                    if let Some(cv) = self.profile.lookup_constant(&compound) {
                        match cv {
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => {
                                self.emit_const(Value::String(Arc::from(s.as_str())))
                            }
                        }
                        return Ok(());
                    }
                    // Constructor call with 0 args: ClassName.Create
                    let ctor_nm = &self.profile.constructor_name;
                    let is_ctor = if self.case_sensitive {
                        field == ctor_nm
                    } else {
                        field.eq_ignore_ascii_case(ctor_nm)
                    };
                    let canon_obj = self.canon(obj_name);
                    let is_known_class = self.defined_classes.contains(&canon_obj)
                        && self.scope().resolve(obj_name).is_none();
                    if is_ctor && is_known_class {
                        self.emit_var_get(obj_name);
                        self.emit_u8(Op::CALL_REF, 0);
                        return Ok(());
                    }

                    // Pascal allows parameterless class functions to be used
                    // without `()`: `TShape.Circle`. Only auto-invoke when the
                    // member resolves to a known static method whose chunk arity
                    // is receiver-only (the class object plus zero user args).
                    if self.profile.name == "pascal" {
                        let is_class = self.defined_classes.contains(&canon_obj)
                            && self.scope().resolve(obj_name).is_none();
                        if is_class {
                            let method_name = self.canon(field);
                            let zero_arg_static = self
                                .pending_classes
                                .get(canon_obj.as_str())
                                .map(|pc| !pc.static_fields.iter().any(|name| name == &method_name))
                                .unwrap_or(false);
                            if zero_arg_static {
                                let cls_idx = self.str_const(&canon_obj);
                                self.emit_u16(Op::GLOBAL_GET, cls_idx);
                                self.emit(Op::DUP);
                                let method_idx = self.str_const(&method_name);
                                self.emit_u16(Op::STRUCT_GET, method_idx);
                                let fn_tmp = self
                                    .scope()
                                    .resolve("__pascal_static_fn")
                                    .unwrap_or_else(|| self.define_local("__pascal_static_fn"));
                                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                self.emit(Op::DROP);
                                let cls_tmp = self
                                    .scope()
                                    .resolve("__pascal_static_cls")
                                    .unwrap_or_else(|| self.define_local("__pascal_static_cls"));
                                self.emit_u16(Op::LOCAL_SET, cls_tmp);
                                self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                self.emit_u16(Op::LOCAL_GET, cls_tmp);
                                self.emit_u8(Op::CALL_REF, 1);
                                return Ok(());
                            }
                        }
                    }
                }

                if self.profile.namespaces.use_dotnet {
                    let parts = self.flatten_member_chain(expr);
                    if !parts.is_empty() {
                        let lower_parts: Vec<String> =
                            parts.iter().map(|part| self.canon(part)).collect();
                        if common::dotnet::is_namespace_root(&lower_parts[0]) {
                            let scope = self.scope();
                            let dotnet_surface = common::dotnet::surface();
                            let mut imports = dotnet_surface.default_imports().to_vec();
                            imports.extend(self.profile.namespaces.extra_imports.clone());
                            let field_set: std::collections::HashSet<String> =
                                if let Some(ref class_name) = self.current_class {
                                    self.pending_classes
                                        .get(class_name.as_str())
                                        .map(|pending| pending.fields.iter().cloned().collect())
                                        .unwrap_or_default()
                                } else {
                                    std::collections::HashSet::new()
                                };
                            let defined_globals = self.defined_globals.clone();
                            let defined_classes = self.defined_classes.clone();
                            let is_user_class_fn = move |name: &str| -> bool {
                                defined_classes.contains(name)
                                    || defined_classes
                                        .iter()
                                        .any(|class_name| class_name.eq_ignore_ascii_case(name))
                            };
                            let is_user_class_for_local = is_user_class_fn.clone();
                            let ctx = common::dotnet::ResolutionContext {
                                is_local: &|name: &str| {
                                    if is_user_class_for_local(name) {
                                        return false;
                                    }
                                    scope.resolve(name).is_some()
                                        || scope.resolve_ci(name).is_some()
                                        || defined_globals.contains(name)
                                        || defined_globals.iter().any(|global_name| {
                                            global_name.eq_ignore_ascii_case(name)
                                        })
                                },
                                is_class_field: &|name: &str| field_set.contains(name),
                                is_user_type: &is_user_class_fn,
                                imports: &imports,
                            };
                            let refs: Vec<&str> =
                                lower_parts.iter().map(|part| part.as_str()).collect();
                            match common::dotnet::resolve_dotted_name(&refs, &ctx) {
                                common::dotnet::DottedResolution::GlobalAccess { name } => {
                                    let idx = self.str_const(&name);
                                    self.emit_u16(Op::GLOBAL_GET, idx);
                                    return Ok(());
                                }
                                common::dotnet::DottedResolution::CommonCall { emit } => {
                                    self.emit_common(&emit, 0, self.line);
                                    return Ok(());
                                }
                                common::dotnet::DottedResolution::HostCall { module, func } => {
                                    let idx = self.import(&module, &func);
                                    self.emit_host_call(idx, 0);
                                    return Ok(());
                                }
                                common::dotnet::DottedResolution::NamespaceAccess {
                                    parts: ns_parts,
                                } => {
                                    if ns_parts.len() >= 2 {
                                        let mut found_window: Option<(usize, usize)> = None;
                                        'outer: for start in 0..ns_parts.len().saturating_sub(1) {
                                            for end in ((start + 2)..=ns_parts.len()).rev() {
                                                let key = ns_parts[start..end].join(".");
                                                if self.profile.lookup_constant(&key).is_some() {
                                                    found_window = Some((start, end));
                                                    break 'outer;
                                                }
                                            }
                                        }
                                        if let Some((_const_start, const_end)) = found_window {
                                            let key = ns_parts[_const_start..const_end].join(".");
                                            let cv = self
                                                .profile
                                                .lookup_constant(&key)
                                                .cloned()
                                                .unwrap();
                                            match cv {
                                                ConstantValue::Float(f) => {
                                                    self.emit_const(Value::F64(f))
                                                }
                                                ConstantValue::Str(s) => self.emit_const(
                                                    Value::String(Arc::from(s.as_str())),
                                                ),
                                            }
                                            for part in &ns_parts[const_end..] {
                                                let idx = self.str_const(part);
                                                self.emit_u16(Op::STRUCT_GET, idx);
                                            }
                                            return Ok(());
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }

                if self.is_js_profile() && matches!(&object.kind, ExprKind::Super) && !*null_safe {
                    if let Some(parent) = self
                        .current_class
                        .as_ref()
                        .and_then(|cn| self.pending_classes.get(cn.as_str()))
                        .and_then(|pc| pc.parent.clone())
                    {
                        let result_slot = self.define_local("__js_super_prop_result");
                        let saved_this = self.save_js_this("__js_prev_this_super_prop");
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
                        self.set_js_this_from_stack();

                        let getter_key = self.str_const(&format!("__get_{}", field));
                        self.emit_var_get(&parent);
                        self.emit_u16(Op::STRUCT_GET, getter_key);
                        let getter_slot = self.define_local("__js_super_prop_getter");
                        self.emit_u16(Op::LOCAL_SET, getter_slot);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, getter_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        common::expressions::emit_undefined(self.chunk(), line);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, getter_slot);
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
                        self.emit_u8(Op::CALL_REF, 1);
                        self.chunk().emit_end(line);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit(Op::DROP);
                        self.restore_js_this(saved_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }
                }

                // Proxy get-trap dispatch (JS profile, only when the
                // module references `Proxy` somewhere). Routes member
                // reads through the inline dispatcher in
                // `emitter/js/proxy_adapter.rs`. Non-proxy code keeps
                // the direct STRUCT_GET path.
                if self.is_js_profile() && self.uses_proxy && !*null_safe {
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    let line = self.line;
                    crate::emitter::js::proxy_adapter::emit_proxy_get_dispatch(
                        &mut self.chunks,
                        self.current,
                        line,
                    );
                    return Ok(());
                }

                if self.is_js_profile() {
                    if *null_safe {
                        self.compile_expr(object)?;
                        let nullsafe_obj_slot = self.define_local("__js_member_nullsafe_obj");
                        self.emit_u16(Op::LOCAL_SET, nullsafe_obj_slot);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, nullsafe_obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        common::expressions::emit_undefined(self.chunk(), line);
                        self.chunk().emit_else(line);
                        let obj_slot = nullsafe_obj_slot;
                        if field == "constructor" {
                            let constructor_of = self.import("ecma:value", "constructorOf");
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_host_call(constructor_of, 1);
                        } else if field == "length" {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            let prop = self.str_const("length");
                            self.emit_u16(Op::STRUCT_GET, prop);
                        } else {
                            let field_name =
                                self.js_member_storage_name_for_receiver(object, field);
                            let prop = self.str_const(&field_name);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u16(Op::STRUCT_GET, prop);
                            let val_slot = self.define_local("__js_member_val");
                            self.emit_u16(Op::LOCAL_SET, val_slot);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.emit(Op::REF_IS_UNDEFINED);
                            let lookup_line = self.line;
                            self.chunk().emit_if_value(lookup_line);
                            self.emit_js_member_fallback_get(obj_slot, &field_name);
                            self.chunk().emit_else(lookup_line);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.chunk().emit_end(lookup_line);
                        }
                        self.chunk().emit_end(line);
                    } else {
                        self.compile_expr(object)?;
                        let obj_slot = self.define_local("__js_member_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        self.emit(Op::DROP);

                        // Bind `__js_this = obj` so a getter installed
                        // by `Object.defineProperty` (which runs via
                        // STRUCT_GET's `__get_<name>` accessor dispatch)
                        // sees the receiver. The VM's accessor-call
                        // path doesn't set `__js_this` itself; doing it
                        // here keeps the semantics consistent with the
                        // explicit method-call path.
                        let saved_this = self.save_js_this("__js_prev_this_member");
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.set_js_this_from_stack();

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let msg_line = self.line;
                        self.chunk().emit_if_value(msg_line);
                        self.emit_const(Value::String(Arc::from(
                            "Cannot read properties of undefined",
                        )));
                        self.chunk().emit_else(msg_line);
                        self.emit_const(Value::String(Arc::from("Cannot read properties of null")));
                        self.chunk().emit_end(msg_line);
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);

                        let result_slot = self.define_local("__js_member_result");
                        if field == "constructor" {
                            let constructor_of = self.import("ecma:value", "constructorOf");
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_host_call(constructor_of, 1);
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.emit(Op::DROP);
                            self.restore_js_this(saved_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            return Ok(());
                        }
                        if field == "length" {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            let prop = self.str_const("length");
                            self.emit_u16(Op::STRUCT_GET, prop);
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.emit(Op::DROP);
                            self.restore_js_this(saved_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            return Ok(());
                        }
                        let symbol_end = if field == "description" {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_TYPEOF);
                            self.emit_const(Value::String(Arc::from("symbol")));
                            {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                            };
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                            self.chunk().emit_if(line);
                            let invoke = self.import("ecma:value", "invokeMethod");
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_const(Value::String(Arc::from("description")));
                            self.emit_host_call(invoke, 2);
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.emit(Op::DROP);
                            self.chunk().emit_else(line);
                            Some(line)
                        } else {
                            None
                        };

                        let field_name = self.js_member_storage_name_for_receiver(object, field);
                        let prop = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let val_slot = self.define_local("__js_member_val");
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let lookup_line = self.line;
                        self.chunk().emit_if_value(lookup_line);
                        self.emit_js_member_fallback_get(obj_slot, &field_name);
                        self.chunk().emit_else(lookup_line);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.chunk().emit_end(lookup_line);
                        // Restore the caller's __js_this — value already
                        // on stack as the access result.
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit(Op::DROP);
                        if let Some(line) = symbol_end {
                            self.chunk().emit_end(line);
                        }
                        self.restore_js_this(saved_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    return Ok(());
                }

                let receiver_type_hint = match &object.kind {
                    ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
                    _ => self.infer_expr_type_hint(object),
                };

                if self.profile.name == "pascal" && !*null_safe {
                    if let Some(class_name) = receiver_type_hint.as_deref().and_then(|type_hint| {
                        self.resolve_pending_class_name_for_type_hint(type_hint)
                    }) {
                        let member_name = self.canon(field);
                        let zero_arg_instance_method = self
                            .pending_classes
                            .get(&class_name)
                            .is_some_and(|pending| {
                                !pending.fields.iter().any(|name| name == &member_name)
                                    && pending
                                        .instance_method_overloads
                                        .get(&member_name)
                                        .is_some_and(|overloads| {
                                            overloads.iter().any(|overload| {
                                                overload.signature.min_arity == 0
                                                    && !overload.signature.has_rest
                                            })
                                        })
                            });
                        if zero_arg_instance_method {
                            self.compile_expr(object)?;
                            let obj_slot = self
                                .scope()
                                .resolve("__pascal_member_obj")
                                .unwrap_or_else(|| self.define_local("__pascal_member_obj"));
                            self.emit_u16(Op::LOCAL_SET, obj_slot);
                            self.emit(Op::DROP);
                            let prop = self.str_const(&member_name);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u16(Op::STRUCT_GET, prop);
                            let fn_slot = self
                                .scope()
                                .resolve("__pascal_member_fn")
                                .unwrap_or_else(|| self.define_local("__pascal_member_fn"));
                            self.emit_u16(Op::LOCAL_SET, fn_slot);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u8(Op::CALL_REF, 1);
                            return Ok(());
                        }
                    }
                }

                let receiver_is_nullable = receiver_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim().ends_with('?'));

                let receiver_array_rank =
                    if matches!(self.profile.name.as_str(), "csharp" | "vb") && field == "Rank" {
                        let inferred = receiver_type_hint.as_deref().and_then(|type_hint| {
                            let normalized = Self::normalize_type_hint(type_hint);
                            let start = normalized.find('[')?;
                            let rest = &normalized[start + 1..];
                            let end = rest.find(']')?;
                            Some(rest[..end].chars().filter(|ch| *ch == ',').count() + 1)
                        });
                        Some(inferred.unwrap_or(1))
                    } else {
                        None
                    };

                if matches!(self.profile.name.as_str(), "csharp" | "vb") && receiver_is_nullable {
                    match field.as_str() {
                        "HasValue" => {
                            self.compile_expr(object)?;
                            self.emit(Op::REF_IS_NULL);
                            {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                            };
                            return Ok(());
                        }
                        "Value" => {
                            self.compile_expr(object)?;
                            return Ok(());
                        }
                        _ => {}
                    }
                }

                if let Some(rank) = receiver_array_rank {
                    self.emit_const(Value::I32(rank as i32));
                    return Ok(());
                }

                let receiver_is_collection_like =
                    if matches!(self.profile.name.as_str(), "csharp" | "vb")
                        && matches!(field.as_str(), "Length" | "Count")
                    {
                        let unknown_receiver_default = field == "Length";
                        let is_collection_like_type = |type_hint: &str| {
                            let normalized = Self::normalize_type_hint(type_hint);
                            let bare = normalized
                                .split('<')
                                .next()
                                .unwrap_or(normalized.as_str())
                                .trim_end_matches('?');
                            let terminal = bare.rsplit('.').next().unwrap_or(bare);
                            Self::is_string_type_hint(type_hint)
                                || matches!(
                                    terminal,
                                    "list"
                                        | "arraylist"
                                        | "dictionary"
                                        | "queue"
                                        | "stack"
                                        | "hashset"
                                        | "set"
                                        | "collection"
                                        | "icollection"
                                        | "readonlycollection"
                                        | "enumerable"
                                        | "ienumerable"
                                        | "readonlylist"
                                        | "ilist"
                                        | "array"
                                )
                                || bare.ends_with("[]")
                        };

                        match &object.kind {
                            ExprKind::Ident(_) | ExprKind::New { .. } | ExprKind::Call { .. } => {
                                receiver_type_hint
                                    .as_deref()
                                    .map(is_collection_like_type)
                                    .unwrap_or(unknown_receiver_default)
                            }
                            ExprKind::Lit(Literal::Str(_))
                            | ExprKind::Interpolation(_)
                            | ExprKind::Array(_) => true,
                            _ => unknown_receiver_default,
                        }
                    } else {
                        false
                    };

                let is_dotnet_dictionary_accessor = !self.case_sensitive
                    && matches!(field.as_str(), "Keys" | "Values")
                    && receiver_type_hint
                        .as_deref()
                        .map(|type_hint| {
                            Self::is_dictionary_type_hint(type_hint)
                                || Self::is_sorted_dictionary_type_hint(type_hint)
                        })
                        .unwrap_or(false)
                    && !matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if name.chars().next().map_or(false, |c| c.is_ascii_uppercase())
                    );

                if is_dotnet_dictionary_accessor {
                    self.compile_expr(object)?;
                    if field == "Keys" {
                        common::collections::emit_iter_keys(
                            &mut self.chunks,
                            self.current,
                            self.line,
                        );
                    } else {
                        self.emit_common("dict.values", 1, self.line);
                    }
                    return Ok(());
                }

                let is_csharp_len_accessor = matches!(self.profile.name.as_str(), "csharp" | "vb")
                    && matches!(field.as_str(), "Length" | "Count")
                    && receiver_is_collection_like
                    && !matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if name.chars().next().map_or(false, |c| c.is_ascii_uppercase())
                    );
                let is_csharp_runtime_count_accessor =
                    matches!(self.profile.name.as_str(), "csharp" | "vb")
                        && self.profile.namespaces.use_dotnet
                        && field == "Count"
                        && !is_csharp_len_accessor
                        && !*null_safe
                        && !matches!(
                            &object.kind,
                            ExprKind::Ident(name)
                                if name.chars().next().map_or(false, |c| c.is_ascii_uppercase())
                        );

                if *null_safe && is_csharp_len_accessor {
                    self.compile_expr(object)?;
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                    self.chunk().emit_else(line);
                    common::collections::emit_len(&mut self.chunks, self.current, self.line);
                    self.chunk().emit_end(line);
                    return Ok(());
                } else {
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                }

                let dotnet_instance_property =
                    if matches!(self.profile.name.as_str(), "csharp" | "vb")
                        && self.profile.namespaces.use_dotnet
                        && !*null_safe
                    {
                        receiver_type_hint.as_deref().and_then(|type_hint| {
                            let class_name = Self::normalize_type_hint(type_hint);
                            common::dotnet::surface().lookup_instance_property(&class_name, field)
                        })
                    } else {
                        None
                    };

                if let Some(target) = dotnet_instance_property {
                    match target {
                        common::dotnet::InstancePropertyTarget::Host { module, func } => {
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, 1);
                            return Ok(());
                        }
                    }
                }

                let dotnet_instance_zero_arg_method =
                    if matches!(self.profile.name.as_str(), "csharp" | "vb")
                        && self.profile.namespaces.use_dotnet
                        && !*null_safe
                        && !is_csharp_len_accessor
                        && !is_csharp_runtime_count_accessor
                    {
                        receiver_type_hint.as_deref().and_then(|type_hint| {
                            let class_name = Self::normalize_type_hint(type_hint);
                            common::dotnet::surface().lookup_instance_method(&class_name, field, 0)
                        })
                    } else {
                        None
                    };

                if let Some(target) = dotnet_instance_zero_arg_method {
                    if let common::dotnet::InstanceMethodTarget::Common { emit, .. } = &target {
                        let line = self.line;
                        self.compile_expr(object)?;
                        self.emit_common(emit, 1, line);
                        return Ok(());
                    }

                    let obj_slot = self.define_local("__dotnet_zero_arg_obj");
                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);

                    let value_slot = self.define_local("__dotnet_zero_arg_value");
                    let field_name = self.canon(field);
                    let canonical_idx = self.str_const(&field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::STRUCT_GET, canonical_idx);
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit(Op::DROP);

                    if matches!(self.profile.name.as_str(), "csharp" | "vb")
                        && self.profile.namespaces.use_dotnet
                        && field.as_str() != field_name
                    {
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let line = self.line;
                        self.chunk().emit_if(line);

                        let exact_idx = self.str_const(field);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, exact_idx);
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit(Op::DROP);

                        self.chunk().emit_end(line);
                    }

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit(Op::REF_IS_UNDEFINED);
                    let line = self.line;
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_else(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    match target {
                        common::dotnet::InstanceMethodTarget::Host { module, func, .. } => {
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, 1);
                        }
                        common::dotnet::InstanceMethodTarget::Common { emit, .. } => {
                            let line = self.line;
                            self.emit_common(&emit, 1, line);
                        }
                    }
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                if is_csharp_len_accessor {
                    common::collections::emit_len(&mut self.chunks, self.current, self.line);
                    return Ok(());
                }

                if is_csharp_runtime_count_accessor {
                    let obj_slot = self.define_local("__count_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);

                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::STRUCT_GET, idx);
                    let value_slot = self.define_local("__count_value");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit(Op::REF_TYPEOF);
                    self.emit_const(Value::String(Arc::from("function")));
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                let static_field_owner = if matches!(self.profile.name.as_str(), "csharp" | "vb") {
                    receiver_type_hint.as_deref().and_then(|type_hint| {
                        let trimmed_type_hint = type_hint.trim().trim_end_matches("()").trim();
                        let metadata_type_hint = self
                            .reflection_type_metadata(type_hint)
                            .map(|_| type_hint)
                            .or_else(|| {
                                self.reflection_type_metadata(trimmed_type_hint)
                                    .map(|_| trimmed_type_hint)
                            })?;
                        self.reflection_type_metadata(metadata_type_hint)
                            .and_then(|meta| {
                                meta.fields.iter().find(|(name, field_meta)| {
                                    name.eq_ignore_ascii_case(field) && field_meta.is_static
                                })
                            })
                            .map(|_| {
                                let short_name =
                                    self.reflection_type_short_name(metadata_type_hint);
                                if self.defined_globals.contains(&self.canon(&short_name)) {
                                    short_name
                                } else {
                                    self.reflection_type_lookup_name(metadata_type_hint)
                                }
                            })
                    })
                } else {
                    None
                };

                if self.profile.name == "vb" {
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    if let Some(type_name) = static_field_owner {
                        let obj_slot = self.define_local("__vb_member_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        let value_slot = self.define_local("__vb_member_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let line = self.line;
                        self.chunk().emit_if_value(line);

                        let class_idx = self.str_const(&self.canon(&type_name));
                        self.emit_u16(Op::GLOBAL_GET, class_idx);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.chunk().emit_end(line);
                        return Ok(());
                    }
                }

                if *null_safe
                    && matches!(self.profile.name.as_str(), "csharp" | "vb")
                    && !is_csharp_len_accessor
                {
                    let obj_slot = self.define_local("__dotnet_nullsafe_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                    self.chunk().emit_else(line);
                    self.emit(Op::NULL);
                    self.chunk().emit_end(line);
                } else if *null_safe {
                    let obj_slot = self.define_local("__member_nullsafe_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_OBJECT);
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                    self.chunk().emit_else(line);
                    self.emit(Op::NULL);
                    self.chunk().emit_end(line);
                } else {
                    let field_name = self.canon(field);
                    if matches!(self.profile.name.as_str(), "csharp" | "vb")
                        && self.profile.namespaces.use_dotnet
                        && field.as_str() != field_name
                    {
                        let obj_slot = self.define_local("__dotnet_member_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        self.emit(Op::DROP);

                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        let value_slot = self.define_local("__dotnet_member_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let line = self.line;
                        self.chunk().emit_if_value(line);

                        let exact_idx = self.str_const(field);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, exact_idx);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.chunk().emit_end(line);
                    } else {
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                }
            }

            // ── Index access ────────────────────────────────────────────
            ExprKind::Index {
                object,
                index,
                null_safe,
            } => {
                // A Range used as the index is a slice operation
                // (C# `arr[1..3]` / `s[0..5]`, Python `arr[1:3]` / `s[0:5]`).
                // Route through compiler_common's polymorphic slice helper so
                // strings and arrays both work uniformly.
                if let ExprKind::Range { start, end, .. } = &index.kind {
                    let line = self.line;
                    if self.is_js_profile() {
                        // Emit an inline polymorphic slice for JS: strings use
                        // STR_SUBSTRING, arrays call ecma:array.slice. Save
                        // operands to locals so we can test the receiver.
                        let obj_slot = self.define_local("__js_range_slice_obj");
                        let start_slot = self.define_local("__js_range_slice_start");
                        let end_slot = self.define_local("__js_range_slice_end");

                        self.compile_expr(object)?;
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        self.emit(Op::DROP);

                        self.compile_expr(start)?;
                        self.emit_u16(Op::LOCAL_SET, start_slot);
                        self.emit(Op::DROP);

                        self.compile_expr(end)?;
                        self.emit_u16(Op::LOCAL_SET, end_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_STRING);
                        self.chunk().emit_if_value(line);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.emit_u16(Op::LOCAL_GET, end_slot);
                        self.emit(Op::STR_SUBSTRING);

                        self.chunk().emit_else(line);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.emit_u16(Op::LOCAL_GET, end_slot);
                        common::collections::emit_slice(&mut self.chunks, self.current, line);

                        self.chunk().emit_end(line);
                    } else {
                        let line = self.line;
                        common::collections::emit_slice_push_func(self.chunk(), line);
                        self.compile_expr(object)?;
                        self.compile_expr(start)?;
                        self.compile_expr(end)?;
                        common::collections::emit_slice_invoke(self.chunk(), line);
                    }
                } else if let ExprKind::Slice { lower, upper, step } = &index.kind {
                    self.compile_expr(object)?;
                    let line = self.line;
                    if step.is_none() {
                        if self.is_js_profile() {
                            // JS: compute start/end into locals, then dispatch
                            // to STR_SUBSTRING for strings or ecma:array.slice
                            // for arrays.
                            let obj_slot = self.define_local("__js_index_slice_obj");
                            let start_slot = self.define_local("__js_index_slice_start");
                            let end_slot = self.define_local("__js_index_slice_end");

                            self.emit_u16(Op::LOCAL_SET, obj_slot);
                            self.emit(Op::DROP);

                            if let Some(l) = lower {
                                self.compile_expr(l)?;
                            } else {
                                self.emit(Op::I32_CONST_0);
                            }
                            self.emit_u16(Op::LOCAL_SET, start_slot);
                            self.emit(Op::DROP);

                            if let Some(u) = upper {
                                self.compile_expr(u)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_slot);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            self.emit_u16(Op::LOCAL_SET, end_slot);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_STRING);
                            self.chunk().emit_if_value(line);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u16(Op::LOCAL_GET, start_slot);
                            self.emit_u16(Op::LOCAL_GET, end_slot);
                            self.emit(Op::STR_SUBSTRING);

                            self.chunk().emit_else(line);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u16(Op::LOCAL_GET, start_slot);
                            self.emit_u16(Op::LOCAL_GET, end_slot);
                            common::collections::emit_slice(&mut self.chunks, self.current, line);

                            self.chunk().emit_end(line);
                        } else if self.profile.name == "fortran" {
                            let obj_slot = self.define_local("__fortran_index_slice_obj");
                            self.emit_u16(Op::LOCAL_SET, obj_slot);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            if let Some(l) = lower {
                                self.compile_expr(l)?;
                            } else {
                                self.emit(Op::I32_CONST_0);
                            }
                            if let Some(u) = upper {
                                self.compile_expr(u)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_slot);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            common::collections::emit_runtime_helper_call(
                                &mut self.chunks,
                                self.current,
                                "__vybe_slice",
                                3,
                                line,
                            );
                        } else {
                            let obj_slot = self.define_local("__py_index_slice_obj");
                            self.emit_u16(Op::LOCAL_SET, obj_slot);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            if let Some(l) = lower {
                                self.compile_expr(l)?;
                            } else {
                                self.emit(Op::I32_CONST_0);
                            }
                            if let Some(u) = upper {
                                self.compile_expr(u)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_slot);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            common::collections::emit_runtime_helper_call(
                                &mut self.chunks,
                                self.current,
                                "__vybe_slice",
                                3,
                                line,
                            );
                        }
                    } else {
                        let step_const = step.as_ref().and_then(|expr| match &expr.kind {
                            ExprKind::Lit(Literal::Int(n)) => Some(*n),
                            ExprKind::Unary {
                                op: UnaryOp::Neg,
                                expr,
                            } => match &expr.kind {
                                ExprKind::Lit(Literal::Int(n)) => Some(-*n),
                                _ => None,
                            },
                            _ => None,
                        });

                        if lower.is_none() && upper.is_none() {
                            if step_const == Some(-1) {
                                self.emit(Op::DUP);
                                self.emit(Op::REF_IS_STRING);
                                let line = self.line;
                                self.chunk().emit_if_value(line);
                                self.emit(Op::STR_REVERSE);
                                self.chunk().emit_else(line);
                                self.emit(Op::NULL);
                                self.emit(Op::NULL);
                                if let Some(s) = step {
                                    self.compile_expr(s)?;
                                } else {
                                    self.emit(Op::NULL);
                                }
                                common::collections::emit_runtime_helper_call(
                                    &mut self.chunks,
                                    self.current,
                                    "__vybe_slicestep",
                                    4,
                                    line,
                                );
                                self.chunk().emit_end(line);
                                return Ok(());
                            }

                            if let Some(step_value) = step_const.filter(|n| *n > 1) {
                                self.emit(Op::DUP);
                                self.emit(Op::REF_IS_STRING);
                                let line = self.line;
                                self.chunk().emit_if_value(line);

                                let str_slot = self.define_local("__py_stride_string");
                                let result_slot = self.define_local("__py_stride_result");
                                let index_slot = self.define_local("__py_stride_index");
                                let len_slot = self.define_local("__py_stride_len");

                                self.emit_u16(Op::LOCAL_SET, str_slot);
                                self.emit(Op::DROP);
                                self.emit_const(Value::String(Arc::from("")));
                                self.emit_u16(Op::LOCAL_SET, result_slot);
                                self.emit(Op::DROP);
                                self.emit(Op::I32_CONST_0);
                                self.emit_u16(Op::LOCAL_SET, index_slot);
                                self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, str_slot);
                                self.emit(Op::STR_LENGTH);
                                self.emit_u16(Op::LOCAL_SET, len_slot);
                                self.emit(Op::DROP);

                                let stride_block = self.chunk().emit_block(line);
                                let (stride_loop, _) = self.chunk().emit_loop_s(line);
                                self.emit_u16(Op::LOCAL_GET, index_slot);
                                self.emit_u16(Op::LOCAL_GET, len_slot);
                                {
                                    let line = self.line;
                                    crate::emitter::ops::emit_dyn_lt(self.chunk(), line);
                                };
                                {
                                    let line = self.line;
                                    crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                                };
                                self.chunk().emit_br_if(1, line);

                                self.emit_u16(Op::LOCAL_GET, result_slot);
                                self.emit_u16(Op::LOCAL_GET, str_slot);
                                self.emit_u16(Op::LOCAL_GET, index_slot);
                                self.emit(Op::F64_FROM_I32);
                                self.emit(Op::STR_CHAR_AT);
                                common::strings::emit_str_concat(self.chunk(), line);
                                self.emit_u16(Op::LOCAL_SET, result_slot);
                                self.emit(Op::DROP);

                                self.emit_u16(Op::LOCAL_GET, index_slot);
                                self.emit_const(Value::I32(step_value as i32));
                                self.emit(Op::I32_ADD);
                                self.emit_u16(Op::LOCAL_SET, index_slot);
                                self.emit(Op::DROP);
                                self.chunk().emit_br(0, line);
                                self.chunk().emit_end(line);
                                self.chunk().patch_loop(stride_loop);
                                self.chunk().emit_end(line);
                                self.chunk().patch_block(stride_block);
                                self.emit_u16(Op::LOCAL_GET, result_slot);

                                self.chunk().emit_else(line);
                                self.emit(Op::NULL);
                                self.emit(Op::NULL);
                                if let Some(s) = step {
                                    self.compile_expr(s)?;
                                } else {
                                    self.emit(Op::NULL);
                                }
                                common::collections::emit_runtime_helper_call(
                                    &mut self.chunks,
                                    self.current,
                                    "__vybe_slicestep",
                                    4,
                                    line,
                                );
                                self.chunk().emit_end(line);
                                return Ok(());
                            }
                        }

                        if let Some(l) = lower {
                            self.compile_expr(l)?;
                        } else {
                            self.emit(Op::NULL);
                        }
                        if let Some(u) = upper {
                            self.compile_expr(u)?;
                        } else {
                            self.emit(Op::NULL);
                        }
                        if let Some(s) = step {
                            self.compile_expr(s)?;
                        } else {
                            self.emit(Op::NULL);
                        }
                        common::collections::emit_runtime_helper_call(
                            &mut self.chunks,
                            self.current,
                            "__vybe_slicestep",
                            4,
                            line,
                        );
                    }
                } else if self.profile.name == "pascal"
                    && self.expr_is_known_string_receiver(object)
                {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    self.emit_const(Value::F64(1.0));
                    self.emit(Op::F64_SUB);
                    self.emit(Op::STR_CHAR_AT);
                } else if self.is_js_profile() && self.uses_proxy && !*null_safe {
                    // Proxy get-trap dispatch on bracket-notation reads.
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    let line = self.line;
                    crate::emitter::js::proxy_adapter::emit_proxy_get_dispatch(
                        &mut self.chunks,
                        self.current,
                        line,
                    );
                } else if self.is_js_profile() && *null_safe {
                    self.compile_expr(object)?;
                    let obj_slot = self.define_local("__js_index_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    common::expressions::emit_undefined(self.chunk(), line);
                    self.chunk().emit_else(line);
                    let key_slot = self.define_local("__js_index_key");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    match &index.kind {
                        ExprKind::Member {
                            object,
                            field,
                            null_safe: false,
                        } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") => {
                            let fallback_key = match field.as_str() {
                                "iterator" => Some("iterator"),
                                "asyncIterator" => Some("asyncIterator"),
                                "toPrimitive" => Some("toprimitive"),
                                "hasInstance" => Some("hasinstance"),
                                _ => None,
                            };
                            if let Some(fallback_key) = fallback_key {
                                self.emit_const(Value::String(Arc::from(fallback_key)));
                            } else {
                                self.compile_expr(index)?;
                            }
                        }
                        _ => self.compile_expr(index)?,
                    }
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    if let Some(ns) = self
                        .infer_expr_type_hint(object)
                        .as_deref()
                        .map(Self::normalize_type_hint)
                        .and_then(|type_hint| match type_hint.as_str() {
                            "bigint64array" => Some("ecma:bigint64array"),
                            "biguint64array" => Some("ecma:biguint64array"),
                            _ => None,
                        })
                    {
                        let idx = self.import(ns, "get");
                        self.emit_host_call(idx, 2);
                    } else {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    let val_slot = self.define_local("__js_index_val");
                    self.emit_u16(Op::LOCAL_SET, val_slot);
                    self.emit(Op::DROP);
                    // String[n] out-of-bounds: ARRAY_GET returns null, but JS spec (§6.1.4.1) needs undefined.
                    {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_STRING);
                        let string_line = self.line;
                        self.chunk().emit_if(string_line);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.emit(Op::REF_IS_NULL);
                        let null_line = self.line;
                        self.chunk().emit_if(null_line);
                        self.emit(Op::UNDEFINED);
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(null_line);
                        self.chunk().emit_end(string_line);
                    }
                    let lookup = self.str_const("__vybe_js_get_method");
                    self.emit_u16(Op::LOCAL_GET, val_slot);
                    self.emit(Op::REF_IS_UNDEFINED);
                    let lookup_line = self.line;
                    self.chunk().emit_if_value(lookup_line);
                    self.emit_u16(Op::GLOBAL_GET, lookup);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    match &index.kind {
                        ExprKind::Member {
                            object,
                            field,
                            null_safe: false,
                        } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") => {
                            let fallback_key = match field.as_str() {
                                "iterator" => Some("iterator"),
                                "asyncIterator" => Some("asyncIterator"),
                                "toPrimitive" => Some("toprimitive"),
                                "hasInstance" => Some("hasinstance"),
                                _ => None,
                            };
                            if let Some(fallback_key) = fallback_key {
                                self.emit_const(Value::String(Arc::from(fallback_key)));
                            } else {
                                self.emit_u16(Op::LOCAL_GET, key_slot);
                            }
                        }
                        _ => self.emit_u16(Op::LOCAL_GET, key_slot),
                    }
                    self.emit_u8(Op::CALL_REF, 2);
                    self.chunk().emit_else(lookup_line);
                    self.emit_u16(Op::LOCAL_GET, val_slot);
                    self.chunk().emit_end(lookup_line);
                    self.chunk().emit_end(line);
                } else if matches!(self.profile.name.as_str(), "csharp" | "vb") {
                    self.compile_expr(object)?;
                    let obj_slot = self.define_local("__index_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);

                    let null_safe_if = if *null_safe {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit(Op::NULL);
                        self.chunk().emit_else(line);
                        Some(line)
                    } else {
                        None
                    };

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let getter_key = self.str_const("__get___index__");
                    self.emit_u16(Op::STRUCT_GET, getter_key);
                    let getter_slot = self.define_local("__index_getter");
                    self.emit_u16(Op::LOCAL_SET, getter_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, getter_slot);
                    self.emit(Op::REF_IS_NULL);
                    let getter_line = self.line;
                    self.chunk().emit_if_value(getter_line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.compile_collection_key(object, index)?;
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    self.chunk().emit_else(getter_line);
                    self.emit_u16(Op::LOCAL_GET, getter_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.compile_collection_key(object, index)?;
                    self.emit_u8(Op::CALL_REF, 2);
                    self.chunk().emit_end(getter_line);
                    if let Some(line) = null_safe_if {
                        self.chunk().emit_end(line);
                    }
                } else {
                    if self.is_js_profile() {
                        self.compile_expr(object)?;
                        let obj_slot = self.define_local("__js_index_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let msg_line = self.line;
                        self.chunk().emit_if_value(msg_line);
                        self.emit_const(Value::String(Arc::from(
                            "Cannot read properties of undefined",
                        )));
                        self.chunk().emit_else(msg_line);
                        self.emit_const(Value::String(Arc::from("Cannot read properties of null")));
                        self.chunk().emit_end(msg_line);
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);

                        let key_slot = self.define_local("__js_index_key");
                        match &index.kind {
                            ExprKind::Member {
                                object,
                                field,
                                null_safe: false,
                            } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") =>
                            {
                                let fallback_key = match field.as_str() {
                                    "iterator" => Some("iterator"),
                                    "asyncIterator" => Some("asyncIterator"),
                                    "toPrimitive" => Some("toprimitive"),
                                    "hasInstance" => Some("hasinstance"),
                                    _ => None,
                                };
                                if let Some(fallback_key) = fallback_key {
                                    self.emit_const(Value::String(Arc::from(fallback_key)));
                                } else {
                                    self.compile_array_index_operand_for_owner(object, index)?;
                                }
                            }
                            _ => self.compile_array_index_operand_for_owner(object, index)?,
                        }
                        if self.profile.negative_index_wraps {
                            self.emit_negative_index_wrap();
                        }
                        self.emit_u16(Op::LOCAL_SET, key_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::LOCAL_GET, key_slot);
                        if let Some(ns) = self
                            .infer_expr_type_hint(object)
                            .as_deref()
                            .map(Self::normalize_type_hint)
                            .and_then(|type_hint| match type_hint.as_str() {
                                "bigint64array" => Some("ecma:bigint64array"),
                                "biguint64array" => Some("ecma:biguint64array"),
                                _ => None,
                            })
                        {
                            let idx = self.import(ns, "get");
                            self.emit_host_call(idx, 2);
                        } else {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        let val_slot = self.define_local("__js_index_val");
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        self.emit(Op::DROP);
                        // String[n] out-of-bounds: ARRAY_GET returns null, but JS spec needs undefined.
                        {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_STRING);
                            let string_line = self.line;
                            self.chunk().emit_if(string_line);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.emit(Op::REF_IS_NULL);
                            let null_line = self.line;
                            self.chunk().emit_if(null_line);
                            self.emit(Op::UNDEFINED);
                            self.emit_u16(Op::LOCAL_SET, val_slot);
                            self.emit(Op::DROP);
                            self.chunk().emit_end(null_line);
                            self.chunk().emit_end(string_line);
                        }
                        let lookup = self.str_const("__vybe_js_get_method");
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let lookup_line = self.line;
                        self.chunk().emit_if_value(lookup_line);
                        self.emit_u16(Op::GLOBAL_GET, lookup);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        match &index.kind {
                            ExprKind::Member {
                                object,
                                field,
                                null_safe: false,
                            } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") =>
                            {
                                let fallback_key = match field.as_str() {
                                    "iterator" => Some("iterator"),
                                    "asyncIterator" => Some("asyncIterator"),
                                    "toPrimitive" => Some("toprimitive"),
                                    "hasInstance" => Some("hasinstance"),
                                    _ => None,
                                };
                                if let Some(fallback_key) = fallback_key {
                                    self.emit_const(Value::String(Arc::from(fallback_key)));
                                } else {
                                    self.emit_u16(Op::LOCAL_GET, key_slot);
                                }
                            }
                            _ => self.emit_u16(Op::LOCAL_GET, key_slot),
                        }
                        self.emit_u8(Op::CALL_REF, 2);
                        self.chunk().emit_else(lookup_line);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.chunk().emit_end(lookup_line);
                        return Ok(());
                    }
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    self.compile_array_index_operand_for_owner(object, index)?;
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                }
            }

            // ── New ─────────────────────────────────────────────────────
            ExprKind::New { class, args } => {
                let reordered_args;
                let args = if args.iter().any(|arg| arg.name.is_some()) {
                    let ctor_key = match &class.kind {
                        ExprKind::Ident(name) => Some(self.canon(name)),
                        ExprKind::Member { field, .. } => Some(self.canon(field)),
                        _ => None,
                    };
                    if let Some(signatures) = ctor_key
                        .as_ref()
                        .and_then(|key| self.constructor_signatures.get(key))
                    {
                        reordered_args = self.reorder_named_args_with_signatures(args, signatures);
                        reordered_args.as_slice()
                    } else {
                        args.as_slice()
                    }
                } else {
                    args.as_slice()
                };
                // ECMA-262 §10.5.2: `new Proxy(target, handler)` creates an
                // exotic object whose property accesses are intercepted by
                // handler traps. We lower to an inline emitter that
                // produces a wrapper Ordinary object stamped with
                // `__vybe_proxy_target` + `__vybe_proxy_handler`. The
                // module-level `uses_proxy` flag (set by the AST scan)
                // tells subsequent Member/Index emits to route through
                // the dispatcher.
                if self.is_js_profile() {
                    if let ExprKind::Ident(name) = &class.kind {
                        if name == "Set" && args.len() <= 1 {
                            if let Some(arg) = args.first() {
                                self.compile_expr(&arg.value)?;
                                let v_slot = self.define_local("__js_set_iterable");
                                self.emit_u16(Op::LOCAL_SET, v_slot);
                                self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, v_slot);
                                let is_gen_idx = self.import("ecma:value", "isGenerator");
                                self.emit_host_call(is_gen_idx, 1);
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                                self.chunk().emit_if_value(line);
                                let drain_key = self.str_const("__vybe_drain_generator");
                                self.emit_u16(Op::GLOBAL_GET, drain_key);
                                self.emit_u16(Op::LOCAL_GET, v_slot);
                                self.emit_u8(Op::CALL_REF, 1);
                                self.chunk().emit_else(line);
                                let iter_drain_key = self.str_const("__vybe_iter_drain");
                                self.emit_u16(Op::GLOBAL_GET, iter_drain_key);
                                self.emit_u16(Op::LOCAL_GET, v_slot);
                                self.emit_u8(Op::CALL_REF, 1);
                                self.chunk().emit_end(line);
                            } else {
                                common::collections::emit_array_new(
                                    &mut self.chunks,
                                    self.current,
                                    0,
                                    self.line,
                                );
                            }
                            let idx = self.import("ecma:set", "fromIterable");
                            self.emit_host_call(idx, 1);
                            return Ok(());
                        }

                        if name == "Function" {
                            for arg in args {
                                self.compile_expr(&arg.value)?;
                            }
                            let idx = self.import("vybe:js", "function_constructor");
                            self.emit_host_call(idx, args.len() as u8);
                            return Ok(());
                        }

                        if name == "Proxy" && args.len() == 2 {
                            self.uses_proxy = true;
                            self.compile_expr(&args[0].value)?;
                            self.compile_expr(&args[1].value)?;
                            let line = self.line;
                            crate::emitter::js::proxy_adapter::emit_proxy_create(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            return Ok(());
                        }
                    }
                }
                let class_parts = self.flatten_member_chain(class);
                let dotted_type_name = match &class.kind {
                    ExprKind::Ident(name) => Some(self.resolve_source_type_alias(name)),
                    ExprKind::Member { .. }
                        if self.profile.namespaces.use_dotnet && !class_parts.is_empty() =>
                    {
                        Some(self.resolve_source_type_alias(&class_parts.join(".")))
                    }
                    _ => None,
                };
                let php_autoload_name = match &class.kind {
                    ExprKind::Ident(name) if self.is_php_profile() => {
                        Some(Self::strip_global_namespace_prefix(name).to_string())
                    }
                    _ => None,
                };

                if let Some(type_name) = dotted_type_name.as_ref() {
                    // User-defined classes take priority over all built-in type mappings.
                    // This ensures `class Point { ... }` followed by `new Point()` calls
                    // the user constructor, not vybe:drawing::pointNew.
                    let canon_type = self.canon(type_name);
                    if self.defined_classes.contains(&canon_type) {
                        let overload_global = format!("{}$arity{}", canon_type, args.len());
                        let ctor_global = if self.defined_globals.contains(&overload_global) {
                            overload_global
                        } else {
                            canon_type.clone()
                        };
                        // Bypass compile_expr to avoid the implicit-self-field
                        // shadowing path: in case-insensitive languages a field
                        // named `inner` and a class named `Inner` both
                        // canonicalize to "inner", and the implicit-self-field
                        // check would mis-route to `me.inner` instead of the
                        // class global. Type names always come from globals.
                        let autoload_name = php_autoload_name.as_deref().unwrap_or(type_name);
                        self.emit_constructor_global_ref(&ctor_global, autoload_name);
                        for a in args {
                            self.compile_expr(&a.value)?;
                        }
                        // §13.3.5: new.target is the invoked constructor for
                        // the whole construction chain (parent ctor bodies
                        // under super() included). Unique save slot — nested
                        // `new` in the same function must not share it.
                        let saved_nt = self.save_js_new_target(&format!(
                            "__js_prev_nt_static_{}",
                            self.chunks[self.current].local_count
                        ));
                        if saved_nt.is_some() {
                            let class_idx = self.global_name_const_idx(&canon_type);
                            self.emit_u16(Op::GLOBAL_GET, class_idx);
                            self.set_js_new_target_from_stack();
                        }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);
                        self.restore_js_new_target(saved_nt);
                        return Ok(());
                    }
                    if self.is_js_profile() && self.defined_functions.contains(&canon_type) {
                        let idx = self.str_const(&canon_type);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        let ctor_slot = self.define_local("__js_ctor");
                        self.emit_u16(Op::LOCAL_SET, ctor_slot);
                        self.emit(Op::DROP);
                        let line = self.line;
                        let saved_js_new_target =
                            self.save_js_new_target("__js_prev_new_target_new");
                        self.emit_u16(Op::LOCAL_GET, ctor_slot);
                        self.set_js_new_target_from_stack();
                        let _ = line;
                        let (args_slot, _) = self.compile_call_args_array(args, "js_new")?;
                        self.emit_u16(Op::LOCAL_GET, ctor_slot);
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        let reflect_construct = self.import("ecma:reflect", "construct");
                        self.emit_host_call(reflect_construct, 2);
                        self.restore_js_new_target(saved_js_new_target);
                        return Ok(());
                    }
                    if self.is_php_profile() {
                        if let Some(autoload_name) = php_autoload_name.as_deref() {
                            if let Some(flattened_name) = autoload_name.rsplit('\\').next() {
                                let flattened_canon = self.canon(flattened_name);
                                if self.defined_classes.contains(&flattened_canon) {
                                    let overload_global =
                                        format!("{}$arity{}", flattened_canon, args.len());
                                    let ctor_global =
                                        if self.defined_globals.contains(&overload_global) {
                                            overload_global
                                        } else {
                                            flattened_canon
                                        };
                                    self.emit_constructor_global_ref(&ctor_global, autoload_name);
                                    for a in args {
                                        self.compile_expr(&a.value)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                                    return Ok(());
                                }
                            }
                        }
                    }
                    // Nested class: `new Outer.Inner()` — the inner type
                    // is registered as a sibling global per ECMA-334 §15.3.
                    // Try the last segment as a type name when the full
                    // dotted form misses.
                    if class_parts.len() > 1 {
                        let last = class_parts.last().unwrap();
                        let canon_last = self.canon(last);
                        if self.defined_classes.contains(&canon_last) {
                            let autoload_name = php_autoload_name.as_deref().unwrap_or(type_name);
                            self.emit_constructor_global_ref(&canon_last, autoload_name);
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            return Ok(());
                        }
                    }

                    let bare = type_name.to_lowercase();
                    let bare = bare.split('(').next().unwrap_or(&bare).trim();
                    let bare_str = bare.rsplit('.').next().unwrap_or(bare);

                    // WASM threading/async — use compiler_common, NOT host calls
                    match bare_str {
                        "thread" => {
                            // New Thread(callback) → cont_new only (Start resumes)
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            return Ok(());
                        }
                        "task" => {
                            // New Task(callback) → cont_new only
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            return Ok(());
                        }
                        "mutex" | "semaphore" => {
                            // New Mutex() → allocate atomic address for lock
                            self.emit_const(Value::I32(0)); // initial lock value
                            return Ok(());
                        }
                        _ => {}
                    }

                    // Built-in exception types — route through compiler_common
                    // so that every language produces the canonical 4-field
                    // shape and the type name is normalized. PHP `RuntimeException`,
                    // Python `RuntimeError`, JS `Error`, etc. all produce identical
                    // bytecode and can catch each other cross-language.
                    if common::errors::is_exception_type(bare_str) {
                        let ctor_args: Vec<&Expression> = args.iter().map(|a| &a.value).collect();
                        self.emit_js_exception_ctor_value(type_name, &ctor_args)?;
                        return Ok(());
                    }

                    // GUI control: Button, TextBox, Label, Timer, etc.
                    // Checked BEFORE dotnet known_types so GUI controls always
                    // route through the canonical gui emitter regardless of
                    // whether they overlap with .NET BCL types (Timer is both
                    // a GUI control and a System.Threading.Timer — the GUI
                    // form takes priority because we're in `New X()` syntax).
                    let dotnet_ctor_registered = self.profile.namespaces.use_dotnet
                        && (self.defined_globals.contains(bare_str)
                            || self.defined_globals.contains(&bare_str.to_lowercase()));
                    let canonical = common::gui::canonical_control_name(bare_str);
                    if !canonical.is_empty() && !dotnet_ctor_registered {
                        let host_name = common::gui::host_fn_new_control(&canonical);
                        let new_idx = self.import("vybe:gui", &host_name);
                        for a in args {
                            self.compile_expr(&a.value)?;
                        }
                        let line = self.line;
                        common::gui::emit_new_control(
                            self.chunk(),
                            new_idx,
                            args.len() as u8,
                            line,
                        );
                        return Ok(());
                    }
                    // Dotnet component descriptor constructors — fallback after
                    // GUI so .NET-only types like Dictionary still work.
                    let dotnet_constructor = common::dotnet::surface().lookup_constructor(bare_str);
                    if let Some(target) = dotnet_constructor.clone() {
                        for a in args {
                            self.compile_expr(&a.value)?;
                        }
                        // Proper-case class name (preserve from source) for
                        // the __type stamp. Some host fns (e.g.
                        // `vybe:types/collectionPeek`) compare __type with
                        // exact case, so a lowercase stamp would clobber.
                        let proper_name: String = type_name
                            .split('(')
                            .next()
                            .unwrap_or(type_name)
                            .trim()
                            .rsplit('.')
                            .next()
                            .unwrap_or(type_name)
                            .to_string();
                        match target {
                            vybe_bytecode::component_model::ConstructorTarget::Host(target) => {
                                let idx = self.import(&target.module, &target.name);
                                self.emit_host_call(idx, args.len() as u8);
                            }
                            vybe_bytecode::component_model::ConstructorTarget::Common(name) => {
                                let line = self.line;
                                self.emit_common(&name, args.len() as u8, line);
                            }
                        }
                        // Stamp `__type` with the .NET class name so the
                        // runtime TypeRegistry dispatches `d.Add(...)` /
                        // `d.ContainsKey(...)` against the .NET adapter
                        // TypeDef (e.g. `Dictionary`) — not against the
                        // underlying ECMA type (e.g. `Map`). The .NET
                        // adapter's methods are aliases pointing at the
                        // same `ecma:*` host fns, so the underlying
                        // implementation is standardized while the
                        // surface stays .NET-shaped.
                        //
                        // Stack: [obj] → [obj, obj, name] → [obj, obj'] → [obj]
                        self.emit(Op::DUP);
                        self.emit_const(Value::String(Arc::from(proper_name.as_str())));
                        let type_key = self.str_const("__type");
                        self.emit_u16(Op::STRUCT_SET, type_key);
                        self.emit(Op::DROP);

                        // .NET List / ArrayList instance calls like `list.Sort()`
                        // should stay on the shared compare-aware frontend path
                        // instead of falling through to the runtime collection
                        // vtable's raw host sort mapping.
                        let bare_proper_name = proper_name
                            .split('<')
                            .next()
                            .map(str::trim)
                            .unwrap_or(proper_name.as_str());
                        if matches!(bare_proper_name, "List" | "ArrayList") {
                            let sort_global = self.str_const("__vybe_sort_in_place");
                            let sort_key = self.str_const("sort");
                            self.emit(Op::DUP);
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.emit_u16(Op::STRUCT_SET, sort_key);
                            self.emit(Op::DROP);

                            let sort_pascal_key = self.str_const("Sort");
                            self.emit(Op::DUP);
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.emit_u16(Op::STRUCT_SET, sort_pascal_key);
                            self.emit(Op::DROP);
                        }
                        return Ok(());
                    }

                    // Profile known types are now a fallback for entries not
                    // yet absorbed into the shared .NET surface.
                    if dotnet_constructor.is_none() {
                        if let Some((module, func)) = self
                            .profile
                            .lookup_known_type(type_name)
                            .map(|(m, f)| (m.to_string(), f.to_string()))
                        {
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            if module == "common" {
                                let line = self.line;
                                self.emit_common(&func, args.len() as u8, line);
                            } else {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, args.len() as u8);
                            }
                            return Ok(());
                        }
                    }

                    if dotnet_ctor_registered {
                        let ctor_name = if self.defined_globals.contains(bare_str) {
                            bare_str.to_string()
                        } else {
                            bare_str.to_lowercase()
                        };
                        if self.is_php_profile() {
                            let autoload_name = php_autoload_name.as_deref().unwrap_or(type_name);
                            self.emit_constructor_global_ref(&ctor_name, autoload_name);
                        } else {
                            let ctor_idx = self.str_const(&ctor_name);
                            self.emit_u16(Op::GLOBAL_GET, ctor_idx);
                        }
                        for a in args {
                            self.compile_expr(&a.value)?;
                        }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);

                        if bare_str.eq_ignore_ascii_case("list")
                            || bare_str.eq_ignore_ascii_case("arraylist")
                        {
                            let sort_global = self.str_const("__vybe_sort_in_place");
                            let sort_key = self.str_const("sort");
                            self.emit(Op::DUP);
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.emit_u16(Op::STRUCT_SET, sort_key);
                            self.emit(Op::DROP);

                            let sort_pascal_key = self.str_const("Sort");
                            self.emit(Op::DUP);
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.emit_u16(Op::STRUCT_SET, sort_pascal_key);
                            self.emit(Op::DROP);
                        }
                        return Ok(());
                    }

                    if self.is_php_profile() {
                        if let ExprKind::Ident(name) = &class.kind {
                            let autoload_name =
                                Self::strip_global_namespace_prefix(name).to_string();
                            let ctor_base = autoload_name
                                .rsplit('\\')
                                .next()
                                .unwrap_or(autoload_name.as_str());
                            let fallback_ctor = self.canon(ctor_base);
                            let primary_ctor = format!("{}$arity{}", fallback_ctor, args.len());
                            self.emit_dynamic_constructor_global_ref(
                                &primary_ctor,
                                Some(&fallback_ctor),
                                &autoload_name,
                            );
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            return Ok(());
                        }
                    }
                }
                if self.is_js_profile() {
                    self.compile_expr(class)?;
                    let ctor_slot = self.define_local("__js_ctor");
                    self.emit_u16(Op::LOCAL_SET, ctor_slot);
                    self.emit(Op::DROP);
                    let saved_js_new_target = self.save_js_new_target("__js_prev_new_target_new");
                    self.emit_u16(Op::LOCAL_GET, ctor_slot);
                    self.set_js_new_target_from_stack();
                    let (args_slot, _) = self.compile_call_args_array(args, "js_new")?;
                    self.emit_u16(Op::LOCAL_GET, ctor_slot);
                    self.emit_u16(Op::LOCAL_GET, args_slot);
                    let reflect_construct = self.import("ecma:reflect", "construct");
                    self.emit_host_call(reflect_construct, 2);
                    self.restore_js_new_target(saved_js_new_target);
                    return Ok(());
                }

                // User-defined class constructor
                self.compile_expr(class)?;
                for a in args {
                    self.compile_expr(&a.value)?;
                }
                self.emit_u8(Op::CALL_REF, args.len() as u8);
            }

            // ── Assignment as expression ────────────────────────────────
            ExprKind::Assign { target, value } => {
                if matches!(value.kind, ExprKind::Lit(Literal::Null)) {
                    if let ExprKind::Ident(name) = &target.kind {
                        self.emit_buffered_generator_close_ident_if_needed(name);
                    }
                }
                self.compile_expr(value)?;
                self.emit(Op::DUP);
                self.compile_assign_target(target)?;
            }

            // ── Lambda ──────────────────────────────────────────────────
            ExprKind::Lambda {
                params,
                body,
                captures,
                is_async,
            } => {
                if self.is_js_profile() {
                    let mut lexical_captures = captures.clone();
                    if !lexical_captures
                        .iter()
                        .any(|capture| capture == "__js_this" || capture == "&__js_this")
                    {
                        lexical_captures.push("__js_this".to_string());
                    }
                    if !lexical_captures.iter().any(|capture| {
                        capture == "__js_new_target" || capture == "&__js_new_target"
                    }) {
                        lexical_captures.push("__js_new_target".to_string());
                    }
                    self.compile_lambda_with_flags(
                        params,
                        body,
                        &lexical_captures,
                        *is_async,
                        false,
                    )?;
                } else {
                    self.compile_lambda_with_flags(params, body, captures, *is_async, false)?;
                }
            }

            // ── Array literal ───────────────────────────────────────────
            ExprKind::Array(elements) => {
                // Array literals funnel through `common::collections` so
                // every language emits the same import shape.
                //
                // Dispatch on whether ANY element has an explicit key:
                //   - no keys  → `ecma:array` path (integer-indexed,
                //                fast, array-y semantics)
                //   - any key  → `ecma:map` path (IndexMap<Value,Value>,
                //                PHP-shaped associative: mixed int + string
                //                keys preserved in insertion order). Once
                //                a Map, accessors `$a[$k]` use
                //                `ecma:array.get/.set` which now
                //                dispatch polymorphically on Map.
                let line = self.line;
                let is_js_profile = self.is_js_profile();
                let is_js_array_elision = |elem: &ArrayElement| {
                    is_js_profile
                        && !elem.spread
                        && matches!(&elem.key, Some(key) if matches!(key.kind, ExprKind::Lit(Literal::Int(-1))))
                        && matches!(elem.value.kind, ExprKind::Lit(Literal::Undefined))
                };
                let has_keys = elements
                    .iter()
                    .any(|e| e.key.is_some() && !is_js_array_elision(e));
                let has_elisions = is_js_profile && elements.iter().any(is_js_array_elision);

                if self.profile.name == "c"
                    && !has_keys
                    && !has_elisions
                    && elements.iter().all(|elem| !elem.spread)
                {
                    for elem in elements {
                        self.compile_expr(&elem.value)?;
                    }
                    self.emit_u16(Op::ARRAY_NEW_FIXED, elements.len() as u16);
                    return Ok(());
                }

                if has_keys {
                    common::collections::emit_map_new(&mut self.chunks, self.current, line);
                    let mut next_auto_idx: i64 = 0;
                    for elem in elements {
                        if elem.spread {
                            // Spread into a keyed literal isn't
                            // meaningful for Maps in PHP/Python/Ruby
                            // semantics in Phase 1 — skip.
                            continue;
                        }
                        // Stack: [map]
                        self.emit(Op::DUP); // [map, map]
                        match &elem.key {
                            Some(k) => {
                                // Explicit key expression. If it's a
                                // numeric literal, bump the auto index
                                // past it so subsequent unkeyed elements
                                // don't collide (PHP semantics).
                                self.compile_expr(k)?;
                                match &k.kind {
                                    ExprKind::Lit(crate::ast::Literal::Int(n)) => {
                                        next_auto_idx = *n + 1;
                                    }
                                    ExprKind::Lit(crate::ast::Literal::Float(n))
                                        if n.fract() == 0.0 =>
                                    {
                                        next_auto_idx = (*n as i64) + 1;
                                    }
                                    _ => {}
                                }
                            }
                            None => {
                                // Unkeyed element inside a keyed literal
                                // (PHP allows `['x' => 1, 'y']`) — auto
                                // assign the next integer key.
                                self.emit_const(Value::I32(next_auto_idx as i32));
                                next_auto_idx += 1;
                            }
                        }
                        // [map, map, key]
                        self.compile_expr(&elem.value)?;
                        // [map, map, key, value]
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        // set returns null per spec — drop it, map stays on TOS
                        self.emit(Op::DROP);
                    }
                } else {
                    if has_elisions {
                        let array_new_idx = self.import("ecma:array", "new");
                        self.emit_const(Value::I32(elements.len() as i32));
                        self.emit_host_call(array_new_idx, 1);
                        for (index, elem) in elements.iter().enumerate() {
                            if is_js_array_elision(elem) {
                                continue;
                            }
                            self.emit(Op::DUP);
                            self.emit_const(Value::I32(index as i32));
                            self.compile_expr(&elem.value)?;
                            common::collections::emit_set(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                        }
                    } else {
                        // All-unkeyed: use the array path (fast, small).
                        common::collections::emit_array_new(
                            &mut self.chunks,
                            self.current,
                            0,
                            line,
                        );
                        for elem in elements {
                            if elem.spread {
                                // Spread: `concat(current, other)` returns a NEW
                                // array which replaces the one on TOS. JS
                                // generators (Continuation values) can't be
                                // spread by the host concat fn — the iterator
                                // protocol needs spec stack-switching `resume`
                                // (emit_next). When we know at compile time that
                                // the spread value is a direct generator call,
                                // use the common emitter (emit_drain_into_array)
                                // which emits an inline GEN_NEXT loop — the same
                                // opcode compile_generator_for_in uses directly.
                                // No runtime helper, no runtime dispatch, no isGenerator check.
                                // For unknown-shape values, fall back to the runtime
                                // isGenerator dispatch via stdlib.
                                // A direct call to a generator function can be
                                // drained at compile time via the common GEN_NEXT
                                // loop (WASM stack switching) — works for ANY
                                // language with generators (JS, PHP), no runtime
                                // isGenerator check, no runtime helper.
                                let is_known_gen = (self.is_js_profile()
                                    || self.profile.buffered_iterator_methods)
                                    && self.is_direct_generator_call(&elem.value);
                                self.compile_expr(&elem.value)?;
                                if is_known_gen {
                                    // continuation is on TOS; drain into array
                                    crate::emitter::generators::emit_drain_into_array(
                                        &mut self.chunks,
                                        self.current,
                                        line,
                                    );
                                    common::collections::emit_concat(
                                        &mut self.chunks,
                                        self.current,
                                        line,
                                    );
                                    continue;
                                }
                                if self.is_js_profile() {
                                    let v_slot = self.define_local("__arr_spread_v");
                                    self.emit_u16(Op::LOCAL_SET, v_slot);
                                    self.emit(Op::DROP);
                                    self.emit_u16(Op::LOCAL_GET, v_slot);
                                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                                    self.emit_host_call(is_gen_idx, 1);
                                    let line = self.line;
                                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                                    self.chunk().emit_if_value(line);
                                    let drain_key = self.str_const("__vybe_drain_generator");
                                    self.emit_u16(Op::GLOBAL_GET, drain_key);
                                    self.emit_u16(Op::LOCAL_GET, v_slot);
                                    self.emit_u8(Op::CALL_REF, 1);
                                    self.chunk().emit_else(line);
                                    // Non-generator branch: route through
                                    // `__vybe_iter_drain` JS polyfill which
                                    // calls `v.iterator()` and drains the
                                    // protocol with `__js_this` correctly
                                    // bound. For Array / built-ins it
                                    // returns the value unchanged so concat
                                    // sees the natural shape.
                                    let iter_drain_key = self.str_const("__vybe_iter_drain");
                                    self.emit_u16(Op::GLOBAL_GET, iter_drain_key);
                                    self.emit_u16(Op::LOCAL_GET, v_slot);
                                    self.emit_u8(Op::CALL_REF, 1);
                                    self.chunk().emit_end(line);
                                }
                                common::collections::emit_concat(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                            } else {
                                // DUP keeps the array on TOS; push returns the
                                // new length, which we drop.
                                self.emit(Op::DUP);
                                self.compile_expr(&elem.value)?;
                                common::collections::emit_push(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.emit(Op::DROP);
                            }
                        }
                    }
                }
            }

            // ── Tuple (Python) ──────────────────────────────────────────
            ExprKind::Tuple(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements {
                    self.compile_expr(elem)?;
                }
                // Allocate N consecutive slots; common::collections::emit_pack_n
                // stashes stack values and re-pushes into a fresh array —
                // same ecma:array.* surface as literals.
                let base = if n == 0 {
                    0
                } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.define_local("__pack");
                        if i == 0 {
                            first = s;
                        }
                    }
                    first
                };
                common::collections::emit_pack_n(
                    &mut self.chunks,
                    self.current,
                    n as u16,
                    base,
                    line,
                );
            }

            // ── Set (Python) ────────────────────────────────────────────
            ExprKind::Set(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements {
                    self.compile_expr(elem)?;
                }
                let base = if n == 0 {
                    0
                } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.define_local("__pack");
                        if i == 0 {
                            first = s;
                        }
                    }
                    first
                };
                common::collections::emit_pack_n(
                    &mut self.chunks,
                    self.current,
                    n as u16,
                    base,
                    line,
                );
                // Convert the packed array to a Set per ECMA-262 §24.2.1.1
                // — `new Set(iterable)`. Same import V8 satisfies natively
                // for `new Set([1,2,3])`.
                let idx = self.import("ecma:set", "fromIterable");
                self.emit_host_call(idx, 1);
            }

            // ── Object literal ──────────────────────────────────────────
            // Uses dict::emit_new to create the object WITH __keys tracking.
            // Each key is set via struct_set AND appended to __keys so that
            // Object.keys/values/entries (which read __keys) return the
            // right answer.
            ExprKind::Object(props) => {
                let line = self.line;
                common::dict::emit_new(&mut self.chunks, self.current, line);
                for prop in props {
                    match prop {
                        ObjectProperty::KeyValue { key, value } => {
                            if let ExprKind::Lit(Literal::Str(k)) = &key.kind {
                                // Static string key — fast path via struct_set.
                                let key_name = if self.case_sensitive {
                                    k.clone()
                                } else {
                                    self.canon(k)
                                };
                                // Track insertion order BEFORE setting the
                                // property. `trackKey` dedups by skipping keys
                                // already present — so `{ ...base, x: 99 }`
                                // (where `x` arrived via the spread) does NOT
                                // append a duplicate `x` to `__keys`. Must run
                                // pre-set: post-set the property always exists
                                // and tracking would be skipped entirely.
                                if self.is_js_profile() {
                                    self.emit(Op::DUP);
                                    self.emit_const(Value::String(Arc::from(key_name.as_str())));
                                    let track_idx = self.import("ecma:object", "trackKey");
                                    self.emit_host_call(track_idx, 2);
                                    self.emit(Op::DROP);
                                }
                                self.emit(Op::DUP);
                                self.compile_expr(value)?;
                                if self.is_js_profile() {
                                    let should_infer_name = match &value.kind {
                                        ExprKind::Lambda { .. } => true,
                                        ExprKind::FunctionExpr(stmt) => {
                                            matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name.is_empty())
                                        }
                                        ExprKind::ClassExpr { name, .. } => name.is_none(),
                                        _ => false,
                                    };
                                    if should_infer_name {
                                        let inferred_name = if let Some(stripped) =
                                            key_name.strip_prefix("__get_")
                                        {
                                            format!("get {}", stripped)
                                        } else if let Some(stripped) =
                                            key_name.strip_prefix("__set_")
                                        {
                                            format!("set {}", stripped)
                                        } else {
                                            key_name.clone()
                                        };
                                        self.emit(Op::DUP);
                                        self.emit_const(Value::String(Arc::from(
                                            inferred_name.as_str(),
                                        )));
                                        let name_key = self.str_const("name");
                                        self.emit_u16(Op::STRUCT_SET, name_key);
                                        self.emit(Op::DROP);
                                    }
                                }
                                let idx = self.str_const(&key_name);
                                self.emit_u16(Op::STRUCT_SET, idx);
                                self.emit(Op::DROP);
                                // Non-JS: append to __keys directly (JS already
                                // tracked it via the deduping `trackKey` above).
                                if !self.is_js_profile() {
                                    self.emit(Op::DUP);
                                    let keys_key = self.str_const("__keys");
                                    self.emit_u16(Op::STRUCT_GET, keys_key);
                                    self.emit_const(Value::String(Arc::from(key_name.as_str())));
                                    let l = self.line;
                                    common::collections::emit_push(
                                        &mut self.chunks,
                                        self.current,
                                        l,
                                    );
                                    self.emit(Op::DROP);
                                }
                            } else {
                                // Dynamic key — emit_set is
                                // `ecma:array.set(obj, key, value) → null`
                                // so we must push key BEFORE value. The
                                // previous impl pushed value then key,
                                // causing `{ 1: "one" }` to be stored
                                // under key "one" with value 1. Fix
                                // matches the canonical emit_set contract.
                                self.emit(Op::DUP); // [dict, dict]
                                self.compile_expr(key)?; // [dict, dict, key]
                                self.emit(Op::DUP); // [dict, dict, key, key]
                                let key_tmp = self.define_local("__obj_dyn_key");
                                self.emit_u16(Op::LOCAL_SET, key_tmp);
                                self.emit(Op::DROP);
                                // [dict, dict, key]
                                self.compile_expr(value)?; // [dict, dict, key, value]
                                let l = self.line;
                                common::collections::emit_set(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP); // drop returned null
                                // Track dynamic key in __keys (stringified)
                                self.emit(Op::DUP);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                let l = self.line;
                                common::collections::emit_push(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP);
                            }
                        }
                        ObjectProperty::Shorthand(name) => {
                            self.emit(Op::DUP);
                            self.emit_var_get(name);
                            let idx = self.str_const(name);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                            // Track key in __keys
                            self.emit(Op::DUP);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_const(Value::String(Arc::from(name.as_str())));
                            let l = self.line;
                            common::collections::emit_push(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Spread(expr) => {
                            // Object spread mutates the in-progress object and must
                            // leave that object on the stack for the remaining literal.
                            let target_tmp = self.define_local("__obj_spread_target");
                            self.emit_u16(Op::LOCAL_SET, target_tmp);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, target_tmp);
                            self.compile_expr(expr)?;
                            let idx = self.import("ecma:object", "assign");
                            self.emit_host_call(idx, 2);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, target_tmp);
                        }
                        ObjectProperty::Method { key, value } => {
                            self.emit(Op::DUP);
                            if let StmtKind::FunctionDecl {
                                params,
                                body,
                                is_generator,
                                is_async,
                                ..
                            } = &value.kind
                            {
                                if self.is_js_profile() {
                                    self.compile_lambda_with_flags(
                                        params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                    )?;
                                } else {
                                    // Object methods receive `this` as implicit first arg
                                    let mut method_params = vec![Param {
                                        name: self.profile.self_keyword.clone(),
                                        type_hint: None,
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false,
                                    }];
                                    method_params.extend(params.iter().cloned());
                                    self.compile_lambda_with_flags(
                                        &method_params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                    )?;
                                }
                            } else {
                                self.emit(Op::NULL);
                            }
                            // Set fn.name = key for Function.prototype.name support.
                            if self.is_js_profile() {
                                self.emit(Op::DUP);
                                self.emit_const(Value::String(Arc::from(key.as_str())));
                                let name_key = self.str_const("name");
                                self.emit_u16(Op::STRUCT_SET, name_key);
                                self.emit(Op::DROP);
                            }
                            let idx = self.str_const(key);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Accessor { kind, key, value } => {
                            self.emit(Op::DUP);
                            if let StmtKind::FunctionDecl {
                                params,
                                body,
                                is_generator,
                                is_async,
                                ..
                            } = &value.kind
                            {
                                if self.is_js_profile() {
                                    self.compile_lambda_with_flags(
                                        params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                    )?;
                                } else {
                                    // Accessors receive `this` as first arg
                                    let mut accessor_params = vec![Param {
                                        name: self.profile.self_keyword.clone(),
                                        type_hint: None,
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false,
                                    }];
                                    accessor_params.extend(params.iter().cloned());
                                    self.compile_lambda_with_flags(
                                        &accessor_params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                    )?;
                                }
                            } else {
                                self.emit(Op::NULL);
                            }
                            let accessor_name = match kind {
                                AccessorKind::Get => format!("get {}", key),
                                AccessorKind::Set => format!("set {}", key),
                            };
                            self.emit(Op::DUP);
                            self.emit_const(Value::String(Arc::from(accessor_name.as_str())));
                            let name_key = self.str_const("name");
                            self.emit_u16(Op::STRUCT_SET, name_key);
                            self.emit(Op::DROP);
                            let accessor_slot = match kind {
                                AccessorKind::Get => format!("__get_{}", key),
                                AccessorKind::Set => format!("__set_{}", key),
                            };
                            let idx = self.str_const(&accessor_slot);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Computed { key, value } => {
                            // ecma:array.set expects [obj, key, val] → null
                            self.emit(Op::DUP);
                            self.compile_expr(key)?;
                            self.emit(Op::DUP); // save key for trackKey
                            let key_tmp = self.define_local("__obj_comp_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp);
                            self.emit(Op::DROP);
                            self.compile_expr(value)?;
                            let l = self.line;
                            common::collections::emit_set(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP); // drop returned null
                            // Track key — host fn checks if it's a
                            // Symbol and routes to `__sym_keys` so
                            // Object.keys excludes it (ECMA-262 §7.3.22).
                            self.emit(Op::DUP);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            let track_idx = self.import("ecma:object", "trackKey");
                            let line = self.line;
                            self.chunk().emit_op_u16(Op::CALL_IMPORT, track_idx, line);
                            self.chunk().emit(2, line);
                            self.emit(Op::DROP);
                        }
                    }
                }
            }

            // ── String interpolation ────────────────────────────────────
            ExprKind::Interpolation(parts) => {
                if parts.is_empty() {
                    self.emit_const(Value::String(Arc::from("")));
                    return Ok(());
                }
                // For JS, ECMA-262 §13.2.8 specifies template literal
                // substitutions go through GetValue → ToString. ToString
                // on Objects calls ToPrimitive(hint=string) which prefers
                // toString over valueOf — exactly what the
                // `__vybe_to_primitive` polyfill implements. Routing
                // through it ensures user `toString` overrides fire
                // (sets `__js_this` via the JS method-call protocol),
                // then `"" + primitive` produces the final string.
                let use_to_primitive = self.is_js_profile();
                self.emit_const(Value::String(Arc::from("")));
                let acc_slot = self.define_local("__interp_acc");
                self.emit_u16(Op::LOCAL_SET, acc_slot);
                self.emit(Op::DROP);
                let part_slot = self.define_local("__interp_part");

                for part in parts.iter() {
                    match part {
                        InterpolPart::Text(s) => {
                            self.emit_const(Value::String(Arc::from(s.as_str())))
                        }
                        InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                            if use_to_primitive {
                                self.compile_expr(e)?;
                                self.emit_to_primitive("string");
                                // After ToPrimitive, the value is a
                                // primitive (string / number / etc).
                                // Concat with "" to coerce to string.
                                self.emit_const(Value::String(Arc::from("")));
                                let line = self.line;
                                common::strings::emit_str_concat(self.chunk(), line);
                            } else {
                                self.compile_expr(e)?;
                                let value_slot = self.define_local("__interp_value");
                                self.emit_u16(Op::LOCAL_SET, value_slot);
                                self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                let line = self.line;
                                common::strings::emit_to_string(self.chunk(), line);
                            }
                        }
                    }

                    self.emit_u16(Op::LOCAL_SET, part_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, acc_slot);
                    self.emit_u16(Op::LOCAL_GET, part_slot);
                    let line = self.line;
                    common::strings::emit_str_concat(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, acc_slot);
                    self.emit(Op::DROP);
                }

                self.emit_u16(Op::LOCAL_GET, acc_slot);
            }

            // ── Type operations ─────────────────────────────────────────
            ExprKind::IsType {
                expr: inner,
                type_name,
            } => {
                let canon_type = self.canon(type_name);
                if matches!(
                    canon_type.as_str(),
                    "IEnumerable"
                        | "ICollection"
                        | "IList"
                        | "IReadOnlyCollection"
                        | "IReadOnlyList"
                ) {
                    if let ExprKind::Ident(name) = &inner.kind {
                        if self.lookup_var_type_hint(name).is_some_and(|hint| {
                            let bare = hint.split('<').next().unwrap_or(hint).trim();
                            matches!(
                                Self::normalize_type_hint(bare).as_str(),
                                "list" | "arraylist" | "queue" | "stack" | "hashset" | "dictionary"
                            )
                        }) {
                            self.emit(Op::TRUE);
                            return Ok(());
                        }
                    }
                }
                self.compile_expr(inner)?;
                let obj_slot = self.define_local("__is_type_obj");
                self.emit_u16(Op::LOCAL_SET, obj_slot);
                self.emit(Op::DROP);

                // `IsType` always produces `Value::Bool` — the WASM `REF_IS_*`
                // opcodes return `i32`, so wrap with `emit_i32_to_bool` here.
                // This is language-agnostic: JS normalizes `typeof x === "string"`
                // and `Array.isArray(x)` to `IsType` in the walker; Python uses
                // `isinstance`; VB uses `TypeOf x Is T`. All share the same
                // ECMA-compatible Bool output.
                {
                    let line = self.line;
                    match canon_type.as_str() {
                        "string" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_STRING);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "number" | "float" | "double" | "int" | "integer" | "long" | "single"
                        | "decimal" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_NUMBER);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "boolean" | "bool" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_BOOL);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "array" | "list" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_ARRAY);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "function" | "callable" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_FUNC);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "undefined" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_UNDEFINED);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "object" | "dict" | "map" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_OBJECT);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        _ => {} // fall through to the general reflection path
                    }
                }

                // VB "integer"/"int" check: IsNumber AND value is integral
                if self.profile.name == "vb" && matches!(canon_type.as_str(), "integer" | "int") {
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_NUMBER);
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::DUP);
                    self.emit(Op::F64_TRUNC);
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                    crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                    self.chunk().emit_else(line);
                    self.emit(Op::FALSE);
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                let line = self.line;
                let type_idx =
                    self.chunk()
                        .add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(
                            canon_type.as_str(),
                        )));
                let matched_slot = self.define_local("__type_test_matched");
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_slot);
                self.chunk()
                    .emit_op_u16(vybe_bytecode::Op::REF_TEST, type_idx, line);
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_GET, obj_slot);
                let type_key = self.str_const("__type");
                self.emit_u16(Op::STRUCT_GET, type_key);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                };
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);

                let reflection_matches: Vec<String> = self
                    .reflection_types
                    .keys()
                    .filter(|candidate| {
                        !candidate.eq_ignore_ascii_case(&canon_type)
                            && self.reflection_is_assignable_from(type_name, candidate)
                    })
                    .cloned()
                    .collect();
                for candidate in &reflection_matches {
                    let candidate_idx = self.chunk().add_constant(vybe_bytecode::Value::String(
                        std::sync::Arc::from(candidate.as_str()),
                    ));
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.chunk()
                        .emit_op_u16(vybe_bytecode::Op::REF_TEST, candidate_idx, line);
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.emit(Op::DROP);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let type_key = self.str_const("__type");
                    self.emit_u16(Op::STRUCT_GET, type_key);
                    self.emit_const(Value::String(Arc::from(candidate.as_str())));
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.emit(Op::DROP);
                    self.chunk().emit_end(line);
                }

                if matches!(
                    canon_type.as_str(),
                    "IEnumerable"
                        | "ICollection"
                        | "IList"
                        | "IReadOnlyCollection"
                        | "IReadOnlyList"
                ) {
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_ARRAY);
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.emit(Op::DROP);
                    self.chunk().emit_end(line);

                    let list_idx = self
                        .chunk()
                        .add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("list")));
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.chunk()
                        .emit_op_u16(vybe_bytecode::Op::REF_TEST, list_idx, line);
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.emit(Op::DROP);
                    self.chunk().emit_end(line);

                    for key_name in ["length", "count"] {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        let key = self.str_const(key_name);
                        self.emit_u16(Op::STRUCT_GET, key);
                        self.emit(Op::REF_IS_NULL);
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);
                    }

                    if canon_type == "IEnumerable" {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_OBJECT);
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(line);
                    }
                }

                self.emit_u16(Op::LOCAL_GET, obj_slot);
                let types_key = self.str_const("__types");
                self.emit_u16(Op::STRUCT_GET, types_key);
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                self.emit(Op::I32_EQZ);
                self.chunk().emit_if(line);
                let types_slot = self.define_local("__type_test_types");
                self.emit_u16(Op::LOCAL_SET, types_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, types_slot);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                common::collections::emit_contains(&mut self.chunks, self.current, line);
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);
                for candidate in &reflection_matches {
                    self.emit_u16(Op::LOCAL_GET, types_slot);
                    self.emit_const(Value::String(Arc::from(candidate.as_str())));
                    common::collections::emit_contains(&mut self.chunks, self.current, line);
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.emit(Op::DROP);
                    self.chunk().emit_end(line);
                }
                self.chunk().emit_else(line);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);
                // matched_slot holds I32(0) or I32(1); convert to Bool for Go type assertions.
                self.emit_u16(Op::LOCAL_GET, matched_slot);
                self.chunk().emit_if_value(line);
                self.emit(Op::TRUE);
                self.chunk().emit_else(line);
                self.emit(Op::FALSE);
                self.chunk().emit_end(line);
            }

            ExprKind::Cast { expr: inner, .. } => {
                // Pascal allows user functions to shadow builtin type names
                // (`function Double(x: Integer)`). When the parser produced a
                // builtin-style cast node for such a name, honour the user
                // function instead of treating the cast as a no-op.
                if let ExprKind::Cast { type_name, .. } = &expr.kind {
                    let canon_type = self.canon(type_name);
                    if self.profile.name == "csharp" {
                        if let Some(names) = self.enum_value_names.get(&canon_type) {
                            if let ExprKind::Lit(Literal::Int(n)) = &inner.kind {
                                if let Some(member_name) = names.get(n) {
                                    self.emit_const(Value::String(Arc::from(member_name.as_str())));
                                    return Ok(());
                                }
                            }
                        }
                    }

                    if self.profile.name == "csharp" && canon_type == "char" {
                        self.compile_expr(inner)?;
                        let num = self.import("ecma:number", "Number");
                        self.emit_host_call(num, 1);
                        self.emit(Op::F64_FLOOR);
                        self.emit(Op::STR_FROM_CHAR_CODE);
                        return Ok(());
                    }

                    if self.profile.name == "csharp" {
                        match canon_type.as_str() {
                            "int" | "long" | "short" | "byte" | "uint" | "ulong" | "ushort"
                            | "sbyte" => {
                                let is_char_like =
                                    matches!(&inner.kind, ExprKind::Lit(Literal::Char(_)))
                                        || self.infer_expr_type_hint(inner).is_some_and(|hint| {
                                            Self::normalize_type_hint(&hint) == "char"
                                        });
                                if is_char_like {
                                    self.emit(Op::I32_CONST_0);
                                    let line = self.line;
                                    common::strings::emit_to_string(self.chunk(), line);
                                    return Ok(());
                                }
                                self.compile_expr(inner)?;
                                let num = self.import("ecma:number", "Number");
                                self.emit_host_call(num, 1);
                                self.emit(Op::F64_TRUNC);
                                return Ok(());
                            }
                            _ => {}
                        }
                    }

                    if self.profile.name == "c" {
                        match canon_type.as_str() {
                            "double" | "float" => {
                                self.compile_expr(inner)?;
                                let num = self.import("ecma:number", "Number");
                                self.emit_host_call(num, 1);
                                return Ok(());
                            }
                            "char" | "uint8" | "int16" | "int" | "long" | "uint32" => {
                                self.compile_expr(inner)?;
                                let num = self.import("ecma:number", "Number");
                                self.emit_host_call(num, 1);
                                self.emit(Op::F64_TRUNC);
                                match canon_type.as_str() {
                                    "uint8" | "char" => {
                                        self.emit_c_unsigned_wrap(256.0);
                                    }
                                    "int16" => {
                                        self.emit_c_unsigned_wrap(65_536.0);
                                        self.emit_c_signed_wrap_from_unsigned(32_768.0, 65_536.0);
                                    }
                                    "uint32" => {
                                        self.emit_c_unsigned_wrap(4_294_967_296.0);
                                    }
                                    "int" => {
                                        self.emit_c_unsigned_wrap(4_294_967_296.0);
                                        self.emit_c_signed_wrap_from_unsigned(
                                            2_147_483_648.0,
                                            4_294_967_296.0,
                                        );
                                    }
                                    _ => {}
                                }
                                return Ok(());
                            }
                            _ => {}
                        }
                    }

                    if self.profile.name == "pascal" {
                        match self.canon(type_name).as_str() {
                            "integer" | "int" | "longint" => {
                                self.compile_expr(inner)?;
                                let line = self.line;
                                common::math::emit_trunc(self.chunk(), line);
                                return Ok(());
                            }
                            _ => {}
                        }
                    }

                    if self.profile.name == "vb" {
                        if let Some((cast_kind, target_type)) = type_name.split_once(':') {
                            if cast_kind.eq_ignore_ascii_case("TryCast") {
                                let resolved_target = self.resolve_source_type_alias(target_type);
                                let trimmed_target = resolved_target
                                    .trim()
                                    .trim_end_matches('?')
                                    .trim()
                                    .to_string();
                                if self.vb_is_reference_type_hint(&trimmed_target) {
                                    let line = self.line;
                                    self.compile_expr(inner)?;
                                    let value_slot = self.define_local("__vb_trycast_value");
                                    self.emit_u16(Op::LOCAL_SET, value_slot);
                                    self.emit(Op::DROP);
                                    let result_slot = self.define_local("__vb_trycast_result");
                                    self.emit(Op::NULL);
                                    self.emit_u16(Op::LOCAL_SET, result_slot);
                                    self.emit(Op::DROP);

                                    self.emit_u16(Op::LOCAL_GET, value_slot);
                                    self.emit(Op::REF_IS_NULL);
                                    self.emit(Op::I32_EQZ);
                                    let non_null_line = self.line;
                                    self.chunk().emit_if(non_null_line);

                                    if Self::is_string_type_hint(&trimmed_target) {
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit(Op::REF_IS_STRING);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.emit(Op::DROP);
                                        self.chunk().emit_end(line);
                                    } else if trimmed_target.ends_with("()") {
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit(Op::REF_IS_ARRAY);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.emit(Op::DROP);
                                        self.chunk().emit_end(line);
                                    } else if self.vb_is_object_type_hint(&trimmed_target) {
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit(Op::REF_IS_STRING);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.emit(Op::DROP);
                                        self.chunk().emit_end(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit(Op::REF_IS_ARRAY);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.emit(Op::DROP);
                                        self.chunk().emit_end(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit(Op::REF_IS_OBJECT);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.emit(Op::DROP);
                                        self.chunk().emit_end(line);
                                    } else {
                                        let mut expected_names = vec![trimmed_target.clone()];
                                        if let Some(class_name) = self
                                            .resolve_pending_class_name_for_type_hint(
                                                &trimmed_target,
                                            )
                                        {
                                            if !expected_names.iter().any(|candidate| {
                                                candidate.eq_ignore_ascii_case(&class_name)
                                            }) {
                                                expected_names.push(class_name.clone());
                                            }
                                            let short_name = class_name
                                                .rsplit('.')
                                                .next()
                                                .unwrap_or(class_name.as_str())
                                                .to_string();
                                            if !expected_names.iter().any(|candidate| {
                                                candidate.eq_ignore_ascii_case(&short_name)
                                            }) {
                                                expected_names.push(short_name);
                                            }
                                        } else if let Some(short_name) =
                                            trimmed_target.rsplit('.').next()
                                        {
                                            if !expected_names.iter().any(|candidate| {
                                                candidate.eq_ignore_ascii_case(short_name)
                                            }) {
                                                expected_names.push(short_name.to_string());
                                            }
                                        }

                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit(Op::REF_IS_OBJECT);
                                        let object_line = self.line;
                                        self.chunk().emit_if(object_line);
                                        let match_slot = self.define_local("__vb_trycast_match");
                                        self.emit_const(Value::I32(0));
                                        self.emit_u16(Op::LOCAL_SET, match_slot);
                                        self.emit(Op::DROP);
                                        for expected in &expected_names {
                                            self.emit_u16(Op::LOCAL_GET, value_slot);
                                            let type_key = self.str_const("__type");
                                            self.emit_u16(Op::STRUCT_GET, type_key);
                                            self.emit_const(Value::String(Arc::from(
                                                expected.as_str(),
                                            )));
                                            {
                                                let line = self.line;
                                                crate::emitter::ops::emit_dyn_eq(
                                                    self.chunk(),
                                                    line,
                                                );
                                            };
                                            crate::emitter::ops::emit_dyn_to_bool(
                                                self.chunk(),
                                                object_line,
                                            );
                                            self.chunk().emit_if(object_line);
                                            self.emit_const(Value::I32(1));
                                            self.emit_u16(Op::LOCAL_SET, match_slot);
                                            self.emit(Op::DROP);
                                            self.chunk().emit_end(object_line);
                                        }
                                        for expected in &expected_names {
                                            self.emit_u16(Op::LOCAL_GET, value_slot);
                                            let types_key = self.str_const("__types");
                                            self.emit_u16(Op::STRUCT_GET, types_key);
                                            self.emit(Op::DUP);
                                            self.emit(Op::REF_IS_NULL);
                                            self.emit(Op::I32_EQZ);
                                            self.chunk().emit_if(object_line);
                                            let types_slot =
                                                self.define_local("__vb_trycast_types");
                                            self.emit_u16(Op::LOCAL_SET, types_slot);
                                            self.emit(Op::DROP);
                                            self.emit_u16(Op::LOCAL_GET, types_slot);
                                            self.emit_const(Value::String(Arc::from(
                                                expected.as_str(),
                                            )));
                                            common::collections::emit_contains(
                                                &mut self.chunks,
                                                self.current,
                                                line,
                                            );
                                            crate::emitter::ops::emit_dyn_to_bool(
                                                self.chunk(),
                                                object_line,
                                            );
                                            self.chunk().emit_if(object_line);
                                            self.emit_const(Value::I32(1));
                                            self.emit_u16(Op::LOCAL_SET, match_slot);
                                            self.emit(Op::DROP);
                                            self.chunk().emit_end(object_line);
                                            self.chunk().emit_else(object_line);
                                            self.emit(Op::DROP);
                                            self.chunk().emit_end(object_line);
                                        }
                                        self.emit_u16(Op::LOCAL_GET, match_slot);
                                        self.chunk().emit_if(object_line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.emit(Op::DROP);
                                        self.chunk().emit_end(object_line);
                                        self.chunk().emit_end(object_line);
                                    }

                                    self.chunk().emit_end(non_null_line);
                                    self.emit_u16(Op::LOCAL_GET, result_slot);
                                    return Ok(());
                                }
                            }
                        }
                    }

                    if let Some(user_type) = self.user_value_type_name_from_hint(type_name) {
                        if matches!(&inner.kind, ExprKind::Object(_)) {
                            let ctor_global = {
                                let overload = format!("{}$arity0", user_type);
                                if self.defined_globals.contains(&overload) {
                                    overload
                                } else {
                                    user_type.clone()
                                }
                            };

                            let source_slot = self.define_local("__cast_struct_source");
                            self.compile_expr(inner)?;
                            self.emit_u16(Op::LOCAL_SET, source_slot);
                            self.emit(Op::DROP);

                            let value_slot = self.define_local("__cast_struct_value");
                            let idx = self.str_const(&ctor_global);
                            self.emit_u16(Op::GLOBAL_GET, idx);
                            self.emit_u8(Op::CALL_REF, 0);
                            self.emit_u16(Op::LOCAL_SET, value_slot);
                            self.emit(Op::DROP);

                            if let Some(fields) = self
                                .pending_classes
                                .get(&user_type)
                                .map(|pending| pending.fields.clone())
                            {
                                for field_name in fields {
                                    let member_slot = self.define_local("__cast_struct_member");
                                    let field_idx = self.str_const(&field_name);
                                    self.emit_u16(Op::LOCAL_GET, source_slot);
                                    self.emit_u16(Op::STRUCT_GET, field_idx);
                                    self.emit_u16(Op::LOCAL_SET, member_slot);
                                    self.emit(Op::DROP);

                                    self.emit_u16(Op::LOCAL_GET, member_slot);
                                    self.emit(Op::REF_IS_UNDEFINED);
                                    self.emit(Op::I32_EQZ);
                                    let set_line = self.line;
                                    self.chunk().emit_if(set_line);
                                    self.emit_u16(Op::LOCAL_GET, value_slot);
                                    self.emit_u16(Op::LOCAL_GET, member_slot);
                                    self.emit_u16(Op::STRUCT_SET, field_idx);
                                    self.emit(Op::DROP);
                                    self.chunk().emit_end(set_line);
                                }
                            }

                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            return Ok(());
                        }
                    }

                    let shadows_cast = self.defined_functions.contains(&canon_type)
                        || (!self.case_sensitive
                            && self
                                .defined_functions
                                .iter()
                                .any(|name| name.eq_ignore_ascii_case(type_name)));
                    if shadows_cast {
                        self.emit_var_get(type_name);
                        self.compile_expr(inner)?;
                        self.emit_u8(Op::CALL_REF, 1);
                        return Ok(());
                    }
                }

                // Cast is otherwise a no-op in our dynamic VM.
                self.compile_expr(inner)?;
            }

            ExprKind::DefaultOf(type_name) => {
                let normalized = Self::normalize_type_hint(type_name);
                match normalized.as_str() {
                    "int" | "long" | "short" | "byte" | "uint" | "ulong" | "ushort" | "sbyte"
                    | "double" | "float" | "decimal" => {
                        self.emit_const(Value::I32(0));
                    }
                    "bool" => {
                        self.emit_const(Value::Bool(false));
                    }
                    "char" => {
                        self.emit_const(Value::String(Arc::from("\0")));
                    }
                    _ => {
                        self.emit(Op::NULL);
                    }
                }
            }

            ExprKind::TypeOf(inner) => {
                if self.is_js_profile() {
                    if let ExprKind::Ident(name) = &inner.kind {
                        let is_callable_global = !self.has_accessible_local_binding(name)
                            && (matches!(
                                name.as_str(),
                                "eval"
                                    | "parseInt"
                                    | "parseFloat"
                                    | "Function"
                                    | "Object"
                                    | "Boolean"
                                    | "Number"
                                    | "String"
                                    | "Array"
                                    | "Symbol"
                                    | "BigInt"
                                    | "Date"
                                    | "RegExp"
                                    | "Promise"
                                    | "Proxy"
                                    | "Map"
                                    | "Set"
                                    | "WeakMap"
                                    | "WeakSet"
                                    | "ArrayBuffer"
                                    | "SharedArrayBuffer"
                                    | "DataView"
                                    | "Int8Array"
                                    | "Uint8Array"
                                    | "Uint8ClampedArray"
                                    | "Int16Array"
                                    | "Uint16Array"
                                    | "Int32Array"
                                    | "Uint32Array"
                                    | "Float32Array"
                                    | "Float64Array"
                                    | "BigInt64Array"
                                    | "BigUint64Array"
                            ) || self.profile.lookup_builtin(name).is_some());
                        if is_callable_global {
                            self.emit_const(Value::String(Arc::from("function")));
                            return Ok(());
                        }
                    }
                }
                self.compile_expr(inner)?;
                if self.is_js_profile() {
                    // ECMA-262 §13.5.3 Table 41: arrays are "object",
                    // not "array". The VM's REF_TYPEOF emits "array"
                    // (Vybe-specific), so JS routes through the host
                    // helper that returns spec-compliant tags.
                    let idx = self.import("ecma:value", "typeof");
                    self.emit_host_call(idx, 1);
                } else {
                    self.emit(Op::REF_TYPEOF);
                }
            }

            // ── NullCoalesce ────────────────────────────────────────────
            ExprKind::NullCoalesce { left, right } => {
                self.compile_expr(left)?;
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.emit(Op::DROP);
                self.compile_expr(right)?;
                self.chunk().emit_else(line);
                self.chunk().emit_end(line);
            }

            // ── Spread ──────────────────────────────────────────────────
            ExprKind::Spread(inner) => {
                self.compile_expr(inner)?;
                // SPREAD is array-only in the VM (matches WASM
                // `array.copy_into` semantics). Iterables that aren't
                // arrays — Set, Map, String — get coerced to an array
                // first via the polymorphic Symbol.iterator helper
                // (`ecma:object.iterForOf`). Generators (Continuation
                // values) need spec WASM stack-switching `resume`
                // (emit_next) to drive their iterator protocol, which a
                // host fn can't do — route them through the
                // `__stdlib_drain_generator` bytecode helper.
                let inner_slot = self.define_local("__spread_iter");
                self.emit_u16(Op::LOCAL_SET, inner_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, inner_slot);
                let is_gen_idx = self.import("ecma:value", "isGenerator");
                self.emit_host_call(is_gen_idx, 1);
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                let drain_key = self.str_const("__vybe_drain_generator");
                self.emit_u16(Op::GLOBAL_GET, drain_key);
                self.emit_u16(Op::LOCAL_GET, inner_slot);
                self.emit_u8(Op::CALL_REF, 1);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, inner_slot);
                let idx = self.import("ecma:object", "iterForOf");
                self.emit_host_call(idx, 1);
                self.chunk().emit_end(line);
                self.emit(Op::SPREAD);
            }

            // ── Await ───────────────────────────────────────────────────
            ExprKind::Await(inner) => {
                // ECMA-262 §27.2: WASM JSPI suspend point, lowered to the spec
                // stack-switching `suspend` (AWAIT_SUSPEND_TAG). The VM unwraps
                // fulfilled, throws rejected, suspends the fiber on pending, and
                // passes non-promise values through unchanged.
                self.compile_expr(inner)?;
                let line = self.line;
                crate::emitter::functions::emit_await(self.chunk(), line);
            }

            // ── Yield ───────────────────────────────────────────────────
            ExprKind::Yield(val) => {
                if let Some(v) = val {
                    if let Some((key_expr, value_expr)) = self.generator_keyed_yield_parts(v) {
                        self.compile_expr(key_expr)?;
                        let key_slot = self.define_local("__yield_key");
                        self.emit_u16(Op::LOCAL_SET, key_slot);
                        self.emit(Op::DROP);

                        self.compile_expr(value_expr)?;
                        let payload_value_slot = self.define_local("__yield_payload_value");
                        self.emit_u16(Op::LOCAL_SET, payload_value_slot);
                        self.emit(Op::DROP);

                        self.emit_next_generator_payload_id();
                        let payload_id_slot = self.define_local("__yield_payload_id");
                        self.emit_u16(Op::LOCAL_SET, payload_id_slot);
                        self.emit(Op::DROP);

                        self.emit_generator_payload_store();
                        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
                        self.emit_u16(Op::LOCAL_GET, payload_value_slot);
                        let line = self.line;
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        common::dict::emit_new(&mut self.chunks, self.current, line);
                        self.emit(Op::DUP);
                        self.emit_const(Value::Bool(true));
                        let marker_key = self.str_const("__vybe_generator_yield");
                        self.emit_u16(Op::STRUCT_SET, marker_key);
                        self.emit(Op::DROP);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::LOCAL_GET, key_slot);
                        let key_key = self.str_const("key");
                        self.emit_u16(Op::STRUCT_SET, key_key);
                        self.emit(Op::DROP);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
                        let payload_id_key = self.str_const("payload_id");
                        self.emit_u16(Op::STRUCT_SET, payload_id_key);
                        self.emit(Op::DROP);
                    } else {
                        self.compile_expr(v)?;
                    }
                } else {
                    self.emit(Op::NULL);
                }
                let line = self.line;
                crate::emitter::generators::emit_suspend(self.chunk(), line);
                self.emit_generator_resume_value()?;
            }

            ExprKind::YieldFrom(inner) => {
                // ECMA-262 §15.5 `yield*`: drain the inner iterable,
                // re-yielding each value through the enclosing
                // generator. Continuations use WASM stack-switching
                // `GEN_NEXT`; plain iterables are first drained through
                // the shared collection surface and yielded value-by-value.
                self.compile_expr(inner)?;
                let gen_slot = self.define_local("__yield_star_gen");
                let val_slot = self.define_local("__yield_star_val");
                let has_more_slot = self.define_local("__yield_star_has_more");
                let result_slot = self.define_local("__yield_star_result");
                self.emit_u16(Op::LOCAL_SET, gen_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, gen_slot);
                let is_gen_idx = self.import("ecma:value", "isGenerator");
                self.emit_host_call(is_gen_idx, 1);
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                let gen_block = self.chunk().emit_block(line);
                let (gen_loop, _) = self.chunk().emit_loop_s(line);
                self.emit_u16(Op::LOCAL_GET, gen_slot);
                let line = self.line;
                crate::emitter::generators::emit_next(self.chunk(), line);
                // After GEN_NEXT: stack top is has_more (i32), under it value.
                self.emit_u16(Op::LOCAL_SET, has_more_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, val_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_EQZ);
                self.chunk().emit_br_if(1, line);
                self.emit_u16(Op::LOCAL_GET, val_slot);
                let line = self.line;
                crate::emitter::generators::emit_suspend(self.chunk(), line);
                self.emit(Op::DROP);
                self.chunk().emit_br(0, line);
                self.chunk().emit_end(line);
                self.chunk().patch_loop(gen_loop);
                self.chunk().emit_end(line);
                self.chunk().patch_block(gen_block);
                // `yield from` / `yield*` evaluates to the delegated
                // generator's completion value. `GEN_NEXT` leaves that
                // final value in `val_slot` on the terminating step.
                self.emit_u16(Op::LOCAL_GET, val_slot);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, gen_slot);
                if self.is_js_profile() {
                    let iter_drain_key = self.str_const("__vybe_iter_drain");
                    self.emit_u16(Op::GLOBAL_GET, iter_drain_key);
                    self.emit_u16(Op::LOCAL_GET, gen_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                } else {
                    common::collections::emit_iter_values(
                        &mut self.chunks,
                        self.current,
                        self.line,
                    );
                }
                let iter_slot = self.define_local("__yield_star_iter");
                let idx_slot = self.define_local("__yield_star_idx");
                let len_slot = self.define_local("__yield_star_len");
                self.emit_u16(Op::LOCAL_SET, iter_slot);
                self.emit(Op::DROP);

                self.emit_const(Value::F64(0.0));
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, iter_slot);
                let line = self.line;
                common::collections::emit_array_length(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, len_slot);
                self.emit(Op::DROP);

                let iter_block = self.chunk().emit_block(line);
                let (iter_loop, _) = self.chunk().emit_loop_s(line);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, len_slot);
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_lt(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                };
                self.chunk().emit_br_if(1, line);
                self.emit_u16(Op::LOCAL_GET, iter_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::ARRAY_GET);
                let line = self.line;
                crate::emitter::generators::emit_suspend(self.chunk(), line);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_const(Value::F64(1.0));
                self.emit(Op::F64_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);
                self.chunk().emit_br(0, line);
                self.chunk().emit_end(line);
                self.chunk().patch_loop(iter_loop);
                self.chunk().emit_end(line);
                self.chunk().patch_block(iter_block);
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_GET, result_slot);
            }

            // ── AddressOf (VB) ──────────────────────────────────────────
            ExprKind::AddressOf(name) => {
                let parts: Vec<&str> = name.split('.').filter(|part| !part.is_empty()).collect();
                if parts.is_empty() {
                    self.emit(Op::NULL);
                    return Ok(());
                }

                let self_kw = self.profile.self_keyword.clone();
                let is_self_qualified = parts
                    .first()
                    .map(|part| {
                        if self.case_sensitive {
                            *part == self_kw || *part == "Me"
                        } else {
                            part.eq_ignore_ascii_case(&self_kw) || part.eq_ignore_ascii_case("Me")
                        }
                    })
                    .unwrap_or(false);

                let bound_parts: Option<&[&str]> = if self.current_class.is_some() {
                    if is_self_qualified && parts.len() > 1 {
                        Some(&parts[1..])
                    } else if parts.len() == 1
                        && self.defined_class_methods.contains(&self.canon(parts[0]))
                    {
                        Some(&parts[..])
                    } else {
                        None
                    }
                } else {
                    None
                };

                if let Some(method_parts) = bound_parts {
                    if self.emit_self_ref() {
                        for part in method_parts {
                            let idx = self.str_const(&self.canon(part));
                            self.emit_u16(Op::STRUCT_GET, idx);
                        }
                        return Ok(());
                    }
                }

                self.emit_var_get(parts[0]);
                for part in &parts[1..] {
                    let idx = self.str_const(&self.canon(part));
                    self.emit_u16(Op::STRUCT_GET, idx);
                }
            }

            // ── SuperCall (VB/Python) ───────────────────────────────────
            ExprKind::SuperCall { method, args } => {
                let self_kw = self.profile.self_keyword.clone();
                let ctor_name = self.profile.constructor_name.clone();
                let is_ctor_call = method.is_none()
                    || method.as_ref().map_or(false, |m| {
                        if self.case_sensitive {
                            m == &ctor_name || m == "new" || m == "__init__"
                        } else {
                            m.eq_ignore_ascii_case(&ctor_name)
                                || m.eq_ignore_ascii_case("new")
                                || m.eq_ignore_ascii_case("__init__")
                        }
                    });

                if is_ctor_call {
                    // super() / MyBase.New(args) → call parent constructor
                    if let Some(ref class_name) = self.current_class.clone() {
                        if let Some(parent_name) = self
                            .pending_classes
                            .get(class_name.as_str())
                            .and_then(|c| c.parent.clone())
                        {
                            if common::errors::is_exception_type(&parent_name) {
                                let arg_exprs: Vec<&Expression> =
                                    args.iter().map(|arg| &arg.value).collect();
                                self.emit_js_exception_ctor_value(&parent_name, &arg_exprs)?;
                                if let Some(slot) = self
                                    .scope()
                                    .resolve(&self_kw)
                                    .or_else(|| self.scope().resolve_ci(&self_kw))
                                {
                                    self.emit(Op::DUP);
                                    self.emit_u16(Op::LOCAL_SET, slot);
                                    self.emit(Op::DROP);
                                }
                                return Ok(());
                            }
                            self.emit_var_get(&parent_name);
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            if let Some(slot) = self
                                .scope()
                                .resolve(&self_kw)
                                .or_else(|| self.scope().resolve_ci(&self_kw))
                            {
                                self.emit(Op::DUP);
                                self.emit_u16(Op::LOCAL_SET, slot);
                                self.emit(Op::DROP);
                            }
                        } else {
                            self.emit(Op::NULL);
                        }
                    } else {
                        self.emit(Op::NULL);
                    }
                } else if let Some(mname) = method {
                    let parent_name = self
                        .current_class
                        .as_ref()
                        .and_then(|class_name| self.pending_classes.get(class_name.as_str()))
                        .and_then(|pc| pc.parent.clone());
                    if let Some(parent_name) = parent_name {
                        let parent_canon = self.canon(&parent_name);
                        let method_name = self.canon(mname);
                        let method_idx = self.str_const(&method_name);
                        self.emit_var_get(&parent_canon);
                        self.emit_u16(Op::STRUCT_GET, method_idx);

                        if self.is_js_profile() {
                            let saved_js_this = self.save_js_this("__js_prev_this_super_expr");
                            if let Some(self_slot) = self
                                .scope()
                                .resolve(&self_kw)
                                .or_else(|| self.scope().resolve_ci(&self_kw))
                            {
                                self.emit_u16(Op::LOCAL_GET, self_slot);
                            } else {
                                let js_this = self.str_const("__js_this");
                                self.emit_u16(Op::GLOBAL_GET, js_this);
                            }
                            self.set_js_this_from_stack();
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            let result_slot = self.define_local("__js_super_expr_result");
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.emit(Op::DROP);
                            self.restore_js_this(saved_js_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else if let Some(self_slot) = self
                            .scope()
                            .resolve(&self_kw)
                            .or_else(|| self.scope().resolve_ci(&self_kw))
                        {
                            self.emit_u16(Op::LOCAL_GET, self_slot);
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, (args.len() + 1) as u8);
                        } else {
                            self.emit(Op::NULL);
                        }
                    } else {
                        self.emit(Op::NULL);
                    }
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── Comprehension (Python) ──────────────────────────────────
            ExprKind::Comprehension {
                kind,
                element,
                generators,
            } => {
                use crate::ast::ComprehensionKind;
                let line = self.line;
                let is_dict = *kind == ComprehensionKind::Dict;
                let is_set = *kind == ComprehensionKind::Set;

                // Build the accumulator: dict → Object, set/list/gen → Array
                if is_dict {
                    common::dict::emit_new(&mut self.chunks, self.current, line);
                } else {
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                }
                let result_slot = self.define_local("__comp_result");
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                use crate::emitter::loops::LoopState;
                // Compile each generator (nested for-clauses)
                let mut cond_blocks = 0usize;
                let mut loop_info: Vec<(u16, LoopState)> = Vec::new();
                for generator in generators.iter() {
                    self.compile_expr(&generator.iter)?;
                    let arr_slot = self.define_local("__comp_iter");
                    self.emit_u16(Op::LOCAL_SET, arr_slot);
                    self.emit(Op::DROP);
                    let idx_slot = self.define_local("__comp_idx");
                    let lp = common::loops::emit_for_in_start(
                        &mut self.chunks,
                        self.current,
                        arr_slot,
                        idx_slot,
                        line,
                    );
                    let var_name = match &generator.target.kind {
                        ExprKind::Ident(n) => n.clone(),
                        _ => "__comp_var".to_string(),
                    };
                    let var_slot = self.define_local(&var_name);
                    self.emit_u16(Op::LOCAL_SET, var_slot);
                    self.emit(Op::DROP);

                    for cond_expr in &generator.conditions {
                        self.compile_expr(cond_expr)?;
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if(line);
                        cond_blocks += 1;
                    }
                    loop_info.push((idx_slot, lp));
                }

                // Emit the body
                if is_dict {
                    // element is Array([key, val]); store into dict and track __keys
                    if let ExprKind::Array(kv) = &element.kind {
                        if kv.len() == 2 {
                            let key_expr = &kv[0].value;
                            let val_expr = &kv[1].value;
                            // Compile key and save a copy for __keys tracking
                            self.compile_expr(key_expr)?;
                            let key_slot = self.define_local("__comp_key");
                            self.emit_u16(Op::LOCAL_SET, key_slot);
                            self.emit(Op::DROP);
                            // [dict, key, val] → ARRAY_SET → drops from stack
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            self.emit_u16(Op::LOCAL_GET, key_slot);
                            self.compile_expr(val_expr)?;
                            let l = self.line;
                            common::collections::emit_set(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP);
                            // Track key in __keys so len() works
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_u16(Op::LOCAL_GET, key_slot);
                            let l = self.line;
                            common::collections::emit_push(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP);
                        }
                    }
                } else {
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.compile_expr(element)?;
                    let l = self.line;
                    common::collections::emit_push(&mut self.chunks, self.current, l);
                    self.emit(Op::DROP);
                }

                for _ in 0..cond_blocks {
                    self.chunk().emit_end(line);
                }
                // Close loops in reverse order
                for (idx_slot, lp) in loop_info.into_iter().rev() {
                    common::loops::emit_for_in_end(
                        &mut self.chunks,
                        self.current,
                        idx_slot,
                        lp,
                        line,
                    );
                }

                self.emit_u16(Op::LOCAL_GET, result_slot);

                // Set comprehension: convert Array → Set via ecma:set.fromIterable
                if is_set {
                    let idx = self.import("ecma:set", "fromIterable");
                    self.emit_host_call(idx, 1);
                }
            }

            // ── Slice (Python) ──────────────────────────────────────────
            ExprKind::Slice { lower, upper, step } => {
                // Stack on entry (from Index parent): [obj]
                let line = self.line;
                if step.is_none() {
                    let obj_slot = self.define_local("__py_slice_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    if let Some(l) = lower {
                        self.compile_expr(l)?;
                    } else {
                        self.emit(Op::I32_CONST_0);
                    }
                    if let Some(u) = upper {
                        self.compile_expr(u)?;
                    } else {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        common::collections::emit_len(&mut self.chunks, self.current, line);
                    }
                    common::collections::emit_runtime_helper_call(
                        &mut self.chunks,
                        self.current,
                        "__vybe_slice",
                        3,
                        line,
                    );
                } else {
                    // Emit slice parts → [obj, lower, upper, step] then call the
                    // bundled `__vybe_slicestep` polyfill via GLOBAL_GET + CALL_REF.
                    if let Some(l) = lower {
                        self.compile_expr(l)?;
                    } else {
                        self.emit(Op::NULL);
                    }
                    if let Some(u) = upper {
                        self.compile_expr(u)?;
                    } else {
                        self.emit(Op::NULL);
                    }
                    if let Some(s) = step {
                        self.compile_expr(s)?;
                    } else {
                        self.emit(Op::NULL);
                    }
                    common::collections::emit_runtime_helper_call(
                        &mut self.chunks,
                        self.current,
                        "__vybe_slicestep",
                        4,
                        line,
                    );
                }
            }

            // ── Walrus (Python :=) ──────────────────────────────────────
            ExprKind::Walrus { target, value } => {
                self.compile_expr(value)?;
                self.emit(Op::DUP);
                self.compile_assign_target(target)?;
            }

            // ── Void (JS) ───────────────────────────────────────────────
            ExprKind::Void(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::DROP);
                self.emit(Op::UNDEFINED); // ECMA-262 §13.5.2: void → undefined
            }

            // ── Delete (JS expression) ──────────────────────────────────
            ExprKind::Delete(inner) => {
                // delete obj.prop → ecma:object.delete(obj, key), returns true.
                // Proxy modules route through ecma:proxy.deleteProperty so the
                // deleteProperty trap fires (non-proxy targets fall through).
                let delete_import: (&str, &str) = if self.is_js_profile() && self.uses_proxy {
                    ("ecma:proxy", "deleteProperty")
                } else {
                    ("ecma:object", "delete")
                };
                if let ExprKind::Member { object, field, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    let idx = self.import(delete_import.0, delete_import.1);
                    self.emit_host_call(idx, 2);
                } else if let ExprKind::Index { object, index, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    let idx = self.import(delete_import.0, delete_import.1);
                    self.emit_host_call(idx, 2);
                } else {
                    self.compile_expr(inner)?;
                    self.emit(Op::DROP);
                    self.emit(Op::TRUE);
                }
            }

            // ── Destructure (JS) ────────────────────────────────────────
            ExprKind::Destructure(_) => {
                // Destructure patterns are handled at assignment/declaration sites
                self.emit(Op::NULL);
            }

            // ── Sequence (JS comma operator) ────────────────────────────
            ExprKind::Sequence(exprs) => {
                for (i, e) in exprs.iter().enumerate() {
                    self.compile_expr(e)?;
                    if i < exprs.len() - 1 {
                        self.emit(Op::DROP);
                    }
                }
            }

            // ── ClassExpr (JS) ──────────────────────────────────────────
            ExprKind::ClassExpr {
                name,
                parent,
                members,
            } => {
                let class_name = name
                    .clone()
                    .unwrap_or_else(|| format!("__anonymous_class_{}", self.chunks.len()));
                let class_name = self.canon(&class_name);
                let parent_name: Option<String> = if let Some(p) = parent {
                    let synth_parent =
                        self.canon(&format!("__extends_{}_{}", class_name, self.chunks.len()));
                    self.defined_globals.insert(synth_parent.clone());
                    self.compile_expr(p)?;
                    let parent_idx = self.global_name_const_idx(&synth_parent);
                    self.emit_u16(Op::GLOBAL_SET, parent_idx);
                    self.emit(Op::DROP);
                    Some(synth_parent)
                } else {
                    None
                };
                self.defined_globals.insert(class_name.clone());
                let parents: Vec<String> = parent_name.into_iter().collect();
                crate::common::classes::emit::emit_class_from_ast(
                    self,
                    expr.span.clone(),
                    &class_name,
                    &parents,
                    &[],
                    members,
                    &crate::ast::ClassModifiers::default(),
                    false,
                )?;
                self.emit_var_get(&class_name);
            }

            // ── FunctionExpr (JS) ───────────────────────────────────────
            ExprKind::FunctionExpr(stmt) => {
                if let StmtKind::FunctionDecl {
                    name,
                    params,
                    return_type,
                    body,
                    is_sub,
                    is_generator,
                    handles,
                    is_async,
                    ..
                } = &stmt.kind
                {
                    let fn_name = if name.is_empty() {
                        format!("__anon_fn_{}", self.chunks.len())
                    } else {
                        name.clone()
                    };
                    self.compile_function_decl(
                        &fn_name,
                        params,
                        return_type,
                        body,
                        *is_sub,
                        *is_generator,
                        handles,
                        *is_async,
                    )?;
                    self.emit_var_get(&fn_name);
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── Range ───────────────────────────────────────────────────
            ExprKind::Range {
                start,
                end,
                inclusive: _,
            } => {
                self.compile_expr(start)?;
                self.compile_expr(end)?;
                let line = self.line;
                common::collections::emit_range(&mut self.chunks, self.current, 2, line);
            }

            // ── StaticAccess (PHP) ──────────────────────────────────────
            ExprKind::StaticAccess { class, member } => {
                if let (ExprKind::Ident(class_name), ExprKind::Ident(member_name)) =
                    (&class.kind, &member.kind)
                {
                    if self.js_private_member_access_forbidden(member_name) {
                        self.emit_js_private_access_denied(member_name)?;
                        return Ok(());
                    }
                    if let Some(value) = self.enum_member_ordinal(class_name, member_name) {
                        self.emit_const(Value::F64(value as f64));
                        return Ok(());
                    }

                    let compound = format!("{}.{}", class_name, member_name);
                    if let Some(cv) = self.profile.lookup_constant(&compound) {
                        match cv {
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => {
                                self.emit_const(Value::String(Arc::from(s.as_str())))
                            }
                        }
                        return Ok(());
                    }
                }

                // class::member → look up class, then get static member
                self.compile_expr(class)?;
                if let ExprKind::Ident(name) = &member.kind {
                    let field_name = match &class.kind {
                        ExprKind::Ident(class_name) => {
                            self.js_member_storage_name_for_class(class_name, name)
                        }
                        _ => self.canon(name),
                    };
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                } else {
                    self.compile_expr(member)?;
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                }
            }

            // ── Match expression (PHP/Rust) ─────────────────────────────
            ExprKind::Match { subject, arms } => {
                self.compile_expr(subject)?;
                let subject_slot = self.define_local("__match_subj");
                self.emit_u16(Op::LOCAL_SET, subject_slot);
                self.emit(Op::DROP);
                let result_slot = self.define_local("__match_result");
                let matched_slot = self.define_local("__match_matched");
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.emit(Op::DROP);

                for arm in arms {
                    self.emit_u16(Op::LOCAL_GET, matched_slot);
                    self.emit(Op::I32_EQZ);
                    let arm_line = self.line;
                    self.chunk().emit_if(arm_line);
                    if let Some(ref conditions) = arm.conditions {
                        let arm_match_slot = self.define_local("__match_arm_matches");
                        self.emit_const(Value::I32(0));
                        self.emit_u16(Op::LOCAL_SET, arm_match_slot);
                        self.emit(Op::DROP);
                        for c in conditions {
                            self.emit_u16(Op::LOCAL_GET, subject_slot);
                            self.compile_expr(c)?;
                            {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                            };
                            let cond_line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), cond_line);
                            self.chunk().emit_if(cond_line);
                            self.emit_const(Value::I32(1));
                            self.emit_u16(Op::LOCAL_SET, arm_match_slot);
                            self.emit(Op::DROP);
                            self.chunk().emit_end(cond_line);
                        }
                        self.emit_u16(Op::LOCAL_GET, arm_match_slot);
                        let body_line = self.line;
                        self.chunk().emit_if(body_line);
                        self.compile_expr(&arm.body)?;
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit(Op::DROP);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.emit(Op::DROP);
                        self.chunk().emit_end(body_line);
                    } else {
                        // Default arm
                        self.compile_expr(&arm.body)?;
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit(Op::DROP);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.emit(Op::DROP);
                    }
                    self.chunk().emit_end(arm_line);
                }
                self.emit_u16(Op::LOCAL_GET, result_slot);
            }
        }
        Ok(())
    }

    fn try_compile_pascal_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if self.profile.name != "pascal" {
            return Ok(false);
        }

        if self.expr_is_pascal_set(left) && self.expr_is_pascal_set(right) {
            let helper = match op {
                BinOp::Add => Some("__vybe_pascal_set_union"),
                BinOp::Mul => Some("__vybe_pascal_set_intersection"),
                BinOp::Sub => Some("__vybe_pascal_set_difference"),
                _ => None,
            };

            if let Some(helper) = helper {
                let helper_idx = self.str_const(helper);
                self.emit_u16(Op::GLOBAL_GET, helper_idx);
                self.compile_expr(left)?;
                self.compile_expr(right)?;
                self.emit_u8(Op::CALL_REF, 2);
                return Ok(true);
            }
        }

        let method_name = match op {
            BinOp::Add => "Add",
            BinOp::Eq | BinOp::NotEq => "Equal",
            _ => return Ok(false),
        };

        let Some(type_name) = self.pascal_binary_operator_type(left, right, method_name) else {
            return Ok(false);
        };

        let callee = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&type_name)),
            field: method_name.to_string(),
            null_safe: false,
        });
        let args = vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
        ];
        self.compile_call(&callee, &args)?;
        if *op == BinOp::NotEq {
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_not(self.chunk(), line);
            };
        }
        Ok(true)
    }

    fn try_compile_python_set_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if !self.is_python_profile() {
            return Ok(false);
        }

        let helper = match op {
            BinOp::BitOr => Some("union"),
            BinOp::BitAnd => Some("intersection"),
            BinOp::Sub => Some("difference"),
            BinOp::BitXor => Some("symmetricDifference"),
            _ => None,
        };

        let Some(helper) = helper else {
            return Ok(false);
        };

        self.compile_expr(left)?;
        self.compile_expr(right)?;
        let rhs_slot = self.define_local("__py_set_rhs");
        let lhs_slot = self.define_local("__py_set_lhs");
        self.emit_u16(Op::LOCAL_SET, rhs_slot);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_SET, lhs_slot);
        self.emit(Op::DROP);

        let size_key = self.str_const("size");
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::STRUCT_GET, size_key);
        self.emit(Op::REF_IS_NULL);
        let lhs_has_size_slot = self.define_local("__py_set_lhs_has_size");
        self.emit(Op::I32_EQZ);
        self.emit_u16(Op::LOCAL_SET, lhs_has_size_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit_u16(Op::STRUCT_GET, size_key);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        self.emit_u16(Op::LOCAL_GET, lhs_has_size_slot);
        self.emit(Op::I32_AND);
        let line = self.line;
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        let idx = self.import("ecma:set", helper);
        self.emit_host_call(idx, 2);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.compile_binop(op);
        self.chunk().emit_end(line);
        Ok(true)
    }

    fn try_compile_csharp_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if self.profile.name != "csharp" {
            return Ok(false);
        }

        let left_type = self.infer_expr_type_hint(left);
        let right_type = self.infer_expr_type_hint(right);
        let (Some(left_type), Some(right_type)) = (left_type, right_type) else {
            return Ok(false);
        };
        if Self::normalize_type_hint(&left_type) != Self::normalize_type_hint(&right_type) {
            return Ok(false);
        }

        let arg_exprs = [left, right];
        let (method_name, negate_result) = match op {
            BinOp::Add => ("op_Addition", false),
            BinOp::Eq => ("op_Equality", false),
            BinOp::NotEq
                if self.pending_class_has_method_name_for_type(&left_type, "op_Inequality") =>
            {
                ("op_Inequality", false)
            }
            BinOp::NotEq
                if self.pending_class_has_method_name_for_type(&left_type, "op_Equality") =>
            {
                ("op_Equality", true)
            }
            _ => return Ok(false),
        };

        if let Some(chunk_idx) =
            self.resolve_static_method_overload_chunk_for_type(&left_type, method_name, &arg_exprs)
        {
            self.emit_direct_static_method_call(chunk_idx, &arg_exprs)?;
            if negate_result {
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                };
            }
            return Ok(true);
        }

        let Some(class_name) = self.resolve_pending_class_name_for_type_hint(&left_type) else {
            return Ok(false);
        };

        let callee = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&class_name)),
            field: method_name.to_string(),
            null_safe: false,
        });
        let args = vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
        ];
        self.compile_call(&callee, &args)?;
        if negate_result {
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_not(self.chunk(), line);
            };
        }
        Ok(true)
    }

    fn try_compile_fortran_interface_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if self.profile.name != "fortran" {
            return Ok(false);
        }

        let arg_exprs = vec![left.clone(), right.clone()];
        let Some(target_name) = self.resolve_fortran_operator_target(op, &arg_exprs) else {
            return Ok(false);
        };

        let callee =
            if let Some(module_name) = self.enum_members.get(&self.canon(&target_name)).cloned() {
                let target_canon = self.canon(&target_name);
                let prefers_direct_module_global = self
                    .pending_classes
                    .get(&module_name)
                    .is_some_and(|pending| {
                        pending
                            .static_method_names
                            .iter()
                            .any(|member| member == &target_canon)
                    });
                if prefers_direct_module_global {
                    Expression::ident(&target_name)
                } else {
                    Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(&module_name)),
                        field: target_name.clone(),
                        null_safe: false,
                    })
                }
            } else {
                Expression::ident(&target_name)
            };
        let args = vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
        ];
        self.compile_call(&callee, &args)?;
        Ok(true)
    }

    fn pascal_binary_operator_type(
        &self,
        left: &Expression,
        right: &Expression,
        method_name: &str,
    ) -> Option<String> {
        let left_type = self.pascal_expr_static_type(left)?;
        let right_type = self.pascal_expr_static_type(right)?;
        if !left_type.eq_ignore_ascii_case(&right_type) {
            return None;
        }

        let bare_type = left_type
            .split('<')
            .next()
            .unwrap_or(left_type.as_str())
            .trim();
        let canon_type = self.canon(bare_type);
        if !self.defined_globals.contains(&canon_type) {
            return None;
        }
        if !self
            .defined_class_methods
            .contains(&self.canon(method_name))
        {
            return None;
        }
        Some(bare_type.to_string())
    }

    pub(super) fn pascal_expr_static_type(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
            _ => None,
        }
    }

    fn dotnet_expr_static_type(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
            ExprKind::New { class, .. } => match &class.kind {
                ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
                ExprKind::Member { field, .. } => Some(field.to_string()),
                _ => None,
            },
            ExprKind::Call { callee, .. } => match &callee.kind {
                ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("Parse") => {
                    match &object.kind {
                        ExprKind::Ident(name) if name.eq_ignore_ascii_case("Version") => {
                            Some("Version".into())
                        }
                        ExprKind::Member { field, .. } if field.eq_ignore_ascii_case("Version") => {
                            Some("Version".into())
                        }
                        _ => None,
                    }
                }
                _ => None,
            },
            _ => None,
        }
    }

    fn is_dotnet_type_name(type_name: &str, expected: &str) -> bool {
        type_name.eq_ignore_ascii_case(expected)
            || type_name
                .rsplit('.')
                .next()
                .is_some_and(|short| short.eq_ignore_ascii_case(expected))
    }

    fn try_compile_dotnet_datetime_timespan_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if !self.profile.namespaces.use_dotnet {
            return Ok(false);
        }

        let Some(left_type) = self.dotnet_expr_static_type(left) else {
            return Ok(false);
        };
        let Some(right_type) = self.dotnet_expr_static_type(right) else {
            return Ok(false);
        };

        let emit = match op {
            BinOp::Add
                if Self::is_dotnet_type_name(&left_type, "TimeSpan")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") =>
            {
                Some("dotnet.timespan_add")
            }
            BinOp::Sub
                if Self::is_dotnet_type_name(&left_type, "TimeSpan")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") =>
            {
                Some("dotnet.timespan_sub")
            }
            BinOp::Add
                if Self::is_dotnet_type_name(&left_type, "DateTime")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") =>
            {
                Some("dotnet.datetime_add_timespan")
            }
            BinOp::Sub
                if Self::is_dotnet_type_name(&left_type, "DateTime")
                    && Self::is_dotnet_type_name(&right_type, "DateTime") =>
            {
                Some("dotnet.datetime_subtract_datetime")
            }
            BinOp::Lt
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_lt")
            }
            BinOp::Gt
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_gt")
            }
            BinOp::Eq
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_eq")
            }
            BinOp::NotEq
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_ne")
            }
            _ => None,
        };

        let Some(emit) = emit else {
            return Ok(false);
        };

        self.compile_expr(left)?;
        self.compile_expr(right)?;
        let line = self.line;
        self.emit_common(emit, 2, line);
        Ok(true)
    }

    pub(super) fn pascal_helper_function_name(&self, type_name: &str, method_name: &str) -> String {
        let sanitize = |text: &str| {
            text.chars()
                .map(|ch| {
                    if ch.is_ascii_alphanumeric() {
                        ch.to_ascii_lowercase()
                    } else {
                        '_'
                    }
                })
                .collect::<String>()
        };
        format!(
            "__pascal_helper_{}_{}",
            sanitize(type_name),
            sanitize(method_name),
        )
    }

    fn emit_c_unsigned_wrap(&mut self, modulus: f64) {
        self.emit_const(Value::F64(modulus));
        self.compile_binop(&BinOp::Mod);
        self.emit_const(Value::F64(modulus));
        self.emit(Op::F64_ADD);
        self.emit_const(Value::F64(modulus));
        self.compile_binop(&BinOp::Mod);
    }

    fn emit_c_signed_wrap_from_unsigned(&mut self, threshold: f64, modulus: f64) {
        self.emit(Op::DUP);
        self.emit_const(Value::F64(threshold));
        self.emit(Op::F64_GE);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_const(Value::F64(modulus));
        self.emit(Op::F64_SUB);
        self.chunk().emit_else(line);
        self.chunk().emit_end(line);
    }
}
