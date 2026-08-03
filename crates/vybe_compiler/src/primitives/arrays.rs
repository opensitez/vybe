//! Collection/array binding, indexing, SetLength, fixed-array initializers, var-decl coercion.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

impl Compiler {
    /// Whether `type_hint` names a dictionary/map.
    ///
    /// Spellings moved to `vybe_ast::builtin_types` (`builtinslotplan.md`
    /// step 4). Still narrow — it catches .NET's `Dictionary<K,V>` but not
    /// Dart's `Map<K,V>` nor a Python dict — and that narrowness is now
    /// visible in one table instead of buried in a predicate.
    pub(super) fn is_dictionary_type_hint(type_hint: &str) -> bool {
        vybe_ast::builtin_types::is(type_hint, vybe_ast::builtin_slots::BuiltinType::Map)
    }

    pub(super) fn is_sorted_dictionary_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).contains("sorteddictionary")
    }

    pub(super) fn is_sorted_set_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).contains("sortedset")
    }

    pub(super) fn is_case_insensitive_string_key_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).contains("#ordinalignorecase")
    }

    pub(super) fn expr_uses_case_insensitive_string_keys(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(Self::is_case_insensitive_string_key_type_hint),
            _ => self
                .infer_expr_type_hint(expr)
                .as_deref()
                .is_some_and(Self::is_case_insensitive_string_key_type_hint) }
    }

    pub(super) fn compile_collection_key(
        &mut self,
        owner: &Expression,
        key: &Expression,
    ) -> Result<(), String> {
        self.compile_array_index_operand_for_owner(owner, key)?;
        if self.expr_uses_case_insensitive_string_keys(owner) {
            let line = self.line;
            common::strings::emit_to_lower(self.chunk(), line);
        }
        Ok(())
    }

    pub(super) fn is_callable_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        if normalized.ends_with("()") {
            return false;
        }
        let lower = normalized.to_ascii_lowercase();
        if lower.starts_with("procedure(") || lower == "procedure" {
            return true;
        }
        let leaf = lower.rsplit('.').next().unwrap_or(lower.as_str());
        let bare = leaf
            .split('<')
            .next()
            .unwrap_or(leaf)
            .split('(')
            .next()
            .unwrap_or(leaf)
            .trim();
        bare.ends_with("eventhandler")
            || matches!(
                bare,
                "func" | "action" | "eventhandler" | "predicate" | "comparison" | "converter"
            )
            || lower.contains(" delegate")
    }

    pub(super) fn callable_return_type_hint(type_hint: &str) -> Option<String> {
        let normalized = Self::normalize_type_hint(type_hint);
        let return_type = normalized.rsplit_once("->")?.1.trim();
        if return_type.is_empty() {
            None
        } else {
            Some(return_type.to_string())
        }
    }

    pub(super) fn lookup_var_type_hint(&self, name: &str) -> Option<&str> {
        if let Some(binding) = self.static_local_binding(name) {
            if let Some(type_hint) = binding.type_hint.as_deref() {
                return Some(type_hint);
            }
        }
        if let Some(type_hint) = self.scope().resolve_type(name) {
            return Some(type_hint);
        }
        for scope in self.scopes.iter().rev().skip(1) {
            if let Some(type_hint) = scope.resolve_type(name) {
                return Some(type_hint);
            }
        }
        if let Some(type_hint) = self.lookup_implicit_self_field_type_hint(name) {
            return Some(type_hint);
        }
        let cname = self.canon(name);
        self.global_type_hints.get(&cname).map(|s| s.as_str())
    }

    pub(super) fn has_accessible_local_binding(&self, name: &str) -> bool {
        if self.static_local_binding(name).is_some() {
            return true;
        }
        self.scopes.iter().rev().any(|scope| {
            scope.resolve(name).is_some()
        })
    }

    pub(super) fn static_local_binding(&self, name: &str) -> Option<&StaticLocalBinding> {
        let canon_name = self.canon(name);
        self.static_local_bindings
            .iter()
            .rev()
            .find_map(|bindings| bindings.get(&canon_name))
    }

    pub(super) fn has_static_local_binding(&self, name: &str) -> bool {
        self.static_local_binding(name).is_some()
    }

    pub(super) fn array_binding_key(&self, name: &str) -> String {
        let canon_name = self.canon(name);
        if self.scopes.len() > 1 {
            let class_name = self.current_class.as_deref().unwrap_or("<module>");
            let func_name = self.current_func_name.as_deref().unwrap_or("<top>");
            format!(
                "{}::{}::{}",
                self.canon(class_name),
                self.canon(func_name),
                canon_name
            )
        } else {
            canon_name
        }
    }

    pub(super) fn record_array_binding(&mut self, name: &str, metadata: ArrayBindingMetadata) {
        let key = self.array_binding_key(name);
        self.array_bindings.insert(key, metadata);
    }

    pub(super) fn lookup_array_binding(&self, name: &str) -> Option<&ArrayBindingMetadata> {
        let key = self.array_binding_key(name);
        self.array_bindings
            .get(&key)
            .or_else(|| self.array_bindings.get(&self.canon(name)))
    }

    /// Declared index bounds for an array binding, when the language records
    /// them.
    ///
    /// No language check: `pascal_bounds` is only ever populated from an
    /// `array[lo..hi]` type hint, so a language that does not write that shape
    /// gets `None` from the lookup anyway.
    pub(super) fn array_index_bounds_for_owner(
        &self,
        owner: &Expression,
    ) -> Option<PascalArrayBoundsMetadata> {
        if let ExprKind::Ident(name) = &owner.kind {
            if let Some(bounds) = self
                .lookup_array_binding(name)
                .and_then(|binding| binding.pascal_bounds.clone())
            {
                return Some(bounds);
            }
        }

        self.infer_expr_type_hint(owner)
            .as_deref()
            .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint))
    }

    pub(super) fn profile_array_index_semantics(&self) -> Option<ArrayIndexSemantics> {
        match self.profile.name.as_str() {
            _ => None }
    }

    pub(super) fn normalized_array_index_operand_for_owner(
        &self,
        owner: &Expression,
        index: Expression,
    ) -> Expression {
        if let Some(bounds) = self.array_index_bounds_for_owner(owner) {
            if let Some(dimension) = bounds.dimensions.first() {
                let normalized_index = if dimension.uses_char_ordinal {
                    Self::pascal_ordinal_index_expr(index)
                } else {
                    index
                };
                return normalize_array_index_operand(
                    normalized_index,
                    ArrayIndexSemantics {
                        first_index: dimension.first_index },
                );
            }
        }

        if let Some(semantics) = self.profile_array_index_semantics() {
            return normalize_array_index_operand(index, semantics);
        }

        index
    }

    #[allow(dead_code)]
    pub(super) fn compile_array_index_operand(&mut self, index: &Expression) -> Result<(), String> {
        if let Some(semantics) = self.profile_array_index_semantics() {
            let normalized = normalize_array_index_operand(index.clone(), semantics);
            self.compile_expr(&normalized)
        } else {
            self.compile_expr(index)
        }
    }

    pub(super) fn compile_array_index_operand_for_owner(
        &mut self,
        owner: &Expression,
        index: &Expression,
    ) -> Result<(), String> {
        let normalized = self.normalized_array_index_operand_for_owner(owner, index.clone());
        self.compile_expr(&normalized)
    }

    pub(super) fn compile_setlength(
        &mut self,
        target: &Expression,
        len_expr: &Expression,
    ) -> Result<(), String> {
        if let ExprKind::Ident(name) = &target.kind {
            self.compile_expr(target)?;
            let arr_slot = self.define_local("__setlength_array");
            self.emit_u16(Op::LOCAL_SET, arr_slot);

            self.emit_u16(Op::LOCAL_GET, arr_slot);
            self.emit(Op::REF_IS_NULL);
            let line = self.line;
            self.chunk().emit_if(line);
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            self.emit_u16(Op::LOCAL_SET, arr_slot);
            self.emit_u16(Op::LOCAL_GET, arr_slot);
            self.emit_var_set(name);
            self.chunk().emit_end(line);

            self.emit_u16(Op::LOCAL_GET, arr_slot);
        } else {
            self.compile_expr(target)?;
        }
        self.compile_expr(len_expr)?;
        let set_length_idx = self.import("ecma:array", "setLength");
        self.emit_host_call(set_length_idx, 2);
        self.emit(Op::DROP);
        self.emit_null();
        Ok(())
    }

    pub(super) fn ensure_static_local_binding(
        &mut self,
        name: &str,
        type_hint: Option<String>,
    ) -> Result<StaticLocalBinding, String> {
        let canon_name = self.canon(name);
        let normalized_type_hint = type_hint.as_deref().map(Self::normalize_type_hint);

        if let Some(existing) = self
            .static_local_bindings
            .last_mut()
            .and_then(|bindings| bindings.get_mut(&canon_name))
        {
            if existing.type_hint.is_none() {
                existing.type_hint = normalized_type_hint;
            }
            return Ok(existing.clone());
        }

        let func_name = self
            .current_func_name
            .as_deref()
            .map(|name| self.canon(name))
            .unwrap_or_else(|| "anon".to_string());
        let Some(bindings) = self.static_local_bindings.last_mut() else {
            return Err(format!("static local `{name}` declared outside a function"));
        };
        let global_name = format!(
            "__staticlocal_{}_{}_{}",
            self.current, func_name, canon_name
        );
        let binding = StaticLocalBinding {
            init_flag_name: format!("{}__init", global_name),
            global_name,
            type_hint: normalized_type_hint };
        bindings.insert(canon_name, binding.clone());
        Ok(binding)
    }

    pub(super) fn emit_multidim_inclusive_array_initializer(
        &mut self,
        bounds: &[Expression],
    ) -> Result<(), String> {
        let line = self.line;
        if bounds.is_empty() {
            self.emit_null();
            return Ok(());
        }

        self.compile_expr(&bounds[0])?;
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };

        if bounds.len() == 1 {
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            return Ok(());
        }

        let len_slot = self.define_local("__vb_md_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__vb_md_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);

        let index_slot = self.define_local("__vb_md_index");
        inst!(self, core_wasm::f64_const, 0.0);
        self.emit_u16(Op::LOCAL_SET, index_slot);

        let fill_block = self.chunk().emit_block(line);
        let (fill_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_multidim_inclusive_array_initializer(&bounds[1..])?;
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(fill_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(fill_block);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        Ok(())
    }

    pub(super) fn emit_pascal_fixed_array_initializer(
        &mut self,
        dimensions: &[PascalArrayDimensionMetadata],
    ) -> Result<(), String> {
        let line = self.line;
        if dimensions.is_empty() {
            self.emit_null();
            return Ok(());
        }

        self.emit_const(Value::F64(dimensions[0].length as f64));

        if dimensions.len() == 1 {
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            return Ok(());
        }

        let len_slot = self.define_local("__pascal_md_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__pascal_md_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);

        let index_slot = self.define_local("__pascal_md_index");
        inst!(self, core_wasm::f64_const, 0.0);
        self.emit_u16(Op::LOCAL_SET, index_slot);

        let fill_block = self.chunk().emit_block(line);
        let (fill_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_pascal_fixed_array_initializer(&dimensions[1..])?;
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(fill_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(fill_block);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        Ok(())
    }

    pub(super) fn emit_fortran_fixed_array_initializer(
        &mut self,
        bounds: &[Expression],
    ) -> Result<(), String> {
        let line = self.line;
        if bounds.is_empty() {
            self.emit_null();
            return Ok(());
        }

        self.compile_expr(&bounds[0])?;

        if bounds.len() == 1 {
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            return Ok(());
        }

        let len_slot = self.define_local("__fortran_md_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__fortran_md_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);

        let index_slot = self.define_local("__fortran_md_index");
        inst!(self, core_wasm::f64_const, 0.0);
        self.emit_u16(Op::LOCAL_SET, index_slot);

        let fill_block = self.chunk().emit_block(line);
        let (fill_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_fortran_fixed_array_initializer(&bounds[1..])?;
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(fill_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(fill_block);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        Ok(())
    }

    pub(super) fn is_fortran_fixed_array_synth_init(expr: &Expression) -> bool {
        matches!(
            &expr.kind,
            ExprKind::Call { callee, args, optional: false }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "Array")
                    && args.len() == 2
                    && matches!(args[1].value.kind, ExprKind::Lit(Literal::Int(0)))
        )
    }

    pub(super) fn emit_var_decl_initializer_value(
        &mut self,
        decl: &VarDeclarator,
        resolved_type_hint: Option<&str>,
    ) -> Result<(), String> {
        if let Some(ref init_expr) = decl.init {
            if self.profile.array_bounds_declare_fixed_shape
                && decl
                    .array_bounds
                    .as_ref()
                    .is_some_and(|bounds| !bounds.is_empty())
                && Self::is_fortran_fixed_array_synth_init(init_expr)
            {
                self.emit_fortran_fixed_array_initializer(
                    decl.array_bounds
                        .as_ref()
                        .expect("checked non-empty bounds"),
                )?;
            } else {
                self.compile_expr_with_value_copy(init_expr)?;
                let effective_type_hint = resolved_type_hint.or(decl.type_hint.as_deref());
                let skip_c_coerce = if self.profile.aggregate_decl_skips_coercion {
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
                self.maybe_promote_array_literal_to_set(
                    decl.type_hint.as_deref(),
                    init_expr,
                );
            }
        } else if let Some(ref bounds) = decl.array_bounds {
            if self.profile.array_bounds_declare_fixed_shape {
                self.emit_fortran_fixed_array_initializer(bounds)?;
            } else if bounds.len() > 1 {
                // Multi-dimensional inclusive-bound declaration (`Dim a(2,3)`).
                // The whole `array_bounds` path is inclusive (single-dim below
                // adds +1 for every language), so any language producing
                // multi-dim bounds here wants the same nested inclusive array —
                // no language-name gate needed. C# arrays carry their size in
                // the `new[...]` expression, not `array_bounds`, so they never
                // reach this path.
                self.emit_multidim_inclusive_array_initializer(bounds)?;
            } else if let Some(size_expr) = bounds.first() {
                let line = self.line;
                self.compile_expr(size_expr)?;
                self.emit_const(Value::F64(1.0));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                };
                common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            } else {
                self.emit_null();
            }
        } else {
            let resolved_type_hint = resolved_type_hint.map(str::to_string).or_else(|| {
                decl.type_hint
                    .as_deref()
                    .map(|type_hint| self.resolve_source_type_alias(type_hint))
            });
            let effective_type_hint = resolved_type_hint.as_deref().or(decl.type_hint.as_deref());

            if let Some(metadata) = effective_type_hint
                .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint))
                .filter(|metadata| metadata.is_fixed)
            {
                if metadata.dimensions.len() > 1 {
                    self.emit_pascal_fixed_array_initializer(&metadata.dimensions)?;
                } else if let Some(dimension) = metadata.dimensions.first() {
                    let line = self.line;
                    self.emit_const(Value::F64(dimension.length as f64));
                    common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                } else {
                    self.emit_null();
                }
            } else if effective_type_hint
                .and_then(Self::vb_fixed_string_len)
                .is_some()
            {
                self.emit_const(Value::String(Arc::from("")));
            } else if let Some(type_name) = decl
                .type_hint
                .as_deref()
                .and_then(|type_hint| self.user_value_type_name_from_hint(type_hint))
            {
                let ctor_global = {
                    let overload = crate::primitives::classes::ctor_global_for(&type_name, 0);
                    if self.defined_globals.contains(&overload) {
                        overload
                    } else {
                        type_name.clone()
                    }
                };
                let idx = self.str_const(&ctor_global);
                self.emit_u16(Op::GLOBAL_GET, idx);
                self.emit_u8(Op::CALL_REF, 0);
                return Ok(());
            } else {
                match effective_type_hint.map(|s| s.to_lowercase()).as_deref() {
                    Some("integer") | Some("int") | Some("longint") | Some("real")
                    | Some("double") | Some("float") => {
                        inst!(self, core_wasm::f64_const, 0.0);
                    }
                    Some("boolean") | Some("bool") => inst!(self, core_wasm::bool_const, false),
                    Some(type_hint) if Self::is_string_type_hint(type_hint) => {
                        self.emit_const(Value::String(Arc::from("")))
                    }
                    _ => self.emit_null() }
            }
        }
        if let Some(target_len) = decl
            .type_hint
            .as_deref()
            .and_then(Self::vb_fixed_string_len)
        {
            self.emit_vb_fixed_string_adjust_from_stack(target_len, false);
        }
        Ok(())
    }

    pub(super) fn coerce_c_value_for_type_hint(
        &mut self,
        type_hint: Option<&str>,
    ) -> Result<(), String> {
        // Dynamically-typed languages infer a type *hint* for dispatch only and
        // must never mutate the value: e.g. JS `let t = true` infers "bool", and
        // C-style value coercion (`_Bool` → i32 0/1, int-width truncation) would
        // flatten the boolean to a number (`typeof` "number", prints "1") on
        // both declaration and later assignment. Driven by the profile
        // capability, not the language name.
        if !self.profile.coerces_value_to_type_hint {
            return Ok(());
        }
        let Some(type_hint) = type_hint else {
            return Ok(());
        };
        let normalized = Self::normalize_type_hint(type_hint);
        match normalized.as_str() {
            "bool" | "boolean" | "_bool" => {
                if self.profile.materialize_bool_results {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    return Ok(());
                }
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            }
            // A language whose `char` holds a CHARACTER, not an 8-bit
            // integer, must not get the modular byte coercion below. Was
            "char" if self.hint_is_builtin_string(&normalized) => {}
            "char" | "uint8" | "unsigned char" | "byte" => {
                self.emit(Op::F64_TRUNC);
                self.emit_const(Value::F64(256.0));
                self.compile_binop(&BinOp::Mod);
                self.emit_const(Value::F64(256.0));
                self.emit(Op::F64_ADD);
                self.emit_const(Value::F64(256.0));
                self.compile_binop(&BinOp::Mod);
            }
            "signed char" | "int8" | "sbyte" => {
                // Signed 8-bit: wrap to 0..255 then sign-extend (>= 128 → −256),
                // mirroring the int16 path. (`i8` range is −128..127.)
                self.emit(Op::F64_TRUNC);
                self.emit_const(Value::F64(256.0));
                self.compile_binop(&BinOp::Mod);
                self.emit_const(Value::F64(256.0));
                self.emit(Op::F64_ADD);
                self.emit_const(Value::F64(256.0));
                self.compile_binop(&BinOp::Mod);
                inst!(self, core_wasm::dup);
                self.emit_const(Value::F64(128.0));
                self.emit(Op::F64_GE);
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.emit_const(Value::F64(256.0));
                self.emit(Op::F64_SUB);
                self.chunk().emit_else(line);
                self.chunk().emit_end(line);
            }
            "int16" => {
                self.emit(Op::F64_TRUNC);
                self.emit_const(Value::F64(65_536.0));
                self.compile_binop(&BinOp::Mod);
                self.emit_const(Value::F64(65_536.0));
                self.emit(Op::F64_ADD);
                self.emit_const(Value::F64(65_536.0));
                self.compile_binop(&BinOp::Mod);
                inst!(self, core_wasm::dup);
                self.emit_const(Value::F64(32_768.0));
                self.emit(Op::F64_GE);
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.emit_const(Value::F64(65_536.0));
                self.emit(Op::F64_SUB);
                self.chunk().emit_else(line);
                self.chunk().emit_end(line);
            }
            "uint32" | "unsigned int" => {
                self.emit(Op::F64_TRUNC);
                self.emit_const(Value::F64(4_294_967_296.0));
                self.compile_binop(&BinOp::Mod);
                self.emit_const(Value::F64(4_294_967_296.0));
                self.emit(Op::F64_ADD);
                self.emit_const(Value::F64(4_294_967_296.0));
                self.compile_binop(&BinOp::Mod);
            }
            "int" => {
                self.emit(Op::F64_TRUNC);
                self.emit_const(Value::F64(4_294_967_296.0));
                self.compile_binop(&BinOp::Mod);
                self.emit_const(Value::F64(4_294_967_296.0));
                self.emit(Op::F64_ADD);
                self.emit_const(Value::F64(4_294_967_296.0));
                self.compile_binop(&BinOp::Mod);
                inst!(self, core_wasm::dup);
                self.emit_const(Value::F64(2_147_483_648.0));
                self.emit(Op::F64_GE);
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.emit_const(Value::F64(4_294_967_296.0));
                self.emit(Op::F64_SUB);
                self.chunk().emit_else(line);
                self.chunk().emit_end(line);
            }
            "float" | "single" => {
                self.emit_const(Value::F64(10_000_000.0));
                self.compile_binop(&BinOp::Mul);
                let idx = self.import("ecma:math", "trunc");
                self.emit_host_call(idx, 1);
                self.emit_const(Value::F64(10_000_000.0));
                self.compile_binop(&BinOp::Div);
            }
            _ => {}
        }
        Ok(())
    }
}
