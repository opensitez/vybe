//! `compile_expr` — the `ExprKind` dispatch. The single largest
//! method in the compiler, split out so edits to other concerns
//! (calls, classes, statements) don't churn a multi-thousand-line
//! file.

use super::*;

impl Compiler {
    // ════════════════════════════════════════════════════════════════════════
    // Expression compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            // ── Literals ────────────────────────────────────────────────
            ExprKind::Lit(lit) => {
                match lit {
                    Literal::Int(n) => self.emit_const(Value::F64(*n as f64)),
                    Literal::Float(n) => self.emit_const(Value::F64(*n)),
                    Literal::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                    Literal::Char(c) => self.emit_const(Value::String(Arc::from(c.to_string().as_str()))),
                    Literal::Bool(b) => if *b { self.emit(Op::TRUE) } else { self.emit(Op::FALSE) },
                    Literal::Null => self.emit(Op::NULL),
                    Literal::Undefined => { let l = self.line; common::expressions::emit_undefined(self.chunk(), l); }
                    Literal::Ellipsis => self.emit(Op::NULL),
                }
            }

            // ── Identifier ──────────────────────────────────────────────
            ExprKind::Ident(name) => {
                // JS global constants that aren't variables
                match name.as_str() {
                    "NaN" => { self.emit_const(Value::F64(f64::NAN)); return Ok(()); }
                    "Infinity" => { self.emit_const(Value::F64(f64::INFINITY)); return Ok(()); }
                    "undefined" if self.case_sensitive => { let l = self.line; common::expressions::emit_undefined(self.chunk(), l); return Ok(()); }
                    _ => {}
                }
                if self.is_python_profile() {
                    match name.as_str() {
                        "__debug__" => { self.emit(Op::TRUE); return Ok(()); }
                        "__name__" => { self.emit_const(Value::String(Arc::from("__main__"))); return Ok(()); }
                        _ => {}
                    }
                }
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                    || self.has_static_local_binding(name);

                // Implicit self field access (only if NOT a local)
                if !is_local && self.is_class_field(name) {
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        self.emit_u16(Op::LOCAL_GET, slot);
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
                            ConstantValue::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
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

                if self.profile.name == "vb" && !is_local && self.defined_functions.contains(name.as_str()) {
                    self.emit_var_get(name);
                    self.emit_u8(Op::CALL_REF, 0);
                    return Ok(());
                }

                self.emit_var_get(name);
            }

            // ── This / Super ────────────────────────────────────────────
            ExprKind::This => {
                let self_kw = &self.profile.self_keyword;
                if let Some(slot) = self.scope().resolve(self_kw)
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
                    if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|pc| pc.parent.clone()) {
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
                let expr_is_csharp_integral = |compiler: &Compiler, expr: &Expression| {
                    matches!(expr.kind, ExprKind::Lit(Literal::Int(_)))
                        || compiler
                            .infer_expr_type_hint(expr)
                            .as_deref()
                            .is_some_and(is_csharp_integral_type)
                };

                // Short-circuit for And/Or
                if *op == BinOp::And {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip = common::expressions::emit_and_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(&mut self.chunks[self.current], skip);
                    return Ok(());
                }
                if *op == BinOp::Or {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip = common::expressions::emit_or_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(&mut self.chunks[self.current], skip);
                    return Ok(());
                }
                // NullCoalesce as binary op
                if *op == BinOp::NullCoalesce {
                    self.compile_expr(left)?;
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::DROP);
                    self.compile_expr(right)?;
                    self.patch_jump(skip);
                    return Ok(());
                }
                // Pow → canonical stdlib path: push func ref BEFORE operands
                // so [func, base, exponent] is on the stack for call_ref.
                if *op == BinOp::Pow {
                    let line = self.line;
                    common::math::emit_pow_push_func(self.chunk(), line);
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    common::math::emit_pow_invoke(self.chunk(), line);
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
                        self.compile_expr(right)?;
                        self.compile_binop(op);
                        return Ok(());
                    }
                    if let crate::ast::ExprKind::Ident(type_name) = &right.kind {
                        self.compile_expr(left)?;
                        let line = self.line;
                        let name_canon = self.canon(type_name);
                        let idx = self.chunk().add_constant(
                            vybe_bytecode::Value::String(std::sync::Arc::from(name_canon.as_str())),
                        );
                        self.chunk().emit_op_u16(vybe_bytecode::Op::REF_TEST, idx, line);
                        return Ok(());
                    }
                }
                if self.profile.name == "pascal" && (*op == BinOp::In || *op == BinOp::NotIn) {
                    if !self.expr_is_pascal_set(right) {
                        let line = self.line;
                        self.compile_expr(right)?;
                        self.compile_expr(left)?;
                        common::collections::emit_contains(&mut self.chunks, self.current, line);
                        if *op == BinOp::NotIn {
                            self.emit(Op::DYN_NOT);
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
                        if *op == UnaryOp::PostInc { self.emit(Op::DUP); }
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        if *op == UnaryOp::PreInc { self.emit(Op::DUP); }
                        self.compile_assign_target(inner)?;
                    }
                    UnaryOp::PreDec | UnaryOp::PostDec => {
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostDec { self.emit(Op::DUP); }
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::F64_SUB);
                        if *op == UnaryOp::PreDec { self.emit(Op::DUP); }
                        self.compile_assign_target(inner)?;
                    }
                    _ => {
                        self.compile_expr(inner)?;
                        match op {
                            UnaryOp::Neg => { let l = self.line; common::math::emit_neg(self.chunk(), l); }
                            UnaryOp::Pos => {
                                // JS `+v` coerces to number — ECMA-262 §7.1.4 ToNumber.
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
                            UnaryOp::Not => self.emit(Op::DYN_NOT),
                            UnaryOp::BitNot => { let l = self.line; common::expressions::emit_i32_not(self.chunk(), l); }
                            UnaryOp::Typeof => self.emit(Op::REF_TYPEOF),
                            UnaryOp::Void => { self.emit(Op::DROP); self.emit(Op::UNDEFINED); }
                            UnaryOp::Delete => { self.emit(Op::DROP); self.emit(Op::TRUE); }
                            UnaryOp::Deref => { let idx = self.str_const("__value"); self.emit_u16(Op::STRUCT_GET, idx); }
                            UnaryOp::AddrOf => {} // no-op in VM
                            UnaryOp::Await => {} // handled below in ExprKind::Await
                            _ => {} // PreInc etc handled above
                        }
                    }
                }
            }

            // ── Ternary ─────────────────────────────────────────────────
            ExprKind::Ternary { cond, then, else_ } => {
                self.compile_expr(cond)?;
                self.emit_python_truthiness_from_stack();
                let else_j = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(then)?;
                let end_j = self.emit_jump(Op::BR);
                self.patch_jump(else_j);
                self.compile_expr(else_)?;
                self.patch_jump(end_j);
            }

            // ── Call ────────────────────────────────────────────────────
            ExprKind::Call { callee, args, optional } => {
                if *optional {
                    // Optional call: callee?.() — short-circuit to null if callee is null/undefined.
                    // Stack: compile callee → [func_or_null].
                    // Dup to check null while preserving the original for the call.
                    self.compile_expr(callee)?;
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_TRUE);
                    // Not null — call it. Stack: [func]. Compile args, call.
                    for a in args { self.compile_expr(&a.value)?; }
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(skip);
                    // Null path: the dup left [null] on stack, use it as result
                    self.patch_jump(end);
                } else {
                    if self.profile.parens_for_index
                        && !args.is_empty()
                        && matches!(&callee.kind,
                            ExprKind::Ident(name) if self.lookup_array_binding(name).is_some()
                        )
                    {
                        self.compile_expr(callee)?;
                        for arg in args {
                            self.compile_expr(&arg.value)?;
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
            ExprKind::Member { object, field, null_safe } => {
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
                            self.emit(if self.reflection_is_enum_type(&type_name) { Op::TRUE } else { Op::FALSE });
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "IsValueType") => {
                            self.emit(if self.reflection_is_value_type(&type_name) { Op::TRUE } else { Op::FALSE });
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
                        (ReflectionBinding::Method { type_name, method_name }, "IsStatic") => {
                            let is_static = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.methods.get(&method_name))
                                .map(|meta| meta.is_static)
                                .unwrap_or(false);
                            self.emit(if is_static { Op::TRUE } else { Op::FALSE });
                            return Ok(());
                        }
                        (ReflectionBinding::Property { type_name, property_name }, "CanWrite") => {
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
                                || self.enum_value_names.keys().any(|known| known.eq_ignore_ascii_case(candidate))
                        });
                        self.emit(if is_enum { Op::TRUE } else { Op::FALSE });
                        return Ok(());
                    }
                }

                // Namespace constant check (Math.PI, etc.)
                if let ExprKind::Ident(obj_name) = &object.kind {
                    let prefers_type_lookup = self.prefers_type_qualified_member_lookup(obj_name, field);
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
                            ConstantValue::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                        }
                        return Ok(());
                    }
                    // Constructor call with 0 args: ClassName.Create
                    let ctor_nm = &self.profile.constructor_name;
                    let is_ctor = if self.case_sensitive { field == ctor_nm } else { field.eq_ignore_ascii_case(ctor_nm) };
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
                            let zero_arg_static = self.pending_classes.get(canon_obj.as_str())
                                .map(|pc| !pc.static_fields.iter().any(|name| name == &method_name))
                                .unwrap_or(false);
                            if zero_arg_static {
                                let cls_idx = self.str_const(&canon_obj);
                                self.emit_u16(Op::GLOBAL_GET, cls_idx);
                                self.emit(Op::DUP);
                                let method_idx = self.str_const(&method_name);
                                self.emit_u16(Op::STRUCT_GET, method_idx);
                                let fn_tmp = self.scope().resolve("__pascal_static_fn")
                                    .unwrap_or_else(|| self.define_local("__pascal_static_fn"));
                                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                self.emit(Op::DROP);
                                let cls_tmp = self.scope().resolve("__pascal_static_cls")
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
                        let lower_parts: Vec<String> = parts.iter().map(|part| self.canon(part)).collect();
                        if common::dotnet::is_namespace_root(&lower_parts[0]) {
                            let scope = self.scope();
                            let dotnet_surface = common::dotnet::surface();
                            let mut imports = dotnet_surface.default_imports().to_vec();
                            imports.extend(self.profile.namespaces.extra_imports.clone());
                            let field_set: std::collections::HashSet<String> = if let Some(ref class_name) = self.current_class {
                                self.pending_classes.get(class_name.as_str())
                                    .map(|pending| pending.fields.iter().cloned().collect())
                                    .unwrap_or_default()
                            } else {
                                std::collections::HashSet::new()
                            };
                            let defined_globals = self.defined_globals.clone();
                            let defined_classes = self.defined_classes.clone();
                            let is_user_class_fn = move |name: &str| -> bool {
                                defined_classes.contains(name)
                                    || defined_classes.iter().any(|class_name| class_name.eq_ignore_ascii_case(name))
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
                                        || defined_globals.iter().any(|global_name| global_name.eq_ignore_ascii_case(name))
                                },
                                is_class_field: &|name: &str| field_set.contains(name),
                                is_user_type: &is_user_class_fn,
                                imports: &imports,
                            };
                            let refs: Vec<&str> = lower_parts.iter().map(|part| part.as_str()).collect();
                            match common::dotnet::resolve_dotted_name(&refs, &ctx) {
                                common::dotnet::DottedResolution::CommonCall { emit } => {
                                    self.emit_common(&emit, 0, self.line);
                                    return Ok(());
                                }
                                common::dotnet::DottedResolution::HostCall { module, func } => {
                                    let idx = self.import(&module, &func);
                                    self.emit_host_call(idx, 0);
                                    return Ok(());
                                }
                                _ => {}
                            }
                        }
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
                        &mut self.chunks, self.current, line,
                    );
                    return Ok(());
                }

                if self.is_js_profile() {
                    if *null_safe {
                        self.compile_expr(object)?;
                        self.emit(Op::DUP);
                        self.emit(Op::REF_IS_NULL);
                        let non_null = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit(Op::DROP);
                        let line = self.line;
                        common::expressions::emit_undefined(self.chunk(), line);
                        let end = self.emit_jump(Op::BR);
                        self.patch_jump(non_null);
                        let obj_slot = self.define_local("__js_member_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        self.emit(Op::DROP);
                        let field_name = self.canon(field);
                        let prop = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let val_slot = self.define_local("__js_member_val");
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.emit(Op::REF_IS_NULL);
                        let have_direct = self.emit_jump(Op::BR_IF_FALSE);
                        let lookup = self.str_const("__vybe_js_get_method");
                        self.emit_u16(Op::GLOBAL_GET, lookup);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_const(Value::String(Arc::from(field_name.as_str())));
                        self.emit_u8(Op::CALL_REF, 2);
                        let end_lookup = self.emit_jump(Op::BR);
                        self.patch_jump(have_direct);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.patch_jump(end_lookup);
                        self.patch_jump(end);
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
                        let non_null = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_UNDEFINED);
                        let is_null = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_const(Value::String(Arc::from("Cannot read properties of undefined")));
                        let have_msg = self.emit_jump(Op::BR);
                        self.patch_jump(is_null);
                        self.emit_const(Value::String(Arc::from("Cannot read properties of null")));
                        self.patch_jump(have_msg);
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        let line = self.line;
                        common::errors::emit_throw(self.chunk(), line);
                        self.patch_jump(non_null);
                        let field_name = self.canon(field);
                        let prop = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let val_slot = self.define_local("__js_member_val");
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.emit(Op::REF_IS_NULL);
                        let have_direct = self.emit_jump(Op::BR_IF_FALSE);
                        let lookup = self.str_const("__vybe_js_get_method");
                        self.emit_u16(Op::GLOBAL_GET, lookup);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_const(Value::String(Arc::from(field_name.as_str())));
                        self.emit_u8(Op::CALL_REF, 2);
                        let end = self.emit_jump(Op::BR);
                        self.patch_jump(have_direct);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.patch_jump(end);
                        // Restore the caller's __js_this — value already
                        // on stack as the access result.
                        let result_slot = self.define_local("__js_member_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        self.restore_js_this(saved_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    return Ok(());
                }

                let receiver_type_hint = match &object.kind {
                    ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
                    _ => self.infer_expr_type_hint(object),
                };

                let receiver_is_nullable = receiver_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim().ends_with('?'));

                let receiver_array_rank = if matches!(self.profile.name.as_str(), "csharp" | "vb")
                    && field == "Rank"
                {
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
                            self.emit(Op::DYN_NOT);
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

                let receiver_is_collection_like = if matches!(self.profile.name.as_str(), "csharp" | "vb")
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
                        ExprKind::Ident(_) | ExprKind::New { .. } | ExprKind::Call { .. } => receiver_type_hint
                            .as_deref()
                            .map(is_collection_like_type)
                            .unwrap_or(unknown_receiver_default),
                        ExprKind::Lit(Literal::Str(_)) | ExprKind::Interpolation(_) | ExprKind::Array(_) => true,
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
                        common::collections::emit_iter_keys(&mut self.chunks, self.current, self.line);
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
                let is_csharp_runtime_count_accessor = matches!(self.profile.name.as_str(), "csharp" | "vb")
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
                    let non_null = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(non_null);
                    common::collections::emit_len(&mut self.chunks, self.current, self.line);
                    self.patch_jump(end);
                    return Ok(());
                } else {
                    self.compile_expr(object)?;
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
                    self.emit(Op::DYN_EQ);
                    let return_value = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    let done = self.emit_jump(Op::BR);

                    self.patch_jump(return_value);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.patch_jump(done);
                    return Ok(());
                }

                let static_field_owner = if matches!(self.profile.name.as_str(), "csharp" | "vb") {
                    receiver_type_hint.as_deref().and_then(|type_hint| {
                        let trimmed_type_hint = type_hint.trim().trim_end_matches("()").trim();
                        let metadata_type_hint = self
                            .reflection_type_metadata(type_hint)
                            .map(|_| type_hint)
                            .or_else(|| self.reflection_type_metadata(trimmed_type_hint).map(|_| trimmed_type_hint))?;
                        self.reflection_type_metadata(metadata_type_hint)
                            .and_then(|meta| {
                                meta.fields
                                    .iter()
                                    .find(|(name, field_meta)| name.eq_ignore_ascii_case(field) && field_meta.is_static)
                            })
                            .map(|_| {
                                let short_name = self.reflection_type_short_name(metadata_type_hint);
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
                        let has_instance_value = self.emit_jump(Op::BR_IF_FALSE);

                        let class_idx = self.str_const(&self.canon(&type_name));
                        self.emit_u16(Op::GLOBAL_GET, class_idx);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        let done = self.emit_jump(Op::BR);

                        self.patch_jump(has_instance_value);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.patch_jump(done);
                        return Ok(());
                    }
                }

                if *null_safe && matches!(self.profile.name.as_str(), "csharp" | "vb") && !is_csharp_len_accessor {
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let non_null = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(non_null);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                    self.patch_jump(end);
                } else if *null_safe {
                    // ?. — null-safe access. PHP `?->` short-circuits to
                    // null whenever the receiver isn't an Object. That
                    // also catches the "called a method on a string"
                    // case (test surface assertion: `"hello"?->length`
                    // === null). REF_IS_OBJECT is true only for actual
                    // Object values, false for null / undefined / string
                    // / number / bool — exactly the right discriminant.
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_OBJECT);
                    let skip = self.emit_jump(Op::BR_IF_TRUE);
                    // Receiver is not an object — drop it, push null.
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(skip);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                    self.patch_jump(end);
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
                        let has_canonical_value = self.emit_jump(Op::BR_IF_FALSE);

                        let exact_idx = self.str_const(field);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::STRUCT_GET, exact_idx);
                        let done = self.emit_jump(Op::BR);

                        self.patch_jump(has_canonical_value);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.patch_jump(done);
                    } else {
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                }
            }

            // ── Index access ────────────────────────────────────────────
            ExprKind::Index { object, index, null_safe } => {
                // A Range used as the index is a slice operation
                // (C# `arr[1..3]` / `s[0..5]`, Python `arr[1:3]` / `s[0:5]`).
                // Route through compiler_common's polymorphic slice helper so
                // strings and arrays both work uniformly.
                if let ExprKind::Range { start, end, .. } = &index.kind {
                    let line = self.line;
                    common::collections::emit_slice_push_func(self.chunk(), line);
                    self.compile_expr(object)?;
                    self.compile_expr(start)?;
                    self.compile_expr(end)?;
                    common::collections::emit_slice_invoke(self.chunk(), line);
                } else if let ExprKind::Slice { lower, upper, step } = &index.kind {
                    self.compile_expr(object)?;
                    let line = self.line;
                    if step.is_none() {
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
                        common::collections::emit_stdlib_call(
                            &mut self.chunks, self.current, "__vybe_slice", 3, line,
                        );
                    } else {
                        let step_const = step.as_ref().and_then(|expr| match &expr.kind {
                            ExprKind::Lit(Literal::Int(n)) => Some(*n),
                            ExprKind::Unary { op: UnaryOp::Neg, expr } => match &expr.kind {
                                ExprKind::Lit(Literal::Int(n)) => Some(-*n),
                                _ => None,
                            },
                            _ => None,
                        });

                        if lower.is_none() && upper.is_none() {
                            if step_const == Some(-1) {
                                self.emit(Op::DUP);
                                self.emit(Op::REF_IS_STRING);
                                let non_string = self.emit_jump(Op::BR_IF_FALSE);
                                self.emit(Op::STR_REVERSE);
                                let end = self.emit_jump(Op::BR);
                                self.patch_jump(non_string);
                                self.emit(Op::NULL);
                                self.emit(Op::NULL);
                                if let Some(s) = step { self.compile_expr(s)?; } else { self.emit(Op::NULL); }
                                common::collections::emit_stdlib_call(
                                    &mut self.chunks, self.current, "__vybe_slicestep", 4, line,
                                );
                                self.patch_jump(end);
                                return Ok(());
                            }

                            if let Some(step_value) = step_const.filter(|n| *n > 1) {
                                self.emit(Op::DUP);
                                self.emit(Op::REF_IS_STRING);
                                let non_string = self.emit_jump(Op::BR_IF_FALSE);

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
                                self.emit(Op::DYN_LT);
                                self.emit(Op::DYN_NOT);
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
                                let end = self.emit_jump(Op::BR);

                                self.patch_jump(non_string);
                                self.emit(Op::NULL);
                                self.emit(Op::NULL);
                                if let Some(s) = step { self.compile_expr(s)?; } else { self.emit(Op::NULL); }
                                common::collections::emit_stdlib_call(
                                    &mut self.chunks, self.current, "__vybe_slicestep", 4, line,
                                );
                                self.patch_jump(end);
                                return Ok(());
                            }
                        }

                        if let Some(l) = lower { self.compile_expr(l)?; } else { self.emit(Op::NULL); }
                        if let Some(u) = upper { self.compile_expr(u)?; } else { self.emit(Op::NULL); }
                        if let Some(s) = step { self.compile_expr(s)?; } else { self.emit(Op::NULL); }
                        common::collections::emit_stdlib_call(
                            &mut self.chunks, self.current, "__vybe_slicestep", 4, line,
                        );
                    }
                } else if self.profile.name == "pascal" && self.expr_is_known_string_receiver(object) {
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
                        &mut self.chunks, self.current, line,
                    );
                } else if self.is_js_profile() && *null_safe {
                    self.compile_expr(object)?;
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let non_null = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::DROP);
                    let line = self.line;
                    common::expressions::emit_undefined(self.chunk(), line);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(non_null);
                    let obj_slot = self.define_local("__js_index_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.compile_expr(index)?;
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    self.patch_jump(end);
                } else if matches!(self.profile.name.as_str(), "csharp" | "vb") {
                    self.compile_expr(object)?;
                    let obj_slot = self.define_local("__index_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit(Op::DROP);

                    let null_safe_end = if *null_safe {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let non_null = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit(Op::NULL);
                        let end = self.emit_jump(Op::BR);
                        self.patch_jump(non_null);
                        Some(end)
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
                    let fallback = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, getter_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.compile_collection_key(object, index)?;
                    self.emit_u8(Op::CALL_REF, 2);
                    let done = self.emit_jump(Op::BR);

                    self.patch_jump(fallback);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.compile_collection_key(object, index)?;
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    self.patch_jump(done);
                    if let Some(end) = null_safe_end {
                        self.patch_jump(end);
                    }
                } else {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
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
                        if name == "Proxy" && args.len() == 2 {
                            self.uses_proxy = true;
                            self.compile_expr(&args[0].value)?;
                            self.compile_expr(&args[1].value)?;
                            let line = self.line;
                            crate::emitter::js::proxy_adapter::emit_proxy_create(
                                &mut self.chunks, self.current, line,
                            );
                            return Ok(());
                        }
                    }
                }
                let class_parts = self.flatten_member_chain(class);
                let dotted_type_name = match &class.kind {
                    ExprKind::Ident(name) => Some(self.resolve_source_type_alias(name)),
                    ExprKind::Member { .. } if self.profile.namespaces.use_dotnet && !class_parts.is_empty() => {
                        Some(self.resolve_source_type_alias(&class_parts.join(".")))
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
                        let idx = self.str_const(&ctor_global);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);
                        return Ok(());
                    }
                    // Nested class: `new Outer.Inner()` — the inner type
                    // is registered as a sibling global per ECMA-334 §15.3.
                    // Try the last segment as a type name when the full
                    // dotted form misses.
                    if class_parts.len() > 1 {
                        let last = class_parts.last().unwrap();
                        let canon_last = self.canon(last);
                        if self.defined_classes.contains(&canon_last) {
                            let idx = self.str_const(&canon_last);
                            self.emit_u16(Op::GLOBAL_GET, idx);
                            for a in args { self.compile_expr(&a.value)?; }
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
                            common::threading::emit_thread_new(&mut self.chunks, self.current, line);
                            return Ok(());
                        }
                        "task" => {
                            // New Task(callback) → cont_new only
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(&mut self.chunks, self.current, line);
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
                        for a in args { self.compile_expr(&a.value)?; }
                        let line = self.line;
                        common::gui::emit_new_control(self.chunk(), new_idx, args.len() as u8, line);
                        return Ok(());
                    }
                    // Dotnet component descriptor constructors — fallback after
                    // GUI so .NET-only types like Dictionary still work.
                    let dotnet_constructor = if !dotnet_ctor_registered {
                        common::dotnet::surface().lookup_constructor(bare_str)
                    } else {
                        None
                    };
                    if let Some(target) = dotnet_constructor.clone() {
                        for a in args { self.compile_expr(&a.value)?; }
                        // Proper-case class name (preserve from source) for
                        // the __type stamp. Some host fns (e.g.
                        // `vybe:types/collectionPeek`) compare __type with
                        // exact case, so a lowercase stamp would clobber.
                        let proper_name: String = type_name
                            .split('(').next().unwrap_or(type_name).trim()
                            .rsplit('.').next().unwrap_or(type_name)
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
                        // Stack: [obj] → [obj, name] → [obj]
                        self.emit_const(Value::String(Arc::from(proper_name.as_str())));
                        let stamp_idx = self.import("vybe:types", "__stamp_type");
                        self.emit_host_call(stamp_idx, 2);

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
                        if let Some((module, func)) = self.profile.lookup_known_type(type_name).map(|(m, f)| (m.to_string(), f.to_string())) {
                            for a in args { self.compile_expr(&a.value)?; }
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
                        let ctor_idx = self.str_const(&ctor_name);
                        self.emit_u16(Op::GLOBAL_GET, ctor_idx);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);

                        if bare_str.eq_ignore_ascii_case("list") || bare_str.eq_ignore_ascii_case("arraylist") {
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
                }
                if self.is_js_profile() {
                    self.compile_expr(class)?;
                    let ctor_slot = self.define_local("__js_ctor");
                    self.emit_u16(Op::LOCAL_SET, ctor_slot); self.emit(Op::DROP);
                    let line = self.line;
                    self.emit_common("object.new", 0, line);

                    let instance_slot = self.define_local("__js_instance");
                    self.emit_u16(Op::LOCAL_SET, instance_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, ctor_slot);
                    let proto_key = self.str_const("prototype");
                    self.emit_u16(Op::STRUCT_GET, proto_key);
                    let proto_slot = self.define_local("__js_proto");
                    self.emit_u16(Op::LOCAL_SET, proto_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, proto_slot);
                    self.emit(Op::REF_IS_NULL);
                    let skip_proto = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit_u16(Op::LOCAL_GET, instance_slot);
                    self.emit_u16(Op::LOCAL_GET, proto_slot);
                    let proto_link = self.str_const("__proto__");
                    self.emit_u16(Op::STRUCT_SET, proto_link);
                    self.emit(Op::DROP);
                    self.patch_jump(skip_proto);

                    let saved_js_this = self.save_js_this("__js_prev_this_new");
                    self.emit_u16(Op::LOCAL_GET, instance_slot);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, ctor_slot);
                    for a in args { self.compile_expr(&a.value)?; }
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                    let result_slot = self.define_local("__js_ctor_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_js_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit(Op::REF_IS_NULL);
                    let use_instance = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(use_instance);
                    self.emit_u16(Op::LOCAL_GET, instance_slot);
                    self.patch_jump(end);
                    return Ok(());
                }

                // User-defined class constructor
                self.compile_expr(class)?;
                for a in args { self.compile_expr(&a.value)?; }
                self.emit_u8(Op::CALL_REF, args.len() as u8);
            }

            // ── Assignment as expression ────────────────────────────────
            ExprKind::Assign { target, value } => {
                self.compile_expr(value)?;
                self.emit(Op::DUP);
                self.compile_assign_target(target)?;
            }

            // ── Lambda ──────────────────────────────────────────────────
            ExprKind::Lambda { params, body, .. } => {
                self.compile_lambda(params, body)?;
            }

            // ── Array literal ───────────────────────────────────────────
            ExprKind::Array(elements) => {
                // Array literals funnel through `common::collections` so
                // every language and every array-literal site emits the
                // same import shape. Changing the provider (ecma:array
                // → vybe:array → polyfill) happens in ONE file, not here.
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
                let has_keys = elements.iter().any(|e| e.key.is_some());

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
                        self.emit(Op::DUP);              // [map, map]
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
                                    ExprKind::Lit(crate::ast::Literal::Float(n)) if n.fract() == 0.0 => {
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
                    // All-unkeyed: use the array path (fast, small).
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                    for elem in elements {
                        if elem.spread {
                            // Spread: `concat(current, other)` returns a NEW
                            // array which replaces the one on TOS. JS
                            // generators (Continuation values) can't be
                            // spread by the host concat fn — the iterator
                            // protocol needs Op::GEN_NEXT (WASM stack
                            // switching). Drain via the stdlib helper
                            // when isGenerator(elem) at runtime.
                            self.compile_expr(&elem.value)?;
                            if self.is_js_profile() {
                                let v_slot = self.define_local("__arr_spread_v");
                                self.emit_u16(Op::LOCAL_SET, v_slot); self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, v_slot);
                                let is_gen_idx = self.import("ecma:value", "isGenerator");
                                self.emit_host_call(is_gen_idx, 1);
                                let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                                let drain_key = self.str_const("__vybe_drain_generator");
                                self.emit_u16(Op::GLOBAL_GET, drain_key);
                                self.emit_u16(Op::LOCAL_GET, v_slot);
                                self.emit_u8(Op::CALL_REF, 1);
                                let done = self.emit_jump(Op::BR);
                                self.patch_jump(not_gen);
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
                                self.patch_jump(done);
                            }
                            common::collections::emit_concat(&mut self.chunks, self.current, line);
                        } else {
                            // DUP keeps the array on TOS; push returns the
                            // new length, which we drop.
                            self.emit(Op::DUP);
                            self.compile_expr(&elem.value)?;
                            common::collections::emit_push(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                        }
                    }
                }
            }

            // ── Tuple (Python) ──────────────────────────────────────────
            ExprKind::Tuple(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements { self.compile_expr(elem)?; }
                // Allocate N consecutive slots; common::collections::emit_pack_n
                // stashes stack values and re-pushes into a fresh array —
                // same ecma:array.* surface as literals.
                let base = if n == 0 { 0 } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.define_local("__pack");
                        if i == 0 { first = s; }
                    }
                    first
                };
                common::collections::emit_pack_n(&mut self.chunks, self.current, n as u16, base, line);
            }

            // ── Set (Python) ────────────────────────────────────────────
            ExprKind::Set(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements { self.compile_expr(elem)?; }
                let base = if n == 0 { 0 } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.define_local("__pack");
                        if i == 0 { first = s; }
                    }
                    first
                };
                common::collections::emit_pack_n(&mut self.chunks, self.current, n as u16, base, line);
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
                                self.emit(Op::DUP);
                                self.compile_expr(value)?;
                                let idx = self.str_const(k);
                                self.emit_u16(Op::STRUCT_SET, idx);
                                self.emit(Op::DROP);
                                // Track key in __keys
                                self.emit(Op::DUP);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_const(Value::String(Arc::from(k.as_str())));
                                let l = self.line;
                                common::collections::emit_push(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP);
                            } else {
                                // Dynamic key — emit_set is
                                // `ecma:array.set(obj, key, value) → null`
                                // so we must push key BEFORE value. The
                                // previous impl pushed value then key,
                                // causing `{ 1: "one" }` to be stored
                                // under key "one" with value 1. Fix
                                // matches the canonical emit_set contract.
                                self.emit(Op::DUP);                // [dict, dict]
                                self.compile_expr(key)?;           // [dict, dict, key]
                                self.emit(Op::DUP);                // [dict, dict, key, key]
                                let key_tmp = self.define_local("__obj_dyn_key");
                                self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                                // [dict, dict, key]
                                self.compile_expr(value)?;         // [dict, dict, key, value]
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
                            if let StmtKind::FunctionDecl { params, body, .. } = &value.kind {
                                if self.is_js_profile() {
                                    self.compile_lambda(params, &LambdaBody::Block(body.clone()))?;
                                } else {
                                    // Object methods receive `this` as implicit first arg
                                    let mut method_params = vec![Param {
                                        name: self.profile.self_keyword.clone(),
                                        type_hint: None, default: None,
                                        pass_by: PassBy::Value, is_rest: false,
                                        is_kwargs: false, is_optional: false, is_nullable: false,
                                    }];
                                    method_params.extend(params.iter().cloned());
                                    self.compile_lambda(&method_params, &LambdaBody::Block(body.clone()))?;
                                }
                            } else {
                                self.emit(Op::NULL);
                            }
                            let idx = self.str_const(key);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Accessor { kind, key, value } => {
                            self.emit(Op::DUP);
                            if let StmtKind::FunctionDecl { params, body, .. } = &value.kind {
                                if self.is_js_profile() {
                                    self.compile_lambda(params, &LambdaBody::Block(body.clone()))?;
                                } else {
                                    // Accessors receive `this` as first arg
                                    let mut accessor_params = vec![Param {
                                        name: self.profile.self_keyword.clone(),
                                        type_hint: None, default: None,
                                        pass_by: PassBy::Value, is_rest: false,
                                        is_kwargs: false, is_optional: false, is_nullable: false,
                                    }];
                                    accessor_params.extend(params.iter().cloned());
                                    self.compile_lambda(&accessor_params, &LambdaBody::Block(body.clone()))?;
                                }
                            } else {
                                self.emit(Op::NULL);
                            }
                            let accessor_name = match kind {
                                AccessorKind::Get => format!("__get_{}", key),
                                AccessorKind::Set => format!("__set_{}", key),
                            };
                            let idx = self.str_const(&accessor_name);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Computed { key, value } => {
                            // ecma:array.set expects [obj, key, val] → null
                            self.emit(Op::DUP);
                            self.compile_expr(key)?;
                            self.emit(Op::DUP); // save key for trackKey
                            let key_tmp = self.define_local("__obj_comp_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
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
                let tostring_global = self.str_const("__vybe_tostring");
                self.emit_const(Value::String(Arc::from("")));
                let acc_slot = self.define_local("__interp_acc");
                self.emit_u16(Op::LOCAL_SET, acc_slot);
                self.emit(Op::DROP);
                let part_slot = self.define_local("__interp_part");

                for part in parts.iter() {
                    match part {
                        InterpolPart::Text(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
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
                                self.emit_u16(Op::GLOBAL_GET, tostring_global);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                self.emit_u8(Op::CALL_REF, 1);
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
            ExprKind::IsType { expr: inner, type_name } => {
                let canon_type = self.canon(type_name);
                if matches!(canon_type.as_str(), "IEnumerable" | "ICollection" | "IList" | "IReadOnlyCollection" | "IReadOnlyList") {
                    if let ExprKind::Ident(name) = &inner.kind {
                        if self.lookup_var_type_hint(name).is_some_and(|hint| {
                            let bare = hint.split('<').next().unwrap_or(hint).trim();
                            matches!(Self::normalize_type_hint(bare).as_str(), "list" | "arraylist" | "queue" | "stack" | "hashset" | "dictionary")
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

                if self.profile.name == "vb" {
                    match canon_type.as_str() {
                        "string" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_STRING);
                            return Ok(());
                        }
                        "boolean" | "bool" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_BOOL);
                            return Ok(());
                        }
                        "integer" | "int" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_NUMBER);
                            let not_number = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::DUP);
                            self.emit(Op::F64_TRUNC);
                            self.emit(Op::DYN_EQ);
                            let done = self.emit_jump(Op::BR);
                            self.patch_jump(not_number);
                            self.emit(Op::FALSE);
                            self.patch_jump(done);
                            return Ok(());
                        }
                        "double" | "single" | "decimal" | "float" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_NUMBER);
                            let not_number = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::DUP);
                            self.emit(Op::F64_TRUNC);
                            self.emit(Op::DYN_EQ);
                            let is_integer = self.emit_jump(Op::BR_IF_TRUE);
                            self.emit(Op::TRUE);
                            let done = self.emit_jump(Op::BR);
                            self.patch_jump(is_integer);
                            self.emit(Op::FALSE);
                            let after_integer = self.emit_jump(Op::BR);
                            self.patch_jump(not_number);
                            self.emit(Op::FALSE);
                            self.patch_jump(done);
                            self.patch_jump(after_integer);
                            return Ok(());
                        }
                        "object" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_STRING);
                            let is_string = self.emit_jump(Op::BR_IF_TRUE);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_ARRAY);
                            let is_array = self.emit_jump(Op::BR_IF_TRUE);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit(Op::REF_IS_OBJECT);
                            let done = self.emit_jump(Op::BR);
                            self.patch_jump(is_string);
                            self.emit(Op::TRUE);
                            let end_true = self.emit_jump(Op::BR);
                            self.patch_jump(is_array);
                            self.emit(Op::TRUE);
                            self.patch_jump(end_true);
                            self.patch_jump(done);
                            return Ok(());
                        }
                        _ => {}
                    }
                }

                let line = self.line;
                let type_idx = self.chunk().add_constant(
                    vybe_bytecode::Value::String(std::sync::Arc::from(canon_type.as_str())),
                );
                let mut match_patches = Vec::new();

                self.emit_u16(Op::LOCAL_GET, obj_slot);
                self.chunk().emit_op_u16(vybe_bytecode::Op::REF_TEST, type_idx, line);
                match_patches.push(self.emit_jump(Op::BR_IF_TRUE));

                self.emit_u16(Op::LOCAL_GET, obj_slot);
                let type_key = self.str_const("__type");
                self.emit_u16(Op::STRUCT_GET, type_key);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                self.emit(Op::DYN_EQ);
                match_patches.push(self.emit_jump(Op::BR_IF_TRUE));

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
                    let candidate_idx = self.chunk().add_constant(
                        vybe_bytecode::Value::String(std::sync::Arc::from(candidate.as_str())),
                    );
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.chunk().emit_op_u16(vybe_bytecode::Op::REF_TEST, candidate_idx, line);
                    match_patches.push(self.emit_jump(Op::BR_IF_TRUE));

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let type_key = self.str_const("__type");
                    self.emit_u16(Op::STRUCT_GET, type_key);
                    self.emit_const(Value::String(Arc::from(candidate.as_str())));
                    self.emit(Op::DYN_EQ);
                    match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                }

                if matches!(canon_type.as_str(), "IEnumerable" | "ICollection" | "IList" | "IReadOnlyCollection" | "IReadOnlyList") {
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_ARRAY);
                    match_patches.push(self.emit_jump(Op::BR_IF_TRUE));

                    let list_idx = self.chunk().add_constant(
                        vybe_bytecode::Value::String(std::sync::Arc::from("list")),
                    );
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.chunk().emit_op_u16(vybe_bytecode::Op::REF_TEST, list_idx, line);
                    match_patches.push(self.emit_jump(Op::BR_IF_TRUE));

                    for key_name in ["length", "count"] {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        let key = self.str_const(key_name);
                        self.emit_u16(Op::STRUCT_GET, key);
                        self.emit(Op::REF_IS_NULL);
                        match_patches.push(self.emit_jump(Op::BR_IF_FALSE));
                    }

                    if canon_type == "IEnumerable" {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_OBJECT);
                        match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                    }
                }

                self.emit_u16(Op::LOCAL_GET, obj_slot);
                let types_key = self.str_const("__types");
                self.emit_u16(Op::STRUCT_GET, types_key);
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let has_types = self.emit_jump(Op::BR_IF_FALSE);
                self.emit(Op::DROP);
                self.emit(Op::FALSE);
                let done = self.emit_jump(Op::BR);

                self.patch_jump(has_types);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                common::collections::emit_contains(&mut self.chunks, self.current, line);
                match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                for candidate in &reflection_matches {
                    self.emit(Op::DUP);
                    self.emit_const(Value::String(Arc::from(candidate.as_str())));
                    common::collections::emit_contains(&mut self.chunks, self.current, line);
                    match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                }
                self.emit(Op::FALSE);
                let end_false = self.emit_jump(Op::BR);

                for patch in match_patches {
                    self.patch_jump(patch);
                }
                self.emit(Op::TRUE);
                self.patch_jump(done);
                self.patch_jump(end_false);
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

                    let shadows_cast = self.defined_functions.contains(&canon_type)
                        || (!self.case_sensitive
                            && self.defined_functions.iter().any(|name| name.eq_ignore_ascii_case(type_name)));
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
                    "int" | "long" | "short" | "byte" | "uint" | "ulong" | "ushort"
                    | "sbyte" | "double" | "float" | "decimal" => {
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
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                self.emit(Op::DROP);
                self.compile_expr(right)?;
                self.patch_jump(skip);
            }

            // ── Spread ──────────────────────────────────────────────────
            ExprKind::Spread(inner) => {
                self.compile_expr(inner)?;
                // SPREAD is array-only in the VM (matches WASM
                // `array.copy_into` semantics). Iterables that aren't
                // arrays — Set, Map, String — get coerced to an array
                // first via the polymorphic Symbol.iterator helper
                // (`ecma:object.iterForOf`). Generators (Continuation
                // values) need WASM stack-switching (`Op::GEN_NEXT`)
                // to drive their iterator protocol, which a host fn
                // can't do — route them through the
                // `__stdlib_drain_generator` bytecode helper.
                let inner_slot = self.define_local("__spread_iter");
                self.emit_u16(Op::LOCAL_SET, inner_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, inner_slot);
                let is_gen_idx = self.import("ecma:value", "isGenerator");
                self.emit_host_call(is_gen_idx, 1);
                let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                let drain_key = self.str_const("__vybe_drain_generator");
                self.emit_u16(Op::GLOBAL_GET, drain_key);
                self.emit_u16(Op::LOCAL_GET, inner_slot);
                self.emit_u8(Op::CALL_REF, 1);
                let done = self.emit_jump(Op::BR);
                self.patch_jump(not_gen);
                self.emit_u16(Op::LOCAL_GET, inner_slot);
                let idx = self.import("ecma:object", "iterForOf");
                self.emit_host_call(idx, 1);
                self.patch_jump(done);
                self.emit(Op::SPREAD);
            }

            // ── Await ───────────────────────────────────────────────────
            ExprKind::Await(inner) => {
                // ECMA-262 §27.2 await semantics. The synchronous-promise
                // model unwraps `__value` directly when the promise is
                // already settled. Rejected promises THROW their reason
                // (the spec semantics — `await Promise.reject(x)` throws
                // x at the await site). Pending promises hit the JSPI
                // suspend path via Op::PROMISE_SUSPEND elsewhere.
                self.compile_expr(inner)?;
                let await_slot = self.define_local("__await");
                self.emit_u16(Op::LOCAL_SET, await_slot); self.emit(Op::DROP);
                // Read __state — if "rejected" we throw __value.
                self.emit_u16(Op::LOCAL_GET, await_slot);
                let sk = self.str_const("__state");
                self.emit_u16(Op::STRUCT_GET, sk);
                self.emit_const(Value::String(Arc::from("rejected")));
                self.emit(Op::STR_EQUALS);
                let not_rejected = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, await_slot);
                let vk = self.str_const("__value");
                self.emit_u16(Op::STRUCT_GET, vk);
                self.emit(Op::THROW);
                self.patch_jump(not_rejected);
                // Fulfilled or non-promise: unwrap __value (or pass-through).
                self.emit_u16(Op::LOCAL_GET, await_slot);
                let vk = self.str_const("__value");
                self.emit_u16(Op::STRUCT_GET, vk);
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let use_original = self.emit_jump(Op::BR_IF_TRUE);
                let done = self.emit_jump(Op::BR);
                self.patch_jump(use_original);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, await_slot);
                self.patch_jump(done);
            }

            // ── Yield ───────────────────────────────────────────────────
            ExprKind::Yield(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::NULL); }
                self.emit_u16(Op::SUSPEND, 0);
            }

            ExprKind::YieldFrom(inner) => {
                // ECMA-262 §15.5 `yield*`: drain the inner iterable,
                // re-yielding each value through the enclosing
                // generator. Uses the WASM stack-switching `GEN_NEXT`
                // (pops cont → pushes value+has_more) then `SUSPEND 0`
                // for the per-value yield.
                self.compile_expr(inner)?;
                let gen_slot = self.define_local("__yield_star_gen");
                let val_slot = self.define_local("__yield_star_val");
                let has_more_slot = self.define_local("__yield_star_has_more");
                self.emit_u16(Op::LOCAL_SET, gen_slot); self.emit(Op::DROP);
                let loop_start = self.chunks[self.current].code.len();
                self.emit_u16(Op::LOCAL_GET, gen_slot);
                self.emit(Op::GEN_NEXT);
                // After GEN_NEXT: stack top is has_more (i32), under it value.
                self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, val_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                self.emit(Op::DYN_TO_BOOL);
                let exit = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, val_slot);
                self.emit_u16(Op::SUSPEND, 0);
                self.emit(Op::DROP);
                let line = self.line;
                self.chunks[self.current].emit_loop(loop_start, line);
                self.patch_jump(exit);
                // yield* expression evaluates to the inner generator's
                // return value (final return statement). We don't track
                // that yet — push `undefined` per ECMA spec default.
                self.emit(Op::UNDEFINED);
            }

            // ── AddressOf (VB) ──────────────────────────────────────────
            ExprKind::AddressOf(name) => {
                let parts: Vec<&str> = name.split('.').filter(|part| !part.is_empty()).collect();
                if parts.is_empty() {
                    self.emit(Op::NULL);
                    return Ok(());
                }

                let self_kw = self.profile.self_keyword.clone();
                let is_self_qualified = parts.first().map(|part| {
                    if self.case_sensitive {
                        *part == self_kw || *part == "Me"
                    } else {
                        part.eq_ignore_ascii_case(&self_kw) || part.eq_ignore_ascii_case("Me")
                    }
                }).unwrap_or(false);

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
                let is_ctor_call = method.is_none() || method.as_ref().map_or(false, |m| {
                    if self.case_sensitive {
                        m == &ctor_name || m == "new" || m == "__init__"
                    } else {
                        m.eq_ignore_ascii_case(&ctor_name) || m.eq_ignore_ascii_case("new") || m.eq_ignore_ascii_case("__init__")
                    }
                });

                if is_ctor_call {
                    // super() / MyBase.New(args) → call parent constructor
                    if let Some(ref class_name) = self.current_class.clone() {
                        if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|c| c.parent.clone()) {
                            if common::errors::is_exception_type(&parent_name) {
                                let arg_exprs: Vec<&Expression> = args.iter().map(|arg| &arg.value).collect();
                                self.emit_js_exception_ctor_value(&parent_name, &arg_exprs)?;
                                if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                                    self.emit(Op::DUP);
                                    self.emit_u16(Op::LOCAL_SET, slot);
                                    self.emit(Op::DROP);
                                }
                                return Ok(());
                            }
                            let pname = self.canon(&parent_name);
                            let pidx = self.str_const(&pname);
                            self.emit_u16(Op::GLOBAL_GET, pidx);
                            for a in args { self.compile_expr(&a.value)?; }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
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
                    let parent_name = self.current_class.as_ref()
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
                            if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                                self.emit_u16(Op::LOCAL_GET, self_slot);
                            } else {
                                let js_this = self.str_const("__js_this");
                                self.emit_u16(Op::GLOBAL_GET, js_this);
                            }
                            self.set_js_this_from_stack();
                            for a in args { self.compile_expr(&a.value)?; }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            let result_slot = self.define_local("__js_super_expr_result");
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.emit(Op::DROP);
                            self.restore_js_this(saved_js_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                            self.emit_u16(Op::LOCAL_GET, self_slot);
                            for a in args { self.compile_expr(&a.value)?; }
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
            ExprKind::Comprehension { kind: _, element, generators } => {
                // Simplified: compile as loop building an array
                let line = self.line;
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                let result_slot = self.define_local("__comp_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);

                // Only handle the first generator for simplicity
                if let Some(generator) = generators.first() {
                    self.compile_expr(&generator.iter)?;
                    let arr_slot = self.define_local("__comp_iter");
                    self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                    let idx_slot = self.define_local("__comp_idx");
                    let lp = common::loops::emit_for_in_start(
                        &mut self.chunks, self.current, arr_slot, idx_slot, line,
                    );
                    // Bind loop var
                    let var_name = match &generator.target.kind {
                        ExprKind::Ident(n) => n.clone(),
                        _ => "__comp_var".to_string(),
                    };
                    let var_slot = self.define_local(&var_name);
                    self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);

                    // Check conditions
                    let mut cond_skip = None;
                    for cond_expr in &generator.conditions {
                        self.compile_expr(cond_expr)?;
                        self.emit(Op::DYN_TO_BOOL);
                        cond_skip = Some(self.emit_jump(Op::BR_IF_FALSE));
                    }

                    // Push element via ecma:array.push.
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.compile_expr(element)?;
                    let l = self.line;
                    common::collections::emit_push(&mut self.chunks, self.current, l);
                    self.emit(Op::DROP);

                    if let Some(skip) = cond_skip { self.patch_jump(skip); }

                    common::loops::emit_for_in_end(
                        &mut self.chunks, self.current, idx_slot, lp, line,
                    );
                }

                self.emit_u16(Op::LOCAL_GET, result_slot);
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
                    common::collections::emit_stdlib_call(
                        &mut self.chunks, self.current, "__vybe_slice", 3, line,
                    );
                } else {
                    // Emit slice parts → [obj, lower, upper, step] then call the
                    // bundled `__vybe_slicestep` polyfill directly (skips the
                    // legacy `vybe:array` host-import indirection).
                    if let Some(l) = lower { self.compile_expr(l)?; } else { self.emit(Op::NULL); }
                    if let Some(u) = upper { self.compile_expr(u)?; } else { self.emit(Op::NULL); }
                    if let Some(s) = step { self.compile_expr(s)?; } else { self.emit(Op::NULL); }
                    common::collections::emit_stdlib_call(
                        &mut self.chunks, self.current, "__vybe_slicestep", 4, line,
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
                // delete obj.prop → call vybe:object::deleteProperty(obj, key)
                // which removes the property and returns true.
                if let ExprKind::Member { object, field, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    let idx = self.import("ecma:object", "delete");
                    self.emit_host_call(idx, 2);
                } else if let ExprKind::Index { object, index, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    let idx = self.import("ecma:object", "delete");
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
                    if i < exprs.len() - 1 { self.emit(Op::DROP); }
                }
            }

            // ── ClassExpr (JS) ──────────────────────────────────────────
            ExprKind::ClassExpr { name, parent, members } => {
                let class_name = name.clone().unwrap_or_else(|| "__anonymous_class".to_string());
                let parent_name: Option<String> = if let Some(p) = parent {
                    if let ExprKind::Ident(n) = &p.kind { Some(n.clone()) } else { None }
                } else { None };
                let class_name = self.canon(&class_name);
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
                if let StmtKind::FunctionDecl { name, params, return_type, body, is_sub, is_generator, handles, is_async, .. } = &stmt.kind {
                    let fn_name = if name.is_empty() {
                        format!("__anon_fn_{}", self.chunks.len())
                    } else {
                        name.clone()
                    };
                    self.compile_function_decl(&fn_name, params, return_type, body, *is_sub, *is_generator, handles, *is_async)?;
                    self.emit_var_get(&fn_name);
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── Range ───────────────────────────────────────────────────
            ExprKind::Range { start, end, inclusive: _ } => {
                self.compile_expr(start)?;
                self.compile_expr(end)?;
                let line = self.line;
                common::collections::emit_range(&mut self.chunks, self.current, 2, line);
            }

            // ── StaticAccess (PHP) ──────────────────────────────────────
            ExprKind::StaticAccess { class, member } => {
                if let (ExprKind::Ident(class_name), ExprKind::Ident(member_name)) = (&class.kind, &member.kind) {
                    if let Some(value) = self.enum_member_ordinal(class_name, member_name) {
                        self.emit_const(Value::F64(value as f64));
                        return Ok(());
                    }
                }

                // class::member → look up class, then get static member
                self.compile_expr(class)?;
                if let ExprKind::Ident(name) = &member.kind {
                    let field_name = self.canon(name);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                } else {
                    self.compile_expr(member)?;
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                }
            }

            // ── Match expression (PHP/Rust) ─────────────────────────────
            ExprKind::Match { subject, arms } => {
                self.compile_expr(subject)?;
                let subject_slot = self.define_local("__match_subj");
                self.emit_u16(Op::LOCAL_SET, subject_slot); self.emit(Op::DROP);
                let mut end_patches = Vec::new();
                for arm in arms {
                    if let Some(ref conditions) = arm.conditions {
                        let mut match_patches = Vec::new();
                        for c in conditions {
                            self.emit_u16(Op::LOCAL_GET, subject_slot);
                            self.compile_expr(c)?;
                            self.emit(Op::DYN_EQ);
                            match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                        }
                        let skip = self.emit_jump(Op::BR);
                        for p in match_patches { self.patch_jump(p); }
                        self.compile_expr(&arm.body)?;
                        end_patches.push(self.emit_jump(Op::BR));
                        self.patch_jump(skip);
                    } else {
                        // Default arm
                        self.compile_expr(&arm.body)?;
                        end_patches.push(self.emit_jump(Op::BR));
                    }
                }
                // If no arm matched, null
                self.emit(Op::NULL);
                for p in end_patches { self.patch_jump(p); }
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
            self.emit(Op::DYN_NOT);
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
        self.emit_u16(Op::LOCAL_SET, rhs_slot); self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_SET, lhs_slot); self.emit(Op::DROP);

        let size_key = self.str_const("size");
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::STRUCT_GET, size_key);
        self.emit(Op::REF_IS_NULL);
        let fallback = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit_u16(Op::STRUCT_GET, size_key);
        self.emit(Op::REF_IS_NULL);
        let rhs_fallback = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        let idx = self.import("ecma:set", helper);
        self.emit_host_call(idx, 2);
        let end = self.emit_jump(Op::BR);

        self.patch_jump(fallback);
        self.patch_jump(rhs_fallback);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.compile_binop(op);
        self.patch_jump(end);
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
            BinOp::NotEq if self.pending_class_has_method_name_for_type(&left_type, "op_Inequality") => {
                ("op_Inequality", false)
            }
            BinOp::NotEq if self.pending_class_has_method_name_for_type(&left_type, "op_Equality") => {
                ("op_Equality", true)
            }
            _ => return Ok(false),
        };

        if let Some(chunk_idx) = self.resolve_static_method_overload_chunk_for_type(&left_type, method_name, &arg_exprs) {
            self.emit_direct_static_method_call(chunk_idx, &arg_exprs)?;
            if negate_result {
                self.emit(Op::DYN_NOT);
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
            self.emit(Op::DYN_NOT);
        }
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

        let bare_type = left_type.split('<').next().unwrap_or(left_type.as_str()).trim();
        let canon_type = self.canon(bare_type);
        if !self.defined_globals.contains(&canon_type) {
            return None;
        }
        if !self.defined_class_methods.contains(&self.canon(method_name)) {
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
                        ExprKind::Ident(name) if name.eq_ignore_ascii_case("Version") => Some("Version".into()),
                        ExprKind::Member { field, .. } if field.eq_ignore_ascii_case("Version") => Some("Version".into()),
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
            || type_name.rsplit('.').next().is_some_and(|short| short.eq_ignore_ascii_case(expected))
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
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") => Some("dotnet.timespan_add"),
            BinOp::Sub
                if Self::is_dotnet_type_name(&left_type, "TimeSpan")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") => Some("dotnet.timespan_sub"),
            BinOp::Add
                if Self::is_dotnet_type_name(&left_type, "DateTime")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") => Some("dotnet.datetime_add_timespan"),
            BinOp::Sub
                if Self::is_dotnet_type_name(&left_type, "DateTime")
                    && Self::is_dotnet_type_name(&right_type, "DateTime") => Some("dotnet.datetime_subtract_datetime"),
            BinOp::Lt
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") => Some("dotnet.version_lt"),
            BinOp::Gt
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") => Some("dotnet.version_gt"),
            BinOp::Eq
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") => Some("dotnet.version_eq"),
            BinOp::NotEq
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") => Some("dotnet.version_ne"),
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
                .map(|ch| if ch.is_ascii_alphanumeric() { ch.to_ascii_lowercase() } else { '_' })
                .collect::<String>()
        };
        format!(
            "__pascal_helper_{}_{}",
            sanitize(type_name),
            sanitize(method_name),
        )
    }

}
