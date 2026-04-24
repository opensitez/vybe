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
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());

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
                    } else {
                        self.emit(Op::NULL);
                    }
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
                                // JS `+v` coerces to number. Route through vybe:convert:toNumber.
                                let idx = self.import("vybe:convert", "toNumber");
                                self.emit_host_call(idx, 1);
                            }
                            UnaryOp::Not => self.emit(Op::DYN_NOT),
                            UnaryOp::BitNot => { let l = self.line; common::expressions::emit_i32_not(self.chunk(), l); }
                            UnaryOp::Typeof => self.emit(Op::REF_TYPEOF),
                            UnaryOp::Void => { self.emit(Op::DROP); self.emit(Op::NULL); }
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
                self.emit(Op::DYN_TO_BOOL);
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
                    self.compile_call(callee, args)?;
                    // Multi-value result repack: when the callee is one of
                    // the pre-scanned multi-return functions, CALL leaves
                    // N values on the stack. A destructure-assign consumes
                    // them directly (see `detect_multi_value_receive` in
                    // compiler/mod.rs, which bypasses this branch); every
                    // other use site — `r = f()`, `print(f())`, `f() + g()`
                    // — expects a single value, so we re-pack here.
                    if let ExprKind::Ident(name) = &callee.kind {
                        let cname = self.canon(name);
                        if let Some(&n) = self.multi_return_functions.get(&cname) {
                            self.pack_multi_value_result(n);
                        }
                    }
                }
            }

            // ── Member access ───────────────────────────────────────────
            ExprKind::Member { object, field, null_safe } => {
                // Namespace constant check (Math.PI, etc.)
                if let ExprKind::Ident(obj_name) = &object.kind {
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
                    if is_ctor && self.defined_globals.contains(obj_name.as_str()) {
                        self.emit_var_get(obj_name);
                        self.emit_u8(Op::CALL_REF, 0);
                        return Ok(());
                    }
                }

                self.compile_expr(object)?;

                if *null_safe {
                    // ?. — check null before accessing
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_FALSE);
                    // Object is null — result is null
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(skip);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                    self.patch_jump(end);
                } else {
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                }
            }

            // ── Index access ────────────────────────────────────────────
            ExprKind::Index { object, index } => {
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
                } else {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                }
            }

            // ── New ─────────────────────────────────────────────────────
            ExprKind::New { class, args } => {
                let class_parts = self.flatten_member_chain(class);
                let dotted_type_name = match &class.kind {
                    ExprKind::Ident(name) => Some(name.clone()),
                    ExprKind::Member { .. } if self.profile.namespaces.use_dotnet && !class_parts.is_empty() => {
                        Some(class_parts.join("."))
                    }
                    _ => None,
                };

                if let Some(type_name) = dotted_type_name.as_ref() {
                    // User-defined classes take priority over all built-in type mappings.
                    // This ensures `class Point { ... }` followed by `new Point()` calls
                    // the user constructor, not vybe:drawing::pointNew.
                    let canon_type = self.canon(type_name);
                    if self.defined_classes.contains(&canon_type) {
                        // Bypass compile_expr to avoid the implicit-self-field
                        // shadowing path: in case-insensitive languages a field
                        // named `inner` and a class named `Inner` both
                        // canonicalize to "inner", and the implicit-self-field
                        // check would mis-route to `me.inner` instead of the
                        // class global. Type names always come from globals.
                        let idx = self.str_const(&canon_type);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);
                        return Ok(());
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
                            common::threading::emit_thread_new(self.chunk(), line);
                            return Ok(());
                        }
                        "task" => {
                            // New Task(callback) → cont_new only
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(self.chunk(), line);
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
                        self.emit_u16(Op::STRUCT_NEW, 0);
                        self.emit(Op::DUP);
                        if let Some(msg_arg) = args.first() {
                            self.compile_expr(&msg_arg.value)?;
                        } else {
                            self.emit_const(Value::String(Arc::from("")));
                        }
                        let line = self.line;
                        common::errors::emit_exception_new_finalize(self.chunk(), type_name, line);
                        // Stamp `stack` = "Name: message" using locals.
                        // Stack after finalize: [obj]
                        let exc_tmp = self.scope_mut().define("__exc_tmp");
                        self.emit_u16(Op::LOCAL_SET, exc_tmp); self.emit(Op::DROP);
                        // Build "Name: " + message
                        self.emit_const(Value::String(Arc::from(format!("{}: ", type_name))));
                        self.emit_u16(Op::LOCAL_GET, exc_tmp);
                        let msg_k = self.str_const("message");
                        self.emit_u16(Op::STRUCT_GET, msg_k);
                        // Stack: ["Name: ", msg]. str_concat: a=prefix, b=msg → prefix+msg
                        self.emit(Op::STR_CONCAT);
                        // Stack: ["Name: msg"]. Save it.
                        let sv = self.scope_mut().define("__stack_val");
                        self.emit_u16(Op::LOCAL_SET, sv); self.emit(Op::DROP);
                        // Stamp: obj.stack = stack_val
                        self.emit_u16(Op::LOCAL_GET, exc_tmp);
                        self.emit_u16(Op::LOCAL_GET, sv);
                        let sk = self.str_const("stack");
                        self.emit_u16(Op::STRUCT_SET, sk);
                        self.emit(Op::DROP);
                        // Result: push obj
                        self.emit_u16(Op::LOCAL_GET, exc_tmp);
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
                        match target {
                            vybe_bytecode::component_model::ConstructorTarget::Host(target) => {
                                let idx = self.import(&target.module, &target.name);
                                self.emit_host_call(idx, args.len() as u8);
                            }
                            vybe_bytecode::component_model::ConstructorTarget::Common(name) => {
                                let line = self.line;
                                self.emit_common(&name, line);
                                // Tag the object with `__type` and install
                                // `__get_count` / `__get_length` auto-getters
                                // via the one-shot `vybe:types/__stamp_type`
                                // import. Common-backed constructors create
                                // raw JS-shape objects (Array via
                                // `collections.new`, Object via
                                // `ecma:object/new`, etc.) — the stamp
                                // adds the .NET metadata runtime dispatch
                                // needs without per-class host fns.
                                // Stack: [obj] → [obj, name] → [obj]
                                self.emit_const(Value::String(Arc::from(bare_str)));
                                let stamp_idx = self.import("vybe:types", "__stamp_type");
                                self.emit_host_call(stamp_idx, 2);
                            }
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
                                self.emit_common(&func, line);
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
                        return Ok(());
                    }
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
                            // array which replaces the one on TOS.
                            self.compile_expr(&elem.value)?;
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
                        let s = self.scope_mut().define("__pack");
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
                        let s = self.scope_mut().define("__pack");
                        if i == 0 { first = s; }
                    }
                    first
                };
                common::collections::emit_pack_n(&mut self.chunks, self.current, n as u16, base, line);
                // Convert to set via host call
                let idx = self.import("vybe:collections", "arrayToSet");
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
                                let key_tmp = self.scope_mut().define("__obj_dyn_key");
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
                            // Object spread: merge properties from expr into current object
                            self.compile_expr(expr)?;
                            let idx = self.import("vybe:object", "assign");
                            self.emit_host_call(idx, 2);
                        }
                        ObjectProperty::Method { key, value } => {
                            self.emit(Op::DUP);
                            if let StmtKind::FunctionDecl { params, body, .. } = &value.kind {
                                // Object methods receive `this` as implicit first arg
                                let mut method_params = vec![Param {
                                    name: self.profile.self_keyword.clone(),
                                    type_hint: None, default: None,
                                    pass_by: PassBy::Value, is_rest: false,
                                    is_kwargs: false, is_optional: false, is_nullable: false,
                                }];
                                method_params.extend(params.iter().cloned());
                                self.compile_lambda(&method_params, &LambdaBody::Block(body.clone()))?;
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
                                // Accessors receive `this` as first arg
                                let mut accessor_params = vec![Param {
                                    name: self.profile.self_keyword.clone(),
                                    type_hint: None, default: None,
                                    pass_by: PassBy::Value, is_rest: false,
                                    is_kwargs: false, is_optional: false, is_nullable: false,
                                }];
                                accessor_params.extend(params.iter().cloned());
                                self.compile_lambda(&accessor_params, &LambdaBody::Block(body.clone()))?;
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
                            self.emit(Op::DUP); // save key for __keys
                            let key_tmp = self.scope_mut().define("__obj_comp_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                            self.compile_expr(value)?;
                            let l = self.line;
                            common::collections::emit_set(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP); // drop returned null
                            // Track in __keys
                            self.emit(Op::DUP);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            let l = self.line;
                            common::collections::emit_push(&mut self.chunks, self.current, l);
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
                // Use stdlib __vybe_tostring (pure WASM, populated by bundle::finalize_with_stdlib)
                let tostring_global = self.str_const("__vybe_tostring");
                for (i, part) in parts.iter().enumerate() {
                    match part {
                        InterpolPart::Text(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                        InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                            // Push func ref FIRST, then the value, then call_ref
                            self.emit_u16(Op::GLOBAL_GET, tostring_global);
                            self.compile_expr(e)?;
                            self.emit_u8(Op::CALL_REF, 1);
                        }
                    }
                    if i > 0 {
                        let line = self.line;
                        common::strings::emit_str_concat(self.chunk(), line);
                    }
                }
            }

            // ── Type operations ─────────────────────────────────────────
            ExprKind::IsType { expr: inner, type_name } => {
                // Compare against canonicalized class name (case-insensitive
                // languages like VB/Pascal store class __type lowercased).
                let canon_type = self.canon(type_name);
                self.compile_expr(inner)?;
                let key = self.str_const("__type");
                self.emit_u16(Op::STRUCT_GET, key);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                self.emit(Op::DYN_EQ);
            }

            ExprKind::Cast { expr: inner, .. } => {
                // Cast is a no-op in our dynamic VM
                self.compile_expr(inner)?;
            }

            ExprKind::TypeOf(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::REF_TYPEOF);
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
                self.emit(Op::SPREAD);
            }

            // ── Await ───────────────────────────────────────────────────
            ExprKind::Await(inner) => {
                // In our synchronous VM, promises are already resolved.
                // `await p` unwraps the promise's `__value` property.
                // If the inner value is not a promise, pass through.
                self.compile_expr(inner)?;
                // Save to local, try to read __value
                let await_slot = self.scope_mut().define("__await");
                self.emit_u16(Op::LOCAL_SET, await_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, await_slot);
                let vk = self.str_const("__value");
                self.emit_u16(Op::STRUCT_GET, vk);
                // If __value is null → not a promise, use original
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let use_original = self.emit_jump(Op::BR_IF_TRUE);
                // __value exists → use it (drop the null-check dup)
                let done = self.emit_jump(Op::BR);
                self.patch_jump(use_original);
                self.emit(Op::DROP); // drop null __value
                self.emit_u16(Op::LOCAL_GET, await_slot); // push original
                self.patch_jump(done);
            }

            // ── Yield ───────────────────────────────────────────────────
            ExprKind::Yield(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::NULL); }
                self.emit_u16(Op::SUSPEND, 0);
            }

            ExprKind::YieldFrom(inner) => {
                self.compile_expr(inner)?;
                // Simplified: yield from → just pass through
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
                    // MyBase.Method(args) → this.__base_method(this, args)
                    let base_name = format!("__base_{}", self.canon(mname));
                    if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        let prop = self.str_const(&base_name);
                        self.emit_u16(Op::LOCAL_GET, self_slot);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        self.emit_u16(Op::LOCAL_GET, self_slot);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::CALL_REF, (args.len() + 1) as u8);
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
                let result_slot = self.scope_mut().define("__comp_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);

                // Only handle the first generator for simplicity
                if let Some(generator) = generators.first() {
                    self.compile_expr(&generator.iter)?;
                    let arr_slot = self.scope_mut().define("__comp_iter");
                    self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                    let idx_slot = self.scope_mut().define("__comp_idx");
                    let lp = common::loops::emit_for_in_start(
                        &mut self.chunks, self.current, arr_slot, idx_slot, line,
                    );
                    // Bind loop var
                    let var_name = match &generator.target.kind {
                        ExprKind::Ident(n) => n.clone(),
                        _ => "__comp_var".to_string(),
                    };
                    let var_slot = self.scope_mut().define(&var_name);
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
                // Emit slice parts for use by Index
                if let Some(l) = lower { self.compile_expr(l)?; } else { self.emit(Op::NULL); }
                if let Some(u) = upper { self.compile_expr(u)?; } else { self.emit(Op::NULL); }
                if let Some(s) = step { self.compile_expr(s)?; } else { self.emit(Op::NULL); }
                let idx = self.import("vybe:array", "sliceStep");
                self.emit_host_call(idx, 4); // obj already on stack from Index parent
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
                self.emit(Op::NULL); // void always evaluates to undefined
            }

            // ── Delete (JS expression) ──────────────────────────────────
            ExprKind::Delete(inner) => {
                // delete obj.prop → call vybe:object::deleteProperty(obj, key)
                // which removes the property and returns true.
                if let ExprKind::Member { object, field, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    let idx = self.import("vybe:object", "deleteProperty");
                    self.emit_host_call(idx, 2);
                } else if let ExprKind::Index { object, index } = &inner.kind {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    let idx = self.import("vybe:object", "deleteProperty");
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
                let parent_name = if let Some(p) = parent {
                    if let ExprKind::Ident(n) = &p.kind { Some(n.clone()) } else { None }
                } else { None };
                let class_name = self.canon(&class_name);
                let parent_name = parent_name.map(|p| self.canon(&p));
                self.defined_globals.insert(class_name.clone());
                self.compile_class(&class_name, &parent_name, members)?;
                self.emit_var_get(&class_name);
            }

            // ── FunctionExpr (JS) ───────────────────────────────────────
            ExprKind::FunctionExpr(stmt) => {
                if let StmtKind::FunctionDecl { name, params, return_type, body, is_sub, is_generator, handles, is_async, .. } = &stmt.kind {
                    let fn_name = if name.is_empty() { "__anon_fn" } else { name };
                    self.compile_function_decl(fn_name, params, return_type, body, *is_sub, *is_generator, handles, *is_async)?;
                    self.emit_var_get(fn_name);
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
                // class::member → look up class, then get static member
                self.compile_expr(class)?;
                if let ExprKind::Ident(name) = &member.kind {
                    let idx = self.str_const(name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                } else {
                    self.compile_expr(member)?;
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                }
            }

            // ── Match expression (PHP/Rust) ─────────────────────────────
            ExprKind::Match { subject, arms } => {
                self.compile_expr(subject)?;
                let subject_slot = self.scope_mut().define("__match_subj");
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

}
