//! Call-expression compilation — `compile_call` (handles named calls,
//! method calls, super-calls, spread, dotted lookups) and
//! `compile_lambda`. This is the primary edit site for the inline
//! refactor (Phase G) where `wasm:js-*` imports get replaced by
//! inline WASM GC sequences.

use super::*;

fn python_is_identifier_literal(value: &str) -> bool {
    let mut chars = value.chars();
    match chars.next() {
        Some(ch) if ch == '_' || ch.is_alphabetic() => {}
        _ => return false,
    }
    chars.all(|ch| ch == '_' || ch.is_alphanumeric())
}

fn python_is_printable_literal(value: &str) -> bool {
    value.chars().all(|ch| !ch.is_control())
}

fn terminal_type_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
        ExprKind::Member { field, .. } => Some(field.clone()),
        _ => None,
    }
}

fn dotnet_factory_return_type(callee: &Expression) -> Option<String> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let class_name = terminal_type_name(object)?;
    if class_name.eq_ignore_ascii_case("TimeSpan")
        && matches!(field.as_str(), "FromDays" | "FromHours" | "FromMinutes" | "FromSeconds" | "FromMilliseconds" | "Zero")
    {
        return Some("TimeSpan".into());
    }
    if class_name.eq_ignore_ascii_case("DateTime")
        && matches!(field.as_str(), "Now" | "UtcNow" | "Today" | "Parse")
    {
        return Some("DateTime".into());
    }
    None
}

fn resolve_receiver_type_hint(compiler: &Compiler, recv: &Expression) -> Option<String> {
    match &recv.kind {
        ExprKind::Ident(local_name) => compiler.scope().resolve_type_ci(local_name).map(|s| s.to_string())
            .or_else(|| {
                let cn = compiler.canon(local_name);
                compiler.global_type_hints.get(&cn).cloned()
            })
            .or_else(|| compiler.is_class_static_field_type_hint(local_name)),
        ExprKind::Member { object, field, .. } => {
            let owner_is_self = matches!(&object.kind, ExprKind::This | ExprKind::Super)
                || matches!(&object.kind, ExprKind::Ident(n)
                    if {
                        let cn = compiler.canon(n);
                        cn == compiler.profile.self_keyword
                            || cn == "me"
                            || cn == "this"
                            || cn == "mybase"
                    });
            if owner_is_self {
                compiler.is_class_static_field_type_hint(field)
            } else if let ExprKind::Ident(owner) = &object.kind {
                let owner_name = owner
                    .split('<')
                    .next()
                    .map(str::trim)
                    .unwrap_or(owner);
                let canon_field = compiler.canon(field);

                let mut owner_candidates = vec![owner_name.to_string()];
                let owner_canon = compiler.canon(owner_name);
                if owner_canon != owner_name {
                    owner_candidates.push(owner_canon);
                }

                for owner_key in owner_candidates {
                    let mut current = Some(owner_key.as_str());
                    while let Some(cn) = current {
                        if let Some(pc) = compiler.pending_classes.get(cn) {
                            if let Some(type_hint) = pc.static_field_types.get(&canon_field) {
                                return Some(type_hint.clone());
                            }
                            current = pc.parent.as_deref();
                        } else {
                            break;
                        }
                    }
                }
                None
            } else {
                None
            }
        }
        ExprKind::New { class, .. } => terminal_type_name(class),
        ExprKind::Call { callee, .. } => dotnet_factory_return_type(callee).or_else(|| match &callee.kind {
            ExprKind::Ident(name) => common::dotnet::surface()
                .lookup_constructor(name)
                .map(|_| name.rsplit('.').next().unwrap_or(name).to_string()),
            ExprKind::Member { field, .. } => common::dotnet::surface()
                .lookup_constructor(field)
                .map(|_| field.to_string()),
            _ => None,
        }),
        _ => None,
    }
}

impl Compiler {
    pub(super) fn js_error_instanceof_chain(type_name: &str) -> &'static [&'static str] {
        match type_name.trim() {
            "Error" => &["Error"],
            "EvalError" => &["EvalError", "Error"],
            "RangeError" => &["RangeError", "Error"],
            "ReferenceError" => &["ReferenceError", "Error"],
            "SyntaxError" => &["SyntaxError", "Error"],
            "TypeError" => &["TypeError", "Error"],
            "URIError" => &["URIError", "Error"],
            "AggregateError" => &["AggregateError", "Error"],
            _ => &[],
        }
    }

    pub(super) fn emit_js_exception_ctor_from_message_value(&mut self, type_name: &str) -> Result<(), String> {
        let msg_val = self.define_local("__exc_msg_val");
        self.emit_u16(Op::LOCAL_SET, msg_val);
        self.emit(Op::DROP);

        self.emit_u16(Op::STRUCT_NEW, 0);
        self.emit(Op::DUP);
        self.emit_u16(Op::LOCAL_GET, msg_val);
        let line = self.line;
        common::errors::emit_exception_new_finalize(self.chunk(), type_name, line);

        let exc_tmp = self.define_local("__exc_tmp");
        self.emit_u16(Op::LOCAL_SET, exc_tmp);
        self.emit(Op::DROP);

        self.emit_const(Value::String(Arc::from(format!("{}: ", type_name))));
        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        let msg_k = self.str_const("message");
        self.emit_u16(Op::STRUCT_GET, msg_k);
        self.emit(Op::STR_CONCAT);
        let stack_val = self.define_local("__stack_val");
        self.emit_u16(Op::LOCAL_SET, stack_val);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        self.emit_u16(Op::LOCAL_GET, stack_val);
        let stack_key = self.str_const("stack");
        self.emit_u16(Op::STRUCT_SET, stack_key);
        self.emit(Op::DROP);

        if self.is_js_profile() {
            for name in Self::js_error_instanceof_chain(type_name) {
                common::classes::emit_instanceof_chain(&mut self.chunks, self.current, exc_tmp, name, line);
            }
        }

        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        Ok(())
    }

    pub(super) fn emit_js_exception_ctor_value(&mut self, type_name: &str, args: &[&Expression]) -> Result<(), String> {
        if let Some(msg_arg) = args.first() {
            self.compile_expr(msg_arg)?;
        } else {
            self.emit_const(Value::String(Arc::from("")));
        }
        self.emit_js_exception_ctor_from_message_value(type_name)?;

        if let Some(opts_arg) = args.get(1) {
            let exc_tmp = self.define_local("__exc_with_cause");
            self.emit_u16(Op::LOCAL_SET, exc_tmp);
            self.emit(Op::DROP);
            self.compile_expr(opts_arg)?;
            let cause_key = self.str_const("cause");
            self.emit_u16(Op::STRUCT_GET, cause_key);
            let cause_val = self.define_local("__cause_val");
            self.emit_u16(Op::LOCAL_SET, cause_val);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
            self.emit_u16(Op::LOCAL_GET, cause_val);
            self.emit_u16(Op::STRUCT_SET, cause_key);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Call compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<(), String> {
        let arg_exprs: Vec<&Expression> = args.iter().map(|a| &a.value).collect();

        if self.is_python_profile() {
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "dict" {
                    let line = self.line;
                    common::dict::emit_new(&mut self.chunks, self.current, line);

                    if args.iter().all(|arg| arg.name.is_some()) {
                        for arg in args {
                            let key = arg.name.as_ref().unwrap();
                            self.emit(Op::DUP);
                            self.compile_expr(&arg.value)?;
                            let key_idx = self.str_const(key);
                            self.emit_u16(Op::STRUCT_SET, key_idx);
                            self.emit(Op::DROP);

                            self.emit(Op::DUP);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_const(Value::String(Arc::from(key.as_str())));
                            common::collections::emit_push(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                        }
                        return Ok(());
                    }

                    if args.len() == 1 && args[0].name.is_none() && !args[0].spread {
                        if let ExprKind::Array(elements) = &args[0].value.kind {
                            for element in elements {
                                let ExprKind::Tuple(items) = &element.value.kind else { continue; };
                                if items.len() != 2 { continue; }

                                self.emit(Op::DUP);
                                self.compile_expr(&items[0])?;
                                let key_tmp = self.define_local("__py_dict_ctor_key");
                                self.emit(Op::DUP);
                                self.emit_u16(Op::LOCAL_SET, key_tmp);
                                self.emit(Op::DROP);
                                self.compile_expr(&items[1])?;
                                common::collections::emit_set(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);

                                self.emit(Op::DUP);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                common::collections::emit_push(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);
                            }
                            return Ok(());
                        }
                    }

                    if args.is_empty() {
                        return Ok(());
                    }
                }
            }
        }

        if self.is_php_profile() {
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "compact" {
                    let line = self.line;
                    common::collections::emit_map_new(&mut self.chunks, self.current, line);
                    for arg in args {
                        let ExprKind::Lit(Literal::Str(var_name)) = &arg.value.kind else {
                            self.emit(Op::NULL);
                            return Ok(());
                        };
                        let php_var_name = format!("${}", var_name);
                        self.emit(Op::DUP);
                        self.emit_const(Value::String(Arc::from(var_name.as_str())));
                        self.emit_var_get(&php_var_name);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                    }
                    return Ok(());
                }

                if name == "extract" && arg_exprs.len() == 1 {
                    if let ExprKind::Array(elements) = &arg_exprs[0].kind {
                        let mut count = 0i64;
                        for elem in elements {
                            let Some(key_expr) = &elem.key else { continue; };
                            let bind_name = match &key_expr.kind {
                                ExprKind::Lit(Literal::Str(s)) => format!("${}", s),
                                ExprKind::Lit(Literal::Int(n)) => format!("${}", n),
                                _ => continue,
                            };
                            self.compile_expr(&elem.value)?;
                            self.emit_var_set(&bind_name);
                            count += 1;
                        }
                        self.emit_const(Value::I64(count));
                        return Ok(());
                    }
                }
            }
        }

        // ── super(args) → call parent constructor, store result as this ──
        if let ExprKind::Super = &callee.kind {
            if let Some(ref class_name) = self.current_class.clone() {
                if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|pc| pc.parent.clone()) {
                    if self.is_js_profile() && common::errors::is_exception_type(&parent_name) {
                        self.emit_js_exception_ctor_value(&parent_name, &arg_exprs)?;
                        let self_kw = self.profile.self_keyword.clone();
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
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    // Store result as this
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        self.emit(Op::DUP);
                        self.emit_u16(Op::LOCAL_SET, slot);
                        self.emit(Op::DROP);
                    }
                    return Ok(());
                }
            }
            // No parent — emit null
            self.emit(Op::NULL);
            return Ok(());
        }

        // ── super.method(args) — static class dispatch ───────────────
        //
        // Resolve the parent class statically at compile time. Inside
        // `class C extends B`, `super.method()` always means B's
        // method (regardless of the runtime instance type) — the spec
        // says super uses [[HomeObject]].[[Prototype]], NOT the
        // instance's prototype chain. Multi-level inheritance (C → B
        // → A) needs B.method when called from C and A.method when
        // called from B; the previous `this.__base_method` lookup
        // collided across levels (C overwriting B's slot) and caused
        // an infinite loop on C's super chain.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if matches!(&object.kind, ExprKind::Super) {
                let canon_field = self.canon(field);
                let class_name = self.current_class.clone();
                let parent_name = class_name.as_ref()
                    .and_then(|cn| self.pending_classes.get(cn.as_str()))
                    .and_then(|pc| pc.parent.clone());
                let self_kw = self.profile.self_keyword.clone();
                let self_slot = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw));

                if let Some(parent) = parent_name {
                    // Look up parent class via emit_var_get so closure-
                    // captured parents (mixin pattern: `(Base) => class
                    // extends Base`) resolve through the upvalue scope.
                    self.emit_var_get(&parent);
                    let method_idx = self.str_const(&canon_field);
                    self.emit_u16(Op::STRUCT_GET, method_idx);

                    if self.is_js_profile() {
                        let saved_js_this = self.save_js_this("__js_prev_this_super_method");
                        if let Some(slot) = self_slot {
                            self.emit_u16(Op::LOCAL_GET, slot);
                        } else {
                            let js_this = self.str_const("__js_this");
                            self.emit_u16(Op::GLOBAL_GET, js_this);
                        }
                        self.set_js_this_from_stack();
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        let result_slot = self.define_local("__js_super_method_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit(Op::DROP);
                        self.restore_js_this(saved_js_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    } else {
                        // Typed-language method ABI passes receiver as arg0.
                        if let Some(slot) = self_slot {
                            self.emit_u16(Op::LOCAL_GET, slot);
                        } else {
                            self.emit(Op::NULL);
                        }
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    }
                    return Ok(());
                }

                // Pascal / VB / C# allow `inherited Foo` / `MyBase.Foo` in a
                // root class even when there is no parent implementation. Treat
                // it as a no-op instead of falling through to the generic member
                // call pipeline and recursing back into the current method.
                self.emit(Op::NULL);
                return Ok(());
            }
        }

        // ── Debug intrinsic: __debug_dump(obj) ──────────────────────
        // Available in all languages. Prints object properties to stderr.
        if let ExprKind::Ident(name) = &callee.kind {
            if name == "__debug_dump" {
                for a in &arg_exprs { self.compile_expr(a)?; }
                let idx = self.import("vybe:debug", "dump");
                self.emit_host_call(idx, arg_exprs.len() as u8);
                return Ok(());
            }

            let canon = self.canon(name);
            let shadows_builtin_exception = self.defined_functions.contains(&canon)
                || self.defined_classes.contains(&canon)
                || self.defined_globals.contains(&canon)
                || (!self.case_sensitive && (
                    self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name))
                    || self.defined_classes.iter().any(|g| g.eq_ignore_ascii_case(name))
                    || self.defined_globals.iter().any(|g| g.eq_ignore_ascii_case(name))
                ));
            if !shadows_builtin_exception && common::errors::is_exception_type(name) {
                self.emit_js_exception_ctor_value(name, &arg_exprs)?;
                return Ok(());
            }
        }

        // ── Typed static-field receiver: counts.ContainsKey(...) ─────
        // Static fields can carry type hints too. Resolve them here so
        // class-level typed state uses the same shared .NET surface as
        // locals with type annotations.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let class_name = resolve_receiver_type_hint(self, object);
            if let Some(class_name) = class_name {
                let class_name = class_name
                    .split('<')
                    .next()
                    .map(str::trim)
                    .unwrap_or(&class_name)
                    .to_string();
                let surface = common::dotnet::surface();
                if let Some(target) = surface.lookup_instance_method(&class_name, field, arg_exprs.len() as u8) {
                    self.compile_expr(object)?;
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let total_argc = (arg_exprs.len() + 1) as u8;
                    match target {
                        common::dotnet::InstanceMethodTarget::Host { module, func, .. } => {
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, total_argc);
                        }
                        common::dotnet::InstanceMethodTarget::Common { emit, .. } => {
                            let line = self.line;
                            self.emit_common(&emit, total_argc, line);
                        }
                    }
                    return Ok(());
                }
            }
        }

        // ── ESM host-module import binding ──────────────────────────
        //
        // `import { createServer } from "wasi:http"` binds
        // `createServer` locally. Calling it here emits a direct
        // `CALL_IMPORT` against the recorded (module, fn) pair — the
        // import statement itself is the compile-time declaration.
        if let ExprKind::Ident(name) = &callee.kind {
            let key = self.canon(name);
            if let Some((module, func)) = self.host_import_bindings.get(&key).cloned() {
                for a in &arg_exprs { self.compile_expr(a)?; }
                let idx = self.import(&module, &func);
                self.emit_host_call(idx, arg_exprs.len() as u8);
                return Ok(());
            }
        }

        // ── Builtin check: Ident("print") ───────────────────────────
        // Skip for user-defined functions: a VB `Function Echo(...)` must
        // dispatch to the user's chunk, not to the cross-language `echo →
        // wasi:cli.log` import shortcut.
        if let ExprKind::Ident(name) = &callee.kind {
            let shadows_builtin = self.defined_functions.contains(name)
                || (!self.case_sensitive
                    && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
            if !shadows_builtin && self.try_compile_builtin(name, &arg_exprs)? {
                return Ok(());
            }
        }

        // ── Builtin check: Member("Console.WriteLine") ─────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(obj_name) = &object.kind {
                // Note: Object.create is handled via the host fn
                // (`ecma:object.create`) so it gets the full ECMA-262
                // §20.1.2.2 behaviour: descriptor second-arg, null
                // prototype gets `toString` etc. stamped as Undefined,
                // and parent properties are copied down for member
                // access. The earlier compiler shortcut here only set
                // `__proto__` and missed both — falling through to
                // `try_compile_builtin` below routes to the host fn.

                let compound = format!("{}.{}", obj_name, field);
                if self.try_compile_builtin(&compound, &arg_exprs)? { return Ok(()); }

                // ── ESM wildcard namespace member call ──────────────
                //
                // Per ECMA-262 §16.2, a Module Namespace Object is a
                // compile-time binding — `ns.field` resolves statically
                // to the `(module, field)` export. Covers both profile
                // defaults (JS `console` → `wasi:cli`) and user wildcard
                // imports (`import * as cli from "wasi:cli"`). The
                // Linker populated both into `host_namespace_aliases`.
                //
                // Runs AFTER `try_compile_builtin(compound)` so profile
                // builtins with custom emit logic (`Array.from`,
                // `Math.max`) still win on the names they claim.
                let key = self.canon(obj_name);
                if let Some(module) = self.host_namespace_aliases.get(&key).cloned() {
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let idx = self.import(&module, field);
                    self.emit_host_call(idx, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Two-level host prefix: `vybe.gui.setProperty(...)` ──────
        //
        // VB / languages without ESM imports reach host functions via
        // a literal namespace chain `<prefix>.<module>.<fn>(args)` where
        // the leading ident is a known host-namespace prefix (`vybe`,
        // `wasi`, `wasm`). Emit as `call_import("<prefix>:<module>",
        // "<fn>", args)` — identical to what JS gets via `import * as
        // gui from "vybe:gui"; gui.setProperty(...)`.
        //
        // Without this, the call falls through to the method-call
        // pattern and injects `vybe.<module>` as a phantom receiver,
        // shifting every argument right by one and silently breaking
        // host functions that don't expect a receiver slot.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Member { object: inner_obj, field: inner_field, .. } = &object.kind {
                if let ExprKind::Ident(prefix) = &inner_obj.kind {
                    let prefix_lc = self.canon(prefix);
                    if matches!(prefix_lc.as_str(), "vybe" | "wasi" | "wasm") {
                        let module = format!("{}:{}", prefix_lc, self.canon(inner_field));
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        let idx = self.import(&module, field);
                        self.emit_host_call(idx, arg_exprs.len() as u8);
                        return Ok(());
                    }
                }
            }
        }

        // ── Dotted name resolution FIRST (uses compiler_common::dotnet when use_dotnet) ──
        // Must run before value methods because value methods like "add" would
        // intercept "Controls.Add" which needs special GUI handling.
        if let ExprKind::Member { .. } = &callee.kind {
            let parts = self.flatten_member_chain(callee);
            if parts.len() >= 2 {
                let lower_parts: Vec<String> = parts.iter().map(|s| self.canon(s)).collect();

                // Use dotnet resolver when enabled
                if self.profile.namespaces.use_dotnet_resolver {
                    let skip_simple_instance_chain = if lower_parts.len() == 2 {
                        let head = &parts[0];
                        self.scope().resolve(head).is_some()
                            || self.scope().resolve_ci(head).is_some()
                            || self.defined_globals.contains(head)
                            || self.defined_globals.iter().any(|g| g.eq_ignore_ascii_case(head))
                            || self.is_class_field(head)
                            || self.is_class_static_field(head).is_some()
                    } else {
                        false
                    };
                    if skip_simple_instance_chain {
                        // Keep 2-part local/global member calls (`x.Method(...)`) on the
                        // normal instance pipeline; the dotted resolver is for namespace/
                        // static chains and can otherwise short-circuit LINQ-style calls.
                    } else {
                    let dotnet_surface = common::dotnet::surface();
                    let imports = {
                        let mut imp = dotnet_surface.default_imports().to_vec();
                        imp.extend(self.profile.namespaces.extra_imports.clone());
                        imp
                    };
                    let scope = self.scope();
                    let defined_globals = self.defined_globals.clone();
                    let field_set: std::collections::HashSet<String> = if let Some(ref cn) = self.current_class {
                        self.pending_classes.get(cn.as_str())
                            .map(|pc| pc.fields.iter().cloned().collect())
                            .unwrap_or_default()
                    } else {
                        std::collections::HashSet::new()
                    };
                    // `is_local` must recognise top-level variables that
                    // live in `defined_globals` (VB `Dim` at the module
                    // level, JS top-level `var`/`let`), but MUST NOT
                    // match user classes there — those go through
                    // `is_user_type` which returns Unresolved so static
                    // dispatch runs the class ctor path, not a bogus
                    // struct_get chain off the ctor function. The union
                    // (`is_local`) minus (`is_user_type`) gives the
                    // right set of "things you can local_get and
                    // struct_get from".
                    let defined_classes = self.defined_classes.clone();
                    let is_user_class_fn = move |name: &str| -> bool {
                        defined_classes.contains(name)
                            || defined_classes.iter().any(|c| c.eq_ignore_ascii_case(name))
                    };
                    let is_user_class_for_local = is_user_class_fn.clone();
                    let ctx = common::dotnet::ResolutionContext {
                        is_local: &|name: &str| {
                            if is_user_class_for_local(name) { return false; }
                            scope.resolve(name).is_some()
                            || scope.resolve_ci(name).is_some()
                            || defined_globals.contains(name)
                            || defined_globals.iter().any(|g| g.eq_ignore_ascii_case(name))
                        },
                        is_class_field: &|name: &str| field_set.contains(name),
                        is_user_type: &is_user_class_fn,
                        imports: &imports,
                    };
                    let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
                    let resolution = common::dotnet::resolve_dotted_name(&refs, &ctx);

                    match resolution {
                        common::dotnet::DottedResolution::CommonCall { emit } => {
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            let line = self.line;
                            self.emit_common(&emit, arg_exprs.len() as u8, line);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::HostCall { module, func } => {
                            if self.profile.name == "csharp"
                                && module.eq_ignore_ascii_case("ecma:number")
                                && func.eq_ignore_ascii_case("parseInt")
                                && arg_exprs.len() == 1
                            {
                                let is_char_like = match &arg_exprs[0].kind {
                                    ExprKind::Lit(Literal::Char(_)) => true,
                                    ExprKind::Ident(name) => self.lookup_var_type_hint(name)
                                        .is_some_and(|hint| Self::normalize_type_hint(hint) == "char"),
                                    _ => false,
                                };
                                if is_char_like {
                                    self.compile_expr(arg_exprs[0])?;
                                    self.emit(Op::I32_CONST_0);
                                    self.emit(Op::STR_CHAR_CODE_AT);
                                    return Ok(());
                                }
                            }
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, arg_exprs.len() as u8);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::NamespaceAccess { parts: ns_parts } => {
                            // If any contiguous sub-window of the chain is a profile namespace
                            // constant (e.g. ["system","math","pi","tostring"] where "math.pi"
                            // is a constant), emit the constant and dispatch remaining as a
                            // value method. Namespace prefix before the constant is discarded.
                            if ns_parts.len() >= 2 {
                                let mut found_window: Option<(usize, usize)> = None;
                                'outer: for start in 0..ns_parts.len().saturating_sub(1) {
                                    for end in ((start + 2)..=ns_parts.len().saturating_sub(0)).rev() {
                                        if end > ns_parts.len() { continue; }
                                        let key = ns_parts[start..end].join(".");
                                        if self.profile.lookup_constant(&key).is_some() {
                                            found_window = Some((start, end));
                                            break 'outer;
                                        }
                                    }
                                }
                                if let Some((_const_start, const_end)) = found_window {
                                    let key = ns_parts[_const_start..const_end].join(".");
                                    let cv = self.profile.lookup_constant(&key).cloned().unwrap();
                                    match &cv {
                                        ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                                        ConstantValue::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                                    }
                                    let remaining = ns_parts[const_end..].to_vec();
                                    if let Some(method_name) = remaining.first() {
                                        let argc = arg_exprs.len() as u8;
                                        let def = self.profile.lookup_value_method(method_name, argc).cloned();
                                        if let Some(def) = def {
                                            for a in &arg_exprs { self.compile_expr(a)?; }
                                            let line = self.line;
                                            match &def.emit {
                                                BuiltinEmit::Stdlib(name) => {
                                                    // For stdlib: func ref must be pushed BEFORE object.
                                                    // But object is already on stack. Save it to a temp.
                                                    let tmp = self.define_local("__const_val");
                                                    self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                                                    let global_name = format!("__vybe_{}", name);
                                                    let name_idx = self.str_const(&global_name);
                                                    self.emit_u16(Op::GLOBAL_GET, name_idx);
                                                    self.emit_u16(Op::LOCAL_GET, tmp);
                                                    for a in &arg_exprs { self.compile_expr(a)?; }
                                                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                                                }
                                                BuiltinEmit::HostCall(module, func) => {
                                                    let idx = self.import(module, func);
                                                    self.emit_host_call(idx, (arg_exprs.len() + 1) as u8);
                                                }
                                                BuiltinEmit::Common(name) => {
                                                    let name = name.clone();
                                                    self.emit_common(&name, (arg_exprs.len() + 1) as u8, line);
                                                }
                                                BuiltinEmit::Opcode(op_name) => {
                                                    self.emit_named_opcode(op_name);
                                                }
                                                _ => {
                                                    // Fallback: STRUCT_GET the method and call_ref
                                                    let idx = self.str_const(method_name);
                                                    self.emit_u16(Op::STRUCT_GET, idx);
                                                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                                }
                                            }
                                        } else {
                                            // No value method — STRUCT_GET and call_ref
                                            let idx = self.str_const(method_name);
                                            self.emit_u16(Op::STRUCT_GET, idx);
                                            for a in &arg_exprs { self.compile_expr(a)?; }
                                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                        }
                                    }
                                    return Ok(());
                                }
                            }

                            if !arg_exprs.is_empty() && ns_parts.len() >= 2 {
                                let method_name = ns_parts.last().cloned().unwrap_or_default();
                                let root_idx = self.str_const(&ns_parts[0]);
                                self.emit_u16(Op::GLOBAL_GET, root_idx);
                                for part in &ns_parts[1..ns_parts.len() - 1] {
                                    let idx = self.str_const(part);
                                    self.emit_u16(Op::STRUCT_GET, idx);
                                }
                                let method_idx = self.str_const(&method_name);
                                self.emit(Op::DUP);
                                self.emit_u16(Op::STRUCT_GET, method_idx);
                                let fn_tmp = self.define_local("__ns_fn");
                                self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                                let obj_tmp = self.define_local("__ns_obj");
                                self.reserve_local_slot(obj_tmp);
                                self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                                return Ok(());
                            }

                            let root_idx = self.str_const(&ns_parts[0]);
                            self.emit_u16(Op::GLOBAL_GET, root_idx);
                            for part in &ns_parts[1..] {
                                let idx = self.str_const(part);
                                self.emit_u16(Op::STRUCT_GET, idx);
                            }
                            let is_const = ns_parts
                                .last()
                                .map(|name| dotnet_surface.is_known_constant(name))
                                .unwrap_or(false);
                            if !is_const {
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                            }
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::InstanceMember { local, members } => {
                            // Intercept `parent.Controls.Add(child)` for GUI.
                            // The .NET WinForms surface is `Form.Controls.Add(ctrl)`,
                            // MAUI is `parent.Children.Add(ctrl)`, etc. — all
                            // resolve to the canonical gui emitter.
                            if members.len() >= 2 && members[members.len()-2] == "controls" && members[members.len()-1] == "add" {
                                let line = self.line;
                                let add_idx = self.import("vybe:gui", common::gui::HOST_FN_ADD_CHILD);
                                self.emit_var_get(&local);
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                common::gui::emit_add_child(self.chunk(), add_idx, line);
                                return Ok(());
                            }
                            // Intercept Thread/Task methods → WASM stack switching opcodes.
                            // Disambiguation by arity: `Thread.Join()` is zero-arg; an
                            // array's `.join(sep)` takes one. Without the arity gate
                            // this branch greedy-matched both and routed string-join
                            // through `thread.join` (which returns the exit code, not
                            // a string).
                            if members.len() == 1 && arg_exprs.is_empty() {
                                let method = members[0].as_str();
                                match method {
                                    "join" => {
                                        self.emit_var_get(&local);
                                        let line = self.line;
                                        common::threading::emit_thread_join(self.chunk(), line);
                                        return Ok(());
                                    }
                                    "waitforexit" => {
                                        self.emit_var_get(&local);
                                        let line = self.line;
                                        common::dotnet::core::process_adapter::emit_process_wait_for_exit(&mut self.chunks, self.current, line);
                                        return Ok(());
                                    }
                                    _ => {}
                                }
                            }
                            let _ = local;
                            let _ = members;
                            // For ordinary local/member calls, fall through to the
                            // shared call pipeline below. That keeps value-method
                            // dispatch (`dict.Add`, `queue.Dequeue`, etc.) and the
                            // generic object member path as the single source of truth.
                        }
                        common::dotnet::DottedResolution::NoOp => {
                            self.emit(Op::NULL);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::Unresolved => {
                            // Fall through to value methods and other resolution
                        }
                    }
                    }
                }

                // Non-dotnet: namespace aliases (JS: console → wasi:cli).
                // Reads from `host_namespace_aliases` (populated by the
                // Linker) instead of `profile.lookup_module_alias` — one
                // source of truth for Member-chain resolution.
                let dotnet_root = self.profile.namespaces.use_dotnet_resolver
                    && common::dotnet::is_namespace_root(&lower_parts[0]);
                if !dotnet_root {
                    let alias_key = self.canon(&lower_parts[0]);
                    if let Some(module) = self.host_namespace_aliases.get(&alias_key).cloned() {
                    let func = if lower_parts.len() == 2 { lower_parts[1].clone() } else { lower_parts[1..].join(".") };
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let idx = self.import(&module, &func);
                    self.emit_host_call(idx, arg_exprs.len() as u8);
                    return Ok(());
                    }
                }

                // Profile namespace roots
                if self.profile.is_namespace_root(&lower_parts[0]) {
                    let root_idx = self.str_const(&lower_parts[0]);
                    self.emit_u16(Op::GLOBAL_GET, root_idx);
                    for part in &lower_parts[1..] {
                        let idx = self.str_const(part);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Static method call on user class: ClassName.Method(args) ─
        // Must run BEFORE value methods so user class names like MathUtils.Add
        // don't get hijacked by the array Add value method.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(obj_name) = &object.kind {
                let canon = self.canon(obj_name);
                let is_class = self.defined_classes.contains(&canon)
                    && self.scope().resolve(obj_name).is_none();
                if is_class {
                    if self.is_js_profile() {
                        let cls_idx = self.str_const(&canon);
                        self.emit_u16(Op::GLOBAL_GET, cls_idx);
                        let cls_tmp = self.scope().resolve("__static_cls")
                            .unwrap_or_else(|| self.define_local("__static_cls"));
                        self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, cls_tmp);
                        let method_idx = self.str_const(&self.canon(field));
                        self.emit_u16(Op::STRUCT_GET, method_idx);
                        let fn_tmp = self.scope().resolve("__static_fn")
                            .unwrap_or_else(|| self.define_local("__static_fn"));
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let saved_js_this = self.save_js_this("__js_prev_this_static_method");
                        self.emit_u16(Op::LOCAL_GET, cls_tmp);
                        self.set_js_this_from_stack();
                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        let result_slot = self.define_local("__js_static_method_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        self.restore_js_this(saved_js_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }

                    // Push class, dup, struct_get(method) → [class, fn]
                    // Then swap so fn is first, class is second (as this)
                    let cls_idx = self.str_const(&canon);
                    self.emit_u16(Op::GLOBAL_GET, cls_idx);
                    self.emit(Op::DUP);
                    let m = self.canon(field);
                    let method_idx = self.str_const(&m);
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                    // Stack: [class, fn] — swap so we have [fn, class, ...args]
                    let fn_tmp = self.scope().resolve("__static_fn")
                        .unwrap_or_else(|| self.define_local("__static_fn"));
                    self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                    let cls_tmp = self.scope().resolve("__static_cls")
                        .unwrap_or_else(|| self.define_local("__static_cls"));
                    self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.emit_u16(Op::LOCAL_GET, cls_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
            }
        }

        // ── Nested static type call: Outer.Inner.Method(args) ───────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Member { object: outer_obj, field: nested_name, .. } = &object.kind {
                if let ExprKind::Ident(outer_name) = &outer_obj.kind {
                    let outer_canon = self.canon(outer_name);
                    let is_outer_class = self.defined_classes.contains(&outer_canon)
                        && self.scope().resolve(outer_name).is_none();
                    if is_outer_class {
                        let nested_ok = self.pending_classes.get(outer_canon.as_str())
                            .map(|pc| pc.nested_types.iter().any(|n| {
                                if self.case_sensitive { n == nested_name } else { n.eq_ignore_ascii_case(nested_name) }
                            }))
                            .unwrap_or(false);
                        if nested_ok {
                            let outer_idx = self.str_const(&outer_canon);
                            self.emit_u16(Op::GLOBAL_GET, outer_idx);
                            let nested_idx = self.str_const(&self.canon(nested_name));
                            self.emit_u16(Op::STRUCT_GET, nested_idx);
                            let cls_tmp = self.scope().resolve("__nested_static_cls")
                                .unwrap_or_else(|| self.define_local("__nested_static_cls"));
                            self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, cls_tmp);
                            let method_idx = self.str_const(&self.canon(field));
                            self.emit_u16(Op::STRUCT_GET, method_idx);
                            let fn_tmp = self.scope().resolve("__nested_static_fn")
                                .unwrap_or_else(|| self.define_local("__nested_static_fn"));
                            self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            self.emit_u16(Op::LOCAL_GET, cls_tmp);
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                            return Ok(());
                        }
                    }
                }
            }
        }

        // ── Function.prototype.call / .apply ────────────────────────
        // `fn.call(thisArg, a, b, ...)` → call `fn` with `[a, b, ...]`
        // `fn.apply(thisArg, [a, b, ...])` → same; the array form is
        // unwrapped at runtime via the spread opcode.
        //
        // We can't route this through value_methods because the standard
        // dispatch path pushes the receiver + ALL args, but here we need
        // to drop arg[0] (`thisArg`) from the middle of the stack. Skip
        // when the field is defined on a user class so user methods
        // named `call`/`apply` keep working.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let canon_field = self.canon(field);
            if !self.defined_class_methods.contains(&canon_field)
                && (field == "call" || field == "apply")
            {
                let saved_js_this = self.save_js_this("__js_prev_this_call");
                if self.is_js_profile() {
                    if let Some(this_arg) = arg_exprs.first() {
                        self.compile_expr(this_arg)?;
                    } else {
                        let line = self.line;
                        common::expressions::emit_undefined(self.chunk(), line);
                    }
                    self.set_js_this_from_stack();
                }
                self.compile_expr(object)?;                       // [fn]
                if field == "call" {
                    // Skip thisArg, compile rest as positional args.
                    for a in arg_exprs.iter().skip(1) {
                        self.compile_expr(a)?;
                    }
                    let n = arg_exprs.len().saturating_sub(1);
                    self.emit_u8(Op::CALL_REF, n as u8);
                } else {
                    // apply(thisArg, argsArray) — spread the array.
                    if let Some(args_expr) = arg_exprs.get(1) {
                        self.compile_expr(args_expr)?;
                        self.emit(Op::SPREAD);
                    }
                    // Use call_ref with 0 — the spread opcode pushes
                    // each array element and bumps the call arity at
                    // runtime via Op::call_spread if available, else
                    // we fall back here. The current VM uses Op::SPREAD
                    // before call_ref to flatten the top array.
                    self.emit_u8(Op::CALL_REF, 0);
                }
                if saved_js_this.is_some() {
                    let result_slot = self.define_local("__js_call_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_js_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                }
                return Ok(());
            }
        }

        // ── Component Model instance-method dispatch ────────────────
        //
        // When `obj` is a local with a known .NET type (from
        // `Dim d As New Dictionary(...)` / `var x : Stack` / etc.),
        // resolve the method against the auto-built component
        // descriptor and emit the import call directly. This is the
        // primary dispatch path per the Component Model + ESM
        // architecture — the .NET adapter at the descriptor level
        // translates `Dictionary.Add` → `ecma:map.set`, so the
        // emitted call hits the standardized primitive without any
        // runtime `__type` lookup. The TypeRegistry-driven runtime
        // dispatch (compilation-hints proposal style) is the
        // fallback for dynamically-typed receivers.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let class_name = resolve_receiver_type_hint(self, object);
            if let Some(class_name) = class_name {
                let class_name = class_name
                    .split('<')
                    .next()
                    .map(str::trim)
                    .unwrap_or(&class_name)
                    .to_string();
                let surface = common::dotnet::surface();
                if let Some(target) = surface.lookup_instance_method(&class_name, field, arg_exprs.len() as u8) {
                    // Compile receiver, then args.
                    self.compile_expr(object)?;
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let total_argc = (arg_exprs.len() + 1) as u8;
                    match target {
                        common::dotnet::InstanceMethodTarget::Host { module, func, .. } => {
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, total_argc);
                        }
                        common::dotnet::InstanceMethodTarget::Common { emit, .. } => {
                            let line = self.line;
                            self.emit_common(&emit, total_argc, line);
                        }
                    }
                    return Ok(());
                }
            }
        }

        // ── Value method: obj.toUpperCase() ─────────────────────────
        //
        // Method name shadowing rule: a value method (e.g. `Array.push`,
        // `String.toUpperCase`) is the default for *member-access*
        // receivers like `this.items.push(x)` — the receiver is
        // structurally a property, almost certainly a built-in collection.
        //
        // For *direct* receivers (`this`, `super`, or a local variable
        // by name), if the field is a known user-class method, prefer
        // the user method via the generic call path. That preserves
        // user overrides like `class Stack { push(x) { ... } }` and
        // `class Holder { size() { ... } }` against built-in
        // `Array.push`/`map_size` shadowing.
        //
        // This is a heuristic — the cleaner fix is per-class method sets
        // plus receiver-type inference, tracked in the user's pending
        // "JS/C# compilers don't use common::classes" migration.
        if let ExprKind::Member { object, field, null_safe } = &callee.kind {
            let canon_field = self.canon(field);
            let receiver_is_direct = matches!(
                object.kind,
                ExprKind::This | ExprKind::Super | ExprKind::Ident(_)
            );
            if self.is_python_profile() && arg_exprs.is_empty() {
                if let ExprKind::Lit(Literal::Str(value)) = &object.kind {
                    match field.as_str() {
                        "isidentifier" => {
                            self.emit_const(Value::Bool(python_is_identifier_literal(value.as_ref())));
                            return Ok(());
                        }
                        "isprintable" => {
                            self.emit_const(Value::Bool(python_is_printable_literal(value.as_ref())));
                            return Ok(());
                        }
                        _ => {}
                    }
                }
            }
            // Skip value-method dispatch on null-safe member calls — the
            // null short-circuit must run BEFORE we apply any built-in
            // operator (e.g. `null?.toUpperCase()` returns null, not "").
            // Falls through to the generic Member-access path which
            // handles null_safe correctly.
            let matched_value_method = if *null_safe {
                None
            } else {
                self.profile.lookup_value_method(field, arg_exprs.len() as u8).cloned()
            };
            let prefer_string_stdlib_value_method = matches!(
                matched_value_method.as_ref().map(|d| &d.emit),
                Some(BuiltinEmit::Stdlib(_))
            ) && self.expr_is_known_string_receiver(object);
            // Keep dotnet adapter value-methods ahead of runtime collection
            // dispatch for untyped receivers (notably plain arrays using
            // LINQ-style extension methods like Select/SelectMany).
            let prefer_dotnet_adapter = match matched_value_method.as_ref().map(|d| &d.emit) {
                Some(BuiltinEmit::Common(name)) => name.starts_with("dotnet."),
                _ => false,
            };
            let user_method_shadow = receiver_is_direct
                && self.defined_class_methods.contains(&canon_field);
            // Also skip value_methods if the field is an array HOF method —
            // the array_methods dispatch handles it with proper HOF semantics.
            // Without this, `[1,2,3].includes(2)` routes through the string
            // `includes` value method instead of the array contains HOF.
            let field_lower_check = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            let is_array_method = self.profile.lookup_array_method(&field_lower_check).is_some();
            if user_method_shadow || is_array_method {
                // Fall through — let the HOF dispatch or generic call path handle it
            } else if self.profile.namespaces.use_dotnet
                && common::dotnet::uses_runtime_collection_dispatch_arity(field, arg_exprs.len() as u8)
                && !prefer_string_stdlib_value_method
                && !prefer_dotnet_adapter
            {
                // Let the generic member-call path consult the runtime type
                // registry for shared .NET collection methods instead of
                // intercepting them via language profile value-method tables.
            } else if let Some(def) = matched_value_method {
                // For Stdlib calls, push func ref BEFORE args (call_ref expects [func, args...])
                if let BuiltinEmit::Stdlib(stdlib_name) = &def.emit {
                    let global_name = format!("__vybe_{}", stdlib_name);
                    let name_idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, name_idx);
                    self.compile_expr(object)?;
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
                // Object is first arg, then explicit args
                self.compile_expr(object)?;
                for a in &arg_exprs { self.compile_expr(a)?; }
                // Some opcodes need default args when called with fewer
                // than required. Push defaults here.
                if let BuiltinEmit::Opcode(op) | BuiltinEmit::Common(op) = &def.emit {
                    match op.as_str() {
                        // array_join / collections.join needs [arr, sep]
                        "array_join" | "collections.join" if arg_exprs.is_empty() => {
                            self.emit_const(Value::String(Arc::from(",")));
                        }
                        // array_fill needs [arr, val, start, end]
                        "array_fill" if arg_exprs.len() < 2 => {
                            // Push start=0 and end=arr.length defaults
                            if arg_exprs.is_empty() {
                                self.emit(Op::NULL); // val
                            }
                            self.emit(Op::I32_CONST_0); // start
                            self.emit_const(Value::I32(i32::MAX)); // end (clamped by VM)
                        }
                        // C# `s.Substring(start)` — 1-arg form means
                        // "from start to end of string". STR_SUBSTRING
                        // wants `[s, start, end]`; default end to a
                        // sentinel large value (VM clamps to s.len()).
                        // Same shape applies to ECMA-262 §22.1.3.16
                        // `String.prototype.slice(start)`.
                        "strings.substring" | "strings.slice"
                            if arg_exprs.len() < 2 => {
                            self.emit_const(Value::I32(i32::MAX));
                        }
                        // C#'s `string.ToCharArray()` lowers to STR_SPLIT
                        // which needs a delimiter on the stack. The .NET
                        // semantics ("each char one element") match
                        // splitting on the empty string.
                        "str_split" if arg_exprs.is_empty() => {
                            self.emit_const(Value::String(Arc::from("")));
                        }
                        _ => {}
                    }
                }
                match &def.emit {
                    BuiltinEmit::HostCall(module, func) => {
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, (arg_exprs.len() + 1) as u8);
                    }
                    BuiltinEmit::Opcode(op_name) => {
                        // Object + args already on stack from above
                        self.emit_named_opcode(op_name);
                    }
                    BuiltinEmit::StrLength => {
                        let line = self.line;
                        common::strings::emit_length(self.chunk(), line);
                    }
                    BuiltinEmit::Common(name) => {
                        let line = self.line;
                        let name = name.clone();
                        self.emit_common(&name, (arg_exprs.len() + 1) as u8, line);
                    }
                    BuiltinEmit::Invoke(method_name) => {
                        let line = self.line;
                        let name = method_name.clone();
                        common::invoke::emit_invoke_method(
                            &mut self.chunks,
                            self.current,
                            &name,
                            arg_exprs.len() as u8,
                            line,
                        );
                    }
                    _ => {}
                }
                return Ok(());
            }


            // Array higher-order methods: arr.map(fn), arr.filter(fn), etc.
            // Use compiler_common::loops which emits proper loop bytecode.
            // BUT: skip when the same name is a user-defined class method
            // (e.g. `QueryBuilder.Where(string)` shouldn't be intercepted
            // by the LINQ HOF dispatch). The compiler can't see receiver
            // types at compile time, but it knows what method names user
            // classes have declared.
            let field_lower = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            let canon_field_for_user_check = self.canon(field);
            let user_class_method = self.defined_class_methods.contains(&canon_field_for_user_check);
            if !user_class_method
                && self.profile.lookup_array_method(&field_lower).is_some()
            {
                // (re-fetch only when we're committed to the HOF path so
                // the method name lookup matches the previous behaviour)
            }
            if let Some(stdlib_name) = self.profile.lookup_array_method(&field_lower)
                .filter(|_| !user_class_method)
                .map(|s| s.to_string())
            {
                // Normalize to the JS-style method name used in match below
                let field_lower = match stdlib_name.as_str() {
                    "__array_map" => "map".to_string(),
                    "__array_filter" => "filter".to_string(),
                    "__array_forEach" => "forEach".to_string(),
                    "__array_reduce" => "reduce".to_string(),
                    "__array_find" => "find".to_string(),
                    "__array_sort" => "sort".to_string(),
                    "__array_sort_by_key" => "sort_by_key".to_string(),
                    "__array_some" => "some".to_string(),
                    "__array_every" => "every".to_string(),
                    "__array_flat_map" => "flatMap".to_string(),
                    "__array_reduce_right" => "reduceRight".to_string(),
                    _ => field_lower,
                };
                // Compile arr and fn(s) into local slots
                self.compile_expr(object)?;
                let arr_slot = self.define_local("__hof_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

                if let Some(fn_expr) = arg_exprs.first() {
                    self.compile_expr(fn_expr)?;
                } else {
                    self.emit(Op::NULL);
                }
                let fn_slot = self.define_local("__hof_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);

                let idx_slot = self.define_local("__hof_idx");
                let result_slot = self.define_local("__hof_result");
                let line = self.line;

                match field_lower.as_str() {
                    "map" => {
                        // emit_map leaves result on stack
                        common::loops::emit_map(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, line);
                    }
                    "filter" => {
                        let elem_slot = self.define_local("__hof_elem");
                        common::loops::emit_filter(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, elem_slot, line);
                    }
                    "reduce" => {
                        // reduce(fn, initial?) — initial is second arg.
                        // When initial IS provided, start from i=0 with
                        // acc=initial. emit_reduce always starts from
                        // i=1 with acc=arr[0], so we only use it for
                        // the no-initial case.
                        if let Some(init_expr) = arg_exprs.get(1) {
                            // acc = initial, i = 0
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                            // Inline reduce loop starting from i=0
                            self.emit(Op::I32_CONST_0);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            let loop_start = self.chunks[self.current].current_offset();
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                            self.emit(Op::DYN_LT);
                            let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                            // acc = fn(acc, arr[i])
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            self.emit_u8(Op::CALL_REF, 2);
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                            // i++
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_const(Value::I32(1));
                            self.emit(Op::DYN_ADD);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            self.emit_loop(loop_start);
                            self.patch_jump(exit_jump);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else {
                            // No initial: emit_reduce starts from arr[0], i=1
                            common::loops::emit_reduce(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, line);
                        }
                    }
                    "forEach" | "foreach" => {
                        // Polymorphic forEach: arrays iterate by index,
                        // Maps iterate (val, key, map) per ECMA-262
                        // §24.1.3.5, Sets iterate (val, val, set). The
                        // compiler can't know the receiver type so route
                        // through `ecma:value.invokeMethod` (each impl
                        // is in dispatch_{array,map,set}). For non-JS
                        // profiles, keep the array-only stdlib loop —
                        // PHP / VB iteration semantics differ.
                        if self.is_js_profile() {
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            common::invoke::emit_invoke_method(
                                &mut self.chunks,
                                self.current,
                                "forEach",
                                1,
                                line,
                            );
                            self.emit(Op::DROP); // forEach returns undefined
                        } else {
                            common::loops::emit_foreach(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, line);
                        }
                    }
                    "some" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, true, line);
                    }
                    "every" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, false, line);
                    }
                    "find" => {
                        // find uses includes pattern but returns element not bool.
                        // JS spec §23.1.3.10: returns undefined when no match;
                        // other languages stick with Null for cross-compat
                        // (Python None / VB Nothing / .NET null match Null).
                        if self.is_js_profile() {
                            self.emit(Op::UNDEFINED);
                        } else {
                            self.emit(Op::NULL);
                        }
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks, self.current, arr_slot, idx_slot, line);
                        let elem_slot = self.define_local("__find_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, lp, line);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findIndex" | "findindex" => {
                        // findIndex: like find but returns the index, not the element
                        self.emit_const(Value::I32(-1));
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks, self.current, arr_slot, idx_slot, line);
                        let elem_slot = self.define_local("__findi_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, lp, line);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "includes" => {
                        // `x.includes(v[, fromIndex])` — polymorphic:
                        // arrays do element membership, strings do
                        // substring search starting from fromIndex,
                        // user objects fall through to their own
                        // method. Route through `ecma:value.invokeMethod`
                        // so emitted wasm stays spec-compliant.
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        // Pass remaining args (fromIndex etc.) directly
                        // — fn_slot already holds args[0].
                        for extra in arg_exprs.iter().skip(1) {
                            self.compile_expr(extra)?;
                        }
                        common::invoke::emit_invoke_method(
                            &mut self.chunks,
                            self.current,
                            "includes",
                            arg_exprs.len() as u8,
                            line,
                        );
                    }
                    "sort" => {
                        // JS sort(comparatorFn?) — 2-arg comparator or default
                        // ECMA-262 §23.1.3.30: default comparator is
                        // ToString-based ("10" < "2"), not numeric.
                        // Comparator path uses the stdlib (works for JS
                        // and for all other languages); no-comparator JS
                        // routes to ecma:array.sort which does the
                        // spec-compliant lexicographic sort. Other
                        // languages keep stdlib's numeric default.
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let no_fn = self.emit_jump(Op::BR_IF_TRUE);
                        let global = self.str_const("__vybe_sort_with_comparator");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(no_fn);
                        if self.is_js_profile() {
                            // ecma:array.sort returns the sorted array
                            // (in-place, returns receiver). One-arg call.
                            let idx = self.import("ecma:array", "sort");
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_host_call(idx, 1);
                        } else {
                            let sort_global = self.str_const("__vybe_sort_in_place");
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u8(Op::CALL_REF, 1);
                        }
                        self.patch_jump(done);
                    }
                    "sort_by_key" => {
                        // .NET OrderBy(keySelector) — 1-arg key extractor
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let no_fn = self.emit_jump(Op::BR_IF_TRUE);
                        let global = self.str_const("__vybe_sort_by_key");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(no_fn);
                        let sort_global = self.str_const("__vybe_sort_in_place");
                        self.emit_u16(Op::GLOBAL_GET, sort_global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.patch_jump(done);
                    }
                    "indexOf" | "indexof" => {
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot); // search value
                        common::collections::emit_index_of(&mut self.chunks, self.current, line);
                    }
                    "flatMap" | "flatmap" => {
                        // arr.flatMap(fn) = arr.map(fn).flat()
                        // First emit map: result[i] = fn(arr[i])
                        let mapped_slot = self.define_local("__flatmap_mapped");
                        common::loops::emit_map(&mut self.chunks, self.current, fn_slot, arr_slot, mapped_slot, idx_slot, line);
                        // Now the mapped array is on stack. Flatten it one level.
                        let flat_idx = self.import("ecma:array", "flat");
                        self.emit_const(Value::I32(1));  // depth = 1
                        self.emit_host_call(flat_idx, 2);
                    }
                    "reduceRight" | "reduceright" => {
                        // reduceRight(fn, initial?) — iterate from end to start.
                        if let Some(init_expr) = arg_exprs.get(1) {
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        } else {
                            // acc = arr[len-1]
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        }
                        // Start from len-1 (or len-2 if no initial)
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        if arg_exprs.get(1).is_none() {
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                        }
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let loop_start = self.chunks[self.current].current_offset();
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                        // acc = fn(acc, arr[i])
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u8(Op::CALL_REF, 2);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        // i--
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit_jump);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findLast" | "findlast" => {
                        // Iterate backward, return last element matching predicate
                        self.emit(Op::NULL);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let loop_start = self.chunks[self.current].current_offset();
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                        let elem_slot = self.define_local("__fl_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit_jump);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findLastIndex" | "findlastindex" => {
                        // Iterate backward, return last index matching predicate
                        self.emit_const(Value::I32(-1));
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let loop_start = self.chunks[self.current].current_offset();
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                        let elem_slot2 = self.define_local("__fli_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u16(Op::LOCAL_SET, elem_slot2); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot2);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit_jump);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "removeAll" | "removeall" => {
                        // Iterate backward over arr, splice each matching element.
                        // Returns count of removed items.
                        let removed_slot = self.define_local("__ra_removed");
                        self.emit_const(Value::I32(0));
                        self.emit_u16(Op::LOCAL_SET, removed_slot); self.emit(Op::DROP);
                        // Start i = arr.len - 1
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let ra_loop = self.chunks[self.current].current_offset();
                        // while i >= 0
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let ra_exit = self.emit_jump(Op::BR_IF_FALSE);
                        // elem = arr[i]
                        let ra_elem = self.define_local("__ra_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u16(Op::LOCAL_SET, ra_elem); self.emit(Op::DROP);
                        // if fn(elem) → remove
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, ra_elem);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let ra_skip = self.emit_jump(Op::BR_IF_FALSE);
                        // splice(arr, i, 1) → drop removed array
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_remove_at(&mut self.chunks, self.current, l); }
                        self.emit(Op::DROP);
                        // removed++
                        self.emit_u16(Op::LOCAL_GET, removed_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::DYN_ADD);
                        self.emit_u16(Op::LOCAL_SET, removed_slot); self.emit(Op::DROP);
                        self.patch_jump(ra_skip);
                        // i--
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(ra_loop);
                        self.patch_jump(ra_exit);
                        self.emit_u16(Op::LOCAL_GET, removed_slot);
                    }
                    _ => {
                        // Fallback: call as regular method
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                    }
                }
                return Ok(());
            }
        }

        // ── Constructor call: ClassName.Create(args) ────────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(class_name) = &object.kind {
                let ctor_nm = &self.profile.constructor_name.clone();
                let is_ctor = if self.case_sensitive { field == ctor_nm } else { field.eq_ignore_ascii_case(ctor_nm) };
                let canon_class = self.canon(class_name);
                let is_known_class = self.defined_classes.contains(&canon_class)
                    && self.scope().resolve(class_name).is_none();
                if is_ctor && is_known_class {
                    self.emit_var_get(class_name);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Pascal builtin helper dispatch: value.Helper(args) ───────
        if self.profile.name == "pascal" {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let Some(type_name) = self.pascal_expr_static_type(object) {
                    let helper_name = self.pascal_helper_function_name(&type_name, field);
                    let helper_canon = self.canon(&helper_name);
                    if self.defined_functions.contains(&helper_canon) {
                        self.emit_var_get(&helper_name);
                        self.compile_expr(object)?;
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        return Ok(());
                    }

                    let canon_type = self.canon(&type_name);
                    let canon_field = self.canon(field);
                    let is_callable_field = self.pending_classes.get(canon_type.as_str())
                        .map(|pc| pc.fields.iter().any(|name| name == &canon_field))
                        .unwrap_or(false);
                    if is_callable_field {
                        self.compile_expr(object)?;
                        let obj_tmp = self.define_local("__pascal_callable_field_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                        let prop = self.str_const(&canon_field);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        return Ok(());
                    }
                }
            }
        }

        // ── Method call: obj.method(args) ───────────────────────────
        if let ExprKind::Member { object, field, null_safe } = &callee.kind {
            if self.is_js_profile() {
                self.compile_expr(object)?;
                let obj_tmp = self.define_local("__js_obj");
                self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                let method_name = self.canon(field);
                let prop = self.str_const(&method_name);
                let receiver_marker = self.str_const("__vybe_method_receiver");

                // Generator `.return(v)`: ECMA-262 §27.5.1.4 — terminate
                // the generator and yield `{value: v, done: true}`. We
                // stamp `__vybe_gen_returned = true` on the cont so
                // subsequent `next()` calls short-circuit to
                // `{value: undefined, done: true}`. Pure compiler
                // bookkeeping — no VM-side state mutation needed.
                let gen_return_skip_patch = if !*null_safe && method_name == "return" && arg_exprs.len() <= 1 {
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                    // Stamp __vybe_gen_returned = true on the cont.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::TRUE);
                    let returned_key = self.str_const("__vybe_gen_returned");
                    self.emit_u16(Op::STRUCT_SET, returned_key);
                    self.emit(Op::DROP);
                    // Build { value: v, done: true }.
                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    self.emit(Op::DUP);
                    if arg_exprs.is_empty() {
                        self.emit(Op::UNDEFINED);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    self.emit(Op::DUP);
                    self.emit(Op::TRUE);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    let skip = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    Some(skip)
                } else { None };

                // Generator `.next()` / `.next(v)`: if receiver is a
                // Continuation, drive via WASM stack-switching opcodes
                // and wrap into spec `{value, done}`.
                //   - `g.next()`     → Op::GEN_NEXT (pushes value+has_more)
                //   - `g.next(v)`    → Op::RESUME with v as resume_val
                //                       (pushes yielded value), then
                //                       check `isGeneratorDone` for the
                //                       done flag.
                // Non-Continuations (Array iterators, custom iterables)
                // fall through to regular method dispatch below.
                let gen_next_skip_patch = if !*null_safe && method_name == "next" && arg_exprs.len() <= 1 {
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                    let value_slot = self.define_local("__gen_value");
                    let done_slot = self.define_local("__gen_done");
                    // If a previous `.return()` stamped the cont as
                    // returned, short-circuit to `{value: undefined,
                    // done: true}` per ECMA-262 §27.5.1.2 step 2.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let returned_key2 = self.str_const("__vybe_gen_returned");
                    self.emit_u16(Op::STRUCT_GET, returned_key2);
                    self.emit(Op::DYN_TO_BOOL);
                    let not_returned = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::UNDEFINED);
                    self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                    self.emit(Op::TRUE);
                    self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);
                    let after_returned_branch = self.emit_jump(Op::BR);
                    self.patch_jump(not_returned);
                    if arg_exprs.is_empty() {
                        // `g.next()` — GEN_NEXT path: pushes value+has_more.
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit(Op::GEN_NEXT);
                        let has_more_slot = self.define_local("__gen_has_more");
                        self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, has_more_slot);
                        self.emit(Op::DYN_TO_BOOL);
                        self.emit(Op::DYN_NOT);
                        self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);
                    } else {
                        // `g.next(v)` — RESUME with the resume value;
                        // the suspended yield expression evaluates to
                        // `v`. Pushes only the yielded value back; we
                        // query `isGeneratorDone` for the spec `done`.
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.compile_expr(&arg_exprs[0])?;
                        self.emit_u16(Op::RESUME, 0);
                        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                        self.emit_host_call(is_done_idx, 1);
                        self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);
                    }
                    // Both the early-`returned` short-circuit and the
                    // GEN_NEXT/RESUME paths converge here to build the
                    // `{value, done}` wrapper.
                    self.patch_jump(after_returned_branch);
                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, done_slot);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    let skip = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    Some(skip)
                } else { None };
                let _ = gen_next_skip_patch;
                let _ = gen_return_skip_patch;
                // gen_next_skip_patch / gen_return_skip_patch are
                // patched at the end of the JS method dispatch (when
                // result is on stack and we'd otherwise `return Ok(())`).

                if *null_safe {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    let fn_slot = self.define_local("__js_method_fn");
                    self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit_u16(Op::STRUCT_GET, receiver_marker);
                    self.emit(Op::REF_IS_NULL);
                    let use_js_path = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    let typed_done = self.emit_jump(Op::BR);

                    self.patch_jump(use_js_path);
                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit(Op::REF_IS_NULL);
                    let need_lookup = self.emit_jump(Op::BR_IF_TRUE);

                    let saved_js_this = self.save_js_this("__js_prev_this_method");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    let result_slot = self.define_local("__js_method_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_js_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    let js_done = self.emit_jump(Op::BR);

                    self.patch_jump(need_lookup);
                    let lookup = self.import("ecma:value", "getMethodForCall");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(method_name.as_str())));
                    self.emit_host_call(lookup, 2);
                    let lookup_slot = self.define_local("__js_lookup_fn");
                    self.emit_u16(Op::LOCAL_SET, lookup_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, lookup_slot);
                    self.emit(Op::REF_IS_NULL);
                    let have_fn = self.emit_jump(Op::BR_IF_FALSE);
                    let invoke = self.import("ecma:value", "invokeMethod");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(method_name.as_str())));
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_host_call(invoke, (arg_exprs.len() + 2) as u8);
                    let after_call = self.emit_jump(Op::BR);
                    self.patch_jump(have_fn);
                    let saved_js_this = self.save_js_this("__js_prev_this_lookup");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, lookup_slot);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    let result_slot = self.define_local("__js_lookup_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_js_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.patch_jump(after_call);
                    self.patch_jump(js_done);
                    self.patch_jump(typed_done);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(skip);
                    self.emit(Op::NULL);
                    self.patch_jump(end);
                    return Ok(());
                }

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_slot = self.define_local("__js_method_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, fn_slot);
                self.emit_u16(Op::STRUCT_GET, receiver_marker);
                self.emit(Op::REF_IS_NULL);
                let use_js_path = self.emit_jump(Op::BR_IF_TRUE);

                self.emit_u16(Op::LOCAL_GET, fn_slot);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                let typed_done = self.emit_jump(Op::BR);

                self.patch_jump(use_js_path);
                self.emit_u16(Op::LOCAL_GET, fn_slot);
                self.emit(Op::REF_IS_NULL);
                let need_lookup = self.emit_jump(Op::BR_IF_TRUE);

                let saved_js_this = self.save_js_this("__js_prev_this_method");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                self.emit_u16(Op::LOCAL_GET, fn_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let result_slot = self.define_local("__js_method_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                let js_done = self.emit_jump(Op::BR);

                self.patch_jump(need_lookup);
                let lookup = self.import("ecma:value", "getMethodForCall");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::String(Arc::from(method_name.as_str())));
                self.emit_host_call(lookup, 2);
                let lookup_slot = self.define_local("__js_lookup_fn");
                self.emit_u16(Op::LOCAL_SET, lookup_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, lookup_slot);
                self.emit(Op::REF_IS_NULL);
                let have_fn = self.emit_jump(Op::BR_IF_FALSE);
                let invoke = self.import("ecma:value", "invokeMethod");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::String(Arc::from(method_name.as_str())));
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_host_call(invoke, (arg_exprs.len() + 2) as u8);
                let after_call = self.emit_jump(Op::BR);
                self.patch_jump(have_fn);
                let saved_js_this = self.save_js_this("__js_prev_this_lookup");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                self.emit_u16(Op::LOCAL_GET, lookup_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let result_slot = self.define_local("__js_lookup_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.patch_jump(after_call);
                self.patch_jump(js_done);
                self.patch_jump(typed_done);
                if let Some(skip) = gen_next_skip_patch {
                    self.patch_jump(skip);
                }
                if let Some(skip) = gen_return_skip_patch {
                    self.patch_jump(skip);
                }
                return Ok(());
            }

            self.compile_expr(object)?;
            let obj_tmp = self.define_local("__obj");
            self.reserve_local_slot(obj_tmp);
            self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

            let field_name = self.canon(field);
            let prop = self.str_const(&field_name);

            if *null_safe {
                // obj?.method() — short-circuit to null if obj is null/undefined.
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit(Op::REF_IS_NULL);
                let obj_not_null = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let end = self.emit_jump(Op::BR);
                self.patch_jump(obj_not_null);
                if field.eq_ignore_ascii_case("Invoke") {
                    // C# delegate null-conditional invocation: `d?.Invoke(args)`
                    // should call the delegate value directly when non-null.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    self.patch_jump(end);
                    return Ok(());
                }
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                self.patch_jump(end);
                return Ok(());
            }

            if self.is_php_profile() {
                let is_php_generator_method = (field_name == "current" && arg_exprs.is_empty())
                    || (field_name == "send" && arg_exprs.len() == 1)
                    || (field_name == "next" && arg_exprs.is_empty())
                    || (field_name == "valid" && arg_exprs.is_empty())
                    || (field_name == "getReturn" && arg_exprs.is_empty());

                if is_php_generator_method {
                let started_key = self.str_const("__php_gen_started");
                let current_key = self.str_const("__php_gen_current");
                let done_key = self.str_const("__php_gen_done");
                let return_key = self.str_const("__php_gen_return");

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let is_gen_idx = self.import("ecma:value", "isGenerator");
                self.emit_host_call(is_gen_idx, 1);
                let not_gen = self.emit_jump(Op::BR_IF_FALSE);

                    match field_name.as_str() {
                        "getReturn" => {
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, return_key);
                        }
                        "valid" => {
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, started_key);
                            self.emit(Op::DYN_TO_BOOL);
                            let need_start = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, done_key);
                            self.emit(Op::DYN_TO_BOOL);
                            self.emit(Op::DYN_NOT);
                            let handled = self.emit_jump(Op::BR);

                            self.patch_jump(need_start);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit(Op::GEN_NEXT);
                            let has_more_slot = self.define_local("__php_gen_has_more");
                            self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                            let value_slot = self.define_local("__php_gen_value");
                            self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::STRUCT_SET, started_key);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, has_more_slot);
                            self.emit(Op::DYN_TO_BOOL);
                            let no_more = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_const(Value::Bool(true));
                            let start_done = self.emit_jump(Op::BR);

                            self.patch_jump(no_more);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_u16(Op::STRUCT_SET, return_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_const(Value::Bool(false));

                            self.patch_jump(handled);
                            self.patch_jump(start_done);
                        }
                        "current" => {
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, started_key);
                            self.emit(Op::DYN_TO_BOOL);
                            let need_start = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, done_key);
                            self.emit(Op::DYN_TO_BOOL);
                            let not_done = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit_const(Value::Bool(false));
                            let current_done = self.emit_jump(Op::BR);

                            self.patch_jump(not_done);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, current_key);
                            let handled = self.emit_jump(Op::BR);

                            self.patch_jump(need_start);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit(Op::GEN_NEXT);
                            let has_more_slot = self.define_local("__php_gen_has_more");
                            self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                            let value_slot = self.define_local("__php_gen_value");
                            self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::STRUCT_SET, started_key);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, has_more_slot);
                            self.emit(Op::DYN_TO_BOOL);
                            let no_more = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            let start_done = self.emit_jump(Op::BR);

                            self.patch_jump(no_more);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_u16(Op::STRUCT_SET, return_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_const(Value::Bool(false));

                            self.patch_jump(current_done);
                            self.patch_jump(handled);
                            self.patch_jump(start_done);
                        }
                        "send" | "next" => {
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, started_key);
                            self.emit(Op::DYN_TO_BOOL);
                            let need_start = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, done_key);
                            self.emit(Op::DYN_TO_BOOL);
                            let can_resume = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit_const(Value::Bool(false));
                            let done_already = self.emit_jump(Op::BR);

                            self.patch_jump(can_resume);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            if field_name == "send" {
                                self.compile_expr(arg_exprs[0])?;
                            } else {
                                self.emit(Op::NULL);
                            }
                            self.emit_u16(Op::RESUME, 0);
                            let value_slot = self.define_local("__php_gen_resume_value");
                            self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                            self.emit_host_call(is_done_idx, 1);
                            let yielded = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_u16(Op::STRUCT_SET, return_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            let handled = self.emit_jump(Op::BR);

                            self.patch_jump(yielded);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            let resume_done = self.emit_jump(Op::BR);

                            self.patch_jump(need_start);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit(Op::GEN_NEXT);
                            let has_more_slot = self.define_local("__php_gen_has_more");
                            self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                            let start_value_slot = self.define_local("__php_gen_value");
                            self.emit_u16(Op::LOCAL_SET, start_value_slot); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::STRUCT_SET, started_key);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, has_more_slot);
                            self.emit(Op::DYN_TO_BOOL);
                            let start_no_more = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_value_slot);
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, start_value_slot);
                            let start_done = self.emit_jump(Op::BR);

                            self.patch_jump(start_no_more);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::STRUCT_SET, done_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_value_slot);
                            self.emit_u16(Op::STRUCT_SET, return_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_const(Value::Bool(false));
                            self.emit_u16(Op::STRUCT_SET, current_key);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, start_value_slot);

                            self.patch_jump(done_already);
                            self.patch_jump(handled);
                            self.patch_jump(resume_done);
                            self.patch_jump(start_done);
                        }
                        _ => unreachable!(),
                    }

                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    let fn_tmp = self.define_local("__fn");
                    self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    let generic_done = self.emit_jump(Op::BR);

                    self.patch_jump(end);
                    self.patch_jump(generic_done);
                    return Ok(());
                }
            }

            self.emit_u16(Op::LOCAL_GET, obj_tmp);
            self.emit_u16(Op::STRUCT_GET, prop);
            let fn_tmp = self.define_local("__fn");
            self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, fn_tmp);
            self.emit_u16(Op::LOCAL_GET, obj_tmp);
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
            return Ok(());
        }

        // ── Simple call: name(args) / expr(args) ────────────────────
        if let ExprKind::Ident(name) = &callee.kind {
            // Inside a class: bare call to a static method should bind to
            // the class object before any generic function lookup. Static
            // methods are also registered as ordinary functions, so this
            // must run ahead of `is_known_func`.
            if self.current_class.is_some() && self.current_class_implicit_self {
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());
                if !is_local {
                    if let Some(class_name) = self.is_class_static_method(name) {
                        let cls_idx = self.str_const(&class_name);
                        self.emit_u16(Op::GLOBAL_GET, cls_idx);
                        let method_idx = self.str_const(&self.canon(name));
                        self.emit_u16(Op::STRUCT_GET, method_idx);
                        let fn_tmp = self.define_local("__bare_static_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let cls_tmp = self.define_local("__bare_static_cls");
                        self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        self.emit_u16(Op::LOCAL_GET, cls_tmp);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        return Ok(());
                    }
                }
            }

            let is_known_func = self.defined_functions.contains(name)
                || (!self.case_sensitive && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
            if !is_known_func && self.try_compile_builtin(name, &arg_exprs)? {
                return Ok(());
            }

            // VB array access: `arr(idx)` when `arr` is a known data variable
            // (local OR top-level global from `Dim arr(5)`) and is NOT a
            // declared function or class. VB syntactically overloads `()` for
            // both calls and indexing — the disambiguator is whether the head
            // is a callable function or a value. We must exclude both
            // `defined_functions` and `defined_classes` from the "looks like
            // a variable" set, otherwise `GetResult()` (function call) and
            // `New Result()` (class) would be mis-identified as indexing.
            if !is_known_func && arg_exprs.len() == 1 && self.profile.parens_for_index {
                let canon_name = self.canon(name);
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());
                let is_global_var = self.defined_globals.contains(&canon_name)
                    && !self.defined_classes.contains(&canon_name)
                    && !self.defined_functions.contains(&canon_name);
                let is_callable_typed = self
                    .lookup_var_type_hint(name)
                    .is_some_and(Self::is_callable_type_hint);
                if (is_local || is_global_var) && !is_callable_typed {
                    self.emit_var_get(name);
                    self.compile_expr(arg_exprs[0])?;
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    return Ok(());
                }
            }

            // Inside a class: bare method call → Me.method(args)
            // If name isn't a local variable and we're inside a class body,
            // resolve as Me.name() (implicit self for method calls).
            if self.current_class.is_some() && self.current_class_implicit_self {
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());
                if !is_local && !is_known_func {
                    if self.emit_self_ref() {
                        // Me.name(args) → load Me, dup, struct_get(name).
                        // Real methods receive `this`/Self as arg0, but callable
                        // fields (Pascal procedure/function members) should be
                        // invoked as plain function values.
                        let field_name = self.canon(name);
                        let prop = self.str_const(&field_name);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.define_local("__bare_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let obj_tmp = self.define_local("__bare_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                        if self.profile.name == "pascal" && self.is_class_field(name) {
                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                            return Ok(());
                        }

                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        return Ok(());
                    }
                }
            }

            let has_spread = args.iter().any(|a| a.spread);
            if has_spread {
                // Spread args: build a flat args array, then spread onto
                // stack and call. Stash the accumulator in a local so
                // `ecma:array.push` (returns new length per
                // ECMA-262) and `ecma:array.concat` (returns new
                // array) can both drive the same pattern.
                let line = self.line;
                let args_slot = self.define_local("__spread_args");
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                let mut known_len: Option<usize> = Some(0);
                for a in args {
                    if a.spread {
                        // new_arr = concat(args, spread)
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_concat(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                        if let ExprKind::Array(elems) = &a.value.kind {
                            if let Some(ref mut k) = known_len { *k += elems.len(); }
                        } else {
                            known_len = None;
                        }
                    } else {
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP); // drop new_length returned by push
                        if let Some(ref mut k) = known_len { *k += 1; }
                    }
                }
                if let Some(arity) = known_len {
                    // All spread sources are statically-sized; safe to
                    // expand on the stack and use a normal CALL_REF.
                    self.emit_var_get(name);
                    self.emit_u16(Op::LOCAL_GET, args_slot);
                    self.emit(Op::SPREAD);
                    self.emit_u8(Op::CALL_REF, arity as u8);
                } else {
                    // Unknown runtime length → use the VM's variadic
                    // calling convention: pad the args array to a fixed
                    // MAX_VARIADIC=16 with NULLs, then SPREAD + CALL_REF
                    // 16. Non-variadic callees see arity > formal, the
                    // VM truncates excess; variadic callees (arity=255)
                    // build their rest array by scanning slots until
                    // they hit a NULL — same convention used for
                    // Python `*args` and JS rest params (see the
                    // rest-arr preamble in `compile_function_decl`).
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, args_slot);   // [args]
                    self.emit_const(Value::I32(16));            // [args, 16]
                    common::collections::emit_new_with_length(&mut self.chunks, self.current, line); // [args, pad]
                    common::collections::emit_concat(&mut self.chunks, self.current, line);          // [args++pad]
                    self.emit_const(Value::F64(0.0));
                    self.emit_const(Value::F64(16.0));
                    common::collections::emit_slice(&mut self.chunks, self.current, line);           // [first16]
                    self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);

                    self.emit_var_get(name);                    // [callee]
                    self.emit_u16(Op::LOCAL_GET, args_slot);    // [callee, args16]
                    self.emit(Op::SPREAD);                      // [callee, e0..e15]
                    self.emit_u8(Op::CALL_REF, 16);
                }
                return Ok(());
            }
            if self.is_python_profile() && !is_known_func {
                let callee_slot = self.define_local("__py_call_target");
                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_SET, callee_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let typeof_idx = self.import("ecma:value", "typeof");
                self.emit_host_call(typeof_idx, 1);
                self.emit_const(Value::String(Arc::from("function")));
                self.emit(Op::DYN_EQ);
                let invoke_dunder = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let end = self.emit_jump(Op::BR);

                self.patch_jump(invoke_dunder);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let call_prop = self.str_const("call");
                self.emit_u16(Op::STRUCT_GET, call_prop);
                let call_slot = self.define_local("__py_call_method");
                self.emit_u16(Op::LOCAL_SET, call_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, call_slot);
                self.emit(Op::REF_IS_NULL);
                let try_dunder_name = self.emit_jump(Op::BR_IF_TRUE);
                self.emit_u16(Op::LOCAL_GET, call_slot);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                let found_end = self.emit_jump(Op::BR);

                self.patch_jump(try_dunder_name);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let dunder_prop = self.str_const("__call__");
                self.emit_u16(Op::STRUCT_GET, dunder_prop);
                let dunder_slot = self.define_local("__py_dunder_call_method");
                self.emit_u16(Op::LOCAL_SET, dunder_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, dunder_slot);
                self.emit(Op::REF_IS_NULL);
                let no_dunder = self.emit_jump(Op::BR_IF_TRUE);
                self.emit_u16(Op::LOCAL_GET, dunder_slot);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                let dunder_end = self.emit_jump(Op::BR);

                self.patch_jump(no_dunder);
                self.emit(Op::UNDEFINED);
                self.patch_jump(found_end);
                self.patch_jump(dunder_end);
                self.patch_jump(end);
                return Ok(());
            }

            self.emit_var_get(name);
            if let Some(param_modes) = self.function_param_modes.get(&self.canon(name)).cloned() {
                if param_modes.iter().any(|mode| matches!(mode, PassBy::Ref | PassBy::Out)) {
                    for (index, arg) in args.iter().enumerate() {
                        match param_modes.get(index).copied().unwrap_or(PassBy::Value) {
                            PassBy::Out => self.emit(Op::NULL),
                            PassBy::Ref | PassBy::Const | PassBy::Value => {
                                if !matches!(param_modes.get(index), Some(PassBy::Out)) {
                                    self.compile_expr(&arg.value)?;
                                }
                            }
                        }
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                    let pack_slot = self.define_local("__ref_call_pack");
                    self.emit_u16(Op::LOCAL_SET, pack_slot);
                    self.emit(Op::DROP);
                    let mut ref_out_index = 1usize;
                    for (index, arg) in args.iter().enumerate() {
                        if !matches!(param_modes.get(index), Some(PassBy::Ref | PassBy::Out)) {
                            continue;
                        }
                        self.emit_u16(Op::LOCAL_GET, pack_slot);
                        self.emit_const(Value::F64(ref_out_index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.compile_assign_target(&arg.value)?;
                        ref_out_index += 1;
                    }
                    self.emit_u16(Op::LOCAL_GET, pack_slot);
                    self.emit_const(Value::F64(0.0));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    return Ok(());
                }
            }
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
            return Ok(());
        }

        // ── Computed-member call: `obj[key](args)` ───────────────────
        // For JS profile, treat this like a method call so `__js_this`
        // is bound to `obj` before invocation. Without this binding the
        // callee body sees a stale __js_this and `this.x` traps. Same
        // semantics as ECMA-262 §13.3.7 (CallMemberExpression).
        if self.is_js_profile() {
            if let ExprKind::Index { object, index, .. } = &callee.kind {
                let obj_tmp = self.define_local("__js_idx_obj");
                self.compile_expr(object)?;
                self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                let saved_js_this = self.save_js_this("__js_prev_this_idx");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(index)?;
                let line = self.line;
                common::collections::emit_get(&mut self.chunks, self.current, line);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let result_slot = self.define_local("__js_idx_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                return Ok(());
            }
        }

        // ── Fallback: general expression call ───────────────────────
        self.compile_expr(callee)?;
        for a in &arg_exprs { self.compile_expr(a)?; }
        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Lambda compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_lambda(&mut self, params: &[Param], body: &LambdaBody) -> Result<(), String> {
        let has_rest = params.last().map_or(false, |p| p.is_rest);
        let arity = if has_rest { 255u8 } else { params.len() as u8 };
        let ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk("<lambda>", arity);
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = ci;
        for p in params {
            self.define_local_typed(&p.name, p.type_hint.clone());
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

        // Rest param preamble (same as compile_function_decl).
        // Accumulator pattern: stash arr in rest_slot and reload each
        // iteration so `ecma:array.push` (returns new length per
        // ECMA-262) cleanly drives the push loop.
        if has_rest {
            let rest_name = &params.last().unwrap().name;
            let rest_slot = self.scope().resolve(rest_name).unwrap();
            let line = self.line;
            let rest_arr = self.define_local("__rest_arr");
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            self.emit_u16(Op::LOCAL_SET, rest_arr); self.emit(Op::DROP);
            let max_rest = 16u16;
            let mut done_patches: Vec<usize> = Vec::new();
            for i in 0..max_rest {
                let slot = rest_slot + i;
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                done_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                self.emit_u16(Op::LOCAL_GET, rest_arr);
                self.emit_u16(Op::LOCAL_GET, slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP); // drop new_length
            }
            for p in done_patches { self.patch_jump(p); }
            self.emit_u16(Op::LOCAL_GET, rest_arr);
            self.emit_u16(Op::LOCAL_SET, rest_slot);
            self.emit(Op::DROP);
        }

        // Result slot for ResultSlot languages
        let result_slot = if self.profile.function_return == ReturnStyle::ResultSlot {
            let rs = self.define_local("Result");
            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
            let saved_fn = self.current_func_name.take();
            let saved_rs = self.current_result_slot.take();
            self.current_func_name = Some("<lambda>".into());
            self.current_result_slot = Some(rs);
            Some((rs, saved_fn, saved_rs))
        } else { None };

        match body {
            LambdaBody::Expr(expr) => {
                self.compile_expr(expr)?;
                self.emit(Op::RETURN);
            }
            LambdaBody::Block(stmts) => {
                for s in stmts { self.compile_stmt(s)?; }
            }
        }

        if let Some((rs, saved_fn, saved_rs)) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
            self.current_func_name = saved_fn;
            self.current_result_slot = saved_rs;
        } else if matches!(body, LambdaBody::Block(_)) {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
        }

        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
        self.chunks[ci].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;
        let l = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], ci, uvs.len() as u8, l);
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, l);
            self.chunks[self.current].emit(uv.index, l);
        }
        Ok(())
    }

}
