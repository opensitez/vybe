//! Call-expression compilation — `compile_call` (handles named calls,
//! method calls, super-calls, spread, dotted lookups) and
//! `compile_lambda`. This is the primary edit site for the inline
//! refactor (Phase G) where `wasm:js-*` imports get replaced by
//! inline WASM GC sequences.

use super::*;

impl Compiler {
    // ════════════════════════════════════════════════════════════════════════
    // Call compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<(), String> {
        let arg_exprs: Vec<&Expression> = args.iter().map(|a| &a.value).collect();

        // ── super(args) → call parent constructor, store result as this ──
        if let ExprKind::Super = &callee.kind {
            if let Some(ref class_name) = self.current_class.clone() {
                if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|pc| pc.parent.clone()) {
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

        // ── super.method(args) → this.__base_method(args) ────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if matches!(&object.kind, ExprKind::Super) {
                let base_name = format!("__base_{}", self.canon(field));
                let self_kw = self.profile.self_keyword.clone();
                if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                    let prop = self.str_const(&base_name);
                    self.emit_u16(Op::LOCAL_GET, self_slot);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    // Call with this as first arg
                    self.emit_u16(Op::LOCAL_GET, self_slot);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
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
                            self.emit_common(&emit, line);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::HostCall { module, func } => {
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
                                                    self.emit_common(&name, line);
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
                            // Intercept Thread/Task methods → WASM stack switching opcodes
                            if members.len() == 1 {
                                let method = members[0].as_str();
                                match method {
                                    "join" => {
                                        // th.Join() → thread_join opcode (blocks until thread
                                        // completes, pushes exit code). Leave the exit code on
                                        // stack — the statement wrapper at StmtKind::Expr adds
                                        // its own DROP.
                                        self.emit_var_get(&local);
                                        let line = self.line;
                                        common::threading::emit_thread_join(self.chunk(), line);
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
                return Ok(());
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
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let canon_field = self.canon(field);
            let receiver_is_direct = matches!(
                object.kind,
                ExprKind::This | ExprKind::Super | ExprKind::Ident(_)
            );
            let user_method_shadow = receiver_is_direct
                && self.defined_class_methods.contains(&canon_field);
            let matched_value_method = self.profile.lookup_value_method(field, arg_exprs.len() as u8).cloned();
            let prefer_string_stdlib_value_method = matches!(
                matched_value_method.as_ref().map(|d| &d.emit),
                Some(BuiltinEmit::Stdlib(_))
            ) && self.expr_is_known_string_receiver(object);
            // Also skip value_methods if the field is an array HOF method —
            // the array_methods dispatch handles it with proper HOF semantics.
            // Without this, `[1,2,3].includes(2)` routes through the string
            // `includes` value method instead of the array contains HOF.
            let field_lower_check = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            let is_array_method = self.profile.lookup_array_method(&field_lower_check).is_some();
            if user_method_shadow || is_array_method {
                // Fall through — let the HOF dispatch or generic call path handle it
            } else if self.profile.namespaces.use_dotnet
                && common::dotnet::uses_runtime_collection_dispatch(field)
                && !prefer_string_stdlib_value_method
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
                        self.emit_common(&name, line);
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
            let field_lower = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            if let Some(stdlib_name) = self.profile.lookup_array_method(&field_lower).map(|s| s.to_string()) {
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
                        common::loops::emit_foreach(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, line);
                    }
                    "some" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, true, line);
                    }
                    "every" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, false, line);
                    }
                    "find" => {
                        // find uses includes pattern but returns element not bool
                        self.emit(Op::NULL);
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
                        // `x.includes(v)` — polymorphic: arrays do element
                        // membership, strings do substring search, user
                        // objects fall through to their own method. Route
                        // through `ecma:value.invokeMethod` so the
                        // emitted wasm stays spec-compliant on v8 where
                        // String.prototype.includes and Array.prototype.includes
                        // are distinct methods on distinct prototypes.
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        common::invoke::emit_invoke_method(
                            &mut self.chunks,
                            self.current,
                            "includes",
                            1,
                            line,
                        );
                    }
                    "sort" => {
                        // JS sort(comparatorFn?) — 2-arg comparator or default
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
                        let sort_global = self.str_const("__vybe_sort_in_place");
                        self.emit_u16(Op::GLOBAL_GET, sort_global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u8(Op::CALL_REF, 1);
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
                if is_ctor && self.defined_globals.contains(class_name.as_str()) {
                    self.emit_var_get(class_name);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Method call: obj.method(args) ───────────────────────────
        if let ExprKind::Member { object, field, null_safe } = &callee.kind {
            self.compile_expr(object)?;

            if *null_safe {
                // obj?.method() — short-circuit to null if obj is null/undefined.
                // Stack: [obj]. Check null, if null leave null on stack and skip call.
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let obj_not_null = self.emit_jump(Op::BR_IF_FALSE);
                // obj IS null — leave null on stack, skip call
                let end = self.emit_jump(Op::BR);
                self.patch_jump(obj_not_null);
                // obj is not null — do the method call
                let field_name = self.canon(field);
                let prop = self.str_const(&field_name);
                self.emit(Op::DUP);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                let obj_tmp = self.define_local("__obj");
                self.reserve_local_slot(obj_tmp);
                self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                self.patch_jump(end);
                return Ok(());
            }

            let field_name = self.canon(field);
            let prop = self.str_const(&field_name);
            self.emit(Op::DUP);
            self.emit_u16(Op::STRUCT_GET, prop);
            let fn_tmp = self.define_local("__fn");
            self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
            let obj_tmp = self.define_local("__obj");
            self.reserve_local_slot(obj_tmp);
            self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, fn_tmp);
            self.emit_u16(Op::LOCAL_GET, obj_tmp);
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
            return Ok(());
        }

        // ── Simple call: name(args) / expr(args) ────────────────────
        if let ExprKind::Ident(name) = &callee.kind {
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
                if is_local || is_global_var {
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
                        // Me.name(args) → load Me, dup, struct_get(name), call with this
                        let field_name = self.canon(name);
                        let prop = self.str_const(&field_name);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.define_local("__bare_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let obj_tmp = self.define_local("__bare_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
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
                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.emit(Op::SPREAD);
                let arity = known_len.unwrap_or(16) as u8;
                self.emit_u8(Op::CALL_REF, arity);
                return Ok(());
            }
            self.emit_var_get(name);
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
            return Ok(());
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
            self.define_local(&p.name);
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
