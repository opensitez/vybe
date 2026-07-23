//! Pascal set promotion, var-set, global-map emit, closure/upvalue binding, class-field checks.
//!
//! Extracted from `compiler/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

impl Compiler {
    pub(super) fn maybe_promote_pascal_array_literal_to_set(
        &mut self,
        type_hint: Option<&str>,
        value: &Expression,
    ) {
        if self.profile.name != "pascal" {
            return;
        }
        if !type_hint.is_some_and(Self::is_pascal_set_type_hint) {
            return;
        }
        if !matches!(value.kind, ExprKind::Array(_)) {
            return;
        }
        let idx = self.import("ecma:set", "fromIterable");
        self.emit_host_call(idx, 1);
    }

    pub(super) fn expr_is_pascal_set(&self, expr: &Expression) -> bool {
        if self.profile.name != "pascal" {
            return false;
        }

        match &expr.kind {
            ExprKind::Set(_) => true,
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(Self::is_pascal_set_type_hint),
            ExprKind::Binary { op, left, right }
                if matches!(op, BinOp::Add | BinOp::Mul | BinOp::Sub) =>
            {
                self.expr_is_pascal_set(left) && self.expr_is_pascal_set(right)
            }
            _ => false,
        }
    }

    pub(crate) fn emit_var_get(&mut self, name: &str) {
        // Shared env: locals captured by inner closures live in a shared
        // array so mutations are visible across all closures.
        if let Some(idx) = self.shared_env_index(name) {
            if let Some(env_slot) = self.shared_env_slot {
                let l = self.line;
                crate::emitter::closures::emit_env_get(self.chunk(), env_slot, idx, l);
                return;
            }
        }
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            self.emit_u16(Op::LOCAL_GET, slot);
            if self.binding_uses_pointer_cell(name) {
                common::references::emit_cell_load(&mut self.chunks, self.current, self.line);
            }
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::LOCAL_GET, slot);
                if self.binding_uses_pointer_cell(name) {
                    common::references::emit_cell_load(&mut self.chunks, self.current, self.line);
                }
                return;
            }
        }
        if self.scopes.len() > 1 {
            if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(name);
                let l = self.line;
                crate::emitter::closures::emit_env_get(self.chunk(), env, idx, l);
                return;
            }
        }
        if let Some(binding) = self.static_local_binding(name) {
            let global_name = binding.global_name.clone();
            let idx = self.global_name_const_idx(&global_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            return;
        }
        // Implicit self field — when inside a class method and the name is a
        // field of the current class, read from `me.<name>`. This is what
        // languages like VB do for unqualified field access. Without this,
        // dotted-name resolution that returns InstanceMember { local: "field" }
        // would fall through to global_get and read null.
        if self.current_class_implicit_self && self.is_class_field(name) {
            if self.emit_self_ref() {
                let cname = self
                    .current_class
                    .as_deref()
                    .and_then(|class_name| {
                        self.visible_instance_field_storage_name_for_class(class_name, name)
                    })
                    .unwrap_or_else(|| self.canon(name));
                let idx = self.str_const(&cname);
                self.emit_u16(Op::STRUCT_GET, idx);
                return;
            }
        }
        // Static field of the current class — `Count++` inside `Counter`
        // ctor reads `Counter.Count` (struct_get on the class global).
        // Without this, the bare name falls through to global_get and
        // returns null because the static field lives on the class
        // struct, not the module's global namespace.
        if let Some(class_name) = self.is_class_static_field(name) {
            let class_idx = self.global_name_const_idx(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            let field_idx = self.str_const(&self.canon(name));
            self.emit_u16(Op::STRUCT_GET, field_idx);
            return;
        }
        // Bare static method in class scope — `Double(x)` inside
        // `class Converter` resolves to `Converter.Double`.
        if let Some(class_name) = self.is_class_static_method(name) {
            let class_idx = self.global_name_const_idx(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            let method_idx = self.str_const(&self.canon(name));
            self.emit_u16(Op::STRUCT_GET, method_idx);
            return;
        }
        let cname = self.canon(name);
        let shadows_named_global = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        if !shadows_named_global && self.emit_with_target_get(name) {
            return;
        }
        // Known type used as a value (e.g. `e instanceof RangeError`) — emit
        // the type name as a string so the instanceof ref.test fallback can
        // look it up by name. Without this, `RangeError` would become
        // `global_get` of a nonexistent global → null.
        // Only do this when the name isn't shadowed by an actual global
        // (e.g. `Dim list As New List(Of String)` shadows the `list` type name).
        let is_js_runtime_global = self.profile.has_ecma_globals
            && (matches!(
                name,
                "Object"
                    | "Boolean"
                    | "Number"
                    | "String"
                    | "Array"
                    | "Function"
                    | "Symbol"
                    | "BigInt"
                    | "Error"
                    | "EvalError"
                    | "RangeError"
                    | "ReferenceError"
                    | "SyntaxError"
                    | "TypeError"
                    | "URIError"
                    | "AggregateError"
                    | "SuppressedError"
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
            ) || self.host_namespace_aliases.contains_key(&cname));
        if self.profile.known_types.contains_key(name)
            && !self.defined_globals.contains(name)
            && !self.defined_globals.contains(&cname)
            && !is_js_runtime_global
        {
            self.emit_const(Value::String(Arc::from(name)));
            return;
        }
        // JS builtin constructor used as a *value* (`x === Array`,
        // `o.constructor === Object`, `Array.prototype`): resolve through the
        // stable, host-owned `__ctor_<Name>` anchor instead of the user-facing
        // global. The latter can be re-bound to a fresh, unwired object by
        // later compile/link passes (ESM import wiring), which would break
        // `constructor`/`prototype` identity; `__ctor_<Name>` always points at
        // the ONE canonical constructor (the same object on the shared
        // prototype's `.constructor`). Skipped when the user shadows the name.
        // Built-in Error constructors are recognised from the profile's
        // `known_types` (their backing module is `ecma:error`) rather than a
        // hardcoded name list, so `e.constructor === TypeError` and
        // `typeof TypeError === "function"` resolve through the same canonical
        // `__ctor_<Name>` anchor the host installs for them.
        let is_error_ctor_value = self
            .profile
            .known_types
            .get(name)
            .is_some_and(|(module, _)| module == "ecma:error");
        if self.profile.has_ecma_globals
            && !shadows_named_global
            && (is_js_builtin_ctor_value(&cname) || is_error_ctor_value)
        {
            let idx = self.str_const(&format!("__ctor_{cname}"));
            self.emit_u16(Op::GLOBAL_GET, idx);
            return;
        }
        if self.php_inside_function()
            && !self.php_current_function_declares_global(name)
            && !self.defined_functions.contains(&cname)
            && !self.defined_classes.contains(&cname)
            && !cname.starts_with("__")
            // `use const Lib\LEVEL;` imported names read the qualified
            // global from inside functions too — fall through to the
            // use-alias consult below instead of PHP's undeclared-null.
            && !self.source_type_aliases.contains_key(&cname)
        {
            self.emit(Op::NULL);
            return;
        }
        // Global — canonicalize name for case-insensitive languages
        // But in strict mode, if this is genuinely undeclared, throw ReferenceError
        //
        // Use-alias consult (namespaceplan.md): `use const Lib\MAX;` binds
        // bare `MAX` to the namespace-qualified global `Lib.MAX` (the same
        // `source_type_aliases` map static-access/instanceof already
        // resolve through). Only when nothing else declared the bare name —
        // a real global/function/class shadows the import per PHP scoping.
        let cname = if self.profile.uses_common_resolver
            && !self.defined_globals.contains(&cname)
            && !self.defined_functions.contains(&cname)
            && !self.defined_classes.contains(&cname)
        {
            // PHP §namespace resolution order for unqualified names: the
            // CURRENT namespace's declaration wins over file-level `use`
            // aliases (which are last-wins across the file), then the
            // aliases, then the global fallback.
            let ns_qualified = self.current_namespace.as_deref().and_then(|ns| {
                let q = self.canon(&format!("{ns}.{cname}"));
                (self.defined_functions.contains(&q) || self.defined_globals.contains(&q))
                    .then_some(q)
            });
            match ns_qualified {
                Some(q) => q,
                None => match self.source_type_aliases.get(&cname) {
                    Some(target) => self.canon(target),
                    None => cname,
                },
            }
        } else {
            cname
        };
        let global_key = self.php_variable_global_key(name, &cname);
        let idx = self.global_name_const_idx(&global_key);

        // ECMA-262 §9.1.1.4.6 / §13.3.2.1 GetValue: reading an *unresolvable*
        // reference (a name bound nowhere in the scope chain or on the global
        // object) is a `ReferenceError`. Reaching this fallback means every
        // compile-time resolution attempt failed — not a local/upvalue/static/
        // class-field, not a declared global/function/class, not a builtin or
        // host target (those returned earlier). So this is a genuine *missing
        // binding*, decided at compile time — NOT a runtime "value is undefined"
        // test (a declared `let x;` legitimately holds `undefined`).
        //
        // Driven by the `unresolved_reference_throws` profile capability, and
        // additionally gated on strict mode: per spec the throw applies in
        // sloppy mode too, but sloppy code leans on lenient access to
        // host-provided globals the compiler does not track as bindings (so a
        // blanket throw there mis-fires). `typeof x` on an undeclared name must
        // yield "undefined", so it is suppressed via `in_typeof_operand`.
        // Builtin global-object aliases are excluded defensively.
        let resolvable = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        // §9.1.1.4.6 GetValue throws for an unresolvable read in BOTH strict
        // and sloppy mode. We gate the general throw to strict mode because a
        // blanket sloppy throw misfires on host globals we don't track as
        // bindings — but a name the program itself lexically declared
        // (`program_lexical_names`) that is unresolvable here is provably an
        // out-of-scope *user* binding, never an untracked host global, so it
        // throws in sloppy mode too.
        if self.profile.unresolved_reference_throws
            && (self.in_strict || self.program_lexical_names.contains(&cname))
            && !self.in_typeof_operand
            && !resolvable
            && !cname.starts_with("__")
            && !is_js_builtin_ctor_value(&cname)
            && !matches!(
                name,
                "globalThis"
                    | "window"
                    | "self"
                    | "global"
                    | "globalObject"
                    | "arguments"
                    | "this"
                    | "undefined"
                    | "NaN"
                    | "Infinity"
            )
        {
            let line = self.line;
            self.emit_u16(Op::STRUCT_NEW, 0);
            inst!(self, core_wasm::dup);
            self.emit_const(Value::String(Arc::from(
                format!("{name} is not defined").as_str(),
            )));
            crate::emitter::errors::emit_exception_new_finalize(
                self.chunk(),
                "ReferenceError",
                line,
            );
            crate::emitter::errors::emit_throw(self.chunk(), line);
            return;
        }
        self.emit_u16(Op::GLOBAL_GET, idx);
        if self.binding_uses_pointer_cell(name) {
            common::references::emit_cell_load(&mut self.chunks, self.current, self.line);
        }
    }

    pub(super) fn emit_ensure_global_map(&mut self, name: &str) {
        let key = self.shared_global_slot(name);
        self.emit_u16(Op::GLOBAL_GET, key);
        inst!(self, core_wasm::dup);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit(Op::DROP);
        common::collections::emit_map_new(&mut self.chunks, self.current, line);
        inst!(self, core_wasm::dup);
        self.emit_u16(Op::GLOBAL_SET, key);

        self.chunk().emit_end(line);
    }
    /// ECMA-262 §11.2.1 Directive Prologue: returns `true` if the leading
    /// run of string-literal expression statements contains `"use strict"`.
    pub(crate) fn stmts_have_use_strict_directive(stmts: &[Statement]) -> bool {
        for s in stmts {
            match &s.kind {
                // The walker emits `Empty` for newlines between statements;
                // they don't terminate the directive prologue.
                StmtKind::Empty => continue,
                // The JS walker HOISTS function declarations above the
                // directive prologue (parse-time reorder), so a decl here
                // says nothing about the prologue's textual position.
                StmtKind::FunctionDecl { .. } => continue,
                StmtKind::Expr(e) => match &e.kind {
                    ExprKind::Lit(Literal::Str(v)) => {
                        if v == "use strict" {
                            return true;
                        }
                        // Another directive — keep scanning the prologue.
                    }
                    _ => break,
                },
                _ => break,
            }
        }
        false
    }

    pub(super) fn emit_var_set(&mut self, name: &str) {
        // ECMA-262 §13.15.2 / §6.2.4.7: assigning to a `const` binding is a
        // runtime `TypeError` ("Assignment to constant variable."). The
        // binding is known to the compiler — a `const` local in scope, or a
        // top-level `const` global — so emit an unconditional throw at the
        // assignment site. (Declaration init and direct loop-variable rebinds
        // use `LOCAL_SET`/`GLOBAL_SET` directly and never reach here.)
        if self.profile.ecma_lexical_declarations {
            let is_const_local = self.scope().resolve_is_const(name);
            let is_const_global = !is_const_local
                && self.scope().resolve(name).is_none()
                && (self.const_globals.contains(name)
                    || self.const_globals.contains(&self.canon(name)));
            if is_const_local || is_const_global {
                let line = self.line;
                self.emit_u16(Op::STRUCT_NEW, 0);
                inst!(self, core_wasm::dup);
                self.emit_const(Value::String(Arc::from("Assignment to constant variable.")));
                crate::emitter::errors::emit_exception_new_finalize(
                    self.chunk(),
                    "TypeError",
                    line,
                );
                crate::emitter::errors::emit_throw(self.chunk(), line);
                return;
            }
        }
        // Shared env: locals captured by inner closures
        if let Some(idx) = self.shared_env_index(name) {
            if let Some(env_slot) = self.shared_env_slot {
                let l = self.line;
                crate::emitter::closures::emit_env_set(self.chunk(), env_slot, idx, l);
                return;
            }
        }
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            if self.binding_uses_pointer_cell(name) {
                let value_slot = self.define_local("__ref_cell_set_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit_u16(Op::LOCAL_GET, slot);
                common::references::emit_cell_store(
                    &mut self.chunks,
                    self.current,
                    value_slot,
                    self.line,
                );
                self.emit(Op::DROP);
            } else if let Some((args_slot, index)) = self.js_arguments_alias_for_name(name) {
                let value_slot = self.define_local("__js_arguments_alias_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::LOCAL_SET, slot);
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.emit_const(Value::F64(index as f64));
                self.emit_u16(Op::LOCAL_GET, value_slot);
                common::collections::emit_set(&mut self.chunks, self.current, self.line);
                self.emit(Op::DROP);
            } else {
                self.emit_u16(Op::LOCAL_SET, slot);
            }
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                if self.binding_uses_pointer_cell(name) {
                    let value_slot = self.define_local("__ref_cell_set_value");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit_u16(Op::LOCAL_GET, slot);
                    common::references::emit_cell_store(
                        &mut self.chunks,
                        self.current,
                        value_slot,
                        self.line,
                    );
                    self.emit(Op::DROP);
                } else if let Some((args_slot, index)) = self.js_arguments_alias_for_name(name) {
                    let value_slot = self.define_local("__js_arguments_alias_value_ci");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit_u16(Op::LOCAL_GET, args_slot);
                    self.emit_const(Value::F64(index as f64));
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    common::collections::emit_set(&mut self.chunks, self.current, self.line);
                    self.emit(Op::DROP);
                } else {
                    self.emit_u16(Op::LOCAL_SET, slot);
                }
                return;
            }
        }
        if self.scopes.len() > 1 {
            if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(name);
                let l = self.line;
                crate::emitter::closures::emit_env_set(self.chunk(), env, idx, l);
                return;
            }
        }
        if let Some(binding) = self.static_local_binding(name) {
            let global_name = binding.global_name.clone();
            let idx = self.global_name_const_idx(&global_name);
            self.emit_u16(Op::GLOBAL_SET, idx);
            return;
        }
        if self.current_class_implicit_self && self.is_class_field(name) {
            let value_slot = self.define_local("__implicit_self_value");
            self.emit_u16(Op::LOCAL_SET, value_slot);
            if self.emit_self_ref() {
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let cname = self
                    .current_class
                    .as_deref()
                    .and_then(|class_name| {
                        self.visible_instance_field_storage_name_for_class(class_name, name)
                    })
                    .unwrap_or_else(|| self.canon(name));
                let idx = self.str_const(&cname);
                self.emit_u16(Op::STRUCT_SET, idx);
                self.emit(Op::DROP);
                return;
            }
            self.emit_u16(Op::LOCAL_GET, value_slot);
        }
        // Static field of the current class — write through to
        // `<ClassName>.<name>` instead of falling to global_set.
        if let Some(class_name) = self.is_class_static_field(name) {
            // Stack: [value]. Need [class_obj, value] for STRUCT_SET.
            let value_slot = self.define_local("__static_set_value");
            self.emit_u16(Op::LOCAL_SET, value_slot);
            let class_idx = self.global_name_const_idx(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            let bare_name = self.canon(name);
            let field_idx = self.str_const(&bare_name);
            self.emit_u16(Op::STRUCT_SET, field_idx);
            self.emit(Op::DROP);
            if self.defined_globals.contains(&bare_name) {
                let global_idx = self.global_name_const_idx(&bare_name);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::GLOBAL_SET, global_idx);
            }
            return;
        }
        let cname = self.canon(name);
        let shadows_named_global = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        // ECMA-262 §6.2.5.6 PutValue / §9.1.1.4.5: in strict mode, assigning
        // to an unresolvable reference (a name bound nowhere in the scope
        // chain or on the global object) is a `ReferenceError` — sloppy mode
        // would silently create a global. Reaching here means the name is not
        // a local/upvalue/static/class-field/declared-global. We can only
        // throw for names that cannot be a host-provided builtin global, so
        // exclude the known builtin constructors and global-object aliases.
        // Gated on `in_strict` (rare) to keep sloppy-mode global-creation,
        // which the rest of the suite relies on, intact.
        if self.in_strict
            && self.profile.ecma_strict_mode
            && !shadows_named_global
            && !cname.starts_with("__")
            && !is_js_builtin_ctor_value(&cname)
            && !matches!(
                name,
                "globalThis" | "window" | "self" | "global" | "globalObject" | "arguments"
            )
        {
            let line = self.line;
            self.emit_u16(Op::STRUCT_NEW, 0);
            inst!(self, core_wasm::dup);
            self.emit_const(Value::String(Arc::from(
                format!("{name} is not defined").as_str(),
            )));
            crate::emitter::errors::emit_exception_new_finalize(
                self.chunk(),
                "ReferenceError",
                line,
            );
            crate::emitter::errors::emit_throw(self.chunk(), line);
            return;
        }
        if !shadows_named_global && self.emit_with_target_set(name) {
            return;
        }
        if self.php_inside_function()
            && !self.php_current_function_declares_global(name)
            && !self.defined_functions.contains(&cname)
            && !self.defined_classes.contains(&cname)
            && !cname.starts_with("__")
        {
            let slot = self.define_local(name);
            self.emit_u16(Op::LOCAL_SET, slot);
            return;
        }
        // Global — canonicalize name for case-insensitive languages
        let global_key = self.php_variable_global_key(name, &cname);
        if self.scopes.len() == 1 {
            self.defined_globals.insert(global_key.clone());
        }
        if self.binding_uses_pointer_cell(name) {
            let value_slot = self.define_local("__ref_global_set_value");
            self.emit_u16(Op::LOCAL_SET, value_slot);
            let idx = self.global_name_const_idx(&global_key);
            self.emit_u16(Op::GLOBAL_GET, idx);
            common::references::emit_cell_store(
                &mut self.chunks,
                self.current,
                value_slot,
                self.line,
            );
            self.emit(Op::DROP);
            return;
        }
        let idx = self.global_name_const_idx(&global_key);
        self.emit_u16(Op::GLOBAL_SET, idx);
    }

    pub(super) fn capture_local_slot(&mut self, uv_idx: u8) -> u16 {
        if let Some(&slot) = self.capture_locals.get(&uv_idx) {
            return slot;
        }
        let slot = self.define_local(&format!("__capture_{}", uv_idx));
        self.capture_locals.insert(uv_idx, slot);
        let c = &mut self.chunks[self.current];
        if c.capture_count <= uv_idx {
            c.capture_count = uv_idx + 1;
        }
        if c.capture_base == 0 || slot < c.capture_base {
            c.capture_base = slot;
        }
        slot
    }

    /// Get or allocate the closure environment slot for the current function.
    /// The env is a GC array holding all captured variables by index.
    /// It arrives as upvalue[0] and is copied to this local by call_function_inner.
    pub(super) fn closure_env_slot(&mut self) -> u16 {
        self.capture_local_slot(0)
    }

    /// Get or register a captured variable's index in the closure env array.
    pub(super) fn closure_env_index(&mut self, name: &str) -> u16 {
        if let Some(idx) = self.closure_env_names.iter().position(|n| n == name) {
            return idx as u16;
        }
        let idx = self.closure_env_names.len();
        self.closure_env_names.push(name.to_string());
        idx as u16
    }

    /// Check if a name is in the current function's shared env.
    pub(super) fn shared_env_index(&self, name: &str) -> Option<u16> {
        self.shared_env_names
            .iter()
            .position(|n| n == name)
            .map(|i| i as u16)
    }

    pub(super) fn resolve_upvalue(&mut self, scope_idx: usize, name: &str) -> Option<u8> {
        if scope_idx == 0 {
            return None;
        }
        let parent = scope_idx - 1;
        // Check parent's locals
        let found_local = if self.case_sensitive {
            self.scopes[parent].resolve(name)
        } else {
            self.scopes[parent]
                .resolve(name)
                .or_else(|| self.scopes[parent].resolve_ci(name))
        };
        if let Some(slot) = found_local {
            self.scopes[parent].mark_captured(slot);
            return Some(self.scopes[scope_idx].add_upvalue(slot, true));
        }
        // Recurse up
        if let Some(uv) = self.resolve_upvalue(parent, name) {
            return Some(self.scopes[scope_idx].add_upvalue(uv as u16, false));
        }
        None
    }

    /// Returns the owning class name when `name` is a static field of
    /// the currently-compiling class (or one of its ancestors). Used by
    /// `emit_var_get` / `emit_var_set` to rewrite bare references to
    /// `ClassName.name` so static state lives on the class struct.
    pub(super) fn is_class_static_field(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        if pc.static_fields.iter().any(|f| {
                            if self.case_sensitive {
                                f == name
                            } else {
                                f.eq_ignore_ascii_case(name)
                            }
                        }) {
                            return Some(cn.to_string());
                        }
                        current = pc.parent.as_deref();
                    } else {
                        break;
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }
}
