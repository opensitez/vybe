//! Class-scope membership checks, self/this/new-target handling, member-chain helpers.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

impl Compiler {
    pub(super) fn is_class_static_field_type_hint(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        let canon = self.canon(name);
                        if let Some(type_hint) = pc.static_field_types.get(&canon) {
                            return Some(type_hint.clone());
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

    #[allow(dead_code)]
    pub(super) fn is_class_nested_type(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        if pc.nested_types.iter().any(|n| {
                            if self.case_sensitive {
                                n == name
                            } else {
                                n.eq_ignore_ascii_case(name)
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

    pub(super) fn generic_static_member_key(&self, type_expr: &str, field: &str) -> Option<String> {
        let expr = type_expr.trim();
        if !expr.contains('<') || !expr.contains('>') {
            return None;
        }

        let base = expr.split('<').next().map(str::trim).unwrap_or(expr);
        let base_canon = self.canon(base);
        if !self.defined_classes.contains(&base_canon) {
            return None;
        }

        let field_canon = self.canon(field);
        let has_static = self
            .pending_classes
            .get(base)
            .or_else(|| self.pending_classes.get(base_canon.as_str()))
            .map(|pc| pc.static_fields.iter().any(|f| f == &field_canon))
            .unwrap_or(false);
        if !has_static {
            return None;
        }

        let compact_type: String = expr.chars().filter(|c| !c.is_whitespace()).collect();
        let type_canon = self.canon(&compact_type);
        Some(format!("__gstatic_{}_{}", type_canon, field_canon))
    }

    /// Returns the owning class when `name` is a static method of the
    /// currently compiling class (or one of its ancestors).
    pub(super) fn is_class_static_method(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        if pc.static_method_names.iter().any(|m| {
                            if self.case_sensitive {
                                m == name
                            } else {
                                m.eq_ignore_ascii_case(name)
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

    pub(super) fn next_enclosing_class_name(&self, class_name: &str) -> Option<String> {
        self.pending_classes
            .get(class_name)
            .and_then(|pc| pc.enclosing_class.clone())
            .or_else(|| {
                class_name
                    .rsplit_once('.')
                    .map(|(outer, _)| outer.to_string())
            })
    }

    pub(super) fn class_extends_builtin(&self, class_name: &str, builtin: &str) -> bool {
        let mut current = Some(self.canon(class_name));
        let target = self.canon(builtin);
        while let Some(name) = current {
            let Some(pc) = self.pending_classes.get(name.as_str()) else {
                return false;
            };
            let Some(parent) = pc.parent.as_ref() else {
                return false;
            };
            let parent_canon = self.canon(parent);
            if parent_canon == target {
                return true;
            }
            current = Some(parent_canon);
        }
        false
    }

    /// Check if a name is a field of the current class (for implicit self resolution).
    pub(super) fn is_class_field(&self, name: &str) -> bool {
        if !self.current_class_implicit_self {
            return false;
        }
        self.current_class
            .as_deref()
            .and_then(|class_name| {
                self.visible_instance_field_storage_name_for_class(class_name, name)
            })
            .is_some()
    }

    pub(super) fn emit_self_ref(&mut self) -> bool {
        let self_kw = self.profile.self_keyword.clone();
        if let Some(self_slot) = self
            .scope()
            .resolve(&self_kw)
        {
            self.emit_u16(Op::LOCAL_GET, self_slot);
            return true;
        }
        if self.scopes.len() > 1 {
            if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, &self_kw) {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(&self_kw);
                let l = self.line;
                crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                return true;
            }
        }
        false
    }

    /// Profile-declared class dispatch model — `class_method_dispatch =
    /// "prototype"` in the language's `[compiler]` section. The shared
    /// class pipeline stays language-agnostic; languages opt in via the
    /// profile, never via name checks.
    pub(crate) fn class_prototype_dispatch(&self) -> bool {
        self.profile.class_method_dispatch == "prototype"
    }

    /// REMAINING language-name check, kept here beside `is_python_profile`
    /// rather than in a `php_lang` module of its own — a file named after a
    /// language in shared code invites more of the same, and the whole point
    /// is that this predicate should keep shrinking. Each surviving call site
    /// is a profile property, a normalization, or an adapter that has not been
    /// written yet, not a permanent fixture.
    pub(crate) fn is_php_profile(&self) -> bool {
        self.profile.name == "php"
    }

    /// Do STATIC methods carry a leading receiver slot (the class object)?
    ///
    /// This is late static binding: `static::` and `get_called_class()` resolve
    /// against the class the call was made THROUGH, not the one the method was
    /// declared in, so the callee cannot recover it and the caller has to pass
    /// it. Languages whose statics are plain functions answer false.
    ///
    /// Still a name check, but now ONE — it was five, each re-deriving the same
    /// condition inline. It wants to become a class-shape trait declared by the
    /// frontend, the way `explicit_self_param` already is in seven of them.
    pub(crate) fn static_methods_take_receiver(&self) -> bool {
        self.profile.name == "php"
    }

    /// Must a method CALL pass the receiver as an explicit leading argument?
    ///
    /// Three models, and this predicate picks the third:
    /// - prototype dispatch (JS/Dart) rides `__js_this` and a bound-receiver
    ///   marker on the callable — see `class_prototype_dispatch`;
    /// - bind-on-access (Python) burns the receiver into a fresh bound method
    ///   when the method is READ — see `methods_bind_on_access`;
    /// - otherwise the callable is the raw function off the class struct and
    ///   carries no receiver, so the call site supplies one.
    ///
    /// NOT `explicit_method_receiver_argument`, which is Lua's and means the
    /// opposite — the walker ALREADY passed a receiver, so shared code must not
    /// add a second one.
    ///
    /// The declaration side already records this per callee as
    /// `chunk.is_method` (`classes.rs`, `has_receiver`); nothing reads it back,
    /// because `CallSignature` carries no receiver flag and a dynamic callee
    /// slot cannot reach its chunk. Thread it there and this predicate goes.
    pub(crate) fn call_supplies_receiver(&self) -> bool {
        self.profile.name == "php"
    }

    /// True for profiles whose comparison/equality operators dispatch to a
    /// user-defined dunder (`__eq__`/`__lt__`/… and their cross-language
    /// aliases) — i.e. the same profiles the `<`/`>` sites already route
    /// through `emit_rich_compare_locals` (Python, Ruby, Dart, C#, VB, …).
    /// Excludes JS (ECMA coercion), PHP (loose comparison) and Pascal.
    ///
    /// Dispatch goes through the `Eq`/`Lt`/`Compare` SLOTS: a language that
    /// binds them gets its own semantics, one that binds nothing falls back to
    /// primitive comparison. Declaring nothing IS the opt-out, so no language
    /// needs excluding here — pascal was, by name, and did not need to be.
    /// PHP stays out via `string_aware_relational`, which it declares anyway.
    pub(crate) fn uses_rich_comparison(&self) -> bool {
        !self.profile.ecma_operator_coercion && !self.profile.string_aware_relational
    }

    /// Operator overloading on the arithmetic/unary operators: a user
    /// `__add__`/`__neg__`/… on the operand wins over the primitive op.
    ///
    /// The same profiles that get rich comparison — a language either
    /// dispatches operators to methods or it coerces operands, and the
    /// two are the same question. Languages whose `+` is ECMA-coerced
    /// reach their operator methods through `ecma:value.add`'s
    /// ToPrimitive/`valueOf` chain instead.
    pub(crate) fn uses_rich_operators(&self) -> bool {
        self.uses_rich_comparison()
    }

    pub(super) fn emit_condition_truthiness_from_stack(&mut self) {
        // Only Python needs a custom rule here — empty str/list/dict/set are
        // falsy, which no primitive coercion expresses.
        //
        // PHP used to branch here too, onto a check that read the REMOVED
        // `__keys`/`vybe$assoc_keys_csv` side-band and corrupted the stack. The
        // side-band went away and array truthiness moved to the `empty()`/
        // `isset()` call sites via the Map-aware emitter, but the dead branch
        // outlived both. `emit_dyn_to_bool` is correct for every language that
        // is not Python.
        if !self.profile.truthiness_via_dunder_or_length {
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            };
            return;
        }

        let line = self.line;

        let value_slot = self.define_local("__py_truth_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        // Python: empty str/list/dict/set are falsy; numbers/bool/None use dyn_to_bool.
        // Reuse collections::emit_len (ecma:array.length / ecma:map.size / string length).
        let typeof_idx = self.import("ecma:value", "typeof");
        let is_object_slot = self.define_local("__py_truth_is_object");

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_host_call(typeof_idx, 1);
        self.emit_const(Value::String(Arc::from("object")));
        fn_call!(self, "wasm:js-string", "equals", 2);
        self.emit_u16(Op::LOCAL_SET, is_object_slot);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_host_call(typeof_idx, 1);
        self.emit_const(Value::String(Arc::from("string")));
        fn_call!(self, "wasm:js-string", "equals", 2);
        self.emit_u16(Op::LOCAL_GET, is_object_slot);
        self.chunk().emit_op(Op::I32_OR, line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::collections::emit_len(&mut self.chunks, self.current, line);
        inst!(self, core_wasm::i32_const, 0);
        self.chunk().emit_op(Op::I32_NE, line);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_end(line);
    }

    pub(super) fn save_js_this(&mut self, local_name: &str) -> Option<u16> {
        if !self.profile.ambient_this_binding {
            return None;
        }
        let slot = self
            .scope()
            .resolve(local_name)
            .unwrap_or_else(|| self.define_local(local_name));
        self.emit_global_read("__js_this");
        self.emit_u16(Op::LOCAL_SET, slot);
        Some(slot)
    }

    pub(super) fn set_js_this_from_stack(&mut self) {
        if !self.profile.ambient_this_binding {
            return;
        }
        self.emit_global_write("__js_this");
    }

    pub(super) fn restore_js_this(&mut self, slot: Option<u16>) {
        let Some(slot) = slot else {
            return;
        };
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_global_write("__js_this");
    }

    pub(super) fn save_js_new_target(&mut self, local_name: &str) -> Option<u16> {
        if !self.profile.ecma_new_dispatch {
            return None;
        }
        let slot = self
            .scope()
            .resolve(local_name)
            .unwrap_or_else(|| self.define_local(local_name));
        self.emit_global_read("__js_new_target");
        self.emit_u16(Op::LOCAL_SET, slot);
        Some(slot)
    }

    pub(super) fn set_js_new_target_from_stack(&mut self) {
        if !self.profile.ecma_new_dispatch {
            return;
        }
        self.emit_global_write("__js_new_target");
    }

    pub(super) fn restore_js_new_target(&mut self, slot: Option<u16>) {
        let Some(slot) = slot else {
            return;
        };
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_global_write("__js_new_target");
    }

    pub(super) fn set_js_new_target_undefined(&mut self) {
        if !self.profile.ecma_new_dispatch {
            return;
        }
        let line = self.line;
        common::expressions::emit_undefined(self.chunk(), line);
        self.emit_global_write("__js_new_target");
    }

    pub(super) fn flatten_member_chain(&self, expr: &Expression) -> Vec<String> {
        match &expr.kind {
            ExprKind::Ident(name) => Self::strip_global_namespace_prefix(name)
                .replace("::", ".")
                .split('.')
                .map(str::trim)
                .filter(|part| !part.is_empty())
                .map(ToString::to_string)
                .collect(),
            ExprKind::This => vec![self.profile.self_keyword.clone()],
            ExprKind::Super => vec![
                self.profile
                    .base_keyword
                    .clone()
                    .unwrap_or_else(|| "super".into()),
            ],
            ExprKind::Member { object, field, .. } => {
                let mut parts = self.flatten_member_chain(object);
                parts.push(field.clone());
                if parts
                    .first()
                    .is_some_and(|part| part.eq_ignore_ascii_case("global"))
                {
                    parts.remove(0);
                }
                parts
            }
            _ => Vec::new() }
    }

    /// Extract plain expressions from Argument slice.
    #[allow(dead_code)]
    pub(super) fn arg_exprs(args: &[Argument]) -> Vec<&Expression> {
        args.iter().map(|a| &a.value).collect()
    }

    // ════════════════════════════════════════════════════════════════════════
    // Statement compilation
    // ════════════════════════════════════════════════════════════════════════
}
