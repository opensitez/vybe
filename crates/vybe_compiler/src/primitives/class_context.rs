//! Class-scope membership checks, self/this/new-target handling, member-chain helpers.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use crate::primitives::class_slots;
use super::*;

impl Compiler {
    pub(super) fn is_class_static_field_type_hint(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                // ⛔ ONE ANSWER: `Compiler::resolution_chain`. A hand-rolled
                // `parent` climb answers a DIAMOND by taking the first base's
                // chain and never seeing the others — the defect that made
                // `D().who()` return `'A'` where C3 says `'C'`.
                for cn in self.resolution_chain(&start) {
                    let Some(pc) = self.pending_classes.get(cn.as_str()) else {
                        continue;
                    };
                    let canon = self.canon(name);
                    if let Some(type_hint) = pc.static_field_types.get(&canon) {
                        return Some(type_hint.clone());
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
                for cn in self.resolution_chain(&start) {
                    let Some(pc) = self.pending_classes.get(cn.as_str()) else {
                        continue;
                    };
                    if pc.nested_types.iter().any(|n| {
                        if self.case_sensitive {
                            n == name
                        } else {
                            n.eq_ignore_ascii_case(name)
                        }
                    }) {
                        return Some(cn.to_string());
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
                for cn in self.resolution_chain(&start) {
                    let Some(pc) = self.pending_classes.get(cn.as_str()) else {
                        continue;
                    };
                    if pc.static_method_names.iter().any(|m| {
                        if self.case_sensitive {
                            m == name
                        } else {
                            m.eq_ignore_ascii_case(name)
                        }
                    }) {
                        return Some(cn.to_string());
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

    /// Whether the current class (or an ancestor) declares `name` as an
    /// INSTANCE member of any kind — field, property, or method. This is the
    /// implicit-self question for members that are NOT plain fields: an
    /// accessor-backed property (`val y: Int` with a synthesized getter) has
    /// no entry in `fields`/`field_storage_names`, so `is_class_field` says
    /// no — yet a bare `y = …` in an init block is a write to `this.y` all
    /// the same, and must reach the property machinery, not a global.
    pub(super) fn is_class_instance_member(&self, name: &str) -> bool {
        if !self.current_class_implicit_self {
            return false;
        }
        let canon = self.canon(name);
        // ⛔ ONE ANSWER. This also carried its own `guard > 64` cycle cap —
        // `resolution_chain` is cycle-safe by construction, so the private
        // counter goes with the private walk.
        let Some(start) = self.current_class.as_deref().map(|c| self.canon(c)) else {
            return false;
        };
        for class_key in self.resolution_chain(&start) {
            let Some(pending) = self.pending_classes.get(class_key.as_str()) else {
                continue;
            };
            if pending.instance_member_names.iter().any(|m| m == &canon) {
                return true;
            }
        }
        false
    }

    pub(super) fn emit_self_ref(&mut self) -> bool {
        let self_kw = self.profile.self_keyword.clone();
        if let Some(self_slot) = self.scope().resolve(&self_kw) {
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
        self.method_receiver_model() == Some(vybe_ast::MethodReceiver::Prototype)
    }

    /// Reading a method produces a fresh callable with the receiver already
    /// bound (Python). One of the three dispatch models.
    pub(crate) fn methods_bind_on_access(&self) -> bool {
        self.method_receiver_model() == Some(vybe_ast::MethodReceiver::BindOnAccess)
    }

    /// How a method call obtains its receiver, for this UNIT.
    ///
    /// The three models are mutually exclusive and were previously spread
    /// across three unrelated spellings — a profile string
    /// (`class_method_dispatch = "prototype"`), a profile bool
    /// (`methods_bind_on_access`), and a language NAME (`profile.name ==
    /// "php"`). One question, three answers, none of which could see the
    /// others; nothing prevented a profile from declaring two of them.
    ///
    /// Now one directive with three variants, stated by the walker on
    /// `Module.directives`, so it travels with the UNIT — a multi-language
    /// bundle answers per unit, which a profile installed once per compilation
    /// structurally cannot do.
    /// Cooperative `super()` over a C3 linearization, vs a static parent.
    /// Is `super()` COOPERATIVE — resolved by walking the C3 linearization from
    /// the class the call textually belongs to — rather than static dispatch to
    /// the declared parent?
    ///
    /// ⛔ This was `class_multiple_inheritance`, and that name made one field
    /// answer THREE different questions (§5, "reusing a field as a marker"):
    /// "does this class have several bases" (which `class.bases` already
    /// answers, so the flag was redundant), "bind ancestors only for a diamond",
    /// and this one. The first meaning colliding with the third is what
    /// truncated python's `__mro__`: C3 was gated on `bases.len() > 1`, and **C3
    /// over a one-parent chain IS the chain**, so every grandparent was lost.
    ///
    /// ⚠ Cooperative super is not the same feature as C-style multiple
    /// inheritance — ruby has the former without the latter — which is the
    /// clearest evidence the old name was describing the wrong thing.
    ///
    /// ⛔ Still the WRONG CARRIER: by directives.md §3 this describes how a
    /// CALLEE is dispatched, which is question 3 — a property of the
    /// declaration, not of a region of code. It reads one channel rather than
    /// two, which is the most that can be fixed without the declaration-side
    /// field existing.
    pub(crate) fn super_is_cooperative(&self) -> bool {
        self.profile.class_multiple_inheritance
    }

    /// A missing argument binds `undefined` (ECMA-262 §10.2.1.1).
    pub(crate) fn missing_arg_is_undefined(&self) -> bool {
        self.directives().missing_arg_is_undefined.unwrap_or(false)
    }

    /// Static fields are own properties of the class object.
    pub(crate) fn static_fields_are_own_properties(&self) -> bool {
        self.directives().static_fields_are_own_properties.unwrap_or(false)
    }

    /// Private members are internal slots, not properties (JS `#x`).
    pub(crate) fn supports_private_fields(&self) -> bool {
        self.profile.supports_private_fields
    }

    /// Properties and methods occupy separate namespaces.
    pub(crate) fn separate_property_method_namespace(&self) -> bool {
        self.profile.separate_property_method_namespace
    }

    /// The class object carries `__name__` / `__mro__` / `__bases__`.
    pub(crate) fn class_introspection_metadata(&self) -> bool {
        self.profile.class_introspection_metadata
    }

    /// Default argument expressions evaluate once at definition time.
    pub(crate) fn default_args_evaluated_once(&self) -> bool {
        self.profile.default_args_evaluated_once
    }

    /// ECMA `new` dispatch (§10.2.2): an explicitly returned object wins.
    pub(crate) fn ecma_new_dispatch(&self) -> bool {
        self.profile.ecma_new_dispatch
    }

    /// An `async` body is implicitly wrapped in try/catch.
    pub(crate) fn async_wraps_body_in_try(&self) -> bool {
        self.profile.async_wraps_body_in_try
    }

    /// Every function has an implicit `arguments` object.
    pub(crate) fn has_arguments_object(&self) -> bool {
        self.profile.has_arguments_object
    }

    /// Thrown errors carry the ECMA `Error` shape.
    pub(crate) fn ecma_error_object_shape(&self) -> bool {
        self.profile.ecma_error_object_shape
    }

    /// Methods are overridable without an explicit marker.
    pub(crate) fn methods_virtual_by_default(&self) -> bool {
        self.profile.methods_virtual_by_default
    }

    /// Wrong-arity calls are accepted rather than an error.
    pub(crate) fn relaxed_call_arity(&self) -> bool {
        self.profile.relaxed_call_arity
    }

    /// The language has `undefined` distinct from `null`.
    pub(crate) fn has_undefined_value(&self) -> bool {
        self.profile.has_undefined_value
    }

    /// Class members carry declared metadata readable at run time.
    pub(crate) fn class_member_metadata(&self) -> bool {
        self.profile.class_member_metadata
    }


    pub(crate) fn method_receiver_model(&self) -> Option<vybe_ast::MethodReceiver> {
        if let Some(model) = self.directives().method_receiver {
            return Some(model);
        }
        None
    }


    // `static_methods_take_receiver` is GONE. It answered "do STATIC methods
    // carry the called class as a receiver" with `profile.name == "php"` — a
    // language NAME standing in for a property of the DECLARATION. It is now
    // `NormalClass::late_static_binding`, set by the php frontend the way seven
    // frontends already set `explicit_self_param`, read directly where the
    // class is in scope and through `classes_with_late_static_binding` where
    // only its name is. One declaration, both ends.

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
        // Stated by the walker on `Module.directives`, so it travels with the
        // UNIT and a multi-language bundle answers per unit. A profile is
        // installed once per compilation and cannot.
        if let Some(model) = self.directives().method_receiver {
            return model == vybe_ast::MethodReceiver::CallSite;
        }
        false
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

    /// Box an i32 comparison result as a `Bool` — unless a CONDITION asked for
    /// the i32, in which case the boxing is skipped and that is reported.
    ///
    /// Skipping and reporting are deliberately the SAME statement. Two
    /// separate booleans would drift, and the two drift directions are not
    /// equally bad: skip-without-report only costs the ladder, but
    /// report-without-skip hands `BR_IF` a boxed `Bool` — which it accepts —
    /// and the loop branches on the wrong thing in total silence.
    pub(super) fn emit_i32_to_bool_or_report(&mut self) {
        if std::mem::take(&mut self.want_i32_condition) {
            self.gave_i32_condition = true;
            return;
        }
        let line = self.line;
        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
    }

    /// Compile `cond` and leave an **i32** 0/1 on the stack.
    ///
    /// The general path is `compile_expr` + `emit_condition_truthiness_from_stack`.
    /// But a relational operator has already produced an i32 — `emit_js_lt` and
    /// friends end in `f64.lt` — and the `emit_i32_to_bool` after them exists
    /// only for VALUE position. In condition position the truthiness ladder
    /// undid it immediately, via its own `js-boolean:test` + `js-boolean:cast`:
    /// three host calls and two branches to turn an i32 into an i32.
    ///
    /// Soundness rests on two things, neither of them a promise made here:
    /// `compile_expr` TAKES the request at entry, so `a < b && c < d` compiles
    /// its comparisons with the request clear and still boxes them; and only
    /// emitters whose result provably came from a WASM compare opcode honour
    /// it. `emit_rich_compare_locals` never does — its dunder arm returns the
    /// user's `__lt__` value, which can be any object — so Python and Pascal
    /// keep the full ladder.
    pub(super) fn compile_condition_to_i32(
        &mut self,
        cond: &vybe_ast::Expression,
    ) -> Result<(), String> {
        self.want_i32_condition = true;
        self.gave_i32_condition = false;
        let result = self.compile_expr(cond);
        self.want_i32_condition = false;
        let gave = std::mem::take(&mut self.gave_i32_condition);
        result?;
        if !gave {
            self.emit_condition_truthiness_from_stack();
        }
        Ok(())
    }

    /// Turn the value on the stack into an i32 truth, by the rule the
    /// `truthiness` DIRECTIVE states.
    ///
    /// This is the ONE place that answers "is this true". Every site that turns
    /// a value into a condition routes here — `if`, `while`, `and`/`or`,
    /// `bool()`, and `Unary{Not}`. They used to decide separately and drifted:
    /// `emit_dyn_not` never asked, so `assert []`, which desugars to a
    /// hand-built `Unary{Not}`, silently passed under a protocol language.
    ///
    /// Under [`Truthiness::Protocol`] the ladder is CPython §3.3.1 verbatim —
    /// [`ProtocolSlot::Bool`], then [`ProtocolSlot::Len`], then the value. A
    /// builtin `[]` is falsy through the SAME `Len` rung a user class with
    /// `__len__` uses, so there is no "empty collections" special case and a
    /// class in any language earns the behaviour by binding the slot.
    pub(super) fn emit_condition_truthiness_from_stack(&mut self) {
        if !self.protocol_truthiness() {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            return;
        }

        let line = self.line;
        let value_slot = self.define_local("__truth_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        // ── rung 0: absent is false ─────────────────────────────────────
        // Nothing below can run on a null: `ecma:value.typeof` traps on one,
        // so the very first probe took the whole ladder down and `if None:`
        // never reached the coercion that would have answered it.
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::REF_IS_NULL);
        self.chunk().emit_if_value(line);
        inst!(self, core_wasm::i32_const, 0);
        self.chunk().emit_else(line);

        // STRUCT_GET traps on a primitive, so every slot probe sits behind the
        // object test — the same gate `emit_rich_unary` uses.
        let typeof_idx = self.import("ecma:value", "typeof");
        let is_object_slot = self.define_local("__truth_is_object");
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_host_call(typeof_idx, 1);
        self.emit_const(Value::String(Arc::from("object")));
        fn_call!(self, "wasm:js-string", "equals", 2);
        self.emit_u16(Op::LOCAL_SET, is_object_slot);

        // ── rung 1: ProtocolSlot::Bool ──────────────────────────────────
        let bool_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal(&vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Bool)));
        let bool_method = self.define_local("__truth_bool_method");
        self.emit_ref_null_local(bool_method);
        self.emit_u16(Op::LOCAL_GET, is_object_slot);
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &bool_key);
        self.emit_u16(Op::LOCAL_SET, bool_method);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, bool_method);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, bool_method);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::callable::emit_direct_invoke_chunk(self.chunk(), 1, line);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_else(line);

        // ── rung 2: ProtocolSlot::Len on a user class ───────────────────
        let len_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal(&vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Len)));
        let len_method = self.define_local("__truth_len_method");
        self.emit_ref_null_local(len_method);
        self.emit_u16(Op::LOCAL_GET, is_object_slot);
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &len_key);
        self.emit_u16(Op::LOCAL_SET, len_method);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, len_method);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, len_method);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::callable::emit_direct_invoke_chunk(self.chunk(), 1, line);
        fn_call!(self, "wasm:js-number", "toI32", 1);
        inst!(self, core_wasm::i32_const, 0);
        self.chunk().emit_op(Op::I32_NE, line);
        self.chunk().emit_else(line);

        // ── rung 3: the builtin length — a str/array/map answers `Len`
        //            intrinsically, so this is the same rung, not a case ──
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

        // ── rung 4: the value itself ────────────────────────────────────
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_end(line);

        self.chunk().emit_end(line); // rung 2 (Len slot)
        self.chunk().emit_end(line); // rung 1 (Bool slot)
        self.chunk().emit_end(line); // rung 0 (null)
    }

    /// `local = null` — a slot probe starts empty so the non-object path skips
    /// the `STRUCT_GET` without leaving the local undefined.
    fn emit_ref_null_local(&mut self, slot: u16) {
        let line = self.line;
        self.chunk()
            .emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        self.emit_u16(Op::LOCAL_SET, slot);
    }

    pub(super) fn save_js_this(&mut self, local_name: &str) -> Option<u16> {
        if !self.ambient_this() {
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

    /// Clear the ambient receiver so a CONSTRUCTION allocates instead of
    /// adopting whatever `this` happens to be live.
    ///
    /// A constructor reads its receiver from `__js_this` and allocates only
    /// when that global is absent — `struct.new_default` sits in the `else` of
    /// a null test at the top of every `__<Class>_ctor_N`. Every other
    /// `save_js_this` site pairs the save with a `bind_js_this_from_local`; the
    /// `New` emit was the one that saved and restored without ever writing a
    /// value in between, so inside an instance method the constructor found the
    /// enclosing receiver, skipped the allocation, and wrote its fields into
    /// it. `One bump() => One(this.v + 1)` answered `identical(a, b) == true`;
    /// a constructor in a class-static field initializer shared one object for
    /// the same reason.
    ///
    /// `new` is unconditional: it always makes a fresh object, whatever the
    /// caller's context. Clearing states that, rather than relying on the
    /// caller happening to have no receiver.
    pub(super) fn clear_js_this(&mut self) {
        if !self.ambient_this() {
            return;
        }
        let line = self.line;
        self.chunk()
            .emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        self.emit_global_write("__js_this");
    }

    /// Write the ambient receiver from the value on the stack.
    ///
    /// **Emits unconditionally.** The `receiver_binding` directive is answered
    /// ONCE, by `save_js_this`, whose `Option` says whether this language has an
    /// ambient receiver at all; a site that got `None` must not compute a
    /// receiver, must not push one, and must not call this. There is no second
    /// decision here and no silent no-op to fall into.
    ///
    /// This used to decide for itself — `if !self.ambient_this() { return; }` —
    /// while every caller pushed unconditionally:
    ///
    /// ```ignore
    /// let saved = self.save_js_this("__js_prev_this_member");
    /// self.emit_u16(Op::LOCAL_GET, obj_slot);   // ALWAYS pushed
    /// self.set_js_this_from_stack();            // popped only if Ambient
    /// ```
    ///
    /// `ReceiverBinding::Ambient` is declared by js and dart ONLY, so the other
    /// fourteen languages leaked one operand per member read, at ten sites. The
    /// stray is invisible while nothing live sits under it — a statement
    /// boundary truncates it — and fatal where something does: `W(self.v)`
    /// inside a method left the ctor ref buried and `CALL_REF` took the stray as
    /// the callee (`Not a function`), and `{'k': self.v}` compiled to
    /// `{<V object>: 5}` — the stray DISPLACED the key, silently.
    ///
    /// A guard that emits nothing is dead code for the language it fires on,
    /// and dead code that is *also* half a pair is how the two halves drifted.
    /// `restore_js_this(None)` reads the same `Option`, so all three members of
    /// the save/bind/restore triple now agree by construction.
    pub(super) fn set_js_this_from_stack(&mut self) {
        self.emit_global_write("__js_this");
    }

    /// **The** answer to "what is `this` here". One function, one answer.
    ///
    /// flexclassplan.md §"Two implementations of `what is this here`, and they
    /// disagreed" (item 2): `ExprKind::This` and `emit_js_current_this_value`
    /// were two independent resolvers, and they did not agree. Verified
    /// divergences, all three in the FALLBACK path — their ambient branches
    /// were condition-for-condition identical, which is why the bug presented
    /// as `super.m()` handing the callee an undefined receiver rather than as
    /// a general breakage:
    ///
    /// | | `ExprKind::This` | `emit_js_current_this_value` |
    /// |---|---|---|
    /// | receiver-is-a-parameter guard | yes | **none** |
    /// | local names tried | `self_kw`, `Self`, `self`, `this` | **`self_kw` only** |
    /// | derived-ctor TDZ (§9.1.1.3.4) | yes | **none** |
    ///
    /// The union is the first column throughout; the second contributed
    /// nothing the union needs.
    ///
    /// ⚠ This resolves the receiver as it is bound TODAY, which still includes
    /// the ambient `__js_this` global. That global is not a WASM concept —
    /// core wasm has no `this`, and a mutable module global standing in for a
    /// parameter needs the `save`/`restore` pair around every call, i.e. a
    /// hand-rolled shadow stack for something the substrate models natively.
    /// Collapsing the two resolvers is the PREREQUISITE for removing it: while
    /// two sites disagreed about what `this` is, nothing could safely change
    /// what it resolves TO.
    pub(super) fn emit_receiver_value(&mut self) {
        // ⛔ A CHUNK THAT TOOK THE RECEIVER AS A PARAMETER ALREADY HAS ONE —
        // the ambient global is not it.
        //
        // A property accessor is compiled as a chunk whose FIRST PARAMETER is
        // the receiver (`classes.rs`, `declare_receiver_first_accessor`). The
        // ambient branch fires first and reads `__js_this`, a global that
        // accessor chunk never set — so `this` inside an accessor answered null
        // while the same `this` in a method on the SAME object in the SAME run
        // was correct.
        //
        // Invisible from parsed source: every walker emits `Ident("this")` for
        // an accessor, which resolves the local and works. It bites only a
        // producer of SYNTHESIZED class AST — dart's `core_classes`, flutter's
        // adapter classes. There is no split to memorise; `This` means the
        // receiver.
        let receiver_is_a_parameter = self.chunks[self.current]
            .handled_call_tags
            .iter()
            .any(|tag| tag == vybe_runtime::RECEIVER_FIRST_ACCESSOR_TAG);

        if !receiver_is_a_parameter
            && self.ambient_this()
            && self.current_class.is_some()
            && self.current_func_name.as_deref() != Some("<lambda>")
            && self.current_func_name.as_deref().is_some_and(|name| {
                !name.eq_ignore_ascii_case(&self.profile.constructor_name)
            })
        {
            self.emit_global_read("__js_this");
            return;
        }

        let self_kw = self.profile.self_keyword.clone();
        if let Some(slot) = self
            .scope()
            .resolve(&self_kw)
            .or_else(|| self.scope().resolve("Self"))
            .or_else(|| self.scope().resolve("self"))
            .or_else(|| self.scope().resolve("this"))
        {
            // §9.1.1.3.4: inside a derived constructor `this` is in TDZ until
            // super() runs — reading it while the slot is still null throws a
            // ReferenceError rather than handing back null.
            if self.js_derived_ctor_ctx == Some((self.current, slot)) {
                let l = self.line;
                crate::primitives::classes::emit_this_initialized_guard(self.chunk(), slot, l);
            }
            self.emit_u16(Op::LOCAL_GET, slot);
            return;
        }

        if self.scopes.len() > 1 {
            // Arrow function: capture `this` from the enclosing scope.
            if self.resolve_upvalue(self.scopes.len() - 1, &self_kw).is_some() {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(&self_kw);
                let l = self.line;
                crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                return;
            }
            if self.ambient_this()
                && self
                    .resolve_upvalue(self.scopes.len() - 1, "__js_this")
                    .is_some()
            {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index("__js_this");
                let l = self.line;
                crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                return;
            }
        }

        if self.ambient_this() {
            self.emit_global_read("__js_this");
        } else {
            self.emit_null();
        }
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
            _ => Vec::new(),
        }
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
