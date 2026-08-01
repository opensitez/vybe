//! Class, constructor, and free-function compilation.
//!
//! Extracted from `compiler.rs` to keep that file navigable. The
//! methods on this `impl Compiler { ... }` block are private by
//! convention (they're only called from other compiler methods) and
//! crate-private for the `dotnet_register` bridge.

use super::*;
use crate::primitives::ArrayBindingMetadata;
use crate::primitives::class_normalize::{BaseCall, NormalConstructor, NormalMethod};

/// Global name of the arity-specialized constructor overload for `prefix`.
///
/// Overloaded constructors emit one global per arity (`Point$arity2`), with the
/// unsuffixed class global as the fallback entry. The `$arity` separator is not
/// any language's spelling — it is this compiler's private encoding, and both
/// the writer (this module) and the readers (`expressions.rs`, `calls.rs`,
/// `arrays.rs`, `link.rs`) must agree on it. It used to be spelled inline in 15
/// places across 6 files.
pub fn ctor_global_for(prefix: &str, arity: usize) -> String {
    format!("{prefix}$arity{arity}")
}

impl Compiler {
    fn qualified_nested_type_name(enclosing: &str, nested: &str) -> String {
        if nested.contains('.') {
            nested.to_string()
        } else {
            format!("{enclosing}.{nested}")
        }
    }

    fn qualify_nested_type_stmt(stmt: &Statement, enclosing: &str) -> Statement {
        let mut stmt = stmt.clone();
        match &mut stmt.kind {
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                *name = Self::qualified_nested_type_name(enclosing, name);
                let owner = name.clone();
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        **nested = Self::qualify_nested_type_stmt(nested, &owner);
                    }
                }
            }
            StmtKind::InterfaceDecl { name, .. } => {
                *name = Self::qualified_nested_type_name(enclosing, name);
            }
            StmtKind::EnumDecl { name, .. } => {
                *name = Self::qualified_nested_type_name(enclosing, name);
            }
            _ => {}
        }
        stmt
    }

    fn fixed_array_zero_expr(type_hint: &str) -> Option<Expression> {
        let trimmed = type_hint.trim();
        if !trimmed.starts_with('[') {
            return None;
        }

        let close = trimmed.find(']')?;
        let head = trimmed.get(1..close)?.trim();
        if head.is_empty() || head == "..." {
            return None;
        }

        let len = head.parse::<usize>().ok()?;
        let element_type = trimmed.get(close + 1..)?.trim();
        let element_expr = Self::fixed_array_zero_expr(element_type)
            .unwrap_or_else(|| Self::array_default_element_expr(Some(element_type)));

        Some(Expression::new(ExprKind::Cast {
            expr: Box::new(Expression::new(ExprKind::Array(
                (0..len)
                    .map(|_| ArrayElement {
                        key: None,
                        value: element_expr.clone(),
                        spread: false,
                        by_ref: false,
                    })
                    .collect(),
            ))),
            type_name: trimmed.to_string(),
        }))
    }

    fn array_default_element_expr(type_hint: Option<&str>) -> Expression {
        match type_hint
            .map(str::trim)
            .map(|hint| hint.strip_suffix("()").unwrap_or(hint))
            .map(Self::normalize_type_hint)
            .as_deref()
        {
            Some("integer") | Some("int") | Some("int32") | Some("longint") | Some("real")
            | Some("double") | Some("float") | Some("single") | Some("decimal") | Some("long")
            | Some("int64") | Some("short") | Some("int16") | Some("uint") | Some("uint32")
            | Some("ulong") | Some("uint64") | Some("ushort") | Some("uint16") | Some("byte")
            | Some("sbyte") | Some("char") => Expression::new(ExprKind::Lit(Literal::Int(0))),
            Some("boolean") | Some("bool") => Expression::new(ExprKind::Lit(Literal::Bool(false))),
            Some(type_hint) if Self::is_string_type_hint(type_hint) => Expression::string(""),
            _ => Expression::null(),
        }
    }

    fn array_bounds_extent_expr(bounds: &[Expression]) -> Option<Expression> {
        let mut iter = bounds.iter().cloned();
        let first = iter.next()?;
        Some(iter.fold(first, |acc, bound| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(acc),
                right: Box::new(bound),
            })
        }))
    }

    fn emit_class_field_initializer(
        &mut self,
        owner_slot: u16,
        field_name: &str,
        type_hint: Option<&str>,
        init: Option<&Expression>,
        array_bounds: Option<&[Expression]>,
        is_value_type: bool,
        line: u32,
    ) -> Result<(), String> {
        let value_slot = self.define_local("__field_init_value");
        if let Some(init_expr) = init {
            self.compile_expr(init_expr)?;
        } else if let Some(extent) = array_bounds.and_then(Self::array_bounds_extent_expr) {
            if let Some(init_expr) = type_hint.and_then(Self::fixed_array_zero_expr) {
                self.compile_expr(&init_expr)?;
            } else {
                let init_expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("Array")),
                    args: vec![
                        Argument::positional(extent),
                        Argument::positional(Self::array_default_element_expr(type_hint)),
                    ],
                    optional: false,
                });
                self.compile_expr(&init_expr)?;
            }
        } else if let Some(init_expr) = type_hint.and_then(Self::fixed_array_zero_expr) {
            self.compile_expr(&init_expr)?;
        } else if let Some(type_name) =
            type_hint.and_then(|type_hint| self.user_value_type_name_from_hint(type_hint))
        {
            let ctor_global = {
                let overload = ctor_global_for(&type_name, 0);
                if self.defined_globals.contains(&overload) {
                    overload
                } else {
                    type_name.clone()
                }
            };
            let idx = self.str_const(&ctor_global);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u8(Op::CALL_REF, 0);
        } else if is_value_type {
            self.emit_default_value_for_type_hint(type_hint);
        } else if self.profile.has_undefined_value {
            // JS spec: declared fields with no initializer default to undefined (not null).
            inst!(self, core_wasm::undefined);
        } else {
            self.emit(Op::NULL);
        }
        self.emit_u16(Op::LOCAL_SET, value_slot);

        crate::primitives::classes::emit_init_field_start(self.chunk(), owner_slot, line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::classes::emit_init_field_end(self.chunk(), field_name, line);
        Ok(())
    }

    fn class_requires_form_identity_stamp(&self, parent: &Option<String>) -> bool {
        let mut current = parent.clone().map(|name| self.canon(&name));
        let mut visited = std::collections::HashSet::new();

        while let Some(name) = current {
            if !visited.insert(name.clone()) {
                break;
            }
            if name.eq_ignore_ascii_case("form")
                || self.reflection_is_assignable_from("Form", &name)
            {
                return true;
            }
            current = self
                .pending_classes
                .get(name.as_str())
                .and_then(|pending| pending.parent.clone())
                .or_else(|| self.reflection_base_type_name(&name));
        }

        false
    }

    fn emit_form_identity_stamp(&mut self, this_slot: u16, class_name: &str, _line: u32) {
        let stamped_name = self.canon(class_name);
        let set_property = self.import("vybe:gui", "controlSetProperty");
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_const(Value::String(Arc::from("Name")));
        self.emit_const(Value::String(Arc::from(stamped_name.as_str())));
        self.emit_host_call(set_property, 3);
        self.emit(Op::DROP);
    }

    fn emit_store_super_ref(&mut self, this_slot: u16, parent_name: &str) {
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_var_get(parent_name);
        let super_key = self.str_const("__super");
        self.emit_u16(Op::STRUCT_SET, super_key);
        self.emit(Op::DROP);
    }

    /// Push the ECMA [[HomeObject]].[[Prototype]] base used by `super`.
    ///
    /// Instance methods resolve through `CurrentClass.prototype.__proto__`;
    /// static methods resolve through `CurrentClass.__proto__`. Looking this
    /// up at runtime keeps `Object.setPrototypeOf(C.prototype, ...)` visible
    /// to already-compiled `super.foo()` calls.
    pub(super) fn emit_js_super_home_base(&mut self) {
        let Some(class_name) = self.current_class.clone() else {
            self.emit(Op::NULL);
            return;
        };

        self.emit_var_get(&class_name);
        if !self.current_member_is_static {
            let prototype_key = self.str_const("prototype");
            self.emit_u16(Op::STRUCT_GET, prototype_key);
        }
        let proto_link_key = self.str_const("__proto__");
        self.emit_u16(Op::STRUCT_GET, proto_link_key);
    }

    /// Instance identity + member stamps a derived constructor applies to
    /// the `this` produced by `super()`: __type / type-id, __super link,
    /// base-method saves, field initializers, instance-method binds, form
    /// identity and auto-init calls. Emitted right after a TOP-LEVEL
    /// `super()` statement, or (JS, when `super()` is nested — e.g.
    /// `try{super();}catch{}`) after the whole body guarded by
    /// `this != null`, since a nested super's completion point isn't
    /// statically known.
    #[allow(clippy::too_many_arguments)]
    /// Emit the parent constructor VALUE for class wiring. Bound parents
    /// (user classes, locals, globals) resolve normally; unbound intrinsic
    /// exception parents (`extends Error` / `extends RangeError` …) resolve
    /// through the canonical `__ctor_<Name>` anchors — the bare names are
    /// compile-time bindings, not vm globals, so a plain GLOBAL_GET yields
    /// null and silently breaks the prototype chain.
    fn emit_parent_ctor_value(&mut self, parent_name: &str) {
        self.emit_parent_value(parent_name, false);
    }

    /// Emit the parent CLASS OBJECT for prototype-chain wiring.
    ///
    /// Same resolution as `emit_parent_ctor_value` except the `$arity0`
    /// arity-helper global is never taken: that global holds a bare
    /// allocate-and-return function, not the class object, so it carries
    /// neither the real `.prototype` nor the statics. Linking
    /// `C.prototype.__proto__` / `C.__proto__` through it silently breaks
    /// `super.m()` and static inheritance for every parent that has a
    /// zero-argument constructor (i.e. every parent that declares none).
    fn emit_parent_class_value(&mut self, parent_name: &str) {
        self.emit_parent_value(parent_name, true);
    }

    fn emit_parent_value(&mut self, parent_name: &str, want_class_object: bool) {
        let pname = self.canon(parent_name);
        let default_ctor = ctor_global_for(&pname, 0);
        let bound = self.scope().resolve(parent_name).is_some()
            || self.defined_globals.contains(&pname)
            || self.defined_classes.contains(&pname);
        if !bound
            && self.profile.ecma_error_object_shape
            && common::errors::is_exception_type(parent_name)
            && !self.shadows_builtin_type(parent_name)
        {
            let key = self.str_const(&format!("__ctor_{parent_name}"));
            self.emit_u16(Op::GLOBAL_GET, key);
        } else if !bound
            && self.profile.has_ecma_globals
            && Self::is_ecma_typed_array_ctor_name(parent_name)
            && !self.shadows_builtin_type(parent_name)
        {
            let key = self.str_const(&format!("__ctor_{parent_name}"));
            self.emit_u16(Op::GLOBAL_GET, key);
        } else if !bound
            && self.profile.has_ecma_globals
            && Self::is_ecma_array_buffer_ctor_name(parent_name)
            && !self.shadows_builtin_type(parent_name)
        {
            let key = self.str_const(&format!("__ctor_{parent_name}"));
            self.emit_u16(Op::GLOBAL_GET, key);
        } else if !bound
            && self.profile.has_ecma_globals
            && Self::is_ecma_collection_ctor_name(parent_name)
            && !self.shadows_builtin_type(parent_name)
        {
            let key = self.str_const(&format!("__ctor_{parent_name}"));
            self.emit_u16(Op::GLOBAL_GET, key);
        } else if !want_class_object && self.defined_globals.contains(&default_ctor) {
            let key = self.str_const(&default_ctor);
            self.emit_u16(Op::GLOBAL_GET, key);
        } else {
            self.emit_var_get(&pname);
        }
    }

    fn is_ecma_typed_array_ctor_name(name: &str) -> bool {
        matches!(
            Self::normalize_type_hint(name).as_str(),
            "int8array"
                | "uint8array"
                | "uint8clampedarray"
                | "int16array"
                | "uint16array"
                | "int32array"
                | "uint32array"
                | "float32array"
                | "float64array"
                | "bigint64array"
                | "biguint64array"
        )
    }

    fn is_ecma_array_buffer_ctor_name(name: &str) -> bool {
        matches!(
            Self::normalize_type_hint(name).as_str(),
            "arraybuffer" | "sharedarraybuffer"
        )
    }

    fn is_ecma_collection_ctor_name(name: &str) -> bool {
        matches!(
            Self::normalize_type_hint(name).as_str(),
            "array" | "map" | "set" | "weakmap" | "weakset" | "object"
        )
    }

    /// True when `parent_name` is a framework GUI control used as a class
    /// parent under a profile that has the dotnet namespace mounted.
    ///
    /// `canonical_control_name` is a global, language-blind string table, so it
    /// matches `"Timer"`/`"Panel"`/… regardless of which language is compiling.
    /// Gating on `use_dotnet` re-establishes the resolver scoping the bare
    /// string match skips: only treat the name as a control when `dotnet.*` is
    /// actually in scope, so a same-named class in a non-GUI language (e.g. a
    /// Python `class X(Timer)`) never misroutes to `vybe:gui`.
    pub(crate) fn is_framework_control_parent(&self, parent_name: &str) -> bool {
        self.profile.namespaces.use_dotnet
            && !common::gui::canonical_control_name(parent_name).is_empty()
            // A user class shadowing a control name (`class Timer { … }`) wins,
            // mirroring the standalone-`new` path's `!dotnet_ctor_registered`
            // guard — otherwise `class X : Timer` over the user's own Timer
            // would misroute to `vybe:gui new_Timer`.
            && !self.defined_classes.contains(&self.canon(parent_name))
    }

    fn dotnet_descriptor_parent_has_no_user_ctor(&self, parent_name: &str) -> bool {
        self.profile.namespaces.use_dotnet
            && vybe_runtime::registry::platform_owns_descriptor_class(parent_name)
            && !self.is_framework_control_parent(parent_name)
            && !self.defined_classes.contains(&self.canon(parent_name))
    }

    /// A framework GUI control (`Button`, `Form`, `Panel`, …) used as a class
    /// parent constructs through the `vybe:gui` host factory directly rather
    /// than through a per-class constructor global. The host `new_<Type>`
    /// builds the complete control object — identity fields plus the
    /// `Controls`/`components` collections — so no emitted ctor chunk is
    /// needed, and control access resolves through the component descriptor at
    /// each call site (see `dotnet_framework_instance_method_owner` /
    /// `lookup_instance_property`). Returns `true` when it emitted the
    /// construction into `this_slot`; `false` for non-control parents, leaving
    /// the caller's normal constructor-ref path to run.
    fn try_emit_framework_control_base(
        &mut self,
        parent_name: &str,
        base_args: &[Expression],
        this_slot: u16,
    ) -> Result<bool, String> {
        let canonical = common::gui::canonical_control_name(parent_name);
        if canonical.is_empty() {
            return Ok(false);
        }
        let host_name = common::gui::host_fn_new_control(&canonical);
        let new_idx = self.import(common::gui::GUI_MODULE, &host_name);
        for a in base_args {
            self.compile_expr(a)?;
        }
        let line = self.line;
        common::gui::emit_new_control(self.chunk(), new_idx, base_args.len() as u8, line);
        self.emit_u16(Op::LOCAL_SET, this_slot);
        Ok(true)
    }

    fn emit_derived_ctor_stamps(
        &mut self,
        name: &str,
        this_slot: u16,
        parent: &Option<String>,
        instance_method_names: &[String],
        field_inits: &[(
            String,
            Option<String>,
            Option<Expression>,
            Option<Vec<Expression>>,
        )],
        instance_methods: &[&(String, usize, bool, bool)],
        method_capture_name_map: &HashMap<usize, Vec<String>>,
        method_rest_fixed_counts: &HashMap<usize, u8>,
        is_value_type: bool,
        should_stamp_form_identity: bool,
        body_stmts: &[Statement],
        user_body: &[Statement],
        auto_init_methods: &[String],
        line: u32,
    ) -> Result<(), String> {
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_const(Value::String(Arc::from(name)));
        let type_key = self.str_const("__type");
        self.emit_u16(Op::STRUCT_SET, type_key);
        self.emit(Op::DROP);
        let tid_key = self.str_const(&format!("__tid_{}", self.canon(name)));
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_u16(Op::GLOBAL_GET, tid_key);
        self.emit(Op::SET_TYPE_ID);
        self.emit(Op::DROP);
        if is_value_type {
            crate::primitives::classes::emit_value_equality_stamp(self.chunk(), this_slot, line);
        }
        if let Some(parent_name) = parent {
            let pname = self.canon(parent_name);
            for method_name in instance_method_names {
                crate::primitives::classes::emit_save_base_method(
                    self.chunk(),
                    this_slot,
                    method_name,
                    line,
                );
            }
            self.emit_store_super_ref(this_slot, &pname);
        }
        for (fname, type_hint, init, array_bounds) in field_inits {
            self.emit_class_field_initializer(
                this_slot,
                fname,
                type_hint.as_deref(),
                init.as_ref(),
                array_bounds.as_deref(),
                is_value_type,
                line,
            )?;
        }
        // Multiple inheritance (opt-in): attach every MRO ancestor's methods in
        // reverse-C3 order (lowest priority first) BEFORE binding self's own
        // methods below, so self stays highest and the diamond's shared base is
        // overridden by the nearer base. No-op for single-inheritance classes
        // and for languages that don't set `class_multiple_inheritance`.
        self.emit_mi_ancestor_methods(name, this_slot, line);
        for (mname, mci, _, _) in instance_methods {
            if mname.starts_with("__get_") {
                let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                crate::primitives::object::emit_bind_getter(
                    self.chunk(),
                    this_slot,
                    prop,
                    *mci,
                    line,
                );
            } else if mname.starts_with("__set_") {
                let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                crate::primitives::object::emit_bind_setter(
                    self.chunk(),
                    this_slot,
                    prop,
                    *mci,
                    line,
                );
            } else {
                let capture_names = method_capture_name_map
                    .get(mci)
                    .map(Vec::as_slice)
                    .unwrap_or(&[]);
                self.emit_bind_instance_method_with_aliases(
                    this_slot,
                    mname,
                    *mci,
                    capture_names,
                    method_rest_fixed_counts.get(mci).copied(),
                    !self.class_prototype_dispatch(),
                )?;
            }
        }
        if should_stamp_form_identity && !body_has_identity_stamp(body_stmts) {
            self.emit_form_identity_stamp(this_slot, name, line);
        }
        for aim in auto_init_methods {
            let has_method = instance_methods
                .iter()
                .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
            if has_method && !body_calls_method(user_body, aim) {
                crate::primitives::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
            }
        }
        Ok(())
    }

    /// C3 linearization of a class's MRO — canonical names, self first, the
    /// implicit `object` tail omitted. Recurses through `PendingClass.bases`;
    /// a name not in the registry (external/builtin) is treated as a leaf.
    /// Only meaningful under `class_multiple_inheritance`; callers gate on it.
    pub(super) fn c3_linearize(&self, name: &str) -> Vec<String> {
        let canon = self.canon(name);
        let bases: Vec<String> = match self.pending_classes.get(&canon) {
            Some(pc) => pc.bases.iter().map(|b| self.canon(b)).collect(),
            None => return vec![canon],
        };
        if bases.is_empty() {
            return vec![canon];
        }
        let mut seqs: Vec<Vec<String>> = bases.iter().map(|b| self.c3_linearize(b)).collect();
        seqs.push(bases);
        let mut result = vec![canon];
        result.extend(c3_merge(seqs));
        result
    }

    /// Attach every MRO-ancestor's instance methods onto a new instance, in
    /// reverse-C3 order (lowest priority first) so a nearer base overrides a
    /// more-distant one and the diamond's shared base is bound once. Called
    /// just before self's own methods are bound (which therefore win). No-op
    /// unless the profile opts into MI and the class has >1 declared base.
    ///
    /// Ancestor methods are re-bound from their chunk index with zero upvalues;
    /// a base method that closes over an enclosing scope is a known gap (rare —
    /// module-level classes capture nothing).
    fn emit_mi_ancestor_methods(&mut self, name: &str, this_slot: u16, line: u32) {
        if !self.profile.class_multiple_inheritance {
            return;
        }
        let canon = self.canon(name);
        let is_multi = self
            .pending_classes
            .get(&canon)
            .is_some_and(|pc| pc.bases.len() > 1);
        if !is_multi {
            return;
        }
        let mut ancestors = self.c3_linearize(name); // [self, B, C, A, …]
        ancestors.remove(0); // drop self — bound separately as highest priority
        ancestors.reverse(); // lowest priority first

        // Collect (method_name, chunk_idx, rest_fixed) up front to avoid holding
        // a `pending_classes` borrow across the emit calls. Sort per class for
        // deterministic bytecode (HashMap order is not stable).
        let mut binds: Vec<(String, usize, Option<u8>)> = Vec::new();
        for cls in &ancestors {
            let Some(pc) = self.pending_classes.get(cls) else {
                continue;
            };
            let mut per_class: Vec<(String, usize, Option<u8>)> = pc
                .instance_method_overloads
                .iter()
                .filter_map(|(mname, overloads)| {
                    overloads.first().map(|ov| {
                        let rest_fixed = ov
                            .signature
                            .has_rest
                            .then(|| ov.signature.param_names.len().saturating_sub(1) as u8);
                        (mname.clone(), ov.chunk_idx, rest_fixed)
                    })
                })
                .collect();
            per_class.sort_by(|a, b| a.0.cmp(&b.0));
            binds.extend(per_class);
        }

        let bind_on_access = self.profile.methods_bind_on_access;
        for (mname, chunk_idx, rest_fixed) in binds {
            if let Some(prop) = mname.strip_prefix("__get_") {
                crate::primitives::object::emit_bind_getter(
                    self.chunk(),
                    this_slot,
                    prop,
                    chunk_idx,
                    line,
                );
            } else if let Some(prop) = mname.strip_prefix("__set_") {
                crate::primitives::object::emit_bind_setter(
                    self.chunk(),
                    this_slot,
                    prop,
                    chunk_idx,
                    line,
                );
            } else {
                crate::primitives::object::emit_bind_bound_method(
                    self.chunk(),
                    this_slot,
                    &mname,
                    chunk_idx,
                    rest_fixed,
                    bind_on_access,
                    line,
                );
            }
        }
    }

    pub(super) fn captured_name_for_upvalue(
        &self,
        scope_idx: usize,
        upvalue_idx: u8,
    ) -> Option<String> {
        let upvalue = self
            .scopes
            .get(scope_idx)?
            .upvalues
            .get(upvalue_idx as usize)?;
        let parent_scope_idx = scope_idx.checked_sub(1)?;
        if upvalue.is_local {
            self.scopes
                .get(parent_scope_idx)?
                .locals
                .iter()
                .find(|local| local.slot == upvalue.index)
                .map(|local| local.name.clone())
        } else {
            self.captured_name_for_upvalue(parent_scope_idx, upvalue.index as u8)
        }
    }

    fn emit_ref_func_with_captures(
        &mut self,
        func_idx: usize,
        capture_names: &[String],
        // When true, tag the funcref with the 0x80 "no-intern" flag so the VM
        // mints a FRESH object rather than the shared interned canonical one.
        // A bound method stamps a per-receiver property on its funcref, so it
        // must be distinct per binding (`C().f is C().f` False, and each
        // receiver stays correct). Does NOT touch upvalues (the closure env
        // lives at upvalue[0], so capturing the receiver there would corrupt it).
        no_intern: bool,
    ) -> Result<(), String> {
        let line = self.line;
        let count_byte = capture_names.len() as u8 | if no_intern { 0x80 } else { 0 };
        common::functions::emit_ref_func(
            &mut self.chunks[self.current],
            func_idx,
            count_byte,
            line,
        );
        for capture_name in capture_names {
            if let Some(slot) = self
                .scope()
                .resolve(capture_name)
                .or_else(|| self.scope().resolve_ci(capture_name))
            {
                common::functions::emit_closure_upvalue(
                    &mut self.chunks[self.current],
                    true,
                    slot,
                    line,
                );
                continue;
            }
            if self.scopes.len() > 1 {
                if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, capture_name) {
                    common::functions::emit_closure_upvalue(
                        &mut self.chunks[self.current],
                        false,
                        uv as u16,
                        line,
                    );
                    continue;
                }
            }
            return Err(format!(
                "failed to resolve captured class method binding '{capture_name}'"
            ));
        }
        Ok(())
    }

    /// Set `__js_new_target` to this class ONLY when currently unset.
    /// §13.3.7.1: `super()` preserves the active new.target — the chain's
    /// outermost `new` already set it; this default covers constructors
    /// invoked without an outer `new` frame.
    fn emit_default_js_new_target(&mut self, name: &str) {
        if !self.profile.ecma_new_dispatch {
            return;
        }
        let nt_key = self.str_const("__js_new_target");
        self.emit_u16(Op::GLOBAL_GET, nt_key);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunks[self.current].emit_if(line);
        self.emit_var_get(name);
        self.emit_u16(Op::GLOBAL_SET, nt_key);
        self.chunks[self.current].emit_end(line);
    }

    /// Push the prototype object for a new instance's `__proto__` link:
    /// `(__js_new_target ?? <OwnClass>).prototype`. §9.1.13
    /// OrdinaryCreateFromConstructor uses new.target's prototype, so
    /// `this.constructor` resolves to the *invoked* class even while a
    /// parent constructor body runs under `super()`.
    fn emit_load_instance_proto(&mut self, class_name: &str) {
        let nt_key = self.str_const("__js_new_target");
        self.emit_u16(Op::GLOBAL_GET, nt_key);
        inst!(self, core_wasm::dup);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunks[self.current].emit_if(line);
        self.emit(Op::DROP);
        let class_global = self.str_const(class_name);
        self.emit_u16(Op::GLOBAL_GET, class_global);
        self.chunks[self.current].emit_end(line);
        let prototype_key = self.str_const("prototype");
        self.emit_u16(Op::STRUCT_GET, prototype_key);
    }

    fn emit_bind_instance_method_with_aliases(
        &mut self,
        this_slot: u16,
        method_name: &str,
        method_chunk_idx: usize,
        capture_names: &[String],
        rest_fixed_count: Option<u8>,
        bind_receiver: bool,
    ) -> Result<(), String> {
        let receiver_key = self.str_const("__vybe_method_receiver");
        let rest_key = self.str_const("__vybe_rest_fixed_arity");

        // Prototype-dispatch profiles: the prototype is the source of truth,
        // so reassignment (`C.prototype.m = wrap(C.prototype.m)`) reaches
        // instances constructed afterwards. Falls back to the compiled ref
        // when the prototype has no entry (capture-carrying methods, class
        // expressions without a class global).
        let proto_class = if self.class_prototype_dispatch() && capture_names.is_empty() {
            self.current_class
                .clone()
                .filter(|c| self.defined_classes.contains(&self.canon(c)))
        } else {
            None
        };

        let mut bind_names = vec![method_name.to_string()];
        // The PROTOCOL SLOT this method fills, published under a key derived
        // from the slot's NUMBER (`__vybe_slot_1`), not from any language's
        // spelling. This is what a cross-language call resolves through.
        if let Some(slot_key) = self.current_class_slot_keys.get(method_name).cloned() {
            bind_names.push(slot_key);
        }
        // A `methods_bind_on_access` language (Python/Ruby) needs a DISTINCT
        // bound-method object per instance (`C().f is C().f` False). Tag these
        // funcrefs "no-intern" so the VM mints a fresh object per binding (the
        // receiver stamped below then stays per-instance). Only when the method
        // carries a receiver.
        let no_intern = self.profile.methods_bind_on_access && bind_receiver;

        for bind_name in bind_names {
            self.emit_u16(Op::LOCAL_GET, this_slot);
            if let Some(class_name) = &proto_class {
                let cname = self.canon(class_name);
                let class_idx = self.global_name_const_idx(&cname);
                self.emit_u16(Op::GLOBAL_GET, class_idx);
                let proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_GET, proto_key);
                let mkey = self.str_const(method_name);
                self.emit_u16(Op::STRUCT_GET, mkey);
                inst!(self, core_wasm::dup);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit(Op::DROP);
                self.emit_ref_func_with_captures(method_chunk_idx, capture_names, no_intern)?;
                self.chunks[self.current].emit_end(line);
            } else {
                self.emit_ref_func_with_captures(method_chunk_idx, capture_names, no_intern)?;
            }
            if bind_receiver {
                inst!(self, core_wasm::dup);
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_u16(Op::STRUCT_SET, receiver_key);
                self.emit(Op::DROP);
            }
            if self.profile.supports_private_fields && bind_name.starts_with("__js_private_") {
                inst!(self, core_wasm::dup);
                let display_name = bind_name
                    .rsplit_once('.')
                    .map(|(_, tail)| format!("#{tail}"))
                    .unwrap_or_else(|| method_name.to_string());
                self.emit_const(Value::String(Arc::from(display_name.as_str())));
                let name_key = self.str_const("name");
                self.emit_u16(Op::STRUCT_SET, name_key);
                let line = self.line;
                crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);
            }
            if let Some(fixed_count) = rest_fixed_count {
                inst!(self, core_wasm::dup);
                self.emit_const(Value::F64(fixed_count as f64));
                self.emit_u16(Op::STRUCT_SET, rest_key);
                self.emit(Op::DROP);
            }
            let method_key = self.str_const(&bind_name);
            self.emit_u16(Op::STRUCT_SET, method_key);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Function declaration compilation
    // ════════════════════════════════════════════════════════════════════════

    /// Detect the JS walker's `wrap_generator` lowering: a plain outer
    /// function whose body binds `__gen_fn` to a generator function
    /// expression. Returns the SOURCE function's `(is_async, is_generator)`
    /// so prototype stamping reflects the original kind (§27.3/§27.4).
    pub(super) fn wrapped_generator_kind(body: &[Statement]) -> Option<(bool, bool)> {
        for stmt in body {
            if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
                for d in declarations {
                    if matches!(&d.pattern, crate::ast::BindingPattern::Ident(n) if n == "__gen_fn")
                    {
                        if let Some(init) = &d.init {
                            if let crate::ast::ExprKind::FunctionExpr(inner) = &init.kind {
                                if let StmtKind::FunctionDecl {
                                    is_async,
                                    is_generator,
                                    ..
                                } = &inner.kind
                                {
                                    if *is_generator {
                                        return Some((*is_async, true));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        None
    }

    pub(super) fn compile_function_decl(
        &mut self,
        name: &str,
        params: &[Param],
        return_type: &Option<String>,
        body: &[Statement],
        _is_sub: bool,
        is_generator: bool,
        handles: &[String],
        is_async: bool,
    ) -> Result<(), String> {
        let cname = self.canon(name);
        self.defined_globals.insert(cname.clone());
        self.defined_functions.insert(cname.clone());
        // Register top-level generator functions so `is_direct_generator_call`
        // detects them (`[...gen()]` spread, `foreach (gen() as ...)`). Scoped
        // to buffered-iterator languages (PHP): JS keeps its runtime
        // `isGenerator` dispatch, so registering here would change its routing.
        if is_generator && self.profile.buffered_iterator_methods {
            self.generator_functions.insert(cname.clone());
        }
        self.function_param_modes.insert(
            cname.clone(),
            params.iter().map(|param| param.pass_by).collect(),
        );
        self.function_param_types.insert(
            cname.clone(),
            params.iter().map(|param| param.type_hint.clone()).collect(),
        );
        self.function_min_arity.insert(
            cname.clone(),
            params
                .iter()
                .take_while(|param| param.default.is_none() && !param.is_rest)
                .count(),
        );
        self.function_signatures
            .entry(cname.clone())
            .or_default()
            .push(CallSignature::from_params(params));
        if let Some(return_type) = return_type.as_ref() {
            self.function_return_types
                .insert(cname.clone(), return_type.clone());
        }
        let name = &cname;

        let uses_js_arguments = self.profile.has_arguments_object
            && !is_generator
            && (params
                .iter()
                .any(|param| param.default.as_ref().is_some_and(expr_uses_js_arguments))
                || body.iter().any(stmt_uses_js_arguments));
        let has_rest = params.last().map_or(false, |p| p.is_rest);
        let lowered_has_rest = has_rest || uses_js_arguments;
        let generator_control_arity = usize::from(is_generator && !lowered_has_rest);
        let arity: u8 = if uses_js_arguments {
            (1 + generator_control_arity) as u8
        } else {
            (params.len() + generator_control_arity) as u8
        };
        if uses_js_arguments {
            self.rest_fixed_arities.insert(0);
        } else if has_rest {
            self.rest_fixed_arities
                .insert(params.len().saturating_sub(1) as u8);
        }
        let func_idx = self.chunks.len();
        let mut chunk = common::functions::create_function_chunk(name, arity);
        self.seed_shared_global_constants(&mut chunk);
        // `is_async` carries SOURCE truth (async fns AND async generators).
        // Consumers refine: the JSPI custom-section writer and the VM's
        // call_async gate both exclude generators (async generators are
        // continuations at call time — their async surface is `.next()`
        // returning a promise, which the protocol attach selects on this
        // flag).
        chunk.is_async = is_async;
        // Generators: when the source marked the function as a
        // generator (Python `yield`, JS `function*`, C# `yield return`),
        // stamp the chunk so the VM wraps invocations in a
        // `Continuation` instead of executing the body inline. The
        // body itself was compiled with `SUSPEND` at each yield site.
        chunk.is_generator = is_generator;
        if is_generator && !is_async {
            self.generator_functions.insert(cname.clone());
            // Track the number of user-visible params so call sites can pad
            // missing optional args with `undefined`.  Without this, a call
            // like `range(1, 6)` to `function* range(start, end, step=1)`
            // passes bound_args=[1,6]; GEN_NEXT then calls the body with
            // [1, 6, null] (argc=3) which lands `null` in the `step` slot,
            // preventing the default from applying and causing an infinite loop.
            // Only register fixed-arity generators (no rest, no arguments).
            if !lowered_has_rest {
                self.generator_param_counts
                    .insert(cname.clone(), params.len());
            }
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
        self.static_local_bindings.push(HashMap::new());
        if self.profile.name == "php" {
            self.php_function_globals.push(HashSet::new());
        }
        let saved = self.current;
        self.current = func_idx;
        // Runtime TRY_END counts are per-FRAME: a nested chunk must not
        // inherit the enclosing async body's try depth, or its returns pop
        // the caller's handlers off the shared runtime handler stack.
        let saved_async_try_depth = std::mem::take(&mut self.active_async_try_depth);
        // Function body opens fresh wrt the runtime label_stack —
        // emit_return drains back to this base. Save+restore so nested
        // function decls compose.
        let saved_label_base = self.function_label_base;
        self.function_label_base = self.label_depth;
        let saved_fn = self.current_func_name.take();
        self.current_func_name = Some(name.to_string());
        // Pre-scan: collect locals/params whose address is taken (`&v`) so they
        // can be promoted to a pointer cell once at declaration / entry, rather
        // than lazily at the `&v` site (which re-wraps every loop iteration).
        // Scoped to C — the only profile exercising this path — so the other
        // AddrOf-using languages (Pascal/Go/C#) keep their current behavior.
        let saved_addr_taken = std::mem::take(&mut self.current_addr_taken_locals);
        if self.profile.name == "c" {
            crate::primitives::collect_addr_taken_idents(body, &mut self.current_addr_taken_locals);
        }
        let saved_closure_captured = std::mem::take(&mut self.current_closure_captured_locals);
        let saved_env_names = std::mem::take(&mut self.closure_env_names);
        let saved_capture_locals = std::mem::take(&mut self.capture_locals);
        // Capture parent shared env for nested function upvalue resolution
        let parent_shared_env_slot = self.shared_env_slot;
        let parent_shared_env_names = self.shared_env_names.clone();
        let saved_shared_env_slot = self.shared_env_slot.take();
        let saved_shared_env_names = std::mem::take(&mut self.shared_env_names);
        crate::primitives::collect_closure_captured_idents(
            body,
            &mut self.current_closure_captured_locals,
        );
        // If parent has a shared env, pre-seed closure_env_names so
        // upvalue indices match the parent's shared env layout.
        if !parent_shared_env_names.is_empty() {
            self.closure_env_names = parent_shared_env_names.clone();
        }
        // ECMA-262 §11.2.2: inherit strict mode and additionally enable it on
        // a `"use strict"` directive prologue in this function's body.
        let saved_strict = self.in_strict;
        if Self::stmts_have_use_strict_directive(body) {
            self.in_strict = true;
        }
        self.js_arguments_bindings.push(None);

        let js_arguments_source_slot = if uses_js_arguments {
            Some(self.define_local("__vybe_js_arguments_array"))
        } else {
            None
        };
        let js_arguments_slot = if uses_js_arguments {
            let slot = self.define_local("arguments");
            self.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
            self.emit_u16(Op::LOCAL_SET, slot);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_var_get(name);
            let callee_key = self.str_const("callee");
            self.emit_u16(Op::STRUCT_SET, callee_key);
            self.emit(Op::DROP);
            // §10.4.4.6: arguments objects report "[object Arguments]" —
            // stamp the tag the host's object_to_string_tag reads.
            self.emit_u16(Op::LOCAL_GET, slot);
            self.chunk().emit_string_const("Arguments", 0);
            let type_key = self.str_const("__type");
            self.emit_u16(Op::STRUCT_SET, type_key);
            self.emit(Op::DROP);
            Some(slot)
        } else {
            None
        };

        let mut aliased_params = HashMap::new();
        let mut aliased_indices = HashMap::new();
        // ECMA-262 §10.4.4: only a *non-strict* function with a simple
        // parameter list gets a mapped `arguments` object whose elements
        // alias the named parameters. Strict functions get an unmapped
        // (independent) copy, so `arguments[0] = …` must NOT change the
        // parameter (and vice versa).
        let simple_arguments_alias = uses_js_arguments
            && !self.in_strict
            && params
                .iter()
                .all(|param| param.default.is_none() && !param.is_rest);

        for (index, p) in params.iter().enumerate() {
            self.define_local_typed(&p.name, p.type_hint.clone());
            let normalized_type_hint = p.type_hint.as_deref().map(Compiler::normalize_type_hint);
            if normalized_type_hint
                .as_deref()
                .is_some_and(|type_hint| type_hint.ends_with("()"))
                || normalized_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.ends_with("()"))
            {
                self.record_array_binding(
                    &p.name,
                    ArrayBindingMetadata {
                        is_fixed: false,
                        type_hint: p.type_hint.clone(),
                        pascal_bounds: p
                            .type_hint
                            .as_deref()
                            .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint)),
                    },
                );
            }
            if simple_arguments_alias {
                let slot = self.scope().resolve(&p.name).unwrap();
                aliased_params.insert(p.name.clone(), (slot, index));
                aliased_indices.insert(index, slot);
            }
        }

        let js_arguments_len_slot = if uses_js_arguments {
            let len_slot = self.define_local("__vybe_js_arguments_length");
            self.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
            common::collections::emit_len(&mut self.chunks, self.current, self.line);
            self.emit_u16(Op::LOCAL_SET, len_slot);
            Some(len_slot)
        } else {
            None
        };

        if let Some(slot) = js_arguments_slot {
            *self.js_arguments_bindings.last_mut().unwrap() = Some(JsArgumentsBinding {
                args_slot: slot,
                aliased_params,
                aliased_indices,
            });
        }

        if uses_js_arguments {
            for (index, p) in params.iter().enumerate() {
                let slot = self.scope().resolve(&p.name).unwrap();
                if p.is_rest {
                    self.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
                    self.emit_const(Value::F64(index as f64));
                    self.emit_u16(Op::LOCAL_GET, js_arguments_len_slot.unwrap());
                    common::collections::emit_slice(&mut self.chunks, self.current, self.line);
                    self.emit_u16(Op::LOCAL_SET, slot);
                } else {
                    self.emit_array_value_or_undefined(
                        js_arguments_source_slot.unwrap(),
                        js_arguments_len_slot.unwrap(),
                        index,
                    );
                    self.emit_u16(Op::LOCAL_SET, slot);
                }
            }
        }

        for (param_index, p) in params.iter().enumerate() {
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
                if self.profile.missing_arg_is_undefined {
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                } else {
                    self.emit(Op::REF_IS_NULL);
                }
                let branch_line = self.line;
                self.chunks[self.current].emit_if(branch_line);
                if self.profile.default_args_evaluated_once {
                    // Evaluate-once (Python/Ruby): the default is computed on the
                    // first call that omits the arg and cached in a per-parameter
                    // module global, so every later call reuses the SAME object
                    // (`def f(a=[]); f() is f()` is True). A separate init-flag
                    // global gates the one-time evaluation (correct even when the
                    // default value itself is null). The cache key is unique per
                    // (function chunk, parameter index).
                    let cache = format!("__vybe_dflt_{}_{}", self.current, param_index);
                    let inited = format!("__vybe_dflt_init_{}_{}", self.current, param_index);
                    let cache_idx = self.global_name_const_idx(&cache);
                    let inited_idx = self.global_name_const_idx(&inited);
                    self.emit_u16(Op::GLOBAL_GET, inited_idx);
                    self.emit(Op::REF_IS_NULL); // uninitialised global reads null
                    let init_line = self.line;
                    self.chunks[self.current].emit_if(init_line);
                    self.compile_expr(default)?;
                    self.emit_u16(Op::GLOBAL_SET, cache_idx);
                    self.emit_const(Value::Bool(true));
                    self.emit_u16(Op::GLOBAL_SET, inited_idx);
                    self.chunks[self.current].emit_end(init_line);
                    self.emit_u16(Op::GLOBAL_GET, cache_idx);
                    self.emit_u16(Op::LOCAL_SET, slot);
                } else {
                    self.compile_expr(default)?;
                    self.emit_u16(Op::LOCAL_SET, slot);
                }
                self.chunks[self.current].emit_end(branch_line);
            }
            if p.is_rest && self.profile.name == "lua" {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.stamp_lua_multi_row_slot(slot);
            }
            self.maybe_initialize_fortran_out_param(p);
        }

        // Promote address-taken params to a pointer cell once, at entry. A later
        // `&param` (e.g. inside a loop) then reuses this cell instead of
        // re-wrapping it each iteration. Reads/writes of the param are already
        // cell-aware once the binding is marked.
        if !self.current_addr_taken_locals.is_empty() {
            for p in params {
                if self.current_addr_taken_locals.contains(&p.name) {
                    self.promote_local_binding_to_pointer_cell(&p.name);
                }
            }
        }

        let generator_control_slot =
            is_generator.then(|| self.define_local("__generator_entry_control"));

        if let Some(control_slot) = generator_control_slot {
            self.emit_generator_entry_control(control_slot)?;
        }

        // Result slot for functions with return type (Pascal/VB Function).
        // The slot name is profile-driven so VB can keep it internal
        // (`__result__`) and avoid shadowing user classes named `Result`,
        // while Pascal keeps it as `Result` (user-visible per Pascal idiom).
        let result_slot =
            if return_type.is_some() && self.profile.function_return == ReturnStyle::ResultSlot {
                let slot_name = self.profile.result_slot_name.clone();
                let rs = self.define_local_typed(&slot_name, return_type.clone());
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, rs);
                Some(rs)
            } else {
                None
            };
        let ref_out_slots: Vec<u16> = params
            .iter()
            .filter(|param| matches!(param.pass_by, PassBy::Ref | PassBy::Out))
            .filter_map(|param| self.scope().resolve(&param.name))
            .collect();

        let saved_rs = self.current_result_slot.take();
        let saved_ref_out = self.current_ref_out_params.take();
        self.current_result_slot = result_slot;
        self.current_ref_out_params = (!ref_out_slots.is_empty()).then_some(ref_out_slots);

        // ECMA-262 async function semantics: throws inside the body
        // become rejected Promises, normal returns become fulfilled
        // Promises (§27.7.5.3). Wrap the body in TRY_START/TRY_END so
        // uncaught exceptions short-circuit to the Promise.reject
        // path. The `await` opcode handles per-await suspension via
        // JSPI; this wrap just covers terminal throw / return.
        let async_try = if is_async && !is_generator && self.profile.async_wraps_body_in_try {
            let line = self.line;
            Some(common::functions::emit_async_body_start(
                &mut self.chunks[self.current],
                line,
            ))
        } else {
            None
        };
        if async_try.is_some() {
            self.active_async_try_depth += 1;
        }

        if self.profile.ambient_this_binding
            && crate::primitives::closures_in_body_reference_this(body)
        {
            let this_idx = self.str_const("__js_this");
            self.emit_u16(Op::GLOBAL_GET, this_idx);
            let this_local = self.define_local("__js_this");
            self.emit_u16(Op::LOCAL_SET, this_local);
            self.current_closure_captured_locals
                .insert("__js_this".to_string());
        }

        if !self.current_closure_captured_locals.is_empty() {
            let mut fn_scope_names: HashSet<String> =
                params.iter().map(|p| p.name.clone()).collect();
            crate::primitives::collect_declared_names(body, &mut fn_scope_names);
            let mut captured_names: Vec<String> = self
                .current_closure_captured_locals
                .iter()
                .filter(|name| {
                    fn_scope_names.contains(name.as_str())
                        || parent_shared_env_names.iter().any(|n| n == name.as_str())
                })
                .cloned()
                .collect();
            captured_names.sort();
            if !captured_names.is_empty() {
                let env_size = captured_names.len() as u16;
                let line = self.line;
                for _ in 0..env_size {
                    self.emit(Op::NULL);
                }
                self.chunks[self.current].emit_array_new_fixed(0, env_size, line);
                let env_slot = self.define_local("__shared_env");
                self.emit_u16(Op::LOCAL_SET, env_slot);
                self.shared_env_slot = Some(env_slot);
                self.shared_env_names = captured_names.clone();
                let mut local_decls: HashSet<String> =
                    params.iter().map(|p| p.name.clone()).collect();
                crate::primitives::collect_declared_names(body, &mut local_decls);
                for (idx, cap_name) in captured_names.iter().enumerate() {
                    if let Some(param_slot) = self.scope().resolve(cap_name) {
                        self.emit_u16(Op::LOCAL_GET, param_slot);
                        crate::primitives::closures::emit_env_set(
                            self.chunk(),
                            env_slot,
                            idx as u16,
                            line,
                        );
                    } else if !local_decls.contains(cap_name) && parent_shared_env_slot.is_some() {
                        if let Some(parent_idx) =
                            parent_shared_env_names.iter().position(|n| n == cap_name)
                        {
                            let closure_env = self.closure_env_slot();
                            crate::primitives::closures::emit_env_get(
                                self.chunk(),
                                closure_env,
                                parent_idx as u16,
                                line,
                            );
                            crate::primitives::closures::emit_env_set(
                                self.chunk(),
                                env_slot,
                                idx as u16,
                                line,
                            );
                        }
                    }
                }
            }
        }

        if self.profile.name == "fortran" {
            for statement in body {
                if matches!(&statement.kind, StmtKind::VarDecl { .. }) {
                    self.compile_stmt(statement)?;
                }
            }
            for statement in body {
                if matches!(&statement.kind, StmtKind::FunctionDecl { .. }) {
                    self.compile_stmt(statement)?;
                }
            }
            for statement in body {
                if matches!(
                    &statement.kind,
                    StmtKind::VarDecl { .. } | StmtKind::FunctionDecl { .. }
                ) {
                    continue;
                }
                self.compile_stmt(statement)?;
            }
        } else {
            for statement in body {
                self.compile_stmt(statement)?;
            }
        }

        if async_try.is_some() {
            self.active_async_try_depth = self.active_async_try_depth.saturating_sub(1);
        }

        if let Some(catch_jump) = async_try {
            let line = self.line;
            // Normal exit: wrap return in Promise.resolve(value).
            // The body's compile_stmt may have already emitted RETURNs
            // (early returns); we still need the fall-through path
            // to leave a fulfilled Promise on the stack.
            let chunk = &mut self.chunks[self.current];
            common::functions::emit_async_body_fallthrough(chunk, catch_jump, line);
            let resolve_idx = self.import("ecma:promise", "resolve");
            self.emit_host_call(resolve_idx, 1);
            self.emit_return();
            // Catch handler — exception value on TOS.
            let chunk = &mut self.chunks[self.current];
            common::functions::patch_async_body_catch(chunk, catch_jump);
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
        self.current_addr_taken_locals = saved_addr_taken;
        self.current_closure_captured_locals = saved_closure_captured;
        self.closure_env_names = saved_env_names;
        self.capture_locals = saved_capture_locals;
        self.shared_env_slot = saved_shared_env_slot;
        self.shared_env_names = saved_shared_env_names;
        self.in_strict = saved_strict;
        self.current_result_slot = saved_rs;
        self.current_ref_out_params = saved_ref_out;

        let ns = self.scope().next_slot;
        self.chunks[func_idx].finalize_local_count(ns);
        self.chunks[func_idx].local_names = self.scope().defined_names.clone();
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        let inner_scope_idx = self.scopes.len() - 1;
        let uv_names: Vec<Option<String>> = (0..uvs.len())
            .map(|i| self.captured_name_for_upvalue(inner_scope_idx, i as u8))
            .collect();
        self.js_arguments_bindings.pop();
        self.scopes.pop();
        self.static_local_bindings.pop();
        if self.profile.name == "php" {
            self.php_function_globals.pop();
        }
        self.current = saved;
        self.active_async_try_depth = saved_async_try_depth;
        self.function_label_base = saved_label_base;

        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, 0, line);
        } else if let Some(shared_slot) = parent_shared_env_slot {
            // Parent has a shared env — pass it directly as the upvalue.
            common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                shared_slot,
                line,
            );
        } else {
            let mut env_slots: Vec<u16> = Vec::new();
            for (i, uv) in uvs.iter().enumerate() {
                if let Some(name) = uv_names[i].clone() {
                    let slot = if uv.is_local {
                        uv.index as u16
                    } else {
                        let parent_env = self.closure_env_slot();
                        let parent_idx = self.closure_env_index(&name);
                        let tmp = self.define_local(&format!("__nested_cap_{}", name));
                        crate::primitives::closures::emit_env_get(
                            self.chunk(),
                            parent_env,
                            parent_idx,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        tmp
                    };
                    env_slots.push(slot);
                }
            }
            crate::primitives::closures::emit_env_new(self.chunk(), &env_slots, line);
            let env_slot = self.define_local(&format!("__closure_env_{}", func_idx));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                env_slot,
                line,
            );
        }
        if uses_js_arguments {
            self.emit_stamp_rest_metadata_on_stack(0);
        } else if has_rest {
            self.emit_stamp_rest_metadata_on_stack(params.len().saturating_sub(1));
        }
        let idx = self.str_const(name);
        self.emit_u16(Op::GLOBAL_SET, idx);
        if let Some(callable_global) = self.source_function_callable_global_name(name) {
            let fn_idx = self.str_const(name);
            self.emit_u16(Op::GLOBAL_GET, fn_idx);
            let callable_idx = self.str_const(&callable_global);
            self.emit_u16(Op::GLOBAL_SET, callable_idx);
        }

        if self.profile.has_function_prototype_bind {
            let line = self.line;
            self.emit_common("object.new", 0, line);
            let proto_slot = self.define_local("__js_fn_proto");
            self.emit_u16(Op::LOCAL_SET, proto_slot);

            self.emit_var_get(name);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);

            // ECMA-262 §10.2.4 `length`: number of params before the
            // first one with a default value or rest. Skip rest entirely.
            let length = params
                .iter()
                .take_while(|p| p.default.is_none() && !p.is_rest)
                .count();
            self.emit_var_get(name);
            self.emit_const(Value::F64(length as f64));
            let length_key = self.str_const("length");
            self.emit_u16(Op::STRUCT_SET, length_key);
            self.emit(Op::DROP);

            // The JS walker's wrap_generator lowers `function*` /
            // `async function*` to a PLAIN outer function holding
            // `const __gen_fn = function*(){...}` — recover the source
            // kind from that contract so the §27.3/§27.4 intrinsic
            // stamp survives the lowering.
            let (eff_async, eff_generator) =
                Self::wrapped_generator_kind(body).unwrap_or((is_async, is_generator));
            self.emit_var_get(name);
            {
                let line = self.line;
                crate::primitives::prototypes::emit_stamp_function_kind_proto(
                    self.chunk(),
                    eff_async,
                    eff_generator,
                    line,
                );
            }

            // §10.2.9/§10.2.10: name/length are non-enumerable.
            self.emit_var_get(name);
            {
                let line = self.line;
                crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);
            }

            // §27.3 / §27.7: generator and async function declarations
            // have no [[Construct]] — `new` on them must TypeError (the
            // host construct path checks this marker).
            if eff_async || eff_generator {
                self.emit_var_get(name);
                self.emit_const(Value::Bool(true));
                let non_ctor_key = self.str_const("__vybe_non_ctor");
                self.emit_u16(Op::STRUCT_SET, non_ctor_key);
                self.emit(Op::DROP);
            }

            // §27.7 / §27.3 (node-verified): async (non-generator)
            // functions have NO own `prototype` property; generator
            // functions have one WITHOUT a `constructor` property; plain
            // functions get the classic prototype/constructor pair.
            if !eff_async || eff_generator {
                if !eff_generator {
                    self.emit_u16(Op::LOCAL_GET, proto_slot);
                    self.emit_var_get(name);
                    let ctor_key = self.str_const("constructor");
                    self.emit_u16(Op::STRUCT_SET, ctor_key);
                    self.emit(Op::DROP);
                }

                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_GET, proto_slot);
                let proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_SET, proto_key);
                self.emit(Op::DROP);
            }
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
                    self.current_class
                        .clone()
                        .map(|c| self.canon(&c))
                        .unwrap_or(ctrl_canon)
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

    /// Recursively register a minimal member surface for every nested
    /// class/struct in `members`, keyed by its (already-qualified) name, so
    /// a reference to it from a sibling method compiled earlier resolves as
    /// a user type instead of falling through to a builtin value-method. The
    /// full registration replaces this when the nested type is compiled.
    fn predeclare_nested_type_surfaces(&mut self, members: &[ClassMember], enclosing: &str) {
        for m in members {
            let ClassMember::NestedType(stmt) = m else {
                continue;
            };
            let (nested_name, nested_members, nested_parent): (
                &str,
                &[ClassMember],
                Option<String>,
            ) = match &stmt.kind {
                StmtKind::ClassDecl {
                    name: nn,
                    members: nm,
                    parents,
                    ..
                } => (nn, nm, parents.first().map(|p| self.canon(p))),
                StmtKind::StructDecl {
                    name: nn,
                    members: nm,
                    ..
                } => (nn, nm, None),
                _ => continue,
            };
            let qualified_nested_name = Self::qualified_nested_type_name(enclosing, nested_name);
            let nested_canon = self.canon(&qualified_nested_name);
            let nested_leaf_canon =
                self.canon(nested_name.rsplit('.').next().unwrap_or(nested_name));
            if !self.pending_classes.contains_key(&nested_canon) {
                let mut static_method_names: Vec<String> = Vec::new();
                let mut static_fields: Vec<String> = Vec::new();
                let mut instance_member_names: Vec<String> = Vec::new();
                let mut fields: Vec<String> = Vec::new();
                let mut field_storage_names: HashMap<String, String> = HashMap::new();
                let field_storage_slot_name =
                    |compiler: &Self, owner_class: &str, field_name: &str| {
                        let field_canon = compiler.canon(field_name);
                        if compiler.profile.field_hiding
                            && compiler.field_hides_ancestor(nested_parent.as_deref(), &field_canon)
                        {
                            format!("__hide_{}${}", compiler.canon(owner_class), field_canon)
                        } else {
                            compiler.js_member_storage_name_for_class(owner_class, field_name)
                        }
                    };
                // (method-return-type key, return type) — registered so a
                // chained call `outer.first().next()` compiled in a sibling
                // method (before this nested type is compiled) can infer the
                // intermediate result's type.
                let mut return_types: Vec<(String, String)> = Vec::new();
                for mem in nested_members {
                    match mem {
                        ClassMember::Method(ms) => {
                            if let StmtKind::FunctionDecl {
                                name: mname,
                                modifiers,
                                return_type,
                                ..
                            } = &ms.kind
                            {
                                if modifiers.is_abstract {
                                    continue;
                                }
                                if modifiers.is_static || modifiers.is_shared {
                                    static_method_names.push(self.canon(mname));
                                } else {
                                    instance_member_names.push(self.canon(mname));
                                }
                                if let Some(rt) = return_type {
                                    return_types.push((
                                        self.canon(&format!("{nested_canon}.{mname}")),
                                        rt.clone(),
                                    ));
                                }
                            }
                        }
                        ClassMember::Property { name: pname, .. } => {
                            instance_member_names.push(self.canon(pname));
                        }
                        ClassMember::Field {
                            name: fname,
                            modifiers,
                            ..
                        } => {
                            if modifiers.is_static || modifiers.is_shared {
                                static_fields.push(self.canon(fname));
                            } else {
                                let field_canon = self.canon(fname);
                                let storage_name =
                                    field_storage_slot_name(self, &nested_canon, fname);
                                if storage_name != field_canon {
                                    field_storage_names
                                        .insert(field_canon.clone(), storage_name.clone());
                                }
                                fields.push(storage_name);
                            }
                        }
                        _ => {}
                    }
                }
                self.defined_globals.insert(nested_canon.clone());
                self.defined_classes.insert(nested_canon.clone());
                self.defined_globals.insert(nested_leaf_canon.clone());
                self.defined_classes.insert(nested_leaf_canon.clone());
                self.note_pending_class(&nested_canon, nested_parent);
                if let Some(pc) = self.pending_classes.get_mut(&nested_canon) {
                    pc.enclosing_class = Some(enclosing.to_string());
                    pc.static_method_names = static_method_names;
                    pc.static_fields = static_fields;
                    pc.instance_member_names = instance_member_names;
                    pc.fields = fields;
                    pc.field_storage_names = field_storage_names;
                }
                for (key, rt) in return_types {
                    self.function_return_types.entry(key).or_insert(rt);
                }
            }
            // Recurse: register deeper nested types (`Outer.Inner.Deep`).
            self.predeclare_nested_type_surfaces(nested_members, &nested_canon);
        }
    }

    pub(crate) fn compile_class(
        &mut self,
        class: &crate::primitives::class_normalize::NormalClass,
    ) -> Result<(), String> {
        // Extract the canonicalised names the orchestration below needs.
        // Canonicalisation happens once here rather than at every caller.
        let cname = self.canon(&class.name);
        let name: &str = &cname;
        let parent_canonical = class.parent.as_ref().map(|p| self.canon(p));
        let parent: &Option<String> = &parent_canonical;

        // The ENCLOSING frame's shared env, captured before any method compile
        // clears it. A class declared inside a function closes over that frame,
        // and a captured local there does not live in a slot — it lives in the
        // env array, which every closure receives as upvalue[0]
        // (`bindings.rs::capture_local_slot`). So a method reading such a local
        // emits `env[idx]`, and its binding must therefore capture the ENV, not
        // the individual name. Empty for a top-level class, which is the gate:
        // no env, nothing to forward.
        let enclosing_shared_env_names = self.shared_env_names.clone();

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
        // When properties and methods share the object namespace only in
        // some languages, a property whose name collides with a method needs
        // a distinct slot (see `separate_property_method_namespace`).
        let colliding_method_names: std::collections::HashSet<String> =
            if self.profile.separate_property_method_namespace {
                class
                    .instance_methods
                    .iter()
                    .map(|method| self.canon(&method.source_name))
                    .collect()
            } else {
                std::collections::HashSet::new()
            };
        let mut field_storage_names: HashMap<String, String> = HashMap::new();
        let field_storage_slot_name = |compiler: &Self,
                                       field_name: &str,
                                       method_names: &std::collections::HashSet<String>|
         -> String {
            let canon = compiler.canon(field_name);
            if compiler.profile.separate_property_method_namespace && method_names.contains(&canon)
            {
                format!("__prop${}", canon)
            } else {
                compiler.js_member_storage_name_for_class(&class.name, field_name)
            }
        };

        let mut fields: Vec<String> = Vec::new();
        let mut field_inits: Vec<(
            String,
            Option<String>,
            Option<Expression>,
            Option<Vec<Expression>>,
        )> = Vec::new();
        let mut static_field_inits: Vec<(
            String,
            Option<String>,
            Option<Expression>,
            Option<Vec<Expression>>,
        )> = Vec::new();
        for f in &class.instance_fields {
            let field_canon = self.canon(&f.name);
            // Field hiding (java/C#/VB): a field that shadows an ancestor's
            // gets a declaring-class-qualified slot so both survive on the
            // object and access resolves by the reference's declared type.
            let fname = if self.profile.field_hiding
                && self.field_hides_ancestor(class.parent.as_deref(), &field_canon)
            {
                format!("__hide_{}${}", self.canon(&class.name), field_canon)
            } else {
                field_storage_slot_name(self, &f.name, &colliding_method_names)
            };
            if fname != field_canon {
                field_storage_names.insert(field_canon.clone(), fname.clone());
            }
            fields.push(fname.clone());
            field_inits.push((
                fname,
                f.type_hint.clone(),
                f.init.clone(),
                f.array_bounds.clone(),
            ));
        }
        for f in &class.static_fields {
            let fname = self.js_member_storage_name_for_class(&class.name, &f.name);
            static_field_inits.push((
                fname,
                f.type_hint.clone(),
                f.init.clone(),
                f.array_bounds.clone(),
            ));
        }
        for p in &class.properties {
            let prop_canon = self.canon(&p.source_name);
            let prop_is_override = p
                .getter
                .as_ref()
                .is_some_and(|getter| getter.is_override || getter.raw_modifiers.is_override)
                || p.setter
                    .as_ref()
                    .is_some_and(|setter| setter.is_override || setter.raw_modifiers.is_override);
            let property_storage_name = if self.profile.field_hiding
                && !p.is_static
                && !prop_is_override
                && self.field_hides_ancestor(class.parent.as_deref(), &prop_canon)
            {
                format!("__hide_{}${}", self.canon(&class.name), prop_canon)
            } else {
                prop_canon.clone()
            };
            if property_storage_name != prop_canon {
                field_storage_names.insert(prop_canon.clone(), property_storage_name.clone());
            }

            // Auto-properties get a backing field named like the property;
            // the runtime reads/writes through auto-emitted __get_/__set_
            // chunks bound later.
            if let Some(auto_field_name) = &p.auto_field {
                let pname_canon = if property_storage_name != prop_canon {
                    property_storage_name.clone()
                } else {
                    field_storage_slot_name(self, auto_field_name, &colliding_method_names)
                };
                if pname_canon != self.canon(auto_field_name) {
                    field_storage_names.insert(self.canon(auto_field_name), pname_canon.clone());
                }
                if p.is_static {
                    if !static_field_inits
                        .iter()
                        .any(|(n, _, _, _)| n == &pname_canon)
                    {
                        static_field_inits.push((pname_canon, None, None, None));
                    }
                } else if !fields.contains(&pname_canon) {
                    fields.push(pname_canon.clone());
                    field_inits.push((pname_canon, None, None, None));
                }
            } else if !p.is_static && !fields.contains(&property_storage_name) {
                fields.push(property_storage_name.clone());
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
                    field_inits.push((fname, None, None, None));
                }
            }
        }

        // Store field list for implicit self resolution
        let static_field_names: Vec<String> = static_field_inits
            .iter()
            .map(|(n, _, _, _)| n.clone())
            .collect();
        let mut instance_field_types: HashMap<String, String> = class
            .instance_fields
            .iter()
            .filter_map(|f| {
                f.type_hint.as_ref().map(|t| {
                    (
                        field_storage_names
                            .get(&self.canon(&f.name))
                            .cloned()
                            .unwrap_or_else(|| {
                                self.js_member_storage_name_for_class(&class.name, &f.name)
                            }),
                        Self::normalize_type_hint(t),
                    )
                })
            })
            .collect();
        for member in &class.raw_extra_members {
            match member {
                ClassMember::Event {
                    name,
                    type_hint: Some(type_hint),
                    ..
                } => {
                    instance_field_types
                        .entry(self.canon(name))
                        .or_insert_with(|| Self::normalize_type_hint(type_hint));
                }
                ClassMember::Property {
                    name,
                    type_hint: Some(type_hint),
                    modifiers,
                    ..
                } if !modifiers.is_static => {
                    instance_field_types
                        .entry(self.canon(name))
                        .or_insert_with(|| Self::normalize_type_hint(type_hint));
                }
                _ => {}
            }
        }
        let mut static_member_names = static_field_names;
        let mut static_const_names: Vec<String> = Vec::new();
        for member in &class.raw_extra_members {
            if let ClassMember::Const { name, .. } = member {
                let const_name = self.canon(name);
                static_member_names.push(const_name.clone());
                static_const_names.push(const_name);
            }
        }

        self.pending_classes.insert(
            name.to_string(),
            PendingClass {
                parent: parent.clone(),
                bases: class.bases.clone(),
                enclosing_class: self.current_class.clone(),
                fields: fields.clone(),
                field_storage_names: field_storage_names.clone(),
                is_value_type: class.is_value_type,
                instance_member_names: class
                    .instance_methods
                    .iter()
                    .map(|method| {
                        self.js_member_storage_name_for_class(&class.name, &method.source_name)
                    })
                    .collect(),
                instance_pointer_method_names: class
                    .instance_methods
                    .iter()
                    .filter(|method| {
                        method
                            .params
                            .first()
                            .and_then(|param| param.type_hint.as_deref())
                            .is_some_and(|type_hint| type_hint.trim_start().starts_with('*'))
                    })
                    .map(|method| method.canonical_name.clone())
                    .collect(),
                instance_field_types,
                static_fields: static_member_names,
                static_field_types: class
                    .static_fields
                    .iter()
                    .filter_map(|f| {
                        f.type_hint
                            .as_ref()
                            .map(|t| (self.canon(&f.name), Self::normalize_type_hint(t)))
                    })
                    .collect(),
                static_method_names: class
                    .static_methods
                    .iter()
                    .map(|m| self.js_member_storage_name_for_class(&class.name, &m.source_name))
                    .collect(),
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: class
                    .raw_extra_members
                    .iter()
                    .filter_map(|m| {
                        if let ClassMember::NestedType(stmt) = m {
                            match &stmt.kind {
                                StmtKind::ClassDecl { name, .. }
                                | StmtKind::StructDecl { name, .. }
                                | StmtKind::InterfaceDecl { name, .. }
                                | StmtKind::EnumDecl { name, .. } => {
                                    Some(Self::qualified_nested_type_name(&class.name, name))
                                }
                                _ => None,
                            }
                        } else {
                            None
                        }
                    })
                    .collect(),
                statics: Vec::new(), // filled after methods are compiled
            },
        );

        // Predeclare nested class/struct member surfaces before compiling
        // this class's methods. The real nested-type compilation happens
        // later (with the rest of `raw_extra_members`, after methods, so
        // chunk indices stay byte-identical), which means a call to a nested
        // class's method from a sibling method — e.g. `Inner.add(...)` or
        // `innerInstance.get()` inside `main` — would otherwise not see
        // `Inner` in `pending_classes` yet and fall through to a builtin
        // value-method of the same name (`get`, `add`, …). Top-level classes
        // avoid this via `predeclare_type_names`; nested classes (java/C#/VB)
        // had no equivalent. Recurses the whole nested tree so a deeply
        // qualified reference (`Outer.Inner`) is registered too. Each
        // placeholder is replaced by the full registration when the nested
        // type is actually compiled below.
        self.predeclare_nested_type_surfaces(&class.raw_extra_members, name);

        // Compile methods (including constructor body)
        // (name, chunk_idx, is_ctor, is_static)
        let mut method_chunks: Vec<(String, usize, bool, bool)> = Vec::new();
        let mut method_capture_name_map: HashMap<usize, Vec<String>> = HashMap::new();
        // Which PROTOCOL SLOT each of this class's methods fills, by the
        // method's canonical name. `special_methods` is what every normalizer
        // already produces and nothing has ever read (§2g); this is its first
        // consumer. Resolved to bound names in the method loop below, where the
        // storage name is computed.
        let mut class_slots: HashMap<&str, vybe_ast::ProtocolSlot> = class
            .special_methods
            .iter()
            .map(|s| (s.canonical_name.as_str(), s.kind))
            .collect();
        // The destructor is held in its own field, so a normalizer that routes
        // it there never adds it to `special_methods`. It fills the slot by
        // CONSTRUCTION — that is what the field means — so state it here once
        // instead of asking twelve normalizers to remember.
        if let Some(destructor) = &class.destructor {
            class_slots.insert(
                destructor.canonical_name.as_str(),
                vybe_ast::ProtocolSlot::Destructor,
            );
        }
        self.current_class_slot_keys.clear();
        let saved_class = self.current_class.take();
        let saved_implicit = self.current_class_implicit_self;
        self.current_class = Some(name.to_string());
        self.current_class_implicit_self = class.implicit_self_fields;

        // Pass 2 (ported to NormalClass): pre-register method + property
        // names in `defined_class_methods` so expression-compilation
        // doesn't hijack a method call via the value-method dispatch
        // table. Walks instance_methods + static_methods + properties
        // directly; no reconstructed member iteration.
        // An index operator makes `x[i]` on this type a method call. Record it
        // against the class so the index site can resolve it from the
        // receiver's static type instead of probing every index at runtime.
        //
        // Asked of the ROLE, not the spelling: a Ruby `[]`, a Dart
        // `operator[]` and a PHP `offsetGet` all fill `GetItem`, and none of
        // them is spelled `__getitem__` — the synonym list this used to
        // consult only ever knew two of the spellings.
        if class
            .special_methods
            .iter()
            .any(|s| s.kind == vybe_ast::ProtocolSlot::GetItem)
        {
            let cname = self.canon(&class.name);
            self.classes_with_indexer.insert(cname);
        }
        for m in class
            .instance_methods
            .iter()
            .chain(class.destructor.iter())
            .chain(class.static_methods.iter())
        {
            // Use `source_name` so existing compile paths that look up
            // `self.defined_class_methods.contains("ToString")` (from VB
            // call-site compilation) still hit. Canonical-name-only
            // lookups are a Phase 2b.3 concern.
            self.defined_class_methods
                .insert(self.canon(&m.source_name));
            if let Some(private_name) =
                self.js_private_member_storage_name_for_class(&class.name, &m.source_name)
            {
                self.defined_class_methods.insert(private_name);
            }
            if let Some(arity) = uniform_tuple_return_arity(&m.body) {
                let bound_name = if let Some(private_name) =
                    self.js_private_member_storage_name_for_class(&class.name, &m.source_name)
                {
                    private_name
                } else if m.source_name.starts_with("Symbol.") && !m.canonical_name.is_empty() {
                    m.canonical_name.clone()
                } else {
                    self.canon(&m.source_name)
                };
                self.multi_return_functions.insert(bound_name, arity);
                self.multi_return_functions.insert(
                    self.canon(&format!("{}.{}", class.name, m.source_name)),
                    arity,
                );
            }
        }
        for p in &class.properties {
            self.defined_class_methods
                .insert(self.canon(&p.source_name));
            if let Some(private_name) =
                self.js_private_member_storage_name_for_class(&class.name, &p.source_name)
            {
                self.defined_class_methods.insert(private_name);
            }
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
        let mut compile_normal_method = |cc: &mut Compiler,
                                         m: &NormalMethod,
                                         is_static: bool|
         -> Result<(), String> {
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
            let param_types: Vec<String> = user_params
                .iter()
                .map(|param| {
                    Compiler::normalize_type_hint(param.type_hint.as_deref().unwrap_or("object"))
                })
                .collect();
            let bound_name = if let Some(private_name) =
                cc.js_private_member_storage_name_for_class(&class.name, mname)
            {
                private_name
            } else if mname.starts_with("Symbol.") && !m.canonical_name.is_empty() {
                m.canonical_name.clone()
            } else {
                cc.canon(mname)
            };
            let storage_name = if !is_static
                && !m.is_override
                && m.raw_modifiers.is_hiding
                && cc.method_hides_ancestor(class.parent.as_deref(), &bound_name)
            {
                format!("__hide_{}${}", cc.canon(&class.name), bound_name)
            } else {
                bound_name.clone()
            };
            let qualified_name = cc.canon(&format!("{}.{}", class.name, mname));
            cc.function_param_modes.insert(
                bound_name.clone(),
                user_params.iter().map(|param| param.pass_by).collect(),
            );
            if storage_name != bound_name {
                cc.function_param_modes.insert(
                    storage_name.clone(),
                    user_params.iter().map(|param| param.pass_by).collect(),
                );
            }
            cc.function_param_modes.insert(
                qualified_name.clone(),
                user_params.iter().map(|param| param.pass_by).collect(),
            );
            cc.function_signatures
                .entry(bound_name.clone())
                .or_default()
                .push(CallSignature::from_params(
                    &user_params
                        .iter()
                        .map(|param| (*param).clone())
                        .collect::<Vec<_>>(),
                ));
            if let Some(return_type) = m.return_type.as_ref() {
                cc.function_return_types
                    .insert(bound_name.clone(), return_type.clone());
                if storage_name != bound_name {
                    cc.function_return_types
                        .insert(storage_name.clone(), return_type.clone());
                }
                cc.function_return_types
                    .insert(qualified_name, return_type.clone());
                cc.function_return_types.insert(
                    cc.canon(&format!("{}.{}", class.name, mname)),
                    return_type.clone(),
                );
            }
            // The receiver (`this`) is bound ambiently from the call context
            // (`__js_this`) rather than passed as an explicit first positional
            // parameter. Capability-driven — not gated on the language name.
            let ambient_this = cc.profile.ambient_this_binding;
            let uses_js_arguments = cc.profile.has_arguments_object
                && ambient_this
                && !m.is_generator
                && (user_params
                    .iter()
                    .any(|param| param.default.as_ref().is_some_and(expr_uses_js_arguments))
                    || m.body.iter().any(stmt_uses_js_arguments));
            let has_rest = user_params.last().map_or(false, |p| p.is_rest);
            let lowered_has_rest = has_rest || uses_js_arguments;
            let generator_control_arity = usize::from(m.is_generator && !lowered_has_rest);
            if uses_js_arguments {
                cc.rest_fixed_arities.insert(0);
            } else if has_rest {
                cc.rest_fixed_arities
                    .insert(user_params.len().saturating_sub(1) as u8);
            }
            // Whether this method carries an implicit leading receiver slot,
            // so its `arity` is `params + 1`. Mirrors the arity branches below.
            let has_receiver = if is_static_init {
                false
            } else if is_static {
                cc.profile.name == "php"
            } else if ambient_this {
                false
            } else {
                true
            };
            let arity = if uses_js_arguments {
                (1 + usize::from(has_receiver) + generator_control_arity) as u8
            } else {
                (user_params.len() + usize::from(has_receiver) + generator_control_arity) as u8
            };

            let ci = cc.chunks.len();
            let mut chunk = common::functions::create_function_chunk(mname, arity);
            chunk.is_method = has_receiver;
            chunk.param_count = user_params.len() as u8;
            // A WASM function's type shape (params → results) is what the
            // `call_indirect` runtime check compares. WASM functions have no
            // implicit receiver, so ALL declared params count — `user_params`
            // drops the phantom `self` that `explicit_self_param` assumes. The
            // result count is encoded as comma-joined placeholders in
            // `return_type` (None = a 0-result/void function, distinct from the
            // default 1-value ABI). Gated on the profile capability, not a name.
            if cc.profile.function_references {
                chunk.param_count = m.params.len() as u8;
                chunk.result_arity = m
                    .return_type
                    .as_ref()
                    .map(|rt| rt.split(',').count() as u8)
                    .unwrap_or(0);
            }
            chunk.is_async = m.is_async;
            chunk.is_generator = m.is_generator;
            // Source-level kind for prototype stamping at the attach sites —
            // survives walker lowerings that clear the outer flags. Generator
            // methods are lowered to a plain wrapper holding `__gen_fn`;
            // recover the source kind from that contract.
            let (src_async, src_gen) =
                Self::wrapped_generator_kind(&m.body).unwrap_or((m.is_async, m.is_generator));
            cc.method_fn_kinds.insert(ci, (src_async, src_gen));
            if m.is_generator && !m.is_async {
                let cname_canon = cc.canon(mname);
                cc.generator_functions.insert(cname_canon);
            }
            if let Some(&n) = cc.multi_return_functions.get(&bound_name) {
                chunk.result_arity = n;
            }
            cc.chunks.push(chunk);
            cc.scopes.push(Scope::new_function());
            cc.static_local_bindings.push(HashMap::new());
            let saved = cc.current;
            cc.current = ci;
            let saved_fn = cc.current_func_name.take();
            cc.current_func_name = Some(bound_name.clone());
            let saved_closure_captured = std::mem::take(&mut cc.current_closure_captured_locals);
            crate::primitives::collect_closure_captured_idents(
                &m.body,
                &mut cc.current_closure_captured_locals,
            );

            if !ambient_this && !is_static_init && (!is_static || cc.profile.name == "php") {
                cc.define_local(&self_kw);
            }
            let js_arguments_source_slot = if uses_js_arguments {
                Some(cc.define_local("__vybe_js_arguments_array"))
            } else {
                None
            };
            for p in &user_params {
                cc.define_local_typed(&p.name, p.type_hint.clone());
                let normalized_type_hint =
                    p.type_hint.as_deref().map(Compiler::normalize_type_hint);
                if normalized_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.ends_with("()"))
                    || normalized_type_hint
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.ends_with("()"))
                {
                    cc.record_array_binding(
                        &p.name,
                        ArrayBindingMetadata {
                            is_fixed: false,
                            type_hint: p.type_hint.clone(),
                            pascal_bounds: p.type_hint.as_deref().and_then(|type_hint| {
                                cc.pascal_array_type_hint_metadata(type_hint)
                            }),
                        },
                    );
                }
            }
            let js_arguments_len_slot = if uses_js_arguments {
                let slot = cc.define_local("arguments");
                cc.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
                cc.emit_u16(Op::LOCAL_SET, slot);
                cc.emit_u16(Op::LOCAL_GET, slot);
                cc.chunk().emit_string_const("Arguments", 0);
                let type_key = cc.str_const("__type");
                cc.emit_u16(Op::STRUCT_SET, type_key);
                cc.emit(Op::DROP);

                let len_slot = cc.define_local("__vybe_js_arguments_length");
                cc.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
                common::collections::emit_len(&mut cc.chunks, cc.current, cc.line);
                cc.emit_u16(Op::LOCAL_SET, len_slot);
                Some(len_slot)
            } else {
                None
            };
            if uses_js_arguments {
                for (index, p) in user_params.iter().enumerate() {
                    let slot = cc.scope().resolve(&p.name).unwrap();
                    if p.is_rest {
                        cc.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
                        cc.emit_const(Value::F64(index as f64));
                        cc.emit_u16(Op::LOCAL_GET, js_arguments_len_slot.unwrap());
                        common::collections::emit_slice(&mut cc.chunks, cc.current, cc.line);
                        cc.emit_u16(Op::LOCAL_SET, slot);
                    } else {
                        cc.emit_array_value_or_undefined(
                            js_arguments_source_slot.unwrap(),
                            js_arguments_len_slot.unwrap(),
                            index,
                        );
                        cc.emit_u16(Op::LOCAL_SET, slot);
                    }
                }
            }
            if ambient_this && !is_static && mname.starts_with('#') {
                let this_slot = cc.define_local("__js_private_method_this");
                let js_this = cc.str_const("__js_this");
                cc.emit_u16(Op::GLOBAL_GET, js_this);
                cc.emit_u16(Op::LOCAL_SET, this_slot);
                cc.emit_js_private_brand_check(this_slot, &bound_name)?;
            }
            if !ambient_this && !is_static_init && (!is_static || cc.profile.name == "php") {
                if class.explicit_self_param {
                    if let Some(self_param) = m.params.first() {
                        if self_param.name != self_kw {
                            let self_slot = cc.scope().resolve(&self_kw).unwrap();
                            let alias_slot = cc
                                .define_local_typed(&self_param.name, self_param.type_hint.clone());
                            cc.emit_u16(Op::LOCAL_GET, self_slot);
                            cc.emit_u16(Op::LOCAL_SET, alias_slot);
                        }
                    }
                }
            }
            let generator_control_slot = m
                .is_generator
                .then(|| cc.define_local("__generator_entry_control"));
            let ref_out_slots: Vec<u16> = user_params
                .iter()
                .filter(|param| matches!(param.pass_by, PassBy::Ref | PassBy::Out))
                .filter_map(|param| cc.scope().resolve(&param.name))
                .collect();
            let saved_rs = cc.current_result_slot.take();
            let saved_ref_out = cc.current_ref_out_params.take();
            let saved_member_static = cc.current_member_is_static;
            // Same handling `compile_function_decl` gives a nested function: the
            // enclosing shared-env SLOT is only meaningful in the enclosing
            // frame, so clear it, and pre-seed `closure_env_names` with that
            // frame's layout so `env[idx]` reads here line up with the array
            // actually being passed.
            let saved_shared_env_slot = cc.shared_env_slot.take();
            let saved_shared_env_names = std::mem::take(&mut cc.shared_env_names);
            let saved_closure_env_names = std::mem::take(&mut cc.closure_env_names);
            if !enclosing_shared_env_names.is_empty() {
                cc.closure_env_names = enclosing_shared_env_names.clone();
            }
            cc.current_result_slot = None;
            cc.current_ref_out_params = (!ref_out_slots.is_empty()).then_some(ref_out_slots);
            cc.current_member_is_static = is_static;

            if let Some(control_slot) = generator_control_slot {
                cc.emit_generator_entry_control(control_slot)?;
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
                    if cc.profile.missing_arg_is_undefined {
                        fn_call!(cc, "wasm:js-undefined", "test", 1);
                    } else {
                        cc.emit(Op::REF_IS_NULL);
                    }
                    let branch_line = cc.line;
                    cc.chunks[cc.current].emit_if(branch_line);
                    cc.compile_expr(default)?;
                    cc.emit_u16(Op::LOCAL_SET, slot);
                    cc.chunks[cc.current].emit_end(branch_line);
                }
            }

            if ambient_this
                && !is_static
                && crate::primitives::closures_in_body_reference_this(&m.body)
            {
                let this_idx = cc.str_const("__js_this");
                cc.emit_u16(Op::GLOBAL_GET, this_idx);
                let this_local = cc.define_local(&self_kw);
                cc.emit_u16(Op::LOCAL_SET, this_local);
                cc.current_closure_captured_locals.insert(self_kw.clone());
            }

            // Shared env for closures inside class methods: if the
            // method body has inner closures that capture the method's
            // locals, create a shared env array so mutations are visible
            // across all closures (same mechanism as compile_lambda_direct).
            if !cc.current_closure_captured_locals.is_empty() {
                let mut captured_names: Vec<String> = cc
                    .current_closure_captured_locals
                    .iter()
                    .filter(|name| !cc.defined_globals.contains(name.as_str()))
                    .cloned()
                    .collect();
                captured_names.sort();
                if !captured_names.is_empty() {
                    let env_size = captured_names.len() as u16;
                    let line = cc.line;
                    for _ in 0..env_size {
                        cc.emit(Op::NULL);
                    }
                    cc.chunks[cc.current].emit_array_new_fixed(0, env_size, line);
                    let env_slot = cc.define_local("__shared_env");
                    cc.emit_u16(Op::LOCAL_SET, env_slot);
                    cc.shared_env_slot = Some(env_slot);
                    cc.shared_env_names = captured_names.clone();

                    let mut local_decls: std::collections::HashSet<String> =
                        user_params.iter().map(|p| p.name.clone()).collect();
                    if !ambient_this && !is_static_init && (!is_static || cc.profile.name == "php")
                    {
                        local_decls.insert(self_kw.clone());
                    }
                    crate::primitives::collect_declared_names(&m.body, &mut local_decls);

                    for (idx, cap_name) in captured_names.iter().enumerate() {
                        if let Some(param_slot) = cc.scope().resolve(cap_name) {
                            cc.emit_u16(Op::LOCAL_GET, param_slot);
                            crate::primitives::closures::emit_env_set(
                                cc.chunk(),
                                env_slot,
                                idx as u16,
                                line,
                            );
                        }
                    }
                }
            }

            let async_try = if m.is_async && !m.is_generator && cc.profile.async_wraps_body_in_try {
                let line = cc.line;
                Some(common::functions::emit_async_body_start(
                    &mut cc.chunks[ci],
                    line,
                ))
            } else {
                None
            };
            if async_try.is_some() {
                cc.active_async_try_depth += 1;
            }

            if is_ctor {
                // §15.7.14: class constructors require `new` (JS only —
                // other languages construct through their own paths).
                if cc.profile.ecma_new_dispatch {
                    let line = cc.line;
                    crate::primitives::classes::emit_class_requires_new_guard(
                        cc.chunk(),
                        &class.name,
                        line,
                    );
                }
                for s in &m.body {
                    cc.compile_stmt(s)?;
                }
                if let Some(slot) = cc
                    .scope()
                    .resolve(&self_kw)
                    .or_else(|| cc.scope().resolve_ci(&self_kw))
                {
                    cc.emit_u16(Op::LOCAL_GET, slot);
                    cc.emit_return_through_finally(1)?;
                }
            } else if m.return_type.is_some() && result_style == ReturnStyle::ResultSlot {
                let slot_name = cc.profile.result_slot_name.clone();
                let rs = cc.define_local(&slot_name);
                let returns_self_type = m
                    .return_type
                    .as_deref()
                    .is_some_and(|rt| rt.eq_ignore_ascii_case(&class.name));
                if returns_self_type && body_has_result_member_assign(&m.body) {
                    cc.emit_var_get(&class.name);
                    cc.emit_u8(Op::CALL_REF, 0);
                } else {
                    cc.emit(Op::NULL);
                }
                cc.emit_u16(Op::LOCAL_SET, rs);
                cc.current_result_slot = Some(rs);
                for s in &m.body {
                    cc.compile_stmt(s)?;
                }
                cc.emit_u16(Op::LOCAL_GET, rs);
                cc.emit_return_through_finally(1)?;
            } else {
                for s in &m.body {
                    cc.compile_stmt(s)?;
                }
                if cc.current_ref_out_params.is_some() {
                    cc.emit(Op::NULL);
                    cc.emit_return_through_finally(1)?;
                } else {
                    let line = cc.line;
                    common::functions::emit_function_epilogue(&mut cc.chunks[ci], line);
                }
            }

            if async_try.is_some() {
                cc.active_async_try_depth = cc.active_async_try_depth.saturating_sub(1);
            }
            if let Some(catch_jump) = async_try {
                let line = cc.line;
                let chunk = &mut cc.chunks[ci];
                common::functions::emit_async_body_fallthrough(chunk, catch_jump, line);
                let resolve_idx = cc.import("ecma:promise", "resolve");
                cc.emit_host_call(resolve_idx, 1);
                cc.emit(Op::RETURN);
                let chunk = &mut cc.chunks[ci];
                common::functions::patch_async_body_catch(chunk, catch_jump);
                let reject_idx = cc.import("ecma:promise", "reject");
                cc.emit_host_call(reject_idx, 1);
                cc.emit(Op::RETURN);
            }

            cc.current_func_name = saved_fn;
            cc.current_result_slot = saved_rs;
            cc.current_ref_out_params = saved_ref_out;
            cc.current_member_is_static = saved_member_static;
            cc.current_closure_captured_locals = saved_closure_captured;
            cc.shared_env_slot = saved_shared_env_slot;
            cc.shared_env_names = saved_shared_env_names;
            cc.closure_env_names = saved_closure_env_names;

            let ns = cc.scope().next_slot;
            cc.chunks[ci].finalize_local_count(ns);
            cc.chunks[ci].local_names = cc.scope().defined_names.clone();
            let method_scope_idx = cc.scopes.len() - 1;
            let mut capture_names: Vec<String> = cc.scopes[method_scope_idx]
                .upvalues
                .iter()
                .enumerate()
                .filter_map(|(index, _)| {
                    cc.captured_name_for_upvalue(method_scope_idx, index as u8)
                })
                .collect();
            // `emit_var_get` registers an upvalue per NAME but emits the read as
            // `env[idx]` (bindings.rs), so these names describe what the body
            // reads, not what it receives: the body receives ONE upvalue, the
            // env array. Binding the names individually would hand the method a
            // raw value where it expects the array. Capture the env instead —
            // the same rule `compile_function_decl` applies when the parent has
            // a shared env ("pass it directly as the upvalue"). Only when every
            // capture lives in that env; anything else still binds by name.
            if !enclosing_shared_env_names.is_empty()
                && !capture_names.is_empty()
                && capture_names
                    .iter()
                    .all(|name| enclosing_shared_env_names.contains(name))
            {
                capture_names = vec!["__shared_env".to_string()];
            }
            cc.scopes.pop();
            cc.static_local_bindings.pop();
            cc.current = saved;
            method_capture_name_map.insert(ci, capture_names);
            if let Some(pending) = cc.pending_classes.get_mut(name) {
                let overloads = if is_static {
                    &mut pending.static_method_overloads
                } else {
                    &mut pending.instance_method_overloads
                };
                // Virtuality is decided HERE, once, so the call path stays
                // language-agnostic. Keyword languages (C#/VB/Pascal) mark the
                // method; virtual-by-default languages (java/python/js/...)
                // carry no keyword and opt in via the profile instead. A
                // `static` or `is_not_overridable` member can never be
                // overridden, so it keeps its direct bind either way.
                let is_virtual = m.is_virtual
                    || m.is_override
                    || m.is_abstract
                    || (cc.profile.methods_virtual_by_default
                        && !is_static
                        && !m.raw_modifiers.is_not_overridable);
                overloads
                    .entry(bound_name.clone())
                    .or_default()
                    .push(PendingMethodOverload {
                        param_types: param_types.clone(),
                        chunk_idx: ci,
                        return_type: m.return_type.clone(),
                        signature: CallSignature::from_params(&if uses_js_arguments {
                            vec![Param {
                                name: "__vybe_js_arguments_array".to_string(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: true,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            }]
                        } else {
                            user_params
                                .iter()
                                .map(|param| (*param).clone())
                                .collect::<Vec<_>>()
                        }),
                        is_virtual,
                    });
            }
            let explicit_overload_extends_ancestor = !is_static
                && !is_ctor
                && storage_name == bound_name
                && m.raw_modifiers.is_overloads
                && cc.method_hides_ancestor(class.parent.as_deref(), &bound_name);
            // Publish this method's slot, if it fills one. Keyed by the name
            // the bind sites will see, because that is all they have left.
            if let Some(slot) = class_slots.get(m.canonical_name.as_str()) {
                cc.current_class_slot_keys
                    .insert(storage_name.clone(), vybe_ast::protocol_slot_key(*slot));
            }
            if !explicit_overload_extends_ancestor {
                method_chunks.push((storage_name.clone(), ci, is_ctor, is_static));
            }
            if !is_static && !is_ctor && storage_name == bound_name {
                let overload_storage_name = cc.overload_storage_name(&bound_name, &param_types);
                if overload_storage_name != storage_name {
                    cc.function_param_modes.insert(
                        overload_storage_name.clone(),
                        user_params.iter().map(|param| param.pass_by).collect(),
                    );
                    cc.function_signatures
                        .entry(overload_storage_name.clone())
                        .or_default()
                        .push(CallSignature::from_params(
                            &user_params
                                .iter()
                                .map(|param| (*param).clone())
                                .collect::<Vec<_>>(),
                        ));
                    if let Some(return_type) = m.return_type.as_ref() {
                        cc.function_return_types
                            .insert(overload_storage_name.clone(), return_type.clone());
                    }
                    method_chunks.push((overload_storage_name, ci, is_ctor, is_static));
                }
            }
            Ok(())
        };

        for m in &class.instance_methods {
            compile_normal_method(self, m, false)?;
        }
        if let Some(destructor) = &class.destructor {
            compile_normal_method(self, destructor, false)?;
        }
        for m in &class.static_methods {
            compile_normal_method(self, m, true)?;
        }

        // --- Events / Consts / NestedTypes (from raw_extra_members) ---
        for m in &class.raw_extra_members {
            match m {
                ClassMember::Const {
                    name: cname, value, ..
                } => {
                    self.compile_expr(value)?;
                    let global_name = self.canon(&format!("{}.{}", name, cname));
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.defined_globals.insert(global_name);
                }
                ClassMember::Event { .. } => { /* type-level only */ }
                ClassMember::NestedType(stmt) => {
                    let nested = Self::qualify_nested_type_stmt(stmt, name);
                    let (qualified_nested, leaf_nested) = match &nested.kind {
                        StmtKind::ClassDecl {
                            name: nested_name, ..
                        }
                        | StmtKind::StructDecl {
                            name: nested_name, ..
                        }
                        | StmtKind::InterfaceDecl {
                            name: nested_name, ..
                        }
                        | StmtKind::EnumDecl {
                            name: nested_name, ..
                        } => (
                            self.canon(nested_name),
                            self.canon(nested_name.rsplit('.').next().unwrap_or(nested_name)),
                        ),
                        _ => (String::new(), String::new()),
                    };
                    self.compile_stmt(&nested)?;
                    if !qualified_nested.is_empty() && qualified_nested != leaf_nested {
                        let qualified_idx = self.str_const(&qualified_nested);
                        let leaf_idx = self.str_const(&leaf_nested);
                        self.emit_u16(Op::GLOBAL_GET, qualified_idx);
                        self.emit_u16(Op::GLOBAL_SET, leaf_idx);
                        for arity in 0..=IMPLICIT_CTOR_FORWARD_ARGS {
                            let qualified_ctor = ctor_global_for(&qualified_nested, arity as usize);
                            if self.defined_globals.contains(&qualified_ctor) {
                                let leaf_ctor = ctor_global_for(&leaf_nested, arity as usize);
                                let qualified_idx = self.str_const(&qualified_ctor);
                                let leaf_idx = self.str_const(&leaf_ctor);
                                self.emit_u16(Op::GLOBAL_GET, qualified_idx);
                                self.emit_u16(Op::GLOBAL_SET, leaf_idx);
                                self.defined_globals.insert(leaf_ctor);
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        // --- Properties: getter → __get_<prop>, setter → __set_<prop> ---
        for p in &class.properties {
            // Auto-properties are handled as plain fields in pass-1.
            if p.auto_field.is_some() {
                continue;
            }
            let pname_canon = if let Some(private_name) =
                self.js_private_member_storage_name_for_class(&class.name, &p.source_name)
            {
                private_name
            } else if let Some(storage_name) = field_storage_names.get(&self.canon(&p.source_name))
            {
                storage_name.clone()
            } else if !p.canonical_name.is_empty() {
                p.canonical_name.clone()
            } else {
                self.canon(&p.source_name)
            };
            let prop_is_static = p.is_static;

            if let Some(getter) = &p.getter {
                let get_name = format!("__get_{}", pname_canon);
                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk(&get_name, 1);
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;
                let saved_member_static = self.current_member_is_static;
                self.current_member_is_static = prop_is_static;
                self.define_local(&self_kw);

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
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, rs);
                    let saved_fn = self.current_func_name.take();
                    let saved_rs = self.current_result_slot.take();
                    self.current_func_name = Some(p.source_name.clone());
                    self.current_result_slot = Some(rs);
                    for s in &getter.body {
                        self.compile_stmt(s)?;
                    }
                    self.current_func_name = saved_fn;
                    self.current_result_slot = saved_rs;
                    self.emit_u16(Op::LOCAL_GET, rs);
                    self.emit(Op::RETURN);
                }

                {
                    let ns = self.scope().next_slot;
                    self.chunks[ci].finalize_local_count(ns);
                    self.chunks[ci].local_names = self.scope().defined_names.clone();
                }
                self.scopes.pop();
                self.current = saved;
                self.current_member_is_static = saved_member_static;
                method_chunks.push((get_name, ci, false, prop_is_static));
            }

            if let Some(setter) = &p.setter {
                let set_name = format!("__set_{}", pname_canon);
                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk(&set_name, 2);
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;
                let saved_member_static = self.current_member_is_static;
                self.current_member_is_static = prop_is_static;
                self.define_local(&self_kw);
                let value_param_name = setter
                    .params
                    .first()
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
                    for s in &setter.body {
                        self.compile_stmt(s)?;
                    }
                }

                let line = self.line;
                common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                {
                    let ns = self.scope().next_slot;
                    self.chunks[ci].finalize_local_count(ns);
                    self.chunks[ci].local_names = self.scope().defined_names.clone();
                }
                self.scopes.pop();
                self.current = saved;
                self.current_member_is_static = saved_member_static;
                method_chunks.push((set_name, ci, false, prop_is_static));
            }
        }

        self.current_class = saved_class;
        self.current_class_implicit_self = saved_implicit;

        const IMPLICIT_CTOR_FORWARD_ARGS: u8 = 16;
        let instance_methods: Vec<&(String, usize, bool, bool)> = method_chunks
            .iter()
            .filter(|(_, _, ic, is_static)| !*ic && !*is_static)
            .collect();
        let static_methods: Vec<&(String, usize, bool, bool)> = method_chunks
            .iter()
            .filter(|(_, _, ic, is_static)| !*ic && *is_static)
            .collect();
        let instance_method_names: Vec<String> = instance_methods
            .iter()
            .map(|(n, _, _, _)| n.clone())
            .collect();
        let method_rest_fixed_counts: HashMap<usize, u8> = self
            .pending_classes
            .values()
            .flat_map(|pc| {
                pc.instance_method_overloads
                    .values()
                    .chain(pc.static_method_overloads.values())
            })
            .flat_map(|overloads| overloads.iter())
            .filter(|overload| overload.signature.has_rest)
            .map(|overload| {
                (
                    overload.chunk_idx,
                    overload.signature.param_names.len().saturating_sub(1) as u8,
                )
            })
            .collect();
        let method_rest_fixed_count =
            |chunk_idx: usize| method_rest_fixed_counts.get(&chunk_idx).copied();

        // `constructors` is THE representation. `constructor` is the single
        // view of the same list and no longer selects a different emit path:
        // every normalizer fills both (`NormalMembers::push_constructor`, or
        // the language's own primary-selection rule where it has one), so a
        // class with a `constructor` and an empty `constructors` cannot be
        // produced. The arm that handled that case is gone.
        let ctor_variants: Vec<Option<&NormalConstructor>> = if class.constructors.is_empty() {
            vec![None]
        } else {
            class.constructors.iter().map(Some).collect()
        };
        let ctor_global_prefix = self.canon(name);
        let should_stamp_form_identity = self.class_requires_form_identity_stamp(parent);
        for ctor_variant in &ctor_variants {
            let explicit_arity = ctor_variant
                .map(|ctor| {
                    let skip = if class.explicit_self_param { 1 } else { 0 };
                    ctor.params.len().saturating_sub(skip)
                })
                .unwrap_or_else(|| {
                    if parent.is_some() {
                        IMPLICIT_CTOR_FORWARD_ARGS as usize
                    } else {
                        0
                    }
                });
            self.defined_globals
                .insert(ctor_global_for(&ctor_global_prefix, explicit_arity));
        }

        // Captures are re-resolved BY NAME in whichever frame the ref is being
        // emitted from — this helper is referenced from two different frames
        // (the constructor's arity dispatcher, and the class's defining scope).
        // An `UpvalueDesc`'s `(is_local, index)` are coordinates into ONE
        // parent frame, so replaying them verbatim in the other frame reads a
        // different slot entirely: a class declared inside a function had its
        // ctor helper capture the enclosing local by slot, then the dispatcher
        // replayed that slot against its own params and got `undefined`.
        // Re-resolving also threads the capture through the dispatcher's own
        // upvalue list, so the class closure is built carrying it.
        let emit_helper_ref = |cc: &mut Compiler,
                               helper_idx: usize,
                               helper_captures: &[String]|
         -> Result<(), String> {
            cc.emit_ref_func_with_captures(helper_idx, helper_captures, false)
        };

        let mut ctor_helpers: Vec<(usize, usize, usize, Vec<String>, Option<String>)> = Vec::new();
        for (ctor_index, ctor_variant) in ctor_variants.iter().enumerate() {
            let helper_name = format!("__{}_ctor_{}", name, ctor_index);
            let ctor_base_args_from_nc: Option<Vec<Expression>> = ctor_variant.and_then(|c| {
                if let BaseCall::Explicit(args) = &c.base_call {
                    Some(args.iter().map(|a| a.value.clone()).collect())
                } else {
                    None
                }
            });
            let ctor_auto_base = ctor_variant
                .map(|c| matches!(c.base_call, BaseCall::Auto))
                .unwrap_or(false);
            let ctor_this_args: Option<Vec<Expression>> = ctor_variant.and_then(|c| {
                if let BaseCall::This(args) = &c.base_call {
                    Some(args.iter().map(|a| a.value.clone()).collect())
                } else {
                    None
                }
            });
            let ctor_body: Option<(&Vec<Statement>, &Vec<Param>, Option<&Vec<Expression>>)> =
                ctor_variant.map(|c| (&c.body, &c.params, ctor_base_args_from_nc.as_ref()));
            if let Some((_, params, _)) = ctor_body {
                let skip = if class.explicit_self_param { 1 } else { 0 };
                let ctor_params: Vec<Param> = params.iter().skip(skip).cloned().collect();
                self.constructor_signatures
                    .entry(self.canon(name))
                    .or_default()
                    .push(CallSignature::from_params(&ctor_params));
            }
            let user_params: Vec<String> = ctor_body
                .map(|(_, params, _)| {
                    if class.explicit_self_param {
                        params.iter().skip(1).map(|p| p.name.clone()).collect()
                    } else {
                        params.iter().map(|p| p.name.clone()).collect()
                    }
                })
                .unwrap_or_default();
            let ctor_min_arity = ctor_body
                .map(|(_, params, _)| {
                    let skip = if class.explicit_self_param { 1 } else { 0 };
                    params
                        .iter()
                        .skip(skip)
                        .take_while(|p| p.default.is_none() && !p.is_rest)
                        .count()
                })
                .unwrap_or(0);
            let synthesized_forward_args = ctor_body.is_none() && parent.is_some();
            let user_arity = if synthesized_forward_args {
                IMPLICIT_CTOR_FORWARD_ARGS
            } else {
                user_params.len() as u8
            };

            let helper_idx = self.chunks.len();
            self.chunks.push(common::functions::create_function_chunk(
                &helper_name,
                user_arity,
            ));
            self.scopes.push(Scope::new_function());
            let saved_cur = self.current;
            let saved_class2 = self.current_class.take();
            let saved_implicit2 = self.current_class_implicit_self;
            // As for methods: the enclosing shared-env slot names a local in the
            // ENCLOSING frame, so it must not leak into this chunk — reading it
            // here would index this frame's slot instead. Cleared, with
            // `closure_env_names` pre-seeded to the enclosing layout so the
            // `env[idx]` reads match the array this ctor is handed.
            let saved_ctor_shared_env_slot = self.shared_env_slot.take();
            let saved_ctor_shared_env_names = std::mem::take(&mut self.shared_env_names);
            let saved_ctor_closure_env_names = std::mem::take(&mut self.closure_env_names);
            if !enclosing_shared_env_names.is_empty() {
                self.closure_env_names = enclosing_shared_env_names.clone();
            }
            self.current = helper_idx;
            self.current_class = Some(name.to_string());
            self.current_class_implicit_self = class.implicit_self_fields;

            let ctor_param_defaults: Vec<Option<Expression>> = ctor_body
                .map(|(_, params, _)| {
                    let skip = if class.explicit_self_param { 1 } else { 0 };
                    params
                        .iter()
                        .skip(skip)
                        .map(|p| p.default.clone())
                        .collect()
                })
                .unwrap_or_default();
            for p in &user_params {
                self.define_local(p);
            }
            // §15.7.14: class constructors require `new` (JS only).
            // `__js_new_target` is null on plain calls; every `new` chain
            // (incl. super()) sets or defaults it before this body runs.
            // Emitted AFTER param slots are claimed — emitter scratch
            // allocation before define_local shifts param slots (the
            // documented alloc_scratch/define_local collision).
            if self.profile.ecma_new_dispatch {
                let line = self.line;
                crate::primitives::classes::emit_class_requires_new_guard(self.chunk(), name, line);
            }
            for (i, p) in user_params.iter().enumerate() {
                if let Some(Some(default)) = ctor_param_defaults.get(i) {
                    let slot = self.scope().resolve(p).unwrap();
                    self.emit_u16(Op::LOCAL_GET, slot);
                    if self.profile.missing_arg_is_undefined {
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                    } else {
                        self.emit(Op::REF_IS_NULL);
                    }
                    // Result is already I32(0/1) — no dyn_to_bool needed
                    let branch_line = self.line;
                    self.chunks[self.current].emit_if(branch_line);
                    self.compile_expr(default)?;
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.chunks[self.current].emit_end(branch_line);
                }
            }
            if synthesized_forward_args {
                for i in 0..IMPLICIT_CTOR_FORWARD_ARGS {
                    self.define_local(&format!("__implicit_arg_{}", i));
                }
            }
            self.define_local(&self_kw);
            let this_slot = user_arity as u16;
            if self.profile.ambient_this_binding {
                let js_this = self.str_const("__js_this");
                self.emit_u16(Op::GLOBAL_GET, js_this);
                self.emit_u16(Op::LOCAL_SET, this_slot);
            }
            // §9.1.1.3.4 (JS): derived-constructor `this` TDZ context.
            // While this chunk's body compiles, `this` reads and `super()`
            // calls emit runtime guards against this_slot (null until
            // super() initializes it). Saved/restored so nested classes
            // compiled mid-body don't leak the context.
            let saved_derived_ctx = self.js_derived_ctor_ctx.take();
            if self.profile.ecma_new_dispatch && parent.is_some() && ctor_body.is_some() {
                self.js_derived_ctor_ctx = Some((self.current, this_slot));
            }

            let line = self.line;
            if let Some(this_args) = ctor_this_args {
                let ctor_global = ctor_global_for(&ctor_global_prefix, this_args.len());
                self.emit_var_get(&ctor_global);
                for expr in &this_args {
                    self.compile_expr(expr)?;
                }
                self.emit_u8(Op::CALL_REF, this_args.len() as u8);
                self.emit_u16(Op::LOCAL_SET, this_slot);
                if let Some((body, _, _)) = ctor_body {
                    for stmt in body {
                        self.compile_stmt(stmt)?;
                    }
                }
                for (mname, mci, _, _) in &instance_methods {
                    if mname.starts_with("__get_") {
                        let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                        crate::primitives::object::emit_bind_getter(
                            self.chunk(),
                            this_slot,
                            prop,
                            *mci,
                            line,
                        );
                    } else if mname.starts_with("__set_") {
                        let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                        crate::primitives::object::emit_bind_setter(
                            self.chunk(),
                            this_slot,
                            prop,
                            *mci,
                            line,
                        );
                    } else {
                        let capture_names = method_capture_name_map
                            .get(mci)
                            .map(Vec::as_slice)
                            .unwrap_or(&[]);
                        if self.class_prototype_dispatch()
                            && capture_names.is_empty()
                            && !mname.starts_with("__js_private_")
                        {
                            continue;
                        }
                        self.emit_bind_instance_method_with_aliases(
                            this_slot,
                            mname,
                            *mci,
                            capture_names,
                            method_rest_fixed_count(*mci),
                            !self.class_prototype_dispatch(),
                        )?;
                    }
                }
                crate::primitives::classes::emit_constructor_return(self.chunk(), this_slot, line);
            } else {
                let is_child = parent.is_some();
                let parent_ctor_is_bound = if let Some(parent_name) = parent {
                    if self.dotnet_descriptor_parent_has_no_user_ctor(parent_name) {
                        false
                    } else {
                        let pname = self.canon(parent_name);
                        let has_local = self
                            .scope()
                            .resolve(parent_name)
                            .or_else(|| {
                                if self.case_sensitive {
                                    None
                                } else {
                                    self.scope().resolve_ci(parent_name)
                                }
                            })
                            .is_some();
                        let has_upvalue = self.scopes.len() > 1
                            && self
                                .resolve_upvalue(self.scopes.len() - 1, parent_name)
                                .is_some();
                        let has_static_local = self.static_local_binding(parent_name).is_some();
                        has_local
                            || has_upvalue
                            || has_static_local
                            || self.defined_globals.contains(&pname)
                            || self.defined_classes.contains(&pname)
                            || (self.profile.has_ecma_globals
                                && Self::is_ecma_typed_array_ctor_name(parent_name)
                                && !self.shadows_builtin_type(parent_name))
                            || (self.profile.has_ecma_globals
                                && Self::is_ecma_array_buffer_ctor_name(parent_name)
                                && !self.shadows_builtin_type(parent_name))
                            || (self.profile.has_ecma_globals
                                && Self::is_ecma_collection_ctor_name(parent_name)
                                && !self.shadows_builtin_type(parent_name))
                    }
                } else {
                    false
                };
                if is_child {
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, this_slot);

                    let has_explicit_base =
                        ctor_body.as_ref().map_or(false, |(_, _, ba)| ba.is_some());
                    let auto_base_needed = !has_explicit_base
                        && ctor_body.is_some()
                        && ctor_auto_base
                        && parent.is_some()
                        && {
                            let stmts = ctor_body
                                .as_ref()
                                .map(|(b, _, _)| b.as_slice())
                                .unwrap_or(&[]);
                            !body_has_super_call(stmts)
                        };

                    // Framework GUI control parent (`class Form1 : Form`):
                    // construct via the `vybe:gui` host factory directly. The
                    // host builds the full control (identity + Controls
                    // collection), replacing the per-class ctor global we no
                    // longer emit for control types.
                    // Only intercept on the paths where the block below
                    // (`has_explicit_base || auto_base_needed ||
                    // ctor_body.is_none()`) consumes our `this_slot`
                    // construction; when a ctor body drives `super()` itself,
                    // leave construction to that statement so we don't build
                    // the instance twice.
                    let framework_ctrl_parent = parent
                        .as_deref()
                        .filter(|p| self.is_framework_control_parent(p))
                        .filter(|_| has_explicit_base || auto_base_needed || ctor_body.is_none());
                    if let Some(fp) = framework_ctrl_parent {
                        let base_args: Vec<Expression> = ctor_body
                            .as_ref()
                            .and_then(|(_, _, ba)| ba.as_ref())
                            .map(|a| a.to_vec())
                            .unwrap_or_default();
                        self.try_emit_framework_control_base(fp, &base_args, this_slot)?;
                    } else if let Some((_, _, base_args)) = &ctor_body {
                        if let Some(bargs) = base_args {
                            if let Some(parent_name) = parent {
                                if parent_ctor_is_bound {
                                    self.emit_default_js_new_target(name);
                                    self.emit_parent_ctor_value(parent_name);
                                    for a in *bargs {
                                        self.compile_expr(a)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, bargs.len() as u8);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                } else {
                                    let canon_name = self.canon(name);
                                    crate::primitives::classes::emit_new_typed_object(
                                        self.chunk(),
                                        this_slot,
                                        &canon_name,
                                        line,
                                    );
                                }
                            }
                        } else if auto_base_needed {
                            if let Some(parent_name) = parent {
                                if parent_ctor_is_bound {
                                    self.emit_default_js_new_target(name);
                                    self.emit_parent_ctor_value(parent_name);
                                    self.emit_u8(Op::CALL_REF, 0);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                } else {
                                    let canon_name = self.canon(name);
                                    crate::primitives::classes::emit_new_typed_object(
                                        self.chunk(),
                                        this_slot,
                                        &canon_name,
                                        line,
                                    );
                                }
                            }
                        }
                    } else if let Some(parent_name) = parent {
                        if parent_ctor_is_bound {
                            self.emit_default_js_new_target(name);
                            self.emit_parent_ctor_value(parent_name);
                            if synthesized_forward_args {
                                let parent_ctor_slot =
                                    self.define_local(&format!("__{}_parent_ctor", helper_name));
                                self.emit_u16(Op::LOCAL_SET, parent_ctor_slot);
                                let parent_called_slot =
                                    self.define_local(&format!("__{}_parent_called", helper_name));
                                inst!(self, core_wasm::i32_const, 0);
                                self.emit_u16(Op::LOCAL_SET, parent_called_slot);
                                for count in (1..=IMPLICIT_CTOR_FORWARD_ARGS).rev() {
                                    self.emit_u16(Op::LOCAL_GET, parent_called_slot);
                                    self.emit(Op::I32_EQZ);
                                    self.chunks[self.current].emit_if(line);
                                    self.emit_u16(Op::LOCAL_GET, (count - 1) as u16);
                                    self.emit(Op::REF_IS_NULL);
                                    self.emit(Op::I32_EQZ);
                                    self.chunks[self.current].emit_if(line);
                                    self.emit_u16(Op::LOCAL_GET, parent_ctor_slot);
                                    for arg_index in 0..count {
                                        self.emit_u16(Op::LOCAL_GET, arg_index as u16);
                                    }
                                    self.emit_u8(Op::CALL_REF, count);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                    inst!(self, core_wasm::i32_const, 1);
                                    self.emit_u16(Op::LOCAL_SET, parent_called_slot);
                                    self.chunks[self.current].emit_end(line);
                                    self.chunks[self.current].emit_end(line);
                                }
                                self.emit_u16(Op::LOCAL_GET, parent_called_slot);
                                self.emit(Op::I32_EQZ);
                                self.chunks[self.current].emit_if(line);
                                self.emit_u16(Op::LOCAL_GET, parent_ctor_slot);
                                self.emit_u8(Op::CALL_REF, 0);
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                                self.chunks[self.current].emit_end(line);
                            } else {
                                for i in 0..user_arity {
                                    self.emit_u16(Op::LOCAL_GET, i as u16);
                                }
                                self.emit_u8(Op::CALL_REF, user_arity);
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                            }
                        } else if self.profile.ecma_error_object_shape
                            && common::errors::is_exception_type(parent_name)
                            && !self.shadows_builtin_type(parent_name)
                        {
                            // §15.7.14 default derived ctor over an intrinsic
                            // error parent: super(message) — construct through
                            // the canonical exception shape so message/chain
                            // match a directly-constructed parent error.
                            self.emit_u16(Op::LOCAL_GET, 0);
                            self.emit(Op::REF_IS_NULL);
                            self.chunks[self.current].emit_if_value(line);
                            self.emit_const(Value::String(Arc::from("")));
                            self.chunks[self.current].emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, 0);
                            self.chunks[self.current].emit_end(line);
                            self.emit_js_exception_ctor_from_message_value(parent_name)?;
                            self.emit_u16(Op::LOCAL_SET, this_slot);
                        } else {
                            let canon_name = self.canon(name);
                            crate::primitives::classes::emit_new_typed_object(
                                self.chunk(),
                                this_slot,
                                &canon_name,
                                line,
                            );
                        }
                    }

                    let no_base_derived_body =
                        ctor_body.as_ref().is_some_and(|(body, _, base_args)| {
                            base_args.is_none() && !body_has_super_call(body)
                        }) && !has_explicit_base
                            && !auto_base_needed
                            && !self.profile.ecma_new_dispatch;
                    if no_base_derived_body {
                        let canon_name = self.canon(name);
                        crate::primitives::classes::emit_new_typed_object(
                            self.chunk(),
                            this_slot,
                            &canon_name,
                            line,
                        );
                    }

                    if has_explicit_base
                        || auto_base_needed
                        || ctor_body.is_none()
                        || no_base_derived_body
                    {
                        self.emit_u16(Op::LOCAL_GET, this_slot);
                        self.emit_const(Value::String(Arc::from(name)));
                        let type_key = self.str_const("__type");
                        self.emit_u16(Op::STRUCT_SET, type_key);
                        self.emit(Op::DROP);
                        if class.is_value_type {
                            crate::primitives::classes::emit_value_equality_stamp(
                                self.chunk(),
                                this_slot,
                                line,
                            );
                        }

                        if self.class_prototype_dispatch() {
                            let proto_link_key = self.str_const("__proto__");
                            let proto_local =
                                self.define_local(&format!("__{}_link_proto", helper_name));
                            self.emit_load_instance_proto(name);
                            self.emit_u16(Op::LOCAL_SET, proto_local);
                            self.emit_u16(Op::LOCAL_GET, proto_local);
                            self.emit(Op::REF_IS_NULL);
                            self.emit(Op::I32_EQZ);
                            self.chunks[self.current].emit_if(line);
                            self.emit_u16(Op::LOCAL_GET, this_slot);
                            self.emit_u16(Op::LOCAL_GET, proto_local);
                            self.emit_u16(Op::STRUCT_SET, proto_link_key);
                            self.emit(Op::DROP);
                            self.chunks[self.current].emit_end(line);
                        }

                        for (fname, type_hint, init, array_bounds) in &field_inits {
                            self.emit_class_field_initializer(
                                this_slot,
                                fname,
                                type_hint.as_deref(),
                                init.as_ref(),
                                array_bounds.as_deref(),
                                class.is_value_type,
                                line,
                            )?;
                        }

                        if let Some(parent_name) = parent {
                            let pname = self.canon(parent_name);
                            for method_name in &instance_method_names {
                                crate::primitives::classes::emit_save_base_method(
                                    self.chunk(),
                                    this_slot,
                                    method_name,
                                    line,
                                );
                            }
                            self.emit_store_super_ref(this_slot, &pname);
                        }

                        // Multiple inheritance (opt-in): attach every MRO
                        // ancestor's methods in reverse-C3 order before self's
                        // own, so a nearer base overrides a farther one. No-op
                        // for single-inheritance classes / non-MI languages.
                        self.emit_mi_ancestor_methods(name, this_slot, line);
                        for (mname, mci, _, _) in &instance_methods {
                            if mname.starts_with("__get_") {
                                let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                                crate::primitives::object::emit_bind_getter(
                                    self.chunk(),
                                    this_slot,
                                    prop,
                                    *mci,
                                    line,
                                );
                            } else if mname.starts_with("__set_") {
                                let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                                crate::primitives::object::emit_bind_setter(
                                    self.chunk(),
                                    this_slot,
                                    prop,
                                    *mci,
                                    line,
                                );
                            } else {
                                let capture_names = method_capture_name_map
                                    .get(mci)
                                    .map(Vec::as_slice)
                                    .unwrap_or(&[]);
                                if self.class_prototype_dispatch()
                                    && capture_names.is_empty()
                                    && !mname.starts_with("__js_private_")
                                {
                                    continue;
                                }
                                self.emit_bind_instance_method_with_aliases(
                                    this_slot,
                                    mname,
                                    *mci,
                                    capture_names,
                                    method_rest_fixed_count(*mci),
                                    !self.class_prototype_dispatch(),
                                )?;
                            }
                        }

                        let ctor_stmts: &[Statement] = ctor_body
                            .as_ref()
                            .map(|(b, _, _)| b.as_slice())
                            .unwrap_or(&[]);
                        if should_stamp_form_identity && !body_has_identity_stamp(ctor_stmts) {
                            self.emit_form_identity_stamp(this_slot, name, line);
                        }
                        for aim in &class.auto_init_methods {
                            let has_method = instance_methods
                                .iter()
                                .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                            if has_method && !body_calls_method(ctor_stmts, aim) {
                                crate::primitives::classes::emit_auto_init_call(
                                    self.chunk(),
                                    this_slot,
                                    aim,
                                    line,
                                );
                            }
                        }

                        if let Some((body, _, _)) = ctor_body {
                            for stmt in body {
                                self.compile_stmt(stmt)?;
                            }
                        }
                    } else {
                        let body_stmts: &[Statement] = ctor_body
                            .as_ref()
                            .map(|(b, _, _)| b.as_slice())
                            .unwrap_or(&[]);
                        let is_super_call = |stmt: &Statement| {
                            if let StmtKind::Expr(expr) = &stmt.kind {
                                matches!(&expr.kind, ExprKind::SuperCall { .. })
                                    || matches!(&expr.kind, ExprKind::Call { callee, .. } if matches!(callee.kind, ExprKind::Super))
                            } else {
                                false
                            }
                        };
                        let super_idx = body_stmts.iter().position(is_super_call);
                        let preamble_end = match super_idx {
                            Some(index) => {
                                let mut end = index + 1;
                                while end < body_stmts.len() && is_identity_stamp(&body_stmts[end])
                                {
                                    end += 1;
                                }
                                end
                            }
                            None => 0,
                        };
                        for stmt in &body_stmts[..preamble_end] {
                            self.compile_stmt(stmt)?;
                        }
                        let user_body = &body_stmts[preamble_end..];
                        // JS: when super() isn't a top-level statement
                        // (e.g. `try{super();}catch{}`), its completion
                        // point isn't statically known — defer the
                        // instance stamps until after the body, guarded by
                        // `this != null` (§9.1.1.3.4: this_slot stays null
                        // when super() never ran; the constructor-return
                        // TDZ guard throws the ReferenceError then).
                        let stamps_deferred = super_idx.is_none() && self.profile.ecma_new_dispatch;
                        if !stamps_deferred {
                            self.emit_derived_ctor_stamps(
                                name,
                                this_slot,
                                parent,
                                &instance_method_names,
                                &field_inits,
                                &instance_methods,
                                &method_capture_name_map,
                                &method_rest_fixed_counts,
                                class.is_value_type,
                                should_stamp_form_identity,
                                body_stmts,
                                user_body,
                                &class.auto_init_methods,
                                line,
                            )?;
                        }
                        for stmt in user_body {
                            self.compile_stmt(stmt)?;
                        }
                        if stamps_deferred {
                            self.emit_u16(Op::LOCAL_GET, this_slot);
                            self.emit(Op::REF_IS_NULL);
                            self.emit(Op::I32_EQZ);
                            self.chunks[self.current].emit_if(line);
                            self.emit_derived_ctor_stamps(
                                name,
                                this_slot,
                                parent,
                                &instance_method_names,
                                &field_inits,
                                &instance_methods,
                                &method_capture_name_map,
                                &method_rest_fixed_counts,
                                class.is_value_type,
                                should_stamp_form_identity,
                                body_stmts,
                                user_body,
                                &class.auto_init_methods,
                                line,
                            )?;
                            self.chunks[self.current].emit_end(line);
                        }
                    }
                } else {
                    let canon_name = self.canon(name);
                    crate::primitives::classes::emit_new_typed_object(
                        self.chunk(),
                        this_slot,
                        &canon_name,
                        line,
                    );
                    if class.is_value_type {
                        crate::primitives::classes::emit_value_equality_stamp(
                            self.chunk(),
                            this_slot,
                            line,
                        );
                    }
                    for (fname, type_hint, init, array_bounds) in &field_inits {
                        self.emit_class_field_initializer(
                            this_slot,
                            fname,
                            type_hint.as_deref(),
                            init.as_ref(),
                            array_bounds.as_deref(),
                            class.is_value_type,
                            line,
                        )?;
                    }
                    for (mname, mci, _, _) in &instance_methods {
                        if mname.starts_with("__get_") {
                            let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                            crate::primitives::object::emit_bind_getter(
                                self.chunk(),
                                this_slot,
                                prop,
                                *mci,
                                line,
                            );
                        } else if mname.starts_with("__set_") {
                            let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                            crate::primitives::object::emit_bind_setter(
                                self.chunk(),
                                this_slot,
                                prop,
                                *mci,
                                line,
                            );
                        } else {
                            let capture_names = method_capture_name_map
                                .get(mci)
                                .map(Vec::as_slice)
                                .unwrap_or(&[]);
                            if self.class_prototype_dispatch()
                                && capture_names.is_empty()
                                && !mname.starts_with("__js_private_")
                            {
                                continue;
                            }
                            self.emit_bind_instance_method_with_aliases(
                                this_slot,
                                mname,
                                *mci,
                                capture_names,
                                method_rest_fixed_count(*mci),
                                !self.class_prototype_dispatch(),
                            )?;
                        }
                    }
                    if self.class_prototype_dispatch() {
                        let proto_link_key = self.str_const("__proto__");
                        let proto_local =
                            self.define_local(&format!("__{}_link_proto_base", helper_name));
                        self.emit_load_instance_proto(name);
                        self.emit_u16(Op::LOCAL_SET, proto_local);
                        self.emit_u16(Op::LOCAL_GET, proto_local);
                        self.emit(Op::REF_IS_NULL);
                        self.emit(Op::I32_EQZ);
                        self.chunks[self.current].emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, this_slot);
                        self.emit_u16(Op::LOCAL_GET, proto_local);
                        self.emit_u16(Op::STRUCT_SET, proto_link_key);
                        self.emit(Op::DROP);
                        self.chunks[self.current].emit_end(line);
                    }
                    let ctor_stmts: &[Statement] = ctor_body
                        .as_ref()
                        .map(|(b, _, _)| b.as_slice())
                        .unwrap_or(&[]);
                    for aim in &class.auto_init_methods {
                        let has_method = instance_methods
                            .iter()
                            .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                        if has_method && !body_calls_method(ctor_stmts, aim) {
                            crate::primitives::classes::emit_auto_init_call(
                                self.chunk(),
                                this_slot,
                                aim,
                                line,
                            );
                        }
                    }
                    if let Some((body, _, _)) = ctor_body {
                        for stmt in body {
                            self.compile_stmt(stmt)?;
                        }
                    }
                }

                if !self.class_prototype_dispatch() {
                    // Runtime class identity for languages that carry it in
                    // instance FIELDS. Where methods dispatch through a
                    // prototype, the chain already answers "which class is
                    // this" and re-stamping would fight it.
                    //
                    // This must run AFTER the
                    // ctor body / parent-ctor call — a child ctor receives
                    // `this` from the parent (synthesized forward OR
                    // `parent::__construct()` in the body) carrying the
                    // PARENT's type_id and constructor, so the child
                    // re-stamps:
                    //  - `constructor` → the class object, for `new static`
                    //    and get_class ($this.constructor.name); the JS
                    //    path gets this via the prototype chain instead.
                    //  - `__type` + WASM GC type_id, so instanceof
                    //    (REF_TEST fast path) sees the runtime class.
                    let class_global = self.str_const(name);
                    let ctor_key = self.str_const("constructor");
                    self.emit_u16(Op::LOCAL_GET, this_slot);
                    self.emit_u16(Op::GLOBAL_GET, class_global);
                    self.emit_u16(Op::STRUCT_SET, ctor_key);
                    self.emit(Op::DROP);
                    let canon_name = self.canon(name);
                    crate::primitives::classes::emit_retype_object(
                        self.chunk(),
                        this_slot,
                        &canon_name,
                        line,
                    );
                }

                if self.profile.class_introspection_metadata {
                    // Link each instance to its class object via `__class__`
                    // so `type(obj)` returns the class (and `type(obj) is Cls`
                    // / `type(obj).__name__` work). Re-stamped derived-last,
                    // like the PHP `constructor` block above, so `type(child)`
                    // is the child class even after a parent ctor ran.
                    let class_global = self.str_const(name);
                    let class_key = self.str_const("__class__");
                    self.emit_u16(Op::LOCAL_GET, this_slot);
                    self.emit_u16(Op::GLOBAL_GET, class_global);
                    self.emit_u16(Op::STRUCT_SET, class_key);
                    self.emit(Op::DROP);
                }

                crate::primitives::reflection::emit_instanceof_chain(
                    &mut self.chunks,
                    self.current,
                    this_slot,
                    name,
                    line,
                );
                if self.profile.class_multiple_inheritance && class.bases.len() > 1 {
                    // The primary parent-ctor chain already stamped its own
                    // lineage into `__types`; add the EXTRA MRO ancestors (the
                    // non-primary bases and their exclusive ancestors) so
                    // `isinstance()` recognises every base in the diamond.
                    let mut primary = std::collections::HashSet::new();
                    let mut cur = Some(self.canon(name));
                    while let Some(c) = cur {
                        cur = self
                            .pending_classes
                            .get(&c)
                            .and_then(|pc| pc.parent.as_ref().map(|p| self.canon(p)));
                        primary.insert(c);
                    }
                    for cls in self.c3_linearize(name).into_iter().skip(1) {
                        if !primary.contains(&cls) {
                            crate::primitives::reflection::emit_instanceof_chain(
                                &mut self.chunks,
                                self.current,
                                this_slot,
                                &cls,
                                line,
                            );
                        }
                    }
                }
                let mut interface_names = class.interfaces.clone();
                interface_names.extend(self.reflection_interfaces(name));
                let mut seen_interfaces = std::collections::HashSet::new();
                for interface_name in interface_names {
                    if !seen_interfaces.insert(self.canon(&interface_name)) {
                        continue;
                    }
                    crate::primitives::reflection::emit_instanceof_chain(
                        &mut self.chunks,
                        self.current,
                        this_slot,
                        &interface_name,
                        line,
                    );
                }
                // Set __proto__ link for prototype-dispatch classes.
                // Done just before return so this_slot is guaranteed valid.
                if self.class_prototype_dispatch() {
                    let proto_key = self.str_const("__proto__");
                    self.emit_load_instance_proto(name);
                    let tmp = self.define_local("__final_proto");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit(Op::REF_IS_NULL);
                    self.emit(Op::I32_EQZ);
                    self.chunks[self.current].emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, this_slot);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u16(Op::STRUCT_SET, proto_key);
                    self.emit(Op::DROP);
                    self.chunks[self.current].emit_end(line);
                }
                // §9.1.1.3.4 (JS): returning from a derived constructor
                // with `this` still uninitialized (super() missing, or its
                // throw was caught) is a ReferenceError.
                if self.js_derived_ctor_ctx == Some((self.current, this_slot)) {
                    crate::primitives::classes::emit_this_initialized_guard(
                        self.chunk(),
                        this_slot,
                        line,
                    );
                }
                crate::primitives::classes::emit_constructor_return(self.chunk(), this_slot, line);
            }

            {
                let ns = self.scope().next_slot;
                self.chunks[helper_idx].finalize_local_count(ns);
                self.chunks[helper_idx].local_names = self.scope().defined_names.clone();
            }
            // Names, not slot coordinates — see `emit_helper_ref`. Resolved
            // while the helper's scope is still on the stack, since
            // `captured_name_for_upvalue` walks it to name each upvalue.
            let helper_scope_idx = self.scopes.len() - 1;
            let mut helper_captures: Vec<String> = (0..self.scopes[helper_scope_idx]
                .upvalues
                .len())
                .filter_map(|index| self.captured_name_for_upvalue(helper_scope_idx, index as u8))
                .collect();
            // Same rule the methods use: a ctor body reading an enclosing
            // captured local emits `env[idx]`, so it must receive the env array
            // rather than the individual values.
            if !enclosing_shared_env_names.is_empty()
                && !helper_captures.is_empty()
                && helper_captures
                    .iter()
                    .all(|name| enclosing_shared_env_names.contains(name))
            {
                helper_captures = vec!["__shared_env".to_string()];
            }
            self.scopes.pop();
            self.current = saved_cur;
            self.current_class = saved_class2;
            self.current_class_implicit_self = saved_implicit2;
            self.js_derived_ctor_ctx = saved_derived_ctx;
            self.shared_env_slot = saved_ctor_shared_env_slot;
            self.shared_env_names = saved_ctor_shared_env_names;
            self.closure_env_names = saved_ctor_closure_env_names;
            ctor_helpers.push((
                user_arity as usize,
                ctor_min_arity,
                helper_idx,
                helper_captures,
                ctor_variant.and_then(|c| c.named_name.clone()),
            ));
        }

        let ctor_idx = self.chunks.len();
        let ctor_arity = ctor_helpers
            .iter()
            .map(|(arity, _, _, _, _)| *arity)
            .max()
            .unwrap_or(0) as u8;
        self.chunks
            .push(common::functions::create_function_chunk(name, ctor_arity));
        self.scopes.push(Scope::new_function());
        let saved_cur = self.current;
        self.current = ctor_idx;
        for i in 0..ctor_arity {
            self.define_local(&format!("__ctor_arg_{}", i));
        }
        let line = self.line;

        // Abstract class: mark for compile-time check in `new` expressions.
        if class.is_abstract {
            self.abstract_classes.insert(self.canon(name));
        }

        let js_ctor_relaxes_min_arity = self.profile.relaxed_call_arity;
        let helper_for_count = |count: usize| {
            ctor_helpers
                .iter()
                .filter(|(arity, min_arity, _, _, _)| {
                    count <= *arity && (js_ctor_relaxes_min_arity || count >= *min_arity)
                })
                .min_by_key(|(arity, _, _, _, _)| *arity)
        };
        for count in (1..=ctor_arity as usize).rev() {
            self.emit_u16(Op::LOCAL_GET, (count - 1) as u16);
            // Argument-presence test: a *missing* trailing arg is `undefined`,
            // so `new C(null)` must still dispatch to the 1-arg constructor
            // (`null` is a value, not an absent argument). Languages without a
            // distinct `undefined` fall back to the null test.
            if self.profile.missing_arg_is_undefined {
                fn_call!(self, "wasm:js-undefined", "test", 1);
            } else {
                self.emit(Op::REF_IS_NULL);
            }
            // Result is already I32(0/1) — no dyn_to_bool needed
            self.emit(Op::I32_EQZ);
            self.chunks[self.current].emit_if(line);
            if let Some((_, _, helper_idx, helper_captures, _)) = helper_for_count(count) {
                emit_helper_ref(self, *helper_idx, helper_captures)?;
                for arg_index in 0..count {
                    self.emit_u16(Op::LOCAL_GET, arg_index as u16);
                }
                self.emit_u8(Op::CALL_REF, count as u8);
                self.emit_return_through_finally(1)?;
            }
            self.chunks[self.current].emit_end(line);
        }
        if let Some((_, _, helper_idx, helper_captures, _)) = helper_for_count(0) {
            emit_helper_ref(self, *helper_idx, helper_captures)?;
            self.emit_u8(Op::CALL_REF, 0);
        } else {
            self.emit(Op::NULL);
        }
        self.emit_return_through_finally(1)?;
        {
            let ns = self.scope().next_slot;
            self.chunks[ctor_idx].finalize_local_count(ns);
            self.chunks[ctor_idx].local_names = self.scope().defined_names.clone();
        }
        let ctor_upvalues = self.scope().upvalues.clone();
        self.scopes.pop();
        self.current = saved_cur;

        let ctor_local = self.define_local(&format!("__{}_ctor", name));
        let uv_pairs: Vec<(bool, u16)> = ctor_upvalues
            .iter()
            .map(|uv| (uv.is_local, uv.index))
            .collect();
        let case_sensitive = self.profile.ecma_new_dispatch;
        crate::primitives::classes::emit_store_constructor_with_upvalues(
            self.chunk(),
            name,
            ctor_idx,
            ctor_local,
            &uv_pairs,
            case_sensitive,
            line,
        );
        if !self.class_prototype_dispatch() {
            // Stamp the declared class name on the ctor function so
            // `get_class($x)` ($x.constructor.name) returns it. The
            // prototype branch below stamps `name` during prototype wiring;
            // languages that skip that block stamp here instead.
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_const(Value::String(Arc::from(name)));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);
        }
        for (arity, _, helper_idx, helper_captures, named) in &ctor_helpers {
            emit_helper_ref(self, *helper_idx, helper_captures)?;
            let helper_global = ctor_global_for(&ctor_global_prefix, *arity);
            let helper_idx_const = self.str_const(&helper_global);
            self.emit_u16(Op::GLOBAL_SET, helper_idx_const);
            // A named constructor (`Point.origin()`) is reached through the
            // class rather than by arity — several of them commonly share an
            // arity with each other and with the unnamed ctor. The helper
            // already allocates and returns the instance, so stamping it on
            // the class object makes `Point.origin(...)` an ordinary call,
            // the same shape a factory constructor already compiles to.
            if let Some(named) = named {
                self.emit_u16(Op::LOCAL_GET, ctor_local);
                emit_helper_ref(self, *helper_idx, helper_captures)?;
                let key = self.str_const(named);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);
            }
        }

        if self.class_prototype_dispatch() {
            self.emit_common("object.new", 0, line);
            let proto_local = self.define_local(&format!("__{}_prototype", name));
            self.emit_u16(Op::LOCAL_SET, proto_local);

            // Framework GUI control parents have no emitted ctor global, so
            // `emit_parent_ctor_value` yields null — skip prototype-chain
            // linking (a `STRUCT_GET prototype` on null would trap). Control
            // members resolve through the component descriptor, not a JS
            // prototype chain.
            let proto_parent = parent
                .as_deref()
                .filter(|p| !self.is_framework_control_parent(p));
            if let Some(parent_name) = proto_parent {
                self.emit_parent_class_value(parent_name);
                let parent_proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_GET, parent_proto_key);
                let parent_proto_local = self.define_local(&format!("__{}_parent_prototype", name));
                self.emit_u16(Op::LOCAL_SET, parent_proto_local);
                self.emit_u16(Op::LOCAL_GET, parent_proto_local);
                self.emit(Op::REF_IS_NULL);
                self.emit(Op::I32_EQZ);
                self.chunks[self.current].emit_if(line);
                self.emit_u16(Op::LOCAL_GET, proto_local);
                self.emit_u16(Op::LOCAL_GET, parent_proto_local);
                let proto_link_key = self.str_const("__proto__");
                self.emit_u16(Op::STRUCT_SET, proto_link_key);
                self.emit(Op::DROP);
                self.chunks[self.current].emit_end(line);
            }

            self.emit_u16(Op::LOCAL_GET, proto_local);
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            let ctor_key = self.str_const("constructor");
            self.emit_u16(Op::STRUCT_SET, ctor_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_const(Value::String(Arc::from("prototype")));
            common::dict::emit_new(&mut self.chunks, self.current, self.line);
            inst!(self, core_wasm::dup);
            self.emit_u16(Op::LOCAL_GET, proto_local);
            let value_key = self.str_const("value");
            self.emit_u16(Op::STRUCT_SET, value_key);
            self.emit(Op::DROP);
            for (flag, value) in [
                ("writable", false),
                ("enumerable", false),
                ("configurable", false),
            ] {
                inst!(self, core_wasm::dup);
                self.emit_const(Value::Bool(value));
                let flag_key = self.str_const(flag);
                self.emit_u16(Op::STRUCT_SET, flag_key);
                self.emit(Op::DROP);
            }
            let define_prop_idx = self.import("ecma:object", "defineProperty");
            self.emit_host_call(define_prop_idx, 3);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_const(Value::String(Arc::from(name)));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_const(Value::F64(0.0));
            let length_key = self.str_const("length");
            self.emit_u16(Op::STRUCT_SET, length_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);

            // §15.7.5 step 7 (JS): the class constructor's own
            // [[Prototype]] — the parent constructor for derived classes
            // (static inheritance walks it), %Function.prototype% for
            // base classes (C.bind / C.call / C.apply resolve through it).
            if self.class_prototype_dispatch() {
                // Control parents (null ctor global) get the base-class
                // %Function.prototype% link, not a null static-inheritance
                // link — matching how a control resolves as a leaf class.
                if let Some(parent_name) = proto_parent {
                    self.emit_u16(Op::LOCAL_GET, ctor_local);
                    self.emit_parent_class_value(parent_name);
                    let proto_link_key = self.str_const("__proto__");
                    self.emit_u16(Op::STRUCT_SET, proto_link_key);
                    self.emit(Op::DROP);
                    if self.profile.has_ecma_globals
                        && Self::is_ecma_typed_array_ctor_name(parent_name)
                    {
                        self.emit_u16(Op::LOCAL_GET, ctor_local);
                        self.emit_const(Value::Bool(true));
                        let marker_key = self.str_const("__vybe_typed_array_ctor");
                        self.emit_u16(Op::STRUCT_SET, marker_key);
                        self.emit(Op::DROP);
                        for static_name in ["from", "of"] {
                            self.emit_u16(Op::LOCAL_GET, ctor_local);
                            self.emit_parent_class_value(parent_name);
                            let static_key = self.str_const(static_name);
                            self.emit_u16(Op::STRUCT_GET, static_key);
                            self.emit_u16(Op::LOCAL_GET, ctor_local);
                            let bind_idx = self.import("ecma:function", "bind");
                            self.emit_host_call(bind_idx, 2);
                            self.emit_u16(Op::STRUCT_SET, static_key);
                            self.emit(Op::DROP);
                        }
                    }
                } else {
                    self.emit_u16(Op::LOCAL_GET, ctor_local);
                    crate::primitives::prototypes::emit_stamp_function_kind_proto(
                        self.chunk(),
                        false,
                        false,
                        line,
                    );
                }
            }

            // The prototype is the class's open method table: every
            // instance method lands on it, so `C.prototype.m` resolves
            // and reassignment has a real target. Capture-carrying
            // methods are skipped — their upvalues bind in the
            // constructor's frame, not at definition scope.
            for (mname, mci, _, _) in &instance_methods {
                // Bind EVERY instance method onto the prototype, including
                // capture-carrying ones (methods that use `super` capture
                // `__shared_env`). Previously these were skipped, which left a
                // derived class's own methods off its prototype — so a
                // grandchild's `super.m()` (which does an own-only `struct.get`
                // on the PARENT prototype) couldn't find a derived parent's
                // method and resolved to undefined (multi-level super broke at
                // 3+ levels). `__shared_env` is a local in this
                // (class-definition) scope, so its upvalues resolve here.
                let capture_names = method_capture_name_map
                    .get(mci)
                    .map(Vec::as_slice)
                    .unwrap_or(&[]);
                let (m_async, m_gen) = self
                    .method_fn_kinds
                    .get(mci)
                    .copied()
                    .unwrap_or((self.chunks[*mci].is_async, self.chunks[*mci].is_generator));
                self.emit_u16(Op::LOCAL_GET, proto_local);
                self.emit_ref_func_with_captures(*mci, capture_names, false)?;
                // ECMA-262 function-kind stamp: async/generator methods'
                // __proto__ is the matching intrinsic prototype (§27.7.1 /
                // §27.3.1 / §27.4.1) — `getPrototypeOf(C.prototype.m)`.
                if m_async || m_gen {
                    inst!(self, core_wasm::dup);
                    let line = self.line;
                    crate::primitives::prototypes::emit_stamp_function_kind_proto(
                        self.chunk(),
                        m_async,
                        m_gen,
                        line,
                    );
                }
                // §10.2.9 SetFunctionName: a class method's `name` is its
                // property key (non-enumerable, like all fn metadata).
                inst!(self, core_wasm::dup);
                self.emit_const(Value::String(Arc::from(mname.as_str())));
                let name_key = self.str_const("name");
                self.emit_u16(Op::STRUCT_SET, name_key);
                {
                    let line = self.line;
                    crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(
                        self.chunk(),
                        line,
                    );
                }
                let key = self.str_const(mname);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);
                // Publish the PROTOCOL SLOT alongside the method's own name, so
                // a prototype-dispatch language (JS, PHP, Dart) reaches its
                // roles through the same numeric key as a bind-dispatch one
                // (Python, Ruby). Both paths install methods, so both have to
                // stamp, or the slot exists in half the languages.
                //
                // `proto[slot] = proto[mname]`, emitted as its own sequence
                // rather than folded into the install above: `STRUCT_SET` pops
                // the VALUE and leaves the TARGET, so stamping mid-sequence
                // would have to push the funcref a second time — the earlier
                // shape dup'd the funcref and let it serve as both target and
                // value, which stamped `fn[slot] = fn` (a cycle on the method
                // object) and left the prototype without the slot entirely.
                if let Some(slot_key) = self.current_class_slot_keys.get(mname.as_str()).cloned() {
                    self.emit_u16(Op::LOCAL_GET, proto_local);
                    self.emit_u16(Op::LOCAL_GET, proto_local);
                    let method_key = self.str_const(mname);
                    self.emit_u16(Op::STRUCT_GET, method_key);
                    let slot_const = self.str_const(&slot_key);
                    self.emit_u16(Op::STRUCT_SET, slot_const);
                    self.emit(Op::DROP);
                }
            }
        }

        // Static field initializers run with the self-reference bound to the
        // class constructor object (ECMA-262 §15.7.10 — `this` inside a static
        // field initializer is the class itself), so `static y = this.x * 2`
        // can read sibling static fields. Bind the self-keyword to `ctor_local`
        // for the duration of the initializer emission.
        let saved_static_init_class = self.current_class.take();
        let saved_static_init_implicit = self.current_class_implicit_self;
        self.current_class = Some(name.to_string());
        self.current_class_implicit_self = class.implicit_self_fields;
        let static_self_kw = self.profile.self_keyword.clone();
        let static_self_slot = self.define_local(&static_self_kw);
        self.emit_u16(Op::LOCAL_GET, ctor_local);
        self.emit_u16(Op::LOCAL_SET, static_self_slot);

        let saved_static_js_this = if self.profile.ambient_this_binding {
            let saved = self.save_js_this("__js_prev_this_static_init");
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.set_js_this_from_stack();
            Some(saved)
        } else {
            None
        };

        // Initialize static fields on the constructor object
        for (fname, type_hint, init, array_bounds) in &static_field_inits {
            let saved_member_static = self.current_member_is_static;
            self.current_member_is_static = true;
            if self.profile.name == "js" && fname.starts_with("__static_block_") {
                self.emit_u16(Op::LOCAL_GET, ctor_local);
                if let Some(init_expr) = init {
                    self.compile_expr(init_expr)?;
                } else {
                    inst!(self, core_wasm::undefined);
                }
                self.current_member_is_static = saved_member_static;
                self.emit(Op::DROP);
                self.emit(Op::DROP);
                continue;
            }
            if self.profile.name == "js" && !fname.starts_with("__js_private_") {
                self.emit_u16(Op::LOCAL_GET, ctor_local);
                if let Some(init_expr) = init {
                    self.compile_expr(init_expr)?;
                } else if let Some(extent) = array_bounds
                    .as_deref()
                    .and_then(Self::array_bounds_extent_expr)
                {
                    let init_expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("Array")),
                        args: vec![
                            Argument::positional(extent),
                            Argument::positional(Self::array_default_element_expr(
                                type_hint.as_deref(),
                            )),
                        ],
                        optional: false,
                    });
                    self.compile_expr(&init_expr)?;
                } else {
                    inst!(self, core_wasm::undefined);
                }
                self.current_member_is_static = saved_member_static;
                let value_slot = self.define_local("__js_static_field_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, ctor_local);
                self.emit_const(Value::String(Arc::from(fname.as_str())));
                common::dict::emit_new(&mut self.chunks, self.current, self.line);
                inst!(self, core_wasm::dup);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let value_key = self.str_const("value");
                self.emit_u16(Op::STRUCT_SET, value_key);
                self.emit(Op::DROP);
                for (flag, value) in [
                    ("writable", true),
                    ("enumerable", true),
                    ("configurable", true),
                ] {
                    inst!(self, core_wasm::dup);
                    self.emit_const(Value::Bool(value));
                    let flag_key = self.str_const(flag);
                    self.emit_u16(Op::STRUCT_SET, flag_key);
                    self.emit(Op::DROP);
                }
                let define_prop_idx = self.import("ecma:object", "defineProperty");
                self.emit_host_call(define_prop_idx, 3);
                self.emit(Op::DROP);
                continue;
            }
            self.emit_class_field_initializer(
                ctor_local,
                fname,
                type_hint.as_deref(),
                init.as_ref(),
                array_bounds.as_deref(),
                class.is_value_type,
                self.line,
            )?;
            self.current_member_is_static = saved_member_static;
        }

        if let Some(saved) = saved_static_js_this {
            self.restore_js_this(saved);
        }
        self.current_class = saved_static_init_class;
        self.current_class_implicit_self = saved_static_init_implicit;

        for const_name in &static_const_names {
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            let global_name = self.canon(&format!("{}.{}", name, const_name));
            let global_idx = self.str_const(&global_name);
            self.emit_u16(Op::GLOBAL_GET, global_idx);
            let field_idx = self.str_const(const_name);
            self.emit_u16(Op::STRUCT_SET, field_idx);
            self.emit(Op::DROP);
        }

        let own_static_member_names: Vec<String> = static_field_inits
            .iter()
            .map(|(fname, _, _, _)| fname.clone())
            .chain(static_const_names.iter().cloned())
            .collect();

        // Inherit static fields/constants onto the child class object so
        // PHP late static binding (`static::$x`, `static::NAME`) resolves
        // against the called class instead of stopping at the declaring class.
        if let Some(parent_name) = parent {
            let mut current_parent = Some(self.canon(parent_name));
            while let Some(ref pname) = current_parent {
                let parent_static_fields = self
                    .pending_classes
                    .get(pname.as_str())
                    .map(|pc| pc.static_fields.clone())
                    .unwrap_or_default();
                let next_parent = self
                    .pending_classes
                    .get(pname.as_str())
                    .and_then(|pc| pc.parent.clone());
                for field_name in &parent_static_fields {
                    if self.profile.supports_private_fields
                        && field_name.starts_with("__js_private_")
                    {
                        continue;
                    }
                    if own_static_member_names
                        .iter()
                        .any(|name| name == field_name)
                    {
                        continue;
                    }
                    self.emit_u16(Op::LOCAL_GET, ctor_local);
                    let parent_idx = self.str_const(pname);
                    self.emit_u16(Op::GLOBAL_GET, parent_idx);
                    let field_idx = self.str_const(field_name);
                    self.emit_u16(Op::STRUCT_GET, field_idx);
                    self.emit_u16(Op::STRUCT_SET, field_idx);
                    self.emit(Op::DROP);
                }
                current_parent = next_parent;
            }
        }

        // Attach nested types onto the constructor object so `Outer.Inner`
        // resolves to the nested class constructor through the same shared
        // class-object path as static methods.
        let nested_types = self
            .pending_classes
            .get(name)
            .map(|pc| pc.nested_types.clone())
            .unwrap_or_default();
        for nested in nested_types {
            let nested_canon = self.canon(&nested);
            let nested_idx = self.str_const(&nested_canon);
            for key_name in [
                nested_canon.clone(),
                self.canon(nested.rsplit('.').next().unwrap_or(&nested)),
            ] {
                self.emit_u16(Op::LOCAL_GET, ctor_local);
                self.emit_u16(Op::GLOBAL_GET, nested_idx);
                let key = self.str_const(&key_name);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);
            }
        }

        // Attach static methods to the constructor object
        let mut all_statics: Vec<(String, usize)> = Vec::new();
        let php_static_receiver = if self.profile.name == "php" {
            Some(ctor_local)
        } else {
            None
        };
        for (mname, mci, _, _) in &static_methods {
            if self.profile.supports_private_fields && mname.starts_with("__js_private_") {
                continue;
            }
            let (m_async, m_gen) = self
                .method_fn_kinds
                .get(mci)
                .copied()
                .unwrap_or((self.chunks[*mci].is_async, self.chunks[*mci].is_generator));
            crate::primitives::classes::emit_attach_static_method_kinded(
                self.chunk(),
                ctor_local,
                mname,
                *mci,
                php_static_receiver,
                method_rest_fixed_count(*mci),
                m_async,
                m_gen,
                line,
            );
            all_statics.push((mname.clone(), *mci));
        }

        if self.profile.supports_private_fields {
            for method in &class.static_methods {
                if !method.source_name.starts_with('#') {
                    continue;
                }
                let bound_name =
                    self.js_member_storage_name_for_class(&class.name, &method.source_name);
                if let Some((_, chunk_idx, _, _)) =
                    method_chunks.iter().find(|(name, _, is_ctor, is_static)| {
                        !*is_ctor && *is_static && name == &bound_name
                    })
                {
                    self.emit_u16(Op::LOCAL_GET, ctor_local);
                    self.emit_u16(Op::REF_FUNC, *chunk_idx as u16);
                    self.chunk().emit(0, line);
                    if let Some(receiver_slot) = php_static_receiver {
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::LOCAL_GET, receiver_slot);
                        let receiver_key = self.str_const("__vybe_method_receiver");
                        self.emit_u16(Op::STRUCT_SET, receiver_key);
                        self.emit(Op::DROP);
                    }
                    if let Some(fixed_count) = method_rest_fixed_count(*chunk_idx) {
                        inst!(self, core_wasm::dup);
                        self.emit_const(Value::F64(fixed_count as f64));
                        let rest_key = self.str_const("__vybe_rest_fixed_arity");
                        self.emit_u16(Op::STRUCT_SET, rest_key);
                        self.emit(Op::DROP);
                    }
                    inst!(self, core_wasm::dup);
                    self.emit_const(Value::String(Arc::from(method.source_name.as_str())));
                    let name_key = self.str_const("name");
                    self.emit_u16(Op::STRUCT_SET, name_key);
                    crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(
                        self.chunk(),
                        line,
                    );
                    let storage_key = self.str_const(&bound_name);
                    self.emit_u16(Op::STRUCT_SET, storage_key);
                    self.emit(Op::DROP);
                }
            }
        }

        // Synthetic static constructor hook from language walkers.
        if let Some((_, static_init_ci, _, _)) = static_methods
            .iter()
            .find(|(mname, _, _, _)| mname.eq_ignore_ascii_case("__static_init__"))
        {
            let line = self.line;
            let saved_js_this = self.save_js_this("__js_prev_static_init_this");
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.set_js_this_from_stack();
            self.emit_u16(Op::REF_FUNC, *static_init_ci as u16);
            self.chunk().emit(0, line);
            self.emit_u8(Op::CALL_REF, 0);
            let result_slot = self.define_local("__js_static_init_result");
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.restore_js_this(saved_js_this);
            self.emit_u16(Op::LOCAL_GET, result_slot);
            self.emit(Op::DROP);
        }

        // Inherit parent's static methods — walk up the chain via PendingClass
        if let Some(parent_name) = parent {
            let mut current_parent = Some(self.canon(parent_name));
            while let Some(ref pname) = current_parent {
                let parent_statics = self
                    .pending_classes
                    .get(pname.as_str())
                    .map(|pc| pc.statics.clone())
                    .unwrap_or_default();
                let next_parent = self
                    .pending_classes
                    .get(pname.as_str())
                    .and_then(|pc| pc.parent.clone());
                for (sname, sci) in &parent_statics {
                    // Only inherit if child doesn't already define it
                    if !all_statics.iter().any(|(n, _)| n == sname) {
                        crate::primitives::classes::emit_attach_static_method(
                            self.chunk(),
                            ctor_local,
                            sname,
                            *sci,
                            php_static_receiver,
                            method_rest_fixed_count(*sci),
                            line,
                        );
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

        // Attach instance methods/accessors to the class object so static
        // `super.method()` and `super.prop` dispatch can reach them. ECMA-262
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
            crate::primitives::classes::emit_attach_static_method(
                self.chunk(),
                ctor_local,
                mname,
                *mci,
                None,
                method_rest_fixed_count(*mci),
                line,
            );
        }

        // What the source actually declared — `interface` / `trait` / `mixin` /
        // `module` all arrive as one `ClassDecl`, so this is the only thing that
        // tells them apart. A value type still reports `struct` when its
        // language did not say otherwise.
        let declared_kind = match class.declared_kind {
            vybe_ast::ClassKind::Class if class.is_value_type => {
                crate::primitives::reflection::ReflectKind::Struct
            }
            kind => crate::primitives::reflection::ReflectKind::from_class_kind(kind),
        };

        // The declared-kind annotation is stamped unconditionally: it is what
        // `interface_exists` / `trait_exists` / `kind_of?` read, and every
        // language with those needs it. Deliberately NOT inside the
        // `class_introspection_metadata` opt-in below — that gate is for
        // Python's `__name__` / `__mro__`, and hiding the kind behind it is
        // what forced languages into compile-time side tables instead.
        crate::primitives::object::stamp_local_string_field(
            self.chunk(),
            ctor_local,
            crate::primitives::reflection::FIELD_KIND,
            declared_kind.as_str(),
            line,
        );

        // The class's own member list, for languages whose reflection surface
        // reads it at runtime rather than deriving it at compile time.
        if self.profile.class_member_metadata {
            let fields: Vec<(String, Option<String>, i64)> = class
                .instance_fields
                .iter()
                .chain(class.static_fields.iter())
                .map(|f| (f.name.clone(), f.type_hint.clone(), 0))
                .collect();
            let methods: Vec<(String, usize, Option<String>, Vec<String>, i64)> = class
                .instance_methods
                .iter()
                .chain(class.static_methods.iter())
                .map(|m| {
                    (
                        m.source_name.clone(),
                        m.params.len(),
                        m.return_type.clone(),
                        m.params
                            .iter()
                            .map(|p| p.type_hint.clone().unwrap_or_default())
                            .collect(),
                        0,
                    )
                })
                .collect();
            crate::primitives::classes::emit_stamp_class_members(
                &mut self.chunks,
                self.current,
                ctor_local,
                name,
                &fields,
                &methods,
                line,
            );
        }

        // Class-introspection metadata on the class object: `__name__` (own
        // name) and `__mro__` (self → bases → `object`). Universal — every
        // class gets it, keyed on class construction, no language check.
        if self.profile.class_introspection_metadata {
            crate::primitives::classes::emit_stamp_class_name(
                self.chunk(),
                ctor_local,
                name,
                declared_kind,
                line,
            );
            // Under multiple inheritance, `__mro__` follows the full C3
            // linearization and `__bases__` lists every declared base; single
            // inheritance keeps the one-parent chain unchanged.
            let mi = self.profile.class_multiple_inheritance && class.bases.len() > 1;
            let (mro_globals, bases_globals): (Vec<String>, Vec<String>) = if mi {
                let mut mro = self.c3_linearize(name); // [self, …ancestors…]
                mro.remove(0); // helper re-adds self + the `object` tail
                let bases = class.bases.iter().map(|b| self.canon(b)).collect();
                (mro, bases)
            } else {
                let bg: Vec<String> = parent.as_ref().map(|p| vec![p.clone()]).unwrap_or_default();
                (bg.clone(), bg)
            };
            crate::primitives::classes::emit_stamp_class_mro(
                &mut self.chunks,
                self.current,
                ctor_local,
                &mro_globals,
                line,
            );
            crate::primitives::classes::emit_stamp_class_bases(
                &mut self.chunks,
                self.current,
                ctor_local,
                &bases_globals,
                line,
            );
        }

        // Prototype-dispatch profiles resolve instance members through the
        // prototype chain — statics live on the constructor object only
        // (§15.7: `instance.staticMethod` is undefined), so keep them out
        // of the type table's instance-method fallback. Other dispatch
        // models (VB-style `Instance.SharedMethod`) keep the full list.
        let all_methods: Vec<(String, usize)> = method_chunks
            .iter()
            .filter(|(_, _, _, is_static)| !self.class_prototype_dispatch() || !*is_static)
            .map(|(n, c, _, _)| (n.clone(), *c))
            .collect();
        // Canonicalise per language case-sensitivity: case-insensitive
        // languages (VB/Pascal/COBOL/PHP) lowercase here, case-sensitive
        // (JS/TS/Python/C#) preserve. Registry stores whatever the walker
        // produced; runtime `Op::REF_TEST` looks up by the same canon.
        let canon_name = ctor_global_prefix;
        let canon_parent = parent.as_ref().map(|p| self.canon(p)).unwrap_or_default();
        crate::primitives::classes::register_type(
            &mut self.chunks,
            &canon_name,
            &canon_parent,
            fields,
            all_methods,
            false,
            Vec::new(),
            Some(ctor_idx),
            std::collections::HashMap::new(),
        );

        Ok(())
    }
}

/// C3 merge — the linearization step shared by every MI language. Takes the
/// parent linearizations plus the direct-base list and interleaves them so a
/// class always precedes its parents and relative order is preserved. On an
/// inconsistent hierarchy (no valid head) it falls back to taking the first
/// remaining head, so compilation never panics.
fn c3_merge(mut seqs: Vec<Vec<String>>) -> Vec<String> {
    let mut result = Vec::new();
    loop {
        seqs.retain(|s| !s.is_empty());
        if seqs.is_empty() {
            return result;
        }
        // A valid head appears at the front of some sequence and in the tail of
        // none. Take the first such; otherwise fall back to the first head.
        let head = seqs
            .iter()
            .map(|s| s[0].clone())
            .find(|cand| !seqs.iter().any(|s| s[1..].contains(cand)))
            .unwrap_or_else(|| seqs[0][0].clone());
        if !result.contains(&head) {
            result.push(head.clone());
        }
        for seq in &mut seqs {
            if seq.first() == Some(&head) {
                seq.remove(0);
            }
        }
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
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            body_has_result_member_assign(then_body)
                || elifs
                    .iter()
                    .any(|(_, body)| body_has_result_member_assign(body))
                || else_body
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::For { init, body, .. } => {
            init.as_ref()
                .is_some_and(|stmt| stmt_has_result_member_assign(stmt))
                || body_has_result_member_assign(body)
        }
        StmtKind::ForIn {
            body, else_body, ..
        }
        | StmtKind::While {
            body, else_body, ..
        } => {
            body_has_result_member_assign(body)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::DoWhile { body, .. } => body_has_result_member_assign(body),
        StmtKind::Switch { cases, default, .. } => {
            cases
                .iter()
                .any(|case| body_has_result_member_assign(&case.body))
                || default
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
            ..
        } => {
            body_has_result_member_assign(body)
                || catches
                    .iter()
                    .any(|catch| body_has_result_member_assign(&catch.body))
                || else_body
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
                || finally
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
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

// ═══════════════════════════════════════════════════════════════════
// Shared class emit helpers — moved here from `crate::emitter/src/classes.rs`
// so that classes live in ONE place. See flexclassplan.md §4a-bis.
// ═══════════════════════════════════════════════════════════════════

// ── Object creation ─────────────────────────────────────────────────────

fn stamp_reflection_type_fields(
    chunk: &mut Chunk,
    slot: u16,
    type_name: &str,
    kind: crate::primitives::reflection::ReflectKind,
    line: u32,
) {
    crate::primitives::object::stamp_local_string_field(
        chunk,
        slot,
        crate::primitives::reflection::FIELD_TYPE,
        type_name,
        line,
    );
    crate::primitives::object::stamp_local_string_field(
        chunk,
        slot,
        crate::primitives::reflection::FIELD_TYPE_NAME,
        type_name,
        line,
    );
    crate::primitives::object::stamp_local_string_field(
        chunk,
        slot,
        crate::primitives::reflection::FIELD_KIND,
        kind.as_str(),
        line,
    );
}

/// Create a new empty object and stamp it with type info.
/// Emits: struct_new 0 → local, __type string stamp, __control_name stamp,
/// set_type_id via __tid_ global.
///
/// Stack: unchanged (object stored in this_slot)
///
/// `__control_name` is set to the lowercased class name. For form classes
/// (a user `Class Form1` in any framework — WinForms, MAUI, etc.) this is
/// the key the GUI host's property registry uses, so `Me.Text = "X"` ends
/// up under `("form1", "text")` and `gui.get_property("form1", "text")`
/// reflects the assignment. For non-form classes the field is dead metadata
/// that nothing reads. Stamping it unconditionally keeps the compiler and
/// the resolver from having to detect "is this class a form?" — the same
/// canonical AST and bytecode shape works for both.
pub fn emit_new_typed_object(chunk: &mut Chunk, this_slot: u16, class_name: &str, line: u32) {
    // Create empty object → store in this_slot
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    // Stamp shared reflection/class type metadata.
    stamp_reflection_type_fields(
        chunk,
        this_slot,
        class_name,
        crate::primitives::reflection::ReflectKind::Object,
        line,
    );

    // Stamp __control_name = lowercased class name (canonical control identity).
    crate::primitives::object::stamp_local_string_field(
        chunk,
        this_slot,
        "__control_name",
        &class_name.to_lowercase(),
        line,
    );

    // Stamp WASM GC type_id via __tid_ global. The caller has already
    // canonicalised `class_name` per the source language's case-
    // sensitivity, and `register_type` stored the type under that
    // same name — VM `load_type_table` populates `__tid_<canon>`,
    // which we look up verbatim here.
    let tid_name = chunk.add_constant(Value::String(Arc::from(
        format!("__tid_{}", class_name).as_str(),
    )));
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::GLOBAL_GET, tid_name, line);
    {
        let tid_key = chunk.add_constant(Value::String(Arc::from(
            crate::primitives::reflection::FIELD_TYPE_ID,
        )));
        chunk.emit_op_u16(Op::STRUCT_SET, tid_key, line);
    }
    chunk.emit_op(Op::DROP, line);
}

/// Re-stamp type identity on an EXISTING object — a child constructor
/// receives `this` from the parent ctor call carrying the PARENT's
/// identity, so the child must overwrite `__type` and the WASM GC
/// type_id with its own (otherwise instanceof/REF_TEST and
/// constructorOf resolve to the parent class). Same stamps as
/// `emit_new_typed_object` minus the allocation. `class_name` must be
/// canonicalised like there.
pub fn emit_retype_object(chunk: &mut Chunk, this_slot: u16, class_name: &str, line: u32) {
    stamp_reflection_type_fields(
        chunk,
        this_slot,
        class_name,
        crate::primitives::reflection::ReflectKind::Object,
        line,
    );

    let tid_name = chunk.add_constant(Value::String(Arc::from(
        format!("__tid_{}", class_name).as_str(),
    )));
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::GLOBAL_GET, tid_name, line);
    {
        let tid_key = chunk.add_constant(Value::String(Arc::from(
            crate::primitives::reflection::FIELD_TYPE_ID,
        )));
        chunk.emit_op_u16(Op::STRUCT_SET, tid_key, line);
    }
    chunk.emit_op(Op::DROP, line);
}

/// Mark a shared class instance as a structural value type. The normal class
/// model sets this for language structs; equality adapters can then compare
/// same-shaped values by contents without special-casing VB/C#/future .NET
/// frontends.
pub fn emit_value_equality_stamp(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_bool_const(true, line);
    let key = chunk.add_constant(Value::String(Arc::from("__value_eq")));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `isinstance(obj, "Class")` / `obj.is_a?(Class)` — the READ side of the type
/// stamp, inheritance-aware. Checks membership in the `__types` ancestry array
/// stamped by [`emit_instanceof_chain`] (so `isinstance(subclass, Parent)` and
/// interfaces work); falls back to the single `__type` for objects without a
/// `__types` chain. Shared by every stamped language (retires the per-language
/// `__vybe_instanceof` chunk and Ruby's byte-identical inline).
/// Stack: `[obj, class_name]` → `[bool]`.
pub fn emit_instanceof(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(3);
    let (obj_s, klass_s, types_s) = (base, base + 1, base + 2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, klass_s, line); // [obj]
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_s, line); // []
    // types = obj["__types"]
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_string_const(crate::primitives::reflection::FIELD_TYPES, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, types_s, line);
    // if types present → types.includes(class_name); else obj["__type"] == class_name
    chunks[current].emit_op_u16(Op::LOCAL_GET, types_s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line); // 1 when __types is present
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, types_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, klass_s, line);
    crate::primitives::collections::emit_contains(chunks, current, line); // __types.includes(class_name)
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_string_const(crate::primitives::reflection::FIELD_TYPE, line);
    crate::primitives::collections::emit_get(chunks, current, line); // obj["__type"]
    chunks[current].emit_op_u16(Op::LOCAL_GET, klass_s, line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

// ── Method binding ──────────────────────────────────────────────────────

/// Stamp `__name__` (the class's own name) and the declared kind on the
/// class/constructor object so `Cls.__name__` and `type(obj).__name__` resolve,
/// and so `interface_exists` / `trait_exists` can be answered at runtime from
/// the annotation rather than a compile-time per-language table.
///
/// `kind` comes from `ClassModifiers::kind` — `class` / `interface` / `trait` /
/// `mixin` / `module` / `struct` all reach here as one `ClassDecl`.
/// Stack: unchanged.
pub fn emit_stamp_class_name(
    chunk: &mut Chunk,
    ctor_slot: u16,
    class_name: &str,
    kind: crate::primitives::reflection::ReflectKind,
    line: u32,
) {
    crate::primitives::object::stamp_local_string_field(
        chunk, ctor_slot, "__name__", class_name, line,
    );
    stamp_reflection_type_fields(chunk, ctor_slot, class_name, kind, line);
}

/// Stamp `__fields` and `__methods` on the class object: this class's own
/// members as reflection member tokens.
///
/// Each token is the same 8-element shape `reflection::member_token_expr`
/// builds — `[kind, owner, name, param_count, type_name, return_type,
/// param_types, modifiers]` — so a consumer can read it with
/// `reflection::member_token` without knowing which language produced it.
///
/// This is the runtime source that did not previously exist. Languages had to
/// derive their member lists at compile time (Pascal's `PascalRttiMetadata`
/// HashMap, Dart's `reflection_adapter`), which cannot answer for a type the
/// walk never saw — an autoloaded class, an `eval`'d one, or one from another
/// compilation unit. Stack: unchanged.
pub fn emit_stamp_class_members(
    chunks: &mut [Chunk],
    current: usize,
    ctor_slot: u16,
    class_name: &str,
    fields: &[(String, Option<String>, i64)],
    methods: &[(String, usize, Option<String>, Vec<String>, i64)],
    line: u32,
) {
    use crate::primitives::reflection;

    // One member token: an 8-element array, all elements compile-time known.
    let push_token = |chunks: &mut [Chunk],
                      kind: &str,
                      name: &str,
                      param_count: usize,
                      type_name: &Option<String>,
                      return_type: &Option<String>,
                      param_types: &[String],
                      modifiers: i64| {
        chunks[current].emit_string_const(kind, line);
        chunks[current].emit_string_const(class_name, line);
        chunks[current].emit_string_const(name, line);
        chunks[current].emit_i32_const(param_count as i32, line);
        match type_name {
            Some(t) => chunks[current].emit_string_const(t, line),
            None => chunks[current].emit_op(Op::NULL, line),
        }
        match return_type {
            Some(t) => chunks[current].emit_string_const(t, line),
            None => chunks[current].emit_op(Op::NULL, line),
        }
        for param_type in param_types {
            chunks[current].emit_string_const(param_type, line);
        }
        chunks[current].emit_array_new_fixed(0, param_types.len() as u16, line);
        chunks[current].emit_i32_const(modifiers as i32, line);
        chunks[current].emit_array_new_fixed(0, 8, line);
    };

    for (name, type_name, modifiers) in fields {
        push_token(
            chunks,
            reflection::MEMBER_KIND_FIELD,
            name,
            0,
            type_name,
            &None,
            &[],
            *modifiers,
        );
    }
    chunks[current].emit_array_new_fixed(0, fields.len() as u16, line);
    let fields_key =
        chunks[current].add_constant(Value::String(Arc::from(reflection::FIELD_FIELDS)));
    let fields_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fields_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ctor_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fields_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, fields_key, line);
    chunks[current].emit_op(Op::DROP, line);

    for (name, param_count, return_type, param_types, modifiers) in methods {
        push_token(
            chunks,
            reflection::MEMBER_KIND_METHOD,
            name,
            *param_count,
            &None,
            return_type,
            param_types,
            *modifiers,
        );
    }
    chunks[current].emit_array_new_fixed(0, methods.len() as u16, line);
    let methods_key =
        chunks[current].add_constant(Value::String(Arc::from(reflection::FIELD_METHODS)));
    let methods_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, methods_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ctor_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, methods_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, methods_key, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Stamp `__mro__` on the class object: an array of the ancestor class objects
/// (self first, then each base up the chain, ending with a synthetic `object`).
/// `self_ctor_slot` holds this class; `base_globals` are the canonical global
/// names of the ancestor classes in method-resolution order (excluding self and
/// `object`). Each entry is loaded by global name; a synthetic `{__name__:
/// "object"}` terminates the list so `[c.__name__ for c in Cls.__mro__]` works.
/// Stack: unchanged.
pub fn emit_stamp_class_mro(
    chunks: &mut [Chunk],
    current: usize,
    self_ctor_slot: u16,
    base_globals: &[String],
    line: u32,
) {
    // Build the MRO array: [self, base0, base1, …, object].
    chunks[current].emit_op_u16(Op::LOCAL_GET, self_ctor_slot, line);
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    // push self
    push_array_value_local(chunks, current, self_ctor_slot, line);
    // push each base by global name
    for g in base_globals {
        let gk = chunks[current].add_constant(Value::String(Arc::from(g.as_str())));
        chunks[current].emit_dup(line); // [ctor, arr, arr]
        chunks[current].emit_op_u16(Op::GLOBAL_GET, gk, line); // [ctor, arr, arr, base]
        crate::primitives::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    // push synthetic `object` (a small stand-in carrying just __name__)
    chunks[current].emit_dup(line); // [ctor, arr, arr]
    emit_object_base_stub(&mut chunks[current], line); // [ctor, arr, arr, objstub]
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // ctor.__mro__ = arr
    let key = chunks[current].add_constant(Value::String(Arc::from("__mro__")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, key, line); // [ctor]
    chunks[current].emit_op(Op::DROP, line);
}

/// Stamp `__bases__` on the class object: an array of the DIRECT parent class
/// objects (no self, no full MRO). `base_globals` are the canonical global names
/// of the immediate bases. A class with no explicit base gets `[object]` (the
/// synthetic stub), matching Python's `C.__bases__ == (object,)`. Stack: unchanged.
pub fn emit_stamp_class_bases(
    chunks: &mut [Chunk],
    current: usize,
    self_ctor_slot: u16,
    base_globals: &[String],
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, self_ctor_slot, line); // [ctor]
    crate::primitives::collections::emit_array_new(chunks, current, 0, line); // [ctor, arr]
    if base_globals.is_empty() {
        chunks[current].emit_dup(line);
        emit_object_base_stub(&mut chunks[current], line);
        crate::primitives::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    } else {
        for g in base_globals {
            let gk = chunks[current].add_constant(Value::String(Arc::from(g.as_str())));
            chunks[current].emit_dup(line);
            chunks[current].emit_op_u16(Op::GLOBAL_GET, gk, line);
            crate::primitives::collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    let key = chunks[current].add_constant(Value::String(Arc::from("__bases__")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, key, line); // [ctor]
    chunks[current].emit_op(Op::DROP, line);
}

const SUPER_LOOKUP_CHUNK: &str = "__mi_super_lookup";

/// Ensure the shared cooperative-`super()` lookup chunk exists; return its index.
///
/// Signature `(self, className, methodName) -> fn | null`. Walks
/// `self.__class__.__mro__` (the full C3 linearization stamped by
/// `emit_stamp_class_mro`), finds the class whose `__name__` equals `className`,
/// and returns the first `methodName` found on a class AFTER it in the MRO — the
/// cooperative next method for multiple inheritance. Generic: uses only the
/// `__class__`/`__mro__`/`__name__` introspection stamps, so every MI language
/// shares one implementation. Returns `null` when the receiver is untyped or no
/// later class defines the method.
pub fn ensure_super_lookup_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == SUPER_LOOKUP_CHUNK) {
        return idx;
    }
    let mut c = crate::primitives::functions::create_function_chunk(SUPER_LOOKUP_CHUNK, 3);
    c.alloc_scratch(3); // arg slots 0=self, 1=className, 2=methodName
    let self_a = 0u16;
    let class_a = 1u16;
    let method_a = 2u16;
    let mro = c.alloc_scratch(1);
    let n = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let passed = c.alloc_scratch(1);
    let elem = c.alloc_scratch(1);
    let m = c.alloc_scratch(1);
    let found = c.alloc_scratch(1);

    let obj_get = c.add_import("ecma:object", "get");
    let arr_get = c.add_import("ecma:array", "get");
    let arr_len = c.add_import("ecma:array", "length");
    let to_f64 = c.add_import("wasm:js-number", "toF64");
    let from_f64 = c.add_import("wasm:js-number", "fromF64");

    // cls = self.__class__; return null if absent
    c.emit_op_u16(Op::LOCAL_GET, self_a, line);
    c.emit_string_const("__class__", line);
    c.emit_call(obj_get, 2, line);
    c.emit_dup(line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    // mro = cls.__mro__; return null if absent
    c.emit_string_const("__mro__", line);
    c.emit_call(obj_get, 2, line);
    c.emit_dup(line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_SET, mro, line);

    // n = len(mro); i = 0; passed = 0; found = null
    // `__mro__` is a growable host array (built with ecma:array.push), so use
    // ecma:array access — GC `array.get`/`array.len` read only fixed GC arrays.
    c.emit_op_u16(Op::LOCAL_GET, mro, line);
    c.emit_call(arr_len, 1, line);
    c.emit_call(to_f64, 1, line);
    c.emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    c.emit_op_u16(Op::LOCAL_SET, n, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, passed, line);
    c.emit_op(Op::NULL, line);
    c.emit_op_u16(Op::LOCAL_SET, found, line);

    let block_patch = c.emit_block(line);
    let (loop_patch, _) = c.emit_loop_s(line);
    // if i >= n: break out of block
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    // elem = mro[i]  (host-array element access; index boxed to a number value)
    c.emit_op_u16(Op::LOCAL_GET, mro, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::F64_FROM_I32, line);
    c.emit_call(from_f64, 1, line);
    c.emit_call(arr_get, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, elem, line);
    // if passed == 0 { look for the current class } else { look for the method }
    c.emit_op_u16(Op::LOCAL_GET, passed, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    {
        c.emit_op_u16(Op::LOCAL_GET, elem, line);
        c.emit_string_const("__name__", line);
        c.emit_call(obj_get, 2, line);
        c.emit_op_u16(Op::LOCAL_GET, class_a, line);
        crate::primitives::ops::emit_dyn_eq(&mut c, line);
        c.emit_if(line);
        c.emit_i32_const(1, line);
        c.emit_op_u16(Op::LOCAL_SET, passed, line);
        c.emit_end(line);
    }
    c.emit_else(line);
    {
        // m = elem[methodName]; if m != null && found == null: found = m
        c.emit_op_u16(Op::LOCAL_GET, elem, line);
        c.emit_op_u16(Op::LOCAL_GET, method_a, line);
        c.emit_call(obj_get, 2, line);
        c.emit_op_u16(Op::LOCAL_SET, m, line);
        c.emit_op_u16(Op::LOCAL_GET, m, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_op(Op::I32_EQZ, line); // 1 if m not null
        c.emit_op_u16(Op::LOCAL_GET, found, line);
        c.emit_op(Op::REF_IS_NULL, line); // 1 if found null
        c.emit_op(Op::I32_AND, line);
        c.emit_if(line);
        c.emit_op_u16(Op::LOCAL_GET, m, line);
        c.emit_op_u16(Op::LOCAL_SET, found, line);
        c.emit_end(line);
    }
    c.emit_end(line); // end if passed==0
    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_br(0, line);
    c.emit_end(line); // end loop
    c.patch_loop(loop_patch);
    c.emit_end(line); // end block
    c.patch_block(block_patch);

    c.emit_op_u16(Op::LOCAL_GET, found, line);
    c.emit_op(Op::RETURN, line);

    let idx = chunks.len();
    chunks.push(c);
    idx
}

/// `[ctor, arr]` → push `LOCAL_GET(slot)` onto `arr`, preserving `[ctor, arr]`.
fn push_array_value_local(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Push a minimal `object` class stand-in carrying `__name__ = "object"`, for the
/// tail of every class's `__mro__`. Stack: `[] -> [obj]`.
fn emit_object_base_stub(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    crate::primitives::object::set_string_field(chunk, "__name__", "object", line);
    chunk.emit_dup(line);
    crate::primitives::object::set_string_field(
        chunk,
        crate::primitives::reflection::FIELD_TYPE,
        "object",
        line,
    );
    chunk.emit_dup(line);
    crate::primitives::object::set_string_field(
        chunk,
        crate::primitives::reflection::FIELD_TYPE_NAME,
        "object",
        line,
    );
    chunk.emit_dup(line);
    crate::primitives::object::set_string_field(
        chunk,
        crate::primitives::reflection::FIELD_KIND,
        crate::primitives::reflection::ReflectKind::Class.as_str(),
        line,
    );
}

// ── Super call (cross-language) ────────────────────────────────────────

/// After calling the parent constructor (result on TOS), store it as `this` and
/// save any parent methods that the child will override.
///
/// The compiler handles the actual call: global_get(parent) → push args → call_ref(argc).
/// This helper stores the result and prepares for child method override.
///
/// Stack before: [parent_return_value]  Stack after: []
pub fn emit_super_call_store_result(
    chunk: &mut Chunk,
    this_slot: u16,
    child_method_names: &[&str],
    line: u32,
) {
    // Store parent-created object as this
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    // Save parent's methods that child will override (for super.method() calls)
    for method_name in child_method_names {
        emit_save_base_method(chunk, this_slot, method_name, line);
    }
}

// ── Inheritance ─────────────────────────────────────────────────────────

/// Save parent's version of a method as __base_<name> before child override.
/// Used for super()/MyBase/base calls.
/// Emits: local_get this → local_get this → struct_get name → struct_set __base_name → drop
/// Stack: unchanged
pub fn emit_save_base_method(chunk: &mut Chunk, this_slot: u16, method_name: &str, line: u32) {
    let base_name = format!("__base_{}", method_name);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // obj for struct_set
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // obj for struct_get
    let prop_idx = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::STRUCT_GET, prop_idx, line); // val = this.method (parent version)
    let base_idx = chunk.add_constant(Value::String(Arc::from(base_name.as_str())));
    chunk.emit_op_u16(Op::STRUCT_SET, base_idx, line); // this.__base_method = val
    chunk.emit_op(Op::DROP, line);
}

/// Store parent constructor ref as __super on the instance.
/// Stack: unchanged
pub fn emit_store_super(chunk: &mut Chunk, this_slot: u16, parent_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let parent_c = chunk.add_constant(Value::String(Arc::from(parent_name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, parent_c, line);
    let super_key = chunk.add_constant(Value::String(Arc::from("__super")));
    chunk.emit_op_u16(Op::STRUCT_SET, super_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Inherit static methods from parent constructor via Object.assign.
/// Caller must have the constructor on TOS (typically via dup before this call).
/// Stack before: [constructor]  Stack after: [constructor]
pub fn emit_inherit_statics(chunk: &mut Chunk, parent_name: &str, line: u32) {
    chunk.emit_dup(line);
    let parent_c = chunk.add_constant(Value::String(Arc::from(parent_name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, parent_c, line);
    let assign_fn = chunk.add_import("ecma:object", "assign");
    chunk.emit_call(assign_fn, 2, line);
    chunk.emit_op(Op::DROP, line);
}

// ── Static methods ──────────────────────────────────────────────────────

/// Attach a static method to the constructor function object.
/// Same pattern as VB Shared, JS static, C# static, Python @staticmethod.
/// Stack: unchanged (reads constructor from local)
/// ECMA-262 §15.7.14 step 2 (ClassCallError): a class constructor invoked
/// without `new` throws a TypeError. `__js_new_target` is null on plain
/// calls (set by `new` chains), so the guard is a simple null check at the
/// constructor body's start.
pub fn emit_class_requires_new_guard(chunk: &mut Chunk, class_name: &str, line: u32) {
    let nt_key = chunk.add_constant(Value::String(Arc::from("__js_new_target")));
    chunk.emit_op_u16(Op::GLOBAL_GET, nt_key, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const(
        &format!(
            "Class constructor {} cannot be invoked without 'new'",
            class_name
        ),
        line,
    );
    let te_idx = chunk.add_import("ecma:error", "TypeError");
    chunk.emit_call(te_idx, 1, line);
    crate::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
}

/// §9.1.1.3.4 GetThisBinding (JS only): in a derived constructor `this` is
/// in TDZ until `super()` runs. this_slot holds null until then — reading
/// `this`, or returning with it still null (missing/failed super), throws
/// a ReferenceError.
/// Stack: [] → [] (throws when this_slot is null)
pub fn emit_this_initialized_guard(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const(
        "Must call super constructor in derived class before accessing 'this' \
         or returning from derived constructor",
        line,
    );
    let re_idx = chunk.add_import("ecma:error", "ReferenceError");
    chunk.emit_call(re_idx, 1, line);
    crate::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
}

/// §13.3.7.2 SuperCall step 6 (JS only): calling `super()` when `this` is
/// already initialized throws a ReferenceError.
/// Stack: [] → [] (throws when this_slot is non-null)
pub fn emit_super_once_guard(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_string_const("Super constructor may only be called once", line);
    let re_idx = chunk.add_import("ecma:error", "ReferenceError");
    chunk.emit_call(re_idx, 1, line);
    crate::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
}

pub fn emit_attach_static_method(
    chunk: &mut Chunk,
    ctor_local: u16,
    method_name: &str,
    method_chunk_idx: usize,
    receiver_slot: Option<u16>,
    rest_fixed_count: Option<u8>,
    line: u32,
) {
    emit_attach_static_method_kinded(
        chunk,
        ctor_local,
        method_name,
        method_chunk_idx,
        receiver_slot,
        rest_fixed_count,
        false,
        false,
        line,
    )
}

/// `emit_attach_static_method` + ECMA-262 function-kind prototype stamp:
/// async/generator methods get `__proto__` = the matching intrinsic's
/// prototype (§27.7.1/§27.3.1/§27.4.1) so `getPrototypeOf(C.m)` and
/// `C.m instanceof AsyncFunction` hold.
#[allow(clippy::too_many_arguments)]
pub fn emit_attach_static_method_kinded(
    chunk: &mut Chunk,
    ctor_local: u16,
    method_name: &str,
    method_chunk_idx: usize,
    receiver_slot: Option<u16>,
    rest_fixed_count: Option<u8>,
    is_async: bool,
    is_generator: bool,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, ctor_local, line);
    chunk.emit_op_u16(Op::REF_FUNC, method_chunk_idx as u16, line);
    chunk.emit(0, line);
    if is_async || is_generator {
        chunk.emit_dup(line);
        crate::primitives::prototypes::emit_stamp_function_kind_proto(
            chunk,
            is_async,
            is_generator,
            line,
        );
    }
    if let Some(receiver_slot) = receiver_slot {
        chunk.emit_dup(line);
        chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
        let receiver_key = chunk.add_constant(Value::String(Arc::from("__vybe_method_receiver")));
        chunk.emit_op_u16(Op::STRUCT_SET, receiver_key, line);
        chunk.emit_op(Op::DROP, line);
    }
    if let Some(fixed_count) = rest_fixed_count {
        crate::primitives::object::emit_stamp_rest_metadata(chunk, fixed_count, line);
    }
    // §10.2.9 SetFunctionName: a static method's `name` is its property
    // key (non-enumerable, like all function metadata).
    chunk.emit_dup(line);
    chunk.emit_string_const(method_name, line);
    let name_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::STRUCT_SET, name_key, line);
    crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(chunk, line);
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

// ── Property accessors ──────────────────────────────────────────────────

// ── Constructor return ──────────────────────────────────────────────────

/// Emit return-this at the end of a constructor.
/// Stack: [] → returns this to caller
pub fn emit_constructor_return(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::RETURN, line);
}

// ── Constructor storage ─────────────────────────────────────────────────

/// Store a constructor function as a local + global variable.
/// Stack: unchanged
pub fn emit_store_constructor(
    chunk: &mut Chunk,
    class_name: &str,
    ctor_chunk_idx: usize,
    local_slot: u16,
    line: u32,
) {
    emit_store_constructor_with_upvalues(
        chunk,
        class_name,
        ctor_chunk_idx,
        local_slot,
        &[],
        false,
        line,
    );
}

/// Store a constructor function with upvalue capture. Used for closure-bound
/// parents (e.g. JS mixin pattern `(Base) => class extends Base`) where the
/// constructor body references variables from an enclosing scope.
///
/// Each upvalue entry is `(is_local, index)` — the same wire format the VM
/// reads after `REF_FUNC`. Pass an empty slice for non-closure constructors.
///
/// `case_sensitive`: when `true` (JS profile), the lowercase alias is NOT
/// emitted — otherwise a `class Range` would overwrite a hoisted
/// `function* range` at runtime, silently draining an empty continuation.
///
/// Stack: unchanged
pub fn emit_store_constructor_with_upvalues(
    chunk: &mut Chunk,
    class_name: &str,
    ctor_chunk_idx: usize,
    local_slot: u16,
    upvalues: &[(bool, u16)],
    case_sensitive: bool,
    line: u32,
) {
    chunk.emit_op_u16(Op::REF_FUNC, ctor_chunk_idx as u16, line);
    chunk.emit(upvalues.len() as u8, line);
    for (is_local, index) in upvalues {
        crate::primitives::functions::emit_closure_upvalue(chunk, *is_local, *index, line);
    }
    chunk.emit_op_u16(Op::LOCAL_TEE, local_slot, line);
    // Store under original name (case-sensitive lookup)
    let global_name = chunk.add_constant(Value::String(Arc::from(class_name)));
    chunk.emit_op_u16(Op::GLOBAL_SET, global_name, line);
    // Also store under lowercase alias for cross-language lookup (VB is case-insensitive).
    // Skip in case-sensitive profiles (JS): a `class Range` must NOT overwrite a hoisted
    // `function* range` — the two names are distinct in a case-sensitive language.
    if !case_sensitive {
        let lower = class_name.to_lowercase();
        if lower != class_name {
            chunk.emit_op_u16(Op::LOCAL_GET, local_slot, line);
            let lower_name = chunk.add_constant(Value::String(Arc::from(lower.as_str())));
            chunk.emit_op_u16(Op::GLOBAL_SET, lower_name, line);
        }
    }
}

// ── Field initialization ────────────────────────────────────────────────

/// Set a field on the object to null (pre-declaration / auto-property init).
/// Stack: unchanged
pub fn emit_init_field_null(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::NULL, line);
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Push `this` onto the stack to start a field initialization.
/// Caller compiles the value expression next, then calls `emit_init_field_end`.
/// This wraps the language-specific value-compilation in a compiler_common pattern.
pub fn emit_init_field_start(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

/// Finish a field initialization started with `emit_init_field_start`.
/// Stack before: [this, value]. Stack after: [].
pub fn emit_init_field_end(chunk: &mut Chunk, field_name: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Get a field value from `this`. Stack before: []. Stack after: [value].
pub fn emit_get_field(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}

/// Set a field value on `this` from a value already on the stack.
/// Stack before: [value]. Stack after: [].
pub fn emit_set_field_from_stack(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    // Need [this, value, value]... actually struct_set expects [obj, val].
    // Caller has [value] — we need to insert this BELOW value on stack.
    // Use a temp local approach: store value, push this, push value, struct_set, drop.
    // Simpler: let the caller use start/end pattern when value isn't pre-computed.
    // For pre-computed value: use a local temp.
    let _ = (chunk, this_slot, field_name, line);
    // This pattern is awkward without a swap opcode.
    // Use emit_init_field_start + emit_init_field_end with value compilation in between.
}

// ── Type registration ───────────────────────────────────────────────────

/// Register a type entry in chunk 0's type table.
pub fn register_type(
    chunks: &mut [Chunk],
    name: &str,
    parent: &str,
    fields: Vec<String>,
    methods: Vec<(String, usize)>,
    is_interface: bool,
    implements: Vec<String>,
    constructor_chunk: Option<usize>,
    field_descriptors: std::collections::HashMap<String, vybe_runtime::chunk::PropertyDescriptor>,
) {
    // The walker is responsible for case-canonicalising the name per
    // its language's case-sensitivity (`Compiler::canon` lowercases
    // for VB/Pascal/COBOL/PHP, preserves case for JS/TS/Python/C#).
    // No forced lowercasing here — that would silently collide
    // distinct types in case-sensitive languages (`B` and `b`). The VM
    // matches names exactly; case-insensitivity is entirely a
    // compile-time concern via `canon()`.
    chunks[0].types.push(vybe_runtime::chunk::TypeEntry {
        name: name.to_string(),
        kind: vybe_runtime::chunk::CompositeKind::Struct,
        parent: parent.to_string(),
        fields,
        methods,
        is_interface,
        implements,
        constructor_chunk,
        field_descriptors,
    });
}

/// Register a WASM GC `(array …)` defined type in chunk 0's type table and
/// return the **1-based index of the entry within that table** — the value the
/// compiler emits as the `array.new` immediate.
///
/// This is deliberately *not* the runtime registry id: the host pre-registers
/// its builtin types ahead of the module's own, so a compile-time table
/// position can't equal the registry id. Instead the VM recovers the type name
/// from this index at run time (`module_type_names[imm - 1]`) and resolves it to
/// the registry id *by name*, then `TypeDef::is_array` reads the `Array` kind
/// back to apply spec trapping `array.get`/`set`/`copy`. Names are unique per
/// declaration so the index↔name mapping is stable.
pub fn register_gc_array_type(chunks: &mut [Chunk], name: &str, elem_type: &str) -> usize {
    let type_index = chunks[0].types.len() + 1;
    // The element storage type (`i32`/`i64`/`f32`/`f64`/`i8`/`i16`/ref) is kept
    // as the array's single "field" so the VM can recover the element byte width
    // (for `array.init_data` / packed `array.get_s`) from the instance's rtt —
    // the value model stores i32/f32/f64 all as f64, so the width can't be read
    // back from the runtime value.
    let fields = if elem_type.is_empty() {
        Vec::new()
    } else {
        vec![elem_type.to_string()]
    };
    chunks[0].types.push(vybe_runtime::chunk::TypeEntry {
        name: name.to_string(),
        kind: vybe_runtime::chunk::CompositeKind::Array,
        parent: String::new(),
        fields,
        methods: Vec::new(),
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new(),
    });
    type_index
}

/// Register an interface/trait/protocol in the type table.
/// Interfaces have no constructor and method entries with chunk_idx=0 (signatures only).
/// This is the same across C# `interface`, VB `Interface`, Dart `abstract class`,
/// Python ABC — different syntax, same TypeEntry shape.
pub fn register_interface(
    chunks: &mut [Chunk],
    name: &str,
    methods: Vec<String>,
    parent_interfaces: Vec<String>,
) {
    // Names arrive pre-canonicalised by the walker — see [`register_type`].
    let method_entries: Vec<(String, usize)> = methods.into_iter().map(|m| (m, 0usize)).collect();
    chunks[0].types.push(vybe_runtime::chunk::TypeEntry {
        name: name.to_string(),
        kind: vybe_runtime::chunk::CompositeKind::Struct,
        parent: String::new(),
        fields: Vec::new(),
        methods: method_entries,
        is_interface: true,
        implements: parent_interfaces,
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new(),
    });
}

/// Register a class that implements one or more interfaces.
/// This is the standard pattern for C# `: IFoo, IBar`, Dart `implements Foo, Bar`,
/// VB `Implements IFoo`, Python `class Foo(IBar)`.
pub fn register_class_with_interfaces(
    chunks: &mut [Chunk],
    name: &str,
    parent: &str,
    fields: Vec<String>,
    methods: Vec<(String, usize)>,
    implements: Vec<String>,
    constructor_chunk: Option<usize>,
) {
    register_type(
        chunks,
        name,
        parent,
        fields,
        methods,
        false,
        implements,
        constructor_chunk,
        std::collections::HashMap::new(),
    );
}

// ── Super call (cross-language) ────────────────────────────────────────

// ── .NET default constructor: auto-call InitializeComponent ─────────────

/// In .NET, if a class defines `InitializeComponent()` (typical of WinForms
/// designer-generated code) and has no explicit constructor, the default
/// constructor must call `InitializeComponent()` automatically. Both VB and
/// C# follow this convention.
///
/// Emits bytecode equivalent to:
///   Me.InitializeComponent()      ' VB
///   this.InitializeComponent();   // C#
///
/// The `this_slot` is the local variable holding the class instance.
/// Call this AFTER instance methods have been attached to `this` (so that
/// `struct_get "initializecomponent"` finds the method).
pub fn emit_auto_init_component(chunk: &mut Chunk, this_slot: u16, line: u32) {
    emit_auto_init_call(chunk, this_slot, "initializecomponent", line);
}

/// Emit a call to `this.<method_name>()` — generalized auto-init for any
/// method listed in the profile's `auto_init_methods`.  The method name is
/// lowercased for the struct_get lookup (all method keys are stored lowercase).
pub fn emit_auto_init_call(chunk: &mut Chunk, this_slot: u16, method_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // [this]
    let name_idx = chunk.add_constant(Value::String(Arc::from(method_name.to_lowercase())));
    chunk.emit_op_u16(Op::STRUCT_GET, name_idx, line); // [method_ref]
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // [method_ref, this]
    chunk.emit_op_u8(Op::CALL_REF, 1, line); // call(1) → [result]
    chunk.emit_op(Op::DROP, line); // []
}

// NOTE: needs_auto_init_component() has moved to type_registry.rs where it
// uses the proper CompileTimeTypes hierarchy instead of string matching.
