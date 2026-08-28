//! Interface signature predeclaration & Fortran overload resolution.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

impl Compiler {
    pub(super) fn predeclare_interface_signatures_in_body(&mut self, body: &[Statement]) {
        for stmt in body {
            self.predeclare_interface_signatures_in_stmt(stmt);
        }
    }

    pub(super) fn predeclare_interface_signatures_in_stmt(&mut self, stmt: &Statement) {
        match &stmt.kind {
            StmtKind::Block(body)
            | StmtKind::NamespaceDecl { body, .. }
            | StmtKind::FunctionDecl { body, .. } => {
                self.predeclare_interface_signatures_in_body(body);
            }
            StmtKind::InterfaceDecl {
                name,
                members,
                parents,
                ..
            } => {
                self.register_interface_method_signatures(name, members);
                self.register_interface_as_pending_class(name, members, parents);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                self.predeclare_interface_signatures_in_members(members);
            }
            _ => {}
        }
    }

    pub(super) fn predeclare_interface_signatures_in_members(&mut self, members: &[ClassMember]) {
        for member in members {
            match member {
                ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                    self.predeclare_interface_signatures_in_stmt(stmt);
                }
                _ => {}
            }
        }
    }

    /// Put an interface in the class table.
    ///
    /// ⛔ `StmtKind::InterfaceDecl` COMPILES TO A NO-OP — "interfaces are
    /// type-level only" — and that is right about EMISSION and wrong about
    /// KNOWLEDGE. An interface declares members, and everything that asks "does
    /// this type declare that name" was answering NO for every interface in
    /// every program: `--dump-classes` on a C# file declaring `IGreet` lists
    /// the classes and no interface at all.
    ///
    /// The immediate consumer is augmentation. C# default interface methods are
    /// copied into implementers today by a walker-private `HashMap` +
    /// `inject_interface_defaults` pass, which is `AugmentationMode::Copy` with
    /// `AugmentationPosition::AfterOwn` written out by hand. It cannot move to
    /// the shared model while `apply_augmentations` has nothing to resolve
    /// `from: "IGreet"` against.
    ///
    /// Registered with members and NO constructor: an interface is not
    /// constructible, and nothing here gives it a ctor global or an allocation.
    /// `declared_kind` already exists on the normalized model for exactly this.
    pub(super) fn register_interface_as_pending_class(
        &mut self,
        interface_name: &str,
        members: &[InterfaceMember],
        parents: &[String],
    ) {
        let canonical = self.canon(interface_name);
        if canonical.is_empty() || self.pending_classes.contains_key(&canonical) {
            return;
        }
        let mut instance_member_names = Vec::new();
        for member in members {
            let name = match member {
                InterfaceMember::Method { name, .. }
                | InterfaceMember::Property { name, .. }
                | InterfaceMember::Event { name, .. } => name,
                _ => continue,
            };
            let stored = self.js_member_storage_name_for_class(&canonical, name);
            if !instance_member_names.contains(&stored) {
                instance_member_names.push(stored);
            }
        }
        let bases: Vec<String> = parents.iter().map(|p| self.canon(p)).collect();
        let parent = bases.first().cloned();
        self.pending_classes.insert(
            canonical,
            PendingClass {
                parent,
                bases,
                enclosing_class: self.current_class.clone(),
                fields: Vec::new(),
                field_storage_names: HashMap::new(),
                is_value_type: false,
                instance_member_names,
                instance_pointer_method_names: Vec::new(),
                instance_field_types: HashMap::new(),
                static_fields: Vec::new(),
                static_field_types: HashMap::new(),
                static_method_names: Vec::new(),
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: Vec::new(),
                statics: Vec::new(),
            },
        );
    }

    pub(super) fn register_interface_method_signatures(
        &mut self,
        interface_name: &str,
        members: &[InterfaceMember],
    ) {
        let interface_canonical = self.canon(interface_name);
        let operator_symbol = self.fortran_interface_operator_symbol(interface_name);

        for member in members {
            let InterfaceMember::Method {
                name,
                params,
                return_type,
                signature_source,
                ..
            } = member
            else {
                continue;
            };

            let target_name = signature_source.as_ref().unwrap_or(name);
            let target_canonical = self.canon(target_name);
            let canonical_names =
                if self.profile.interface_block_is_generic_alias && !interface_name.is_empty() {
                    vec![target_canonical.clone(), interface_canonical.clone()]
                } else {
                    vec![self.canon(name)]
                };

            if let Some(source_name) = signature_source.as_ref() {
                let source_canonical = self.canon(source_name);
                for canonical in &canonical_names {
                    if let Some(source_modes) =
                        self.function_param_modes.get(&source_canonical).cloned()
                    {
                        self.function_param_modes
                            .entry(canonical.clone())
                            .or_insert(source_modes);
                    }
                    if let Some(source_types) =
                        self.function_param_types.get(&source_canonical).cloned()
                    {
                        self.function_param_types
                            .entry(canonical.clone())
                            .or_insert(source_types);
                    }
                    if let Some(min_arity) = self.function_min_arity.get(&source_canonical).copied()
                    {
                        self.function_min_arity
                            .entry(canonical.clone())
                            .or_insert(min_arity);
                    }
                    if let Some(signatures) =
                        self.function_signatures.get(&source_canonical).cloned()
                    {
                        self.function_signatures
                            .entry(canonical.clone())
                            .or_insert(signatures);
                    }
                    if let Some(source_return_type) =
                        self.function_return_types.get(&source_canonical).cloned()
                    {
                        self.function_return_types
                            .entry(canonical.clone())
                            .or_insert(source_return_type);
                    }
                }
            } else {
                let param_modes: Vec<PassBy> = params.iter().map(|param| param.pass_by).collect();
                let param_types: Vec<Option<String>> = params
                    .iter()
                    .map(|param| param.type_hint.as_deref().map(str::to_string))
                    .collect();
                let min_arity = params
                    .iter()
                    .take_while(|param| param.default.is_none() && !param.is_rest)
                    .count();
                let signature = CallSignature::from_params(params);

                for canonical in &canonical_names {
                    self.function_param_modes
                        .entry(canonical.clone())
                        .or_insert_with(|| param_modes.clone());
                    self.function_param_types
                        .entry(canonical.clone())
                        .or_insert_with(|| param_types.clone());
                    self.function_min_arity
                        .entry(canonical.clone())
                        .or_insert(min_arity);

                    let signatures = self
                        .function_signatures
                        .entry(canonical.clone())
                        .or_default();
                    if !signatures.iter().any(|existing| {
                        existing.param_names == signature.param_names
                            && existing.min_arity == signature.min_arity
                            && existing.has_rest == signature.has_rest
                    }) {
                        signatures.push(signature.clone());
                    }

                    if let Some(return_type) = return_type.as_ref() {
                        self.function_return_types
                            .entry(canonical.clone())
                            .or_insert_with(|| return_type.clone());
                    }
                }
            }

            if self.profile.interface_block_is_generic_alias && !interface_name.is_empty() {
                let overload = FortranInterfaceOverload {
                    target_name: target_canonical,
                    min_arity: params
                        .iter()
                        .take_while(|param| param.default.is_none() && !param.is_rest)
                        .count(),
                    param_types: params
                        .iter()
                        .map(|param| param.type_hint.as_deref().map(str::to_string))
                        .collect(),
                };

                if let Some(symbol) = operator_symbol.as_ref() {
                    let overloads = self
                        .fortran_operator_overloads
                        .entry(symbol.clone())
                        .or_default();
                    if !overloads
                        .iter()
                        .any(|existing| existing.target_name == overload.target_name)
                    {
                        overloads.push(overload);
                    }
                } else {
                    let overloads = self
                        .fortran_interface_overloads
                        .entry(interface_canonical.clone())
                        .or_default();
                    if !overloads
                        .iter()
                        .any(|existing| existing.target_name == overload.target_name)
                    {
                        overloads.push(overload);
                    }
                }
            }
        }
    }

    pub(super) fn fortran_interface_operator_symbol(&self, name: &str) -> Option<String> {
        let trimmed = name.trim();
        let lower = trimmed.to_ascii_lowercase();
        if !lower.starts_with("operator(") || !trimmed.ends_with(')') {
            return None;
        }
        let start = trimmed.find('(')? + 1;
        let end = trimmed.rfind(')')?;
        Some(trimmed[start..end].trim().to_string())
    }

    /// The key two types must share to be the SAME type for overload dispatch.
    ///
    /// Was `normalize_fortran_dispatch_type`, which hardcoded Fortran's
    /// spellings — `integer`/`int`, `real`/`float`/`double`/`double precision`,
    /// `logical`/`bool` — inside a shared crate. Those now live in Fortran's own
    /// `[builtin_types]`, and this asks the shared spelling table, language
    /// entries first. Any language whose profile declares its spellings gets
    /// overload dispatch for free; none has to be named here.
    ///
    /// Two kinds of key, and they cannot collide: a built-in is keyed by the
    /// `BuiltinType` it classifies to, so `real` and `double precision` agree
    /// without either being canonical, while a user type keys on its canonical
    /// name. The `builtin:` prefix keeps a type actually NAMED `Int` from
    /// colliding with the built-in.
    pub(super) fn overload_dispatch_key(&self, type_hint: &str) -> String {
        let resolved = self.resolve_source_type_alias(type_hint);
        let normalized = Self::normalize_type_hint(&resolved);
        let trimmed = normalized.trim();

        // `type(T)` / `class(T)` wrap a USER type name — check before the
        // spelling table so a declared spelling cannot capture the wrapper.
        if let Some(inner) = super::calls::strip_parametric_type_wrapper(trimmed) {
            return self.canon(inner);
        }

        if let Some(ty) = vybe_ast::builtin_types::classify_with(
            &self.profile.builtin_type_spellings,
            trimmed,
        ) {
            return format!("builtin:{ty:?}");
        }

        self.canon(trimmed)
    }

    pub(super) fn fortran_overload_target_param_types(
        &self,
        overload: &FortranInterfaceOverload,
    ) -> Vec<Option<String>> {
        self.function_param_types
            .get(&overload.target_name)
            .cloned()
            .filter(|param_types| !param_types.is_empty())
            .unwrap_or_else(|| overload.param_types.clone())
    }

    pub(super) fn resolve_fortran_overload_target_with_fallback(
        &self,
        overloads: &[FortranInterfaceOverload],
        arg_exprs: &[Expression],
        allow_unknown_fallback: bool,
    ) -> Option<String> {
        let arg_types: Vec<Option<String>> = arg_exprs
            .iter()
            .map(|expr| {
                self.infer_expr_type_hint(expr)
                    .map(|hint| self.overload_dispatch_key(&hint))
            })
            .collect();
        let has_known_arg_types = arg_types.iter().any(Option::is_some);

        let mut best: Option<(&FortranInterfaceOverload, usize)> = None;
        let mut ambiguous = false;

        for overload in overloads {
            if arg_exprs.len() < overload.min_arity {
                continue;
            }

            let param_types = self.fortran_overload_target_param_types(overload);
            if !param_types.is_empty() && param_types.len() != arg_exprs.len() {
                continue;
            }

            let mut score = 0usize;
            let mut compatible = true;
            // Did a DECLARED parameter actually match a arg whose type we can
            // name (`evidence`), or did it face one we cannot (`unproven`)?
            let mut evidence = false;
            let mut unproven = false;
            for (arg_type, param_type) in arg_types.iter().zip(param_types.iter()) {
                let Some(param_type) = param_type.as_ref() else {
                    continue;
                };
                let param_key = self.overload_dispatch_key(param_type);
                let Some(arg_key) = arg_type.as_ref() else {
                    unproven = true;
                    continue;
                };
                if arg_key == &param_key {
                    score += 2;
                    evidence = true;
                    continue;
                }
                compatible = false;
                break;
            }

            if !compatible {
                continue;
            }

            // An unknown arg type used to `continue` past the check above, which
            // made it compatible with EVERY declared parameter at score 0 — so a
            // sole overload won here and returned below, long before
            // `allow_unknown_fallback` was ever consulted. For an operator that
            // is not a harmless guess: the builtin is the correct alternative,
            // and picking the overload inside its own implementation is how
            // `c%v = a%v + b%v` recursed into itself.
            //
            // So a winner resting ONLY on unknowns needs positive evidence when
            // a builtin could serve instead. Generic-NAME dispatch
            // (`allow_unknown_fallback`) keeps the old behaviour on purpose:
            // there is no builtin to fall back to, so a single overload must
            // still be selected for an argument that does not infer.
            //
            // Gated on `unproven` rather than `!evidence` alone so an overload
            // whose target records no parameter types at all — nothing to prove
            // anything against — resolves exactly as it did before.
            if unproven && !evidence && !allow_unknown_fallback {
                continue;
            }

            match best {
                None => {
                    best = Some((overload, score));
                    ambiguous = false;
                }
                Some((_, best_score)) if score > best_score => {
                    best = Some((overload, score));
                    ambiguous = false;
                }
                Some((_, best_score)) if score == best_score => {
                    ambiguous = true;
                }
                _ => {}
            }
        }

        if let Some((overload, _)) = best {
            if !ambiguous || overloads.len() == 1 {
                return Some(overload.target_name.clone());
            }
        }

        (allow_unknown_fallback && !has_known_arg_types && overloads.len() == 1)
            .then(|| overloads[0].target_name.clone())
    }

    pub(super) fn resolve_fortran_overload_target(
        &self,
        overloads: &[FortranInterfaceOverload],
        arg_exprs: &[Expression],
    ) -> Option<String> {
        self.resolve_fortran_overload_target_with_fallback(overloads, arg_exprs, true)
    }

    pub(super) fn resolve_fortran_interface_target(
        &self,
        name: &str,
        arg_exprs: &[Expression],
    ) -> Option<String> {
        // Redundant name check removed: `fortran_interface_overloads` is only
        // populated under `interface_block_is_generic_alias`.
        Some(self.canon(name))
            .and_then(|canonical| self.fortran_interface_overloads.get(&canonical))
            .and_then(|overloads| self.resolve_fortran_overload_target(overloads, arg_exprs))
    }

    pub(super) fn resolve_fortran_operator_target(
        &self,
        op: &BinOp,
        arg_exprs: &[Expression],
    ) -> Option<String> {
        let symbol = match op {
            BinOp::Add => "+",
            BinOp::Sub => "-",
            BinOp::Mul => "*",
            BinOp::Div => "/",
            BinOp::Pow => "**",
            _ => return None,
        };

        self.fortran_operator_overloads
            .get(symbol)
            .and_then(|overloads| {
                self.resolve_fortran_overload_target_with_fallback(overloads, arg_exprs, false)
            })
    }
}
