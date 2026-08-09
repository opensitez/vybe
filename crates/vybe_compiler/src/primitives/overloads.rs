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
            StmtKind::InterfaceDecl { name, members, .. } => {
                self.register_interface_method_signatures(name, members);
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

    pub(super) fn normalize_fortran_dispatch_type(&self, type_hint: &str) -> String {
        let resolved = self.resolve_source_type_alias(type_hint);
        let normalized = Self::normalize_type_hint(&resolved);
        let trimmed = normalized.trim();

        if let Some(inner) = trimmed
            .strip_prefix("type(")
            .and_then(|rest| rest.strip_suffix(')'))
            .or_else(|| {
                trimmed
                    .strip_prefix("class(")
                    .and_then(|rest| rest.strip_suffix(')'))
            })
        {
            return self.canon(inner.trim());
        }

        if trimmed == "int" || trimmed.starts_with("integer") {
            return "integer".to_string();
        }
        if matches!(trimmed, "real" | "float" | "double" | "double precision")
            || trimmed.starts_with("real(")
        {
            return "real".to_string();
        }
        if trimmed == "bool" || trimmed.starts_with("logical") {
            return "logical".to_string();
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
                    .map(|hint| self.normalize_fortran_dispatch_type(&hint))
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
            for (arg_type, param_type) in arg_types.iter().zip(param_types.iter()) {
                let Some(param_type) = param_type.as_ref() else {
                    continue;
                };
                let param_key = self.normalize_fortran_dispatch_type(param_type);
                let Some(arg_key) = arg_type.as_ref() else {
                    continue;
                };
                if arg_key == &param_key {
                    score += 2;
                    continue;
                }
                compatible = false;
                break;
            }

            if !compatible {
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
