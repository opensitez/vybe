//! Reflection/attribute metadata, member-storage names, private-access, WinForms inference.
//!
//! Extracted from `compiler/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

fn dotnet_descriptor_runtime_type_name(interface: &str, name: &str) -> String {
    let trimmed = name.trim();
    if trimmed.starts_with("System.") {
        return trimmed.to_string();
    }
    let namespace = interface
        .strip_prefix("dotnet.")
        .unwrap_or(interface)
        .trim_end_matches('.');
    if namespace.is_empty() {
        trimmed.to_string()
    } else {
        format!("{namespace}.{trimmed}")
    }
}

impl Compiler {
    pub(super) fn parse_pascal_array_bound_token(token: &str) -> Option<(i64, bool)> {
        let trimmed = token.trim();
        if let Ok(value) = trimmed.parse::<i64>() {
            return Some((value, false));
        }

        if trimmed.starts_with('\'') && trimmed.ends_with('\'') && trimmed.len() >= 3 {
            let inner = &trimmed[1..trimmed.len() - 1];
            let unescaped = inner.replace("''", "'");
            let mut chars = unescaped.chars();
            if let (Some(ch), None) = (chars.next(), chars.next()) {
                return Some((ch as i64, true));
            }
        }

        None
    }

    pub(super) fn pascal_array_type_hint_metadata(
        &self,
        type_hint: &str,
    ) -> Option<PascalArrayBoundsMetadata> {
        let trimmed = type_hint.trim().trim_end_matches('?').trim();
        let lowered = trimmed.to_ascii_lowercase();
        if !lowered.starts_with("array") {
            return None;
        }

        let Some(bracket_start) = trimmed.find('[') else {
            return Some(PascalArrayBoundsMetadata {
                is_fixed: false,
                dimensions: Vec::new(),
            });
        };
        let bracket_end = trimmed[bracket_start + 1..].find(']')? + bracket_start + 1;
        let mut dimensions = Vec::new();
        for dim in trimmed[bracket_start + 1..bracket_end]
            .split(',')
            .map(str::trim)
            .filter(|dim| !dim.is_empty())
        {
            let (lower, upper) = dim.split_once("..")?;
            let (lower, lower_is_char) = Self::parse_pascal_array_bound_token(lower)?;
            let (upper, upper_is_char) = Self::parse_pascal_array_bound_token(upper)?;
            if lower_is_char != upper_is_char {
                return None;
            }
            let length = if upper >= lower {
                (upper - lower + 1) as usize
            } else {
                0
            };
            dimensions.push(PascalArrayDimensionMetadata {
                first_index: lower,
                length,
                uses_char_ordinal: lower_is_char,
            });
        }
        Some(PascalArrayBoundsMetadata {
            is_fixed: !dimensions.is_empty(),
            dimensions,
        })
    }

    pub(super) fn pascal_ordinal_index_expr(index: Expression) -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Ord")),
            args: vec![Argument::positional(index)],
            optional: false,
        })
    }

    pub(super) fn pascal_indexed_type_hint(type_hint: &str) -> Option<String> {
        let trimmed = type_hint.trim().trim_end_matches('?').trim();
        if !trimmed.to_ascii_lowercase().starts_with("array") {
            return None;
        }

        let bracket_start = trimmed.find('[')?;
        let bracket_end = trimmed[bracket_start + 1..].find(']')? + bracket_start + 1;
        let dims = trimmed[bracket_start + 1..bracket_end].trim();
        let after_bracket = trimmed[bracket_end + 1..].trim();
        let of_pos = after_bracket.to_ascii_lowercase().find("of")?;
        let element_type = after_bracket[of_pos + 2..].trim();
        if element_type.is_empty() {
            return None;
        }

        if let Some((_, remaining_dims)) = dims.split_once(',') {
            Some(format!(
                "array[{}] of {}",
                remaining_dims.trim(),
                element_type
            ))
        } else {
            Some(element_type.to_string())
        }
    }

    pub(super) fn collect_reflection_metadata(&mut self, body: &[Statement]) {
        self.reflection_types.clear();
        self.attribute_usage.clear();
        self.reflection_bindings.clear();
        self.seed_dotnet_reflection_metadata();
        for stmt in body {
            self.collect_reflection_stmt(stmt, None);
        }
    }

    fn seed_dotnet_reflection_metadata(&mut self) {
        use crate::profile::ReflectionTypeNaming;
        use vybe_bytecode::component_model::ComponentItemKind;

        if self.profile.reflection_type_naming != ReflectionTypeNaming::Dotnet {
            return;
        }

        for export in crate::platforms::dotnet::emitter::dotnet_component_descriptor().exports {
            let ComponentItemKind::Class(class) = export.kind else {
                continue;
            };
            let runtime_name = dotnet_descriptor_runtime_type_name(&export.interface, &class.name);
            let parents = class
                .parent
                .as_deref()
                .map(|parent| dotnet_descriptor_runtime_type_name(&export.interface, parent))
                .into_iter()
                .collect();
            let interfaces = class
                .interfaces
                .iter()
                .map(|interface| dotnet_descriptor_runtime_type_name(&export.interface, interface))
                .collect();
            let mut metadata = ReflectionTypeMetadata {
                parents,
                interfaces,
                decorators: Vec::new(),
                is_value_type: runtime_name.eq_ignore_ascii_case("System.ValueType")
                    || runtime_name.eq_ignore_ascii_case("System.Enum"),
                ..ReflectionTypeMetadata::default()
            };
            if let Some(constructor) = class.constructor {
                metadata.constructors.push(ReflectionConstructorMetadata {
                    param_types: vec!["System.Object".to_string(); constructor.arity as usize],
                });
            }
            for property in class.properties {
                metadata.properties.insert(
                    property.name,
                    ReflectionMemberMetadata {
                        decorators: Vec::new(),
                        is_static: false,
                        can_write: property.setter.is_some(),
                    },
                );
            }
            for method in class.methods {
                metadata.methods.insert(
                    method.name,
                    ReflectionMethodMetadata {
                        decorators: Vec::new(),
                        params: (0..method.arity)
                            .map(|index| ReflectionParamMetadata {
                                name: format!("arg{index}"),
                                decorators: Vec::new(),
                            })
                            .collect(),
                        is_static: method.is_static,
                    },
                );
            }
            self.reflection_types.insert(runtime_name.clone(), metadata);
            self.attribute_usage
                .insert(runtime_name, AttributeUsageMetadata::default());
        }
    }

    pub(super) fn collect_reflection_stmt(
        &mut self,
        stmt: &Statement,
        parent_runtime_name: Option<&str>,
    ) {
        match &stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                interfaces,
                members,
                decorators,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    parents,
                    interfaces,
                    decorators,
                    members,
                    false,
                );
            }
            StmtKind::StructDecl {
                name,
                interfaces,
                members,
                decorators,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    &[],
                    interfaces,
                    decorators,
                    members,
                    true,
                );
            }
            StmtKind::InterfaceDecl {
                name,
                parents,
                decorators,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                let metadata = ReflectionTypeMetadata {
                    parents: parents
                        .iter()
                        .map(|parent| self.reflection_runtime_type_name(parent, None))
                        .collect(),
                    decorators: decorators.clone(),
                    interfaces: parents
                        .iter()
                        .map(|parent| self.reflection_runtime_type_name(parent, None))
                        .collect(),
                    ..ReflectionTypeMetadata::default()
                };
                self.reflection_types.insert(runtime_name.clone(), metadata);
                let usage = self.extract_attribute_usage(decorators);
                self.attribute_usage.insert(runtime_name, usage);
            }
            StmtKind::EnumDecl {
                name,
                interfaces,
                decorators,
                body_members,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    &[],
                    interfaces,
                    decorators,
                    body_members,
                    true,
                );
            }
            StmtKind::NamespaceDecl { name, body } => {
                let namespace_runtime =
                    self.reflection_runtime_type_name(name, parent_runtime_name);
                for nested in body {
                    self.collect_reflection_stmt(nested, Some(namespace_runtime.as_str()));
                }
            }
            StmtKind::Block(body) => {
                for nested in body {
                    self.collect_reflection_stmt(nested, parent_runtime_name);
                }
            }
            _ => {}
        }
    }

    pub(super) fn record_reflection_type(
        &mut self,
        runtime_name: &str,
        parents: &[String],
        interfaces: &[String],
        decorators: &[Expression],
        members: &[ClassMember],
        is_value_type: bool,
    ) {
        let mut metadata = ReflectionTypeMetadata {
            parents: parents
                .iter()
                .map(|parent| self.reflection_runtime_type_name(parent, None))
                .collect(),
            interfaces: interfaces
                .iter()
                .map(|parent| self.reflection_runtime_type_name(parent, None))
                .collect(),
            decorators: decorators.to_vec(),
            is_value_type,
            ..ReflectionTypeMetadata::default()
        };
        let mut nested_types: Vec<&Statement> = Vec::new();

        for member in members {
            match member {
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name,
                        params,
                        modifiers,
                        ..
                    } = &stmt.kind
                    {
                        let mut method_decorators = Vec::new();
                        let mut param_decorators: HashMap<usize, Vec<Expression>> = HashMap::new();
                        for decorator in &modifiers.decorators {
                            if let Some((index, attr)) =
                                self.unpack_param_decorator_carrier(decorator)
                            {
                                param_decorators.entry(index).or_default().push(attr);
                            } else {
                                method_decorators.push(decorator.clone());
                            }
                        }
                        metadata.methods.insert(
                            name.clone(),
                            ReflectionMethodMetadata {
                                decorators: method_decorators,
                                is_static: modifiers.is_static,
                                params: params
                                    .iter()
                                    .enumerate()
                                    .map(|(index, param)| ReflectionParamMetadata {
                                        name: param.name.clone(),
                                        decorators: param_decorators
                                            .remove(&index)
                                            .unwrap_or_default(),
                                    })
                                    .collect(),
                            },
                        );
                    }
                }
                ClassMember::Property {
                    name,
                    setter,
                    modifiers,
                    ..
                } => {
                    metadata.properties.insert(
                        name.clone(),
                        ReflectionMemberMetadata {
                            decorators: modifiers.decorators.clone(),
                            is_static: modifiers.is_static,
                            can_write: setter.is_some(),
                        },
                    );
                }
                ClassMember::Field {
                    name, modifiers, ..
                } => {
                    metadata.fields.insert(
                        name.clone(),
                        ReflectionMemberMetadata {
                            decorators: modifiers.decorators.clone(),
                            is_static: modifiers.is_static,
                            can_write: true,
                        },
                    );
                }
                ClassMember::Constructor { params, .. } => {
                    metadata.constructors.push(ReflectionConstructorMetadata {
                        param_types: params
                            .iter()
                            .map(|param| {
                                self.reflection_runtime_type_name(
                                    param.type_hint.as_deref().unwrap_or("Object"),
                                    None,
                                )
                            })
                            .collect(),
                    });
                }
                ClassMember::NestedType(stmt) => {
                    let nested_runtime = match &stmt.kind {
                        StmtKind::ClassDecl { name, .. }
                        | StmtKind::StructDecl { name, .. }
                        | StmtKind::InterfaceDecl { name, .. }
                        | StmtKind::EnumDecl { name, .. } => {
                            Some(self.reflection_runtime_type_name(name, Some(runtime_name)))
                        }
                        _ => None,
                    };
                    if let Some(nested_runtime) = nested_runtime {
                        metadata.nested_types.push(nested_runtime);
                    }
                    nested_types.push(stmt);
                }
                _ => {}
            }
        }

        self.reflection_types
            .insert(runtime_name.to_string(), metadata);
        let usage = self.extract_attribute_usage(decorators);
        self.attribute_usage.insert(runtime_name.to_string(), usage);
        for stmt in nested_types {
            self.collect_reflection_stmt(stmt, Some(runtime_name));
        }
    }

    pub(super) fn extract_attribute_usage(
        &self,
        decorators: &[Expression],
    ) -> AttributeUsageMetadata {
        let mut usage = AttributeUsageMetadata::default();

        for decorator in decorators {
            let ExprKind::New { args, .. } = &decorator.kind else {
                continue;
            };
            let Some(attr_type) = self.reflection_attribute_type_name(decorator) else {
                continue;
            };
            if !attr_type.eq_ignore_ascii_case("System.AttributeUsageAttribute") {
                continue;
            }

            for arg in args {
                match arg.name.as_deref() {
                    Some("AllowMultiple") => {
                        usage.allow_multiple =
                            matches!(arg.value.kind, ExprKind::Lit(Literal::Bool(true)));
                    }
                    Some("Inherited") => {
                        usage.inherited =
                            matches!(arg.value.kind, ExprKind::Lit(Literal::Bool(true)));
                    }
                    _ => {}
                }
            }
        }

        usage
    }

    pub(crate) fn reflection_runtime_type_name(
        &self,
        type_name: &str,
        parent_runtime_name: Option<&str>,
    ) -> String {
        use crate::profile::ReflectionTypeNaming;

        let global_stripped = Self::strip_global_namespace_prefix(type_name);
        let trimmed = global_stripped.trim().trim_end_matches('?').trim();
        let mut without_generics = String::with_capacity(trimmed.len());
        let mut depth = 0usize;
        for ch in trimmed.chars() {
            match ch {
                '<' => depth += 1,
                '>' => depth = depth.saturating_sub(1),
                _ if depth == 0 => without_generics.push(ch),
                _ => {}
            }
        }
        let base = without_generics.trim();

        match self.profile.reflection_type_naming {
            // Native: the type keeps its own name. This is what preserves
            // each language's real hierarchy — PHP `Throwable` stays
            // `Throwable`, not `System.Throwable`, so its exception `__types`
            // chain reflects PHP's model (sibling `Error`/`Exception`), never
            // .NET's single `System.Exception` root. Nested types still
            // qualify under their declaring type.
            ReflectionTypeNaming::Native => {
                if let Some(parent) = parent_runtime_name {
                    let leaf = base.rsplit('.').next().unwrap_or(base).trim();
                    format!("{parent}.{leaf}")
                } else {
                    base.to_string()
                }
            }
            // .NET BCL scheme (C#/VB): map primitives and root under `System.`.
            ReflectionTypeNaming::Dotnet => {
                let normalized = match base {
                    "int" | "Int32" => "Int32",
                    "uint" | "UInt32" => "UInt32",
                    "long" | "Int64" => "Int64",
                    "ulong" | "UInt64" => "UInt64",
                    "short" | "Int16" => "Int16",
                    "ushort" | "UInt16" => "UInt16",
                    "byte" | "Byte" => "Byte",
                    "sbyte" | "SByte" => "SByte",
                    "float" | "Single" => "Single",
                    "double" | "Double" => "Double",
                    "decimal" | "Decimal" => "Decimal",
                    "bool" | "Boolean" => "Boolean",
                    "char" | "Char" => "Char",
                    "string" | "String" => "String",
                    "object" | "Object" => "Object",
                    other => other,
                };
                let normalized = normalized
                    .strip_prefix("System.System.")
                    .unwrap_or(normalized);
                if let Some(parent) = parent_runtime_name {
                    let leaf = normalized.rsplit('.').next().unwrap_or(normalized).trim();
                    format!("{parent}.{leaf}")
                } else if normalized.starts_with("System.") {
                    normalized.to_string()
                } else {
                    format!("System.{}", normalized)
                }
            }
        }
    }

    pub(crate) fn reflection_attribute_type_name(&self, expr: &Expression) -> Option<String> {
        let class = match &expr.kind {
            ExprKind::New { class, .. } => class.as_ref(),
            _ => return None,
        };

        let raw_name = match &class.kind {
            ExprKind::Ident(name) => name.clone(),
            ExprKind::Member { .. } => self.flatten_member_chain(class).join("."),
            _ => return None,
        };

        if !raw_name.contains('.') {
            let mut matches: Vec<String> = self
                .reflection_types
                .keys()
                .filter(|known| {
                    known
                        .rsplit('.')
                        .next()
                        .is_some_and(|leaf| leaf == raw_name)
                })
                .cloned()
                .collect();
            matches.sort();
            matches.dedup();
            if matches.len() == 1 {
                return matches.into_iter().next();
            }
        }

        Some(self.reflection_runtime_type_name(&raw_name, None))
    }

    pub(super) fn unpack_param_decorator_carrier(
        &self,
        expr: &Expression,
    ) -> Option<(usize, Expression)> {
        let ExprKind::New { class, args } = &expr.kind else {
            return None;
        };
        let ExprKind::Ident(name) = &class.kind else {
            return None;
        };
        if name != "__vybe_param_attribute" || args.len() != 2 {
            return None;
        }
        let ExprKind::Lit(Literal::Int(index)) = args[0].value.kind else {
            return None;
        };
        Some((index.max(0) as usize, args[1].value.clone()))
    }

    pub(crate) fn reflection_type_lookup_name(&self, type_name: &str) -> String {
        self.reflection_runtime_type_name(type_name, None)
    }

    pub(crate) fn reflection_type_metadata(
        &self,
        type_name: &str,
    ) -> Option<&ReflectionTypeMetadata> {
        let lookup = self.reflection_type_lookup_name(type_name);
        self.reflection_types.get(&lookup)
    }

    pub(crate) fn reflection_type_short_name(&self, type_name: &str) -> String {
        let trimmed = type_name.trim().trim_end_matches('?').trim();
        let without_namespace = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
        let without_generics = without_namespace
            .split('<')
            .next()
            .unwrap_or(without_namespace)
            .trim();
        without_generics.to_string()
    }

    pub(crate) fn reflection_type_full_name(&self, type_name: &str) -> String {
        self.reflection_runtime_type_name(type_name, None)
    }

    pub(crate) fn reflection_is_enum_type(&self, type_name: &str) -> bool {
        let lookup = self.reflection_type_lookup_name(type_name);
        let short_name = self.canon(&self.reflection_type_short_name(type_name));
        self.enum_value_names.contains_key(&lookup)
            || self.enum_value_names.contains_key(&short_name)
            || self.enum_value_names.keys().any(|known| {
                known.eq_ignore_ascii_case(&lookup) || known.eq_ignore_ascii_case(&short_name)
            })
    }

    pub(crate) fn reflection_is_value_type(&self, type_name: &str) -> bool {
        let lookup = self.reflection_type_lookup_name(type_name);
        if self
            .reflection_types
            .get(&lookup)
            .is_some_and(|meta| meta.is_value_type)
        {
            return true;
        }
        matches!(
            lookup.as_str(),
            "System.Boolean"
                | "System.Byte"
                | "System.SByte"
                | "System.Int16"
                | "System.UInt16"
                | "System.Int32"
                | "System.UInt32"
                | "System.Int64"
                | "System.UInt64"
                | "System.Single"
                | "System.Double"
                | "System.Decimal"
                | "System.Char"
                | "System.DateTime"
                | "System.TimeSpan"
                | "System.Guid"
        )
    }

    pub(crate) fn reflection_base_type_name(&self, type_name: &str) -> Option<String> {
        self.reflection_type_metadata(type_name)
            .and_then(|meta| meta.parents.first().cloned())
    }

    pub(crate) fn reflection_nested_type_name(
        &self,
        type_name: &str,
        nested_name: &str,
    ) -> Option<String> {
        let parent = self.reflection_type_lookup_name(type_name);
        let desired = nested_name.trim();
        self.reflection_types
            .keys()
            .find(|candidate| {
                candidate
                    .strip_prefix(&(parent.clone() + "."))
                    .is_some_and(|leaf| leaf.eq_ignore_ascii_case(desired))
            })
            .cloned()
    }

    pub(crate) fn reflection_generic_argument_types(&self, type_name: &str) -> Vec<String> {
        let trimmed = type_name.trim();
        let Some(start) = trimmed.find('<') else {
            return Vec::new();
        };
        let Some(end) = trimmed.rfind('>') else {
            return Vec::new();
        };
        let inner = &trimmed[start + 1..end];
        let mut parts = Vec::new();
        let mut current = String::new();
        let mut depth = 0usize;
        for ch in inner.chars() {
            match ch {
                '<' => {
                    depth += 1;
                    current.push(ch);
                }
                '>' => {
                    depth = depth.saturating_sub(1);
                    current.push(ch);
                }
                ',' if depth == 0 => {
                    let part = current.trim();
                    if !part.is_empty() {
                        parts.push(self.reflection_type_full_name(part));
                    }
                    current.clear();
                }
                _ => current.push(ch),
            }
        }
        let part = current.trim();
        if !part.is_empty() {
            parts.push(self.reflection_type_full_name(part));
        }
        parts
    }

    pub(crate) fn reflection_open_generic_type_name(&self, type_name: &str) -> String {
        self.reflection_type_lookup_name(type_name)
    }

    pub(crate) fn reflection_interfaces(&self, type_name: &str) -> Vec<String> {
        self.reflection_type_metadata(type_name)
            .map(|meta| meta.interfaces.clone())
            .unwrap_or_default()
    }

    pub(crate) fn reflection_is_assignable_from(
        &self,
        target_type: &str,
        candidate_type: &str,
    ) -> bool {
        let target = self.reflection_type_lookup_name(target_type);
        let mut pending = vec![self.reflection_type_lookup_name(candidate_type)];
        let mut visited = HashSet::new();

        while let Some(current) = pending.pop() {
            if !visited.insert(current.clone()) {
                continue;
            }
            if current.eq_ignore_ascii_case(&target) {
                return true;
            }
            if let Some(meta) = self.reflection_types.get(&current) {
                pending.extend(meta.parents.iter().cloned());
                pending.extend(meta.interfaces.iter().cloned());
            }
        }

        false
    }

    pub(super) fn module_imports_namespace(&self, namespace: &str) -> bool {
        self.current_module_imports
            .iter()
            .any(|import| match &import.kind {
                ImportKind::Simple { path, .. }
                | ImportKind::Named { path, .. }
                | ImportKind::Wildcard { path, .. }
                | ImportKind::Default { path, .. } => path.eq_ignore_ascii_case(namespace),
            })
    }

    pub(super) fn should_infer_winforms_form(&self, name: &str, parents: &[String]) -> bool {
        if !parents.is_empty()
            || !self.profile.namespaces.use_dotnet
            || !self.module_imports_namespace("System.Windows.Forms")
        {
            return false;
        }

        // Real VB/C# WinForms code commonly omits the explicit base type in
        // the user-authored partial while the surrounding project/designer
        // model still treats the class as a form. Keep the inference narrow:
        // only classes that follow the standard *Form / FormN naming shape
        // opt into the existing Form adapter wrapper.
        name.to_ascii_lowercase().contains("form")
    }

    pub(crate) fn reflection_constructor_for_types(
        &self,
        type_name: &str,
        param_types: &[String],
    ) -> Option<ReflectionBinding> {
        let lookup = self.reflection_type_lookup_name(type_name);
        let normalized_params: Vec<String> = param_types
            .iter()
            .map(|param| self.reflection_type_lookup_name(param))
            .collect();
        let meta = self.reflection_types.get(&lookup)?;
        let ctor = meta.constructors.iter().find(|ctor| {
            ctor.param_types.len() == normalized_params.len()
                && ctor
                    .param_types
                    .iter()
                    .zip(normalized_params.iter())
                    .all(|(left, right)| left.eq_ignore_ascii_case(right))
        })?;
        Some(ReflectionBinding::Constructor {
            type_name: lookup,
            param_types: ctor.param_types.clone(),
        })
    }

    // ════════════════════════════════════════════════════════════════════════
    // Multi-value tuple returns (opt-in via `multi_value_tuple_returns`)
    // ════════════════════════════════════════════════════════════════════════

    /// `true` if `iter` is an `ExprKind::Call` to an ident that names
    /// a function previously compiled with `is_generator = true`. Used
    /// by `for v in gen():` to pick the stack-switching iterator path.
    pub(super) fn is_direct_generator_call(&self, iter: &Expression) -> bool {
        if let ExprKind::Call { callee, .. } = &iter.kind {
            if let ExprKind::Ident(n) = &callee.kind {
                return self.generator_functions.contains(&self.canon(n));
            }
        }
        false
    }

    pub(super) fn js_private_member_storage_name_for_class(
        &self,
        owner_class: &str,
        field: &str,
    ) -> Option<String> {
        if !self.profile.supports_private_fields || !field.starts_with('#') {
            return None;
        }
        Some(format!(
            "__js_private_{}.{}",
            self.canon(owner_class),
            field.trim_start_matches('#')
        ))
    }

    pub(super) fn js_member_storage_name_for_class(
        &self,
        owner_class: &str,
        field: &str,
    ) -> String {
        self.js_private_member_storage_name_for_class(owner_class, field)
            .unwrap_or_else(|| self.canon(field))
    }

    pub(super) fn js_member_storage_name(&self, field: &str) -> String {
        self.current_class
            .as_deref()
            .map(|class_name| self.js_member_storage_name_for_class(class_name, field))
            .unwrap_or_else(|| self.canon(field))
    }

    pub(super) fn js_member_storage_name_for_receiver(
        &self,
        receiver: &Expression,
        field: &str,
    ) -> String {
        if !self.profile.supports_private_fields || !field.starts_with('#') {
            return self.js_member_storage_name(field);
        }

        if let Some(class_name) = self.current_class.as_deref() {
            if matches!(receiver.kind, ExprKind::This | ExprKind::Super) {
                return self.js_member_storage_name_for_class(class_name, field);
            }
        }

        let parts = self.flatten_member_chain(receiver);
        if !parts.is_empty() {
            let full_canon = self.canon(&parts.join("."));
            if self.defined_classes.contains(&full_canon)
                || self.pending_classes.contains_key(&full_canon)
            {
                return self.js_member_storage_name_for_class(&full_canon, field);
            }

            if let Some(short_name) = parts.last() {
                let short_canon = self.canon(short_name);
                if self.defined_classes.contains(&short_canon)
                    || self.pending_classes.contains_key(&short_canon)
                {
                    return self.js_member_storage_name_for_class(&short_canon, field);
                }
            }
        }

        self.js_member_storage_name(field)
    }

    /// Whether an instance field named `field_canon` is already declared by
    /// some ancestor (walking `parent` up the already-compiled
    /// `pending_classes` chain) — i.e. this declaration HIDES it. Used only
    /// under `profile.field_hiding`.
    pub(super) fn field_hides_ancestor(&self, parent: Option<&str>, field_canon: &str) -> bool {
        let mut current = parent.map(|p| self.canon(p));
        let mut guard = 0;
        while let Some(class_key) = current {
            guard += 1;
            if guard > 64 {
                break;
            }
            let Some(pending) = self.pending_classes.get(&class_key) else {
                break;
            };
            if pending.fields.iter().any(|f| f == field_canon)
                || pending
                    .field_storage_names
                    .keys()
                    .any(|orig| orig == field_canon)
            {
                return true;
            }
            current = pending.parent.as_ref().map(|p| self.canon(p));
        }
        false
    }

    /// Whether `class_key` is a descendant of `ancestor_key`.
    fn class_descends_from(&self, class_key: &str, ancestor_key: &str) -> bool {
        let mut current = self
            .pending_classes
            .get(class_key)
            .and_then(|pending| pending.parent.as_ref())
            .map(|p| self.canon(p));
        let mut guard = 0;
        while let Some(key) = current {
            guard += 1;
            if guard > 64 {
                break;
            }
            if key == ancestor_key {
                return true;
            }
            let Some(pending) = self.pending_classes.get(&key) else {
                break;
            };
            current = pending.parent.as_ref().map(|p| self.canon(p));
        }
        false
    }

    /// Whether any descendant of `class_key` declares `method_key`
    /// NON-virtually — hiding the ancestor's method (C# `new`, VB `Shadows`)
    /// instead of overriding it.
    ///
    /// A hiding method shares the ancestor's runtime slot, so a call through a
    /// declared-ancestor reference cannot be left to dynamic dispatch: the slot
    /// holds the hiding body, but the language says the DECLARED type's body
    /// must run (`Base b = new Derived(); b.Speak()` → Base's). Such a call
    /// keeps its direct bind. The method analogue of
    /// [`Self::field_hides_ancestor`] — same static-type rule, same reason.
    pub(super) fn method_hidden_by_descendant(&self, class_key: &str, method_key: &str) -> bool {
        self.pending_classes.iter().any(|(name, pending)| {
            name != class_key
                && pending
                    .instance_method_overloads
                    .get(method_key)
                    .is_some_and(|overloads| overloads.iter().any(|ov| !ov.is_virtual))
                && self.class_descends_from(name, class_key)
        })
    }

    /// The storage-slot name a class uses for `field`, when it differs from
    /// the plain field name. Data-driven via `PendingClass.field_storage_names`
    /// — a class with no remapped fields returns `None` (so this is a no-op
    /// for the common case, no language gate needed). Populated for PHP
    /// private properties and for statically-typed field hiding
    /// (java/C#/VB: `Parent.value` and a hiding `Child.value` occupy distinct
    /// slots, resolved by the reference's declared type).
    #[allow(dead_code)]
    pub(super) fn field_storage_name_for_class(
        &self,
        class_name: &str,
        field: &str,
    ) -> Option<String> {
        let class_key = self.canon(class_name);
        self.pending_classes
            .get(&class_key)
            .and_then(|pending| pending.field_storage_names.get(&self.canon(field)).cloned())
    }

    /// Resolve the storage slot for a source-level instance field visible
    /// through `class_name`'s static type. Unlike
    /// `field_storage_name_for_class`, this also walks ancestors and returns
    /// the plain field name for non-remapped fields.
    pub(super) fn visible_instance_field_storage_name_for_class(
        &self,
        class_name: &str,
        field: &str,
    ) -> Option<String> {
        let field_canon = self.canon(field);
        let mut current = Some(self.canon(class_name));
        let mut guard = 0;
        while let Some(class_key) = current {
            guard += 1;
            if guard > 64 {
                break;
            }
            let Some(pending) = self.pending_classes.get(&class_key) else {
                break;
            };
            if let Some(storage) = pending.field_storage_names.get(&field_canon) {
                return Some(storage.clone());
            }
            if pending.fields.iter().any(|stored| stored == &field_canon) {
                return Some(field_canon);
            }
            current = pending.parent.as_ref().map(|p| self.canon(p));
        }
        None
    }

    /// Resolve `receiver.field`'s storage slot by the receiver's STATIC type:
    /// `this` / `self` → the current class; `super` → its parent; a typed
    /// local → its declared type; otherwise the inferred type. This is what
    /// makes field hiding pick the declared-type field rather than the
    /// runtime one.
    pub(super) fn field_storage_name_for_receiver(
        &self,
        receiver: &Expression,
        field: &str,
    ) -> Option<String> {
        // JS `#private` fields have their own receiver-storage path.
        if field.starts_with('#') {
            return None;
        }
        let self_kw = self.profile.self_keyword.as_str();
        match &receiver.kind {
            ExprKind::This => self.current_class.as_deref().and_then(|class_name| {
                self.visible_instance_field_storage_name_for_class(class_name, field)
            }),
            ExprKind::Super => self
                .current_class
                .as_deref()
                .and_then(|class_name| self.pending_classes.get(&self.canon(class_name)))
                .and_then(|pending| pending.parent.clone())
                .and_then(|parent| {
                    self.visible_instance_field_storage_name_for_class(&parent, field)
                }),
            ExprKind::Ident(name)
                if name == self_kw || name == "$this" || name.eq_ignore_ascii_case(self_kw) =>
            {
                self.current_class.as_deref().and_then(|class_name| {
                    self.visible_instance_field_storage_name_for_class(class_name, field)
                })
            }
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .and_then(|type_hint| self.resolve_pending_class_name_for_type_hint(type_hint))
                .and_then(|class_name| {
                    self.visible_instance_field_storage_name_for_class(&class_name, field)
                }),
            _ => self
                .infer_expr_type_hint(receiver)
                .and_then(|type_hint| self.resolve_pending_class_name_for_type_hint(&type_hint))
                .and_then(|class_name| {
                    self.visible_instance_field_storage_name_for_class(&class_name, field)
                }),
        }
    }

    pub(super) fn private_member_access_forbidden(&self, field: &str) -> bool {
        self.profile.supports_private_fields
            && field.starts_with('#')
            && self.current_class.is_none()
    }

    pub(super) fn emit_private_access_denied(&mut self, field: &str) -> Result<(), String> {
        let message = format!("Cannot access private member {}", field);
        self.emit_const(Value::String(Arc::from(message.as_str())));
        self.emit_js_exception_ctor_from_message_value("TypeError")?;
        let line = self.line;
        common::errors::emit_throw(self.chunk(), line);
        Ok(())
    }

    pub(super) fn emit_js_private_brand_check(
        &mut self,
        object_slot: u16,
        storage_name: &str,
    ) -> Result<(), String> {
        self.emit_u16(Op::LOCAL_GET, object_slot);
        self.emit_const(Value::String(Arc::from(storage_name)));
        let has_own_idx = self.import("ecma:object", "hasOwn");
        let line = self.line;
        self.emit_host_call(has_own_idx, 2);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);
        self.chunk().emit_else(line);
        self.emit_const(Value::String(Arc::from(
            "Cannot read private member from an object whose class did not declare it",
        )));
        self.emit_js_exception_ctor_from_message_value("TypeError")?;
        common::errors::emit_throw(self.chunk(), line);
        self.chunk().emit_end(line);
        Ok(())
    }

    pub(super) fn emit_js_current_this_value(&mut self) {
        let self_kw = self.profile.self_keyword.clone();
        if let Some(slot) = self
            .scope()
            .resolve(&self_kw)
            .or_else(|| self.scope().resolve_ci(&self_kw))
        {
            self.emit_u16(Op::LOCAL_GET, slot);
        } else if self.scopes.len() > 1
            && self
                .resolve_upvalue(self.scopes.len() - 1, &self_kw)
                .is_some()
        {
            let env = self.closure_env_slot();
            let idx = self.closure_env_index(&self_kw);
            let line = self.line;
            crate::emitter::closures::emit_env_get(self.chunk(), env, idx, line);
        } else if self.scopes.len() > 1
            && self
                .resolve_upvalue(self.scopes.len() - 1, "__js_this")
                .is_some()
        {
            let env = self.closure_env_slot();
            let idx = self.closure_env_index("__js_this");
            let line = self.line;
            crate::emitter::closures::emit_env_get(self.chunk(), env, idx, line);
        } else if self.profile.ambient_this_binding {
            let js_this = self.str_const("__js_this");
            self.emit_u16(Op::GLOBAL_GET, js_this);
        } else {
            self.emit(Op::NULL);
        }
    }

    pub(super) fn emit_with_target_get(&mut self, name: &str) -> bool {
        let Some(slot) = self.with_targets.last().copied() else {
            return false;
        };
        self.emit_u16(Op::LOCAL_GET, slot);
        let idx = self.str_const(&self.canon(name));
        self.emit_u16(Op::STRUCT_GET, idx);
        true
    }

    pub(super) fn emit_with_target_set(&mut self, name: &str) -> bool {
        let Some(slot) = self.with_targets.last().copied() else {
            return false;
        };
        let value_slot = self.define_local("__with_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        let idx = self.str_const(&self.canon(name));
        self.emit_u16(Op::STRUCT_SET, idx);
        true
    }
}
