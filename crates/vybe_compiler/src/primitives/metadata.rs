//! Reflection/attribute metadata, member-storage names, private-access, WinForms inference.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
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

fn reflection_generic_params_from_method(
    name: &str,
    params: &[Param],
    return_type: Option<&str>,
) -> Vec<String> {
    let mut names = Vec::new();
    let mut add_candidate = |candidate: &str| {
        let trimmed = candidate.trim().trim_end_matches('?').trim();
        if trimmed.is_empty() {
            return;
        }
        let leaf = trimmed.rsplit('.').next().unwrap_or(trimmed);
        let is_generic_param = leaf.len() <= 3
            && leaf.chars().next().is_some_and(|ch| ch == 'T' || ch == 'U')
            && leaf
                .chars()
                .all(|ch| ch.is_ascii_uppercase() || ch.is_ascii_digit());
        if is_generic_param && !names.iter().any(|existing| existing == leaf) {
            names.push(leaf.to_string());
        }
    };

    if let Some(inner) = name
        .split_once("(Of")
        .and_then(|(_, rest)| rest.rsplit_once(')').map(|(inner, _)| inner))
    {
        for part in inner.split(',') {
            add_candidate(part);
        }
    }
    if let Some(return_type) = return_type {
        add_candidate(return_type);
    }
    for param in params {
        if let Some(type_hint) = param.type_hint.as_deref() {
            add_candidate(type_hint);
        }
    }
    names
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
        use vybe_runtime::component_model::ComponentItemKind;

        if self.profile.reflection_type_naming != ReflectionTypeNaming::Dotnet {
            return;
        }

        for export in vybe_runtime::registry::platform_component_descriptors()
            .into_iter()
            .flat_map(|d| d.exports)
        {
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
                    params: (0..constructor.arity)
                        .map(|index| ReflectionParamMetadata {
                            name: format!("arg{index}"),
                            decorators: Vec::new(),
                            type_name: Some("System.Object".to_string()),
                        })
                        .collect(),
                    decorators: Vec::new(),
                    visibility: Visibility::Public,
                    is_static: false,
                });
            }
            for property in class.properties {
                metadata.properties.insert(
                    property.name,
                    ReflectionMemberMetadata {
                        decorators: Vec::new(),
                        is_static: false,
                        can_write: property.setter.is_some(),
                        type_name: None,
                        params: Vec::new(),
                        visibility: Visibility::Public,
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
                                type_name: Some("System.Object".to_string()),
                            })
                            .collect(),
                        is_static: method.is_static,
                        return_type: None,
                        visibility: Visibility::Public,
                        is_abstract: false,
                        is_virtual: false,
                        generic_params: Vec::new(),
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
                modifiers,
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
                    modifiers.is_sealed,
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
                    false,
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
                    false,
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
        is_sealed: bool,
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
            is_sealed,
            ..ReflectionTypeMetadata::default()
        };
        let mut nested_types: Vec<&Statement> = Vec::new();

        for member in members {
            match member {
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name,
                        params,
                        return_type,
                        modifiers,
                        ..
                    } = &stmt.kind
                    {
                        if name == "__static_init__" {
                            metadata.constructors.push(ReflectionConstructorMetadata {
                                param_types: vec!["<static>".to_string()],
                                params: params
                                    .iter()
                                    .map(|param| ReflectionParamMetadata {
                                        name: param.name.clone(),
                                        decorators: Vec::new(),
                                        type_name: param.type_hint.as_deref().map(|type_name| {
                                            self.reflection_runtime_type_name(type_name, None)
                                        }),
                                    })
                                    .collect(),
                                decorators: modifiers.decorators.clone(),
                                visibility: modifiers.visibility,
                                is_static: true,
                            });
                        }
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
                                return_type: return_type.as_deref().map(|type_name| {
                                    self.reflection_runtime_type_name(type_name, None)
                                }),
                                visibility: modifiers.visibility,
                                is_abstract: modifiers.is_abstract,
                                is_virtual: modifiers.is_virtual || modifiers.is_override,
                                generic_params: reflection_generic_params_from_method(
                                    name,
                                    params,
                                    return_type.as_deref(),
                                ),
                                params: params
                                    .iter()
                                    .enumerate()
                                    .map(|(index, param)| ReflectionParamMetadata {
                                        name: param.name.clone(),
                                        decorators: param_decorators
                                            .remove(&index)
                                            .unwrap_or_default(),
                                        type_name: param.type_hint.as_deref().map(|type_name| {
                                            self.reflection_runtime_type_name(type_name, None)
                                        }),
                                    })
                                    .collect(),
                            },
                        );
                    }
                }
                ClassMember::Property {
                    name,
                    type_hint,
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
                            type_name: type_hint.as_deref().map(|type_name| {
                                self.reflection_runtime_type_name(type_name, None)
                            }),
                            params: Vec::new(),
                            visibility: modifiers.visibility,
                        },
                    );
                }
                ClassMember::Field {
                    name,
                    type_hint,
                    modifiers,
                    ..
                } => {
                    metadata.fields.insert(
                        name.clone(),
                        ReflectionMemberMetadata {
                            decorators: modifiers.decorators.clone(),
                            is_static: modifiers.is_static,
                            can_write: !modifiers.is_readonly,
                            type_name: type_hint.as_deref().map(|type_name| {
                                self.reflection_runtime_type_name(type_name, None)
                            }),
                            params: Vec::new(),
                            visibility: modifiers.visibility,
                        },
                    );
                }
                ClassMember::Constructor {
                    params, visibility, ..
                } => {
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
                        params: params
                            .iter()
                            .map(|param| ReflectionParamMetadata {
                                name: param.name.clone(),
                                decorators: Vec::new(),
                                type_name: param.type_hint.as_deref().map(|type_name| {
                                    self.reflection_runtime_type_name(type_name, None)
                                }),
                            })
                            .collect(),
                        decorators: Vec::new(),
                        visibility: *visibility,
                        is_static: false,
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

        let indexer_method = metadata
            .methods
            .iter()
            .find(|(name, _)| {
                let canon = self.canon(name);
                canon == "__call__"
                    || canon == "__getitem__"
                    || canon.starts_with("__call__")
                    || canon.starts_with("__getitem__")
            })
            .map(|(_, meta)| meta.clone());
        if let Some(method_meta) = indexer_method {
            if let Some(property_meta) = metadata.properties.get_mut("Item") {
                property_meta.params = method_meta.params;
            } else {
                metadata.properties.insert(
                    "Item".to_string(),
                    ReflectionMemberMetadata {
                        decorators: Vec::new(),
                        is_static: false,
                        can_write: metadata.methods.keys().any(|name| {
                            let canon = self.canon(name);
                            canon == "__setitem__" || canon.starts_with("__setitem__")
                        }),
                        type_name: None,
                        params: method_meta.params,
                        visibility: Visibility::Public,
                    },
                );
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
        let erased = common::generics::erased_type_name(trimmed);
        let base = erased.trim();
        let source_base = base.strip_prefix("System.").unwrap_or(base);
        if parent_runtime_name.is_none() {
            if let Some(resolved) = self
                .resolve_source_namespace_type(base)
                .or_else(|| self.resolve_source_namespace_type(source_base))
            {
                return resolved;
            }
        }

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
                    "int" | "Integer" | "Int32" => "Int32",
                    "uint" | "UInt32" => "UInt32",
                    "long" | "Long" | "Int64" => "Int64",
                    "ulong" | "UInt64" => "UInt64",
                    "short" | "Short" | "Int16" => "Int16",
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
            if let Some(resolved) = self.resolve_source_namespace_type(&raw_name) {
                return Some(self.reflection_runtime_type_name(&resolved, None));
            }
            if !raw_name.ends_with("Attribute") {
                let attr_name = format!("{raw_name}Attribute");
                if let Some(resolved) = self.resolve_source_namespace_type(&attr_name) {
                    return Some(self.reflection_runtime_type_name(&resolved, None));
                }
            }
            let attr_leaf = if raw_name.ends_with("Attribute") {
                raw_name.clone()
            } else {
                format!("{raw_name}Attribute")
            };
            let attr_suffix = format!(".{attr_leaf}");
            let mut class_matches = self.defined_classes.iter().filter(|known| {
                known.eq_ignore_ascii_case(&raw_name)
                    || known.eq_ignore_ascii_case(&attr_leaf)
                    || known.ends_with(&attr_suffix)
            });
            if let Some(resolved) = class_matches.next().cloned() {
                if class_matches.next().is_none() {
                    return Some(self.reflection_runtime_type_name(&resolved, None));
                }
            }
            let mut matches: Vec<String> = self
                .reflection_types
                .keys()
                .filter(|known| {
                    known.rsplit('.').next().is_some_and(|leaf| {
                        leaf.eq_ignore_ascii_case(&raw_name)
                            || leaf.eq_ignore_ascii_case(&format!("{raw_name}Attribute"))
                    })
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
        let trimmed = type_name.trim().trim_end_matches('?').trim();
        if self.reflection_types.contains_key(trimmed) {
            return trimmed.to_string();
        }
        if let Some(existing) = self
            .reflection_types
            .keys()
            .find(|known| known.eq_ignore_ascii_case(trimmed))
        {
            return existing.clone();
        }
        if let Some(system_stripped) = trimmed.strip_prefix("System.") {
            if self.reflection_types.contains_key(system_stripped) {
                return system_stripped.to_string();
            }
            if let Some(existing) = self
                .reflection_types
                .keys()
                .find(|known| known.eq_ignore_ascii_case(system_stripped))
            {
                return existing.clone();
            }
        }
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
        let erased = common::generics::erased_type_name(type_name);
        let trimmed = erased.trim().trim_end_matches('?').trim();
        let short = trimmed
            .rsplit('.')
            .next()
            .unwrap_or(trimmed)
            .trim()
            .to_string();
        if !trimmed.starts_with("System.")
            && short.chars().all(|ch| !ch.is_ascii_uppercase())
            && short.chars().any(|ch| ch.is_ascii_alphabetic())
        {
            let mut chars = short.chars();
            if let Some(first) = chars.next() {
                return first.to_ascii_uppercase().to_string() + chars.as_str();
            }
        }
        short
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

    pub(crate) fn reflection_is_sealed_type(&self, type_name: &str) -> bool {
        let lookup = self.reflection_type_lookup_name(type_name);
        self.reflection_types
            .get(&lookup)
            .is_some_and(|meta| meta.is_sealed)
    }

    pub(crate) fn reflection_declaring_type_name(&self, type_name: &str) -> Option<String> {
        let lookup = self.reflection_type_lookup_name(type_name);
        let parent = lookup.rsplit_once('.')?.0;
        self.reflection_types
            .contains_key(parent)
            .then(|| parent.to_string())
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
        common::generics::generic_argument_type_refs(type_name)
            .into_iter()
            .map(|arg| self.reflection_type_full_name(&common::generics::display_type_ref(&arg)))
            .collect()
    }

    pub(crate) fn reflection_open_generic_type_name(&self, type_name: &str) -> String {
        let open = common::generics::open_generic_type_name(type_name);
        self.reflection_type_lookup_name(&open)
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
        // The IMPORT is the signal. A unit that imports `System.Windows.Forms`
        // has the WinForms surface in scope by saying so; `use_dotnet` added
        // nothing but a language-family name in front of a question the import
        // already answers.
        if !parents.is_empty() || !self.module_imports_namespace("System.Windows.Forms") {
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

    pub(super) fn method_hides_ancestor(&self, parent: Option<&str>, method_canon: &str) -> bool {
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
            if pending.instance_method_overloads.contains_key(method_canon)
                || pending
                    .instance_member_names
                    .iter()
                    .any(|name| name == method_canon)
            {
                return true;
            }
            current = pending.parent.as_ref().map(|p| self.canon(p));
        }
        false
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
        // Probe with `ecma:object.has` (own + prototype-chain walk, raw key
        // lookup) rather than `hasOwn`. The private storage key is
        // `__js_private_<Class>.<name>` — a `__`-prefixed key that the
        // user-facing `hasOwn` deliberately hides (it returns false for any
        // `__`-prefixed key so internal markers don't leak through
        // `Object.hasOwn`). `has` does a raw lookup and also walks the
        // prototype, which is required under `class_method_dispatch =
        // "prototype"` where private *methods* live on the class prototype,
        // not as own instance properties. This makes the brand check see the
        // private member for both fields (own) and methods (proto).
        let has_idx = self.import("ecma:object", "has");
        let line = self.line;
        self.emit_host_call(has_idx, 2);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
        // Inside a class method under ambient `this`, the live receiver is the
        // `__js_this` global — exactly what `ExprKind::This` reads
        // (expressions.rs). Falling through to the local/upvalue search here
        // instead synthesizes an upvalue over a closure env the method never
        // allocated, so `super.m()` handed the callee an undefined receiver
        // whenever the method body had no closure to force a real `this`
        // local. Constructors keep the slot-based path (derived-ctor TDZ);
        // lambdas keep the upvalue path (they capture the enclosing `this`).
        if self.ambient_this()
            && self.current_class.is_some()
            && self.current_func_name.as_deref() != Some("<lambda>")
            && self
                .current_func_name
                .as_deref()
                .is_some_and(|name| !name.eq_ignore_ascii_case(&self.profile.constructor_name))
        {
            self.emit_global_read("__js_this");
            return;
        }

        let self_kw = self.profile.self_keyword.clone();
        if let Some(slot) = self.scope().resolve(&self_kw) {
            self.emit_u16(Op::LOCAL_GET, slot);
        } else if self.scopes.len() > 1
            && self
                .resolve_upvalue(self.scopes.len() - 1, &self_kw)
                .is_some()
        {
            let env = self.closure_env_slot();
            let idx = self.closure_env_index(&self_kw);
            let line = self.line;
            crate::primitives::closures::emit_env_get(self.chunk(), env, idx, line);
        } else if self.scopes.len() > 1
            && self
                .resolve_upvalue(self.scopes.len() - 1, "__js_this")
                .is_some()
        {
            let env = self.closure_env_slot();
            let idx = self.closure_env_index("__js_this");
            let line = self.line;
            crate::primitives::closures::emit_env_get(self.chunk(), env, idx, line);
        } else if self.ambient_this() {
            self.emit_global_read("__js_this");
        } else {
            self.emit_null();
        }
    }

    /// `Caption` inside `with L do …` IS `L.Caption`.
    ///
    /// So it compiles as exactly that — a synthesised member access on the
    /// with-binding, handed to the same `compile_expr` / `compile_assign_target`
    /// every explicit `L.Caption` goes through. One primitive, so a declared
    /// property, a GUI role, an accessor pair and a plain struct field all
    /// behave identically inside and outside the block, and nothing has to be
    /// taught about `with` twice.
    ///
    /// Both directions previously emitted a RAW `STRUCT_GET`/`STRUCT_SET` of the
    /// name. That is right for a plain object with real fields — which is why
    /// it survived — and silently wrong for everything else: the value landed
    /// on a phantom field of that name and the matching read read it straight
    /// back, so the block round-tripped and looked correct while the real
    /// property was never touched. Measured: `with L do Caption := 'x'`
    /// followed by `L.Caption` outside answered the OLD value. In a form it
    /// meant `Parent := Self` wrote a phantom `parent` instead of taking the
    /// role's `appendChild`, so a control built inside a `with` block was
    /// created, configured, and never inserted — two labels in
    /// `examples/pascal/delphi_project` are missing for exactly this reason.
    ///
    /// Shared, not Pascal-only: `StmtKind::With` is the same node VB's `With`
    /// block lowers to.
    fn with_target_member(&self, name: &str) -> Option<Expression> {
        let binding = self.with_targets.last()?;
        Some(Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::Ident(binding.clone()))),
            field: name.to_string(),
            null_safe: false,
        }))
    }

    pub(super) fn emit_with_target_get(&mut self, name: &str) -> bool {
        let Some(member) = self.with_target_member(name) else {
            return false;
        };
        // The member path is infallible for a synthesised access — the object
        // is a binding this compiler just defined — so a failure here is a
        // compiler bug, not a program error, and reporting `false` would
        // silently fall through to a global read.
        self.compile_expr(&member).is_ok()
    }

    pub(super) fn emit_with_target_set(&mut self, name: &str) -> bool {
        let Some(member) = self.with_target_member(name) else {
            return false;
        };
        // The value is already on the stack, which is the contract
        // `compile_assign_target` expects.
        self.compile_assign_target(&member).is_ok()
    }
}
