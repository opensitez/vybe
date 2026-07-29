//! .NET reflection compilation — Type/Method/Property/Field/Constructor/
//! Parameter reflection, custom attributes, and Invoke. Compile-time
//! resolution against class/attribute metadata. Moved out of the former dotnet_calls.rs.

use super::*;
use crate::primitives::calls::{strip_generic_suffix, terminal_type_name};

impl Compiler {
    pub(crate) fn resolve_reflection_binding_expr(
        &self,
        expr: &Expression,
    ) -> Option<ReflectionBinding> {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(type_name)) if type_name.starts_with("System.") => {
                Some(ReflectionBinding::Type(type_name.clone()))
            }
            ExprKind::TypeOf(inner) => {
                let raw_name = match &inner.kind {
                    ExprKind::Ident(name) => name.clone(),
                    ExprKind::Member { .. } => self.flatten_member_chain(inner).join("."),
                    _ => return None,
                };
                Some(ReflectionBinding::Type(
                    self.reflection_runtime_type_name(&raw_name, None),
                ))
            }
            ExprKind::Ident(name) => self.reflection_bindings.get(&self.canon(name)).cloned(),
            ExprKind::Member { object, field, .. } => {
                let receiver = self.resolve_reflection_binding_expr(object)?;
                match (receiver, strip_generic_suffix(field.as_str())) {
                    (ReflectionBinding::Type(type_name), "BaseType") => self
                        .reflection_base_type_name(&type_name)
                        .map(ReflectionBinding::Type),
                    (ReflectionBinding::Type(type_name), "DeclaringType") => self
                        .reflection_declaring_type_name(&type_name)
                        .map(ReflectionBinding::Type),
                    (ReflectionBinding::Type(_), "Assembly") => Some(ReflectionBinding::Assembly),
                    _ => None,
                }
            }
            ExprKind::Call { callee, args, .. } => {
                let ExprKind::Member { object, field, .. } = &callee.kind else {
                    return None;
                };
                let receiver = self.resolve_reflection_binding_expr(object)?;
                match (receiver, strip_generic_suffix(field.as_str())) {
                    (ReflectionBinding::Type(type_name), "GetMethod") => {
                        let method_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Method {
                            type_name,
                            method_name,
                        })
                    }
                    (ReflectionBinding::Type(type_name), "GetProperty") => {
                        let property_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Property {
                            type_name,
                            property_name,
                        })
                    }
                    (ReflectionBinding::Type(type_name), "GetField") => {
                        let field_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Field {
                            type_name,
                            field_name,
                        })
                    }
                    (ReflectionBinding::Type(type_name), "GetNestedType") => {
                        let nested_name = self.resolve_reflection_string_arg(args.first()?)?;
                        self.reflection_nested_type_name(&type_name, &nested_name)
                            .map(ReflectionBinding::Type)
                    }
                    (ReflectionBinding::Type(type_name), "GetGenericTypeDefinition") => Some(
                        ReflectionBinding::Type(self.reflection_open_generic_type_name(&type_name)),
                    ),
                    (ReflectionBinding::Type(type_name), "GetConstructor") => {
                        let param_types =
                            self.resolve_reflection_type_array_expr(&args.first()?.value)?;
                        self.reflection_constructor_for_types(&type_name, &param_types)
                    }
                    (ReflectionBinding::Assembly, "GetName") if args.is_empty() => {
                        Some(ReflectionBinding::AssemblyName)
                    }
                    _ => None,
                }
            }
            ExprKind::Index { object, index, .. } => {
                let ExprKind::Call { callee, args, .. } = &object.kind else {
                    return None;
                };
                if !args.is_empty() {
                    return None;
                }
                let ExprKind::Member {
                    object: method_object,
                    field,
                    ..
                } = &callee.kind
                else {
                    return None;
                };
                if strip_generic_suffix(field.as_str()) != "GetParameters" {
                    return None;
                }
                let ReflectionBinding::Method {
                    type_name,
                    method_name,
                } = self.resolve_reflection_binding_expr(method_object)?
                else {
                    return None;
                };
                let ExprKind::Lit(Literal::Int(position)) = &index.kind else {
                    return None;
                };
                Some(ReflectionBinding::Parameter {
                    type_name,
                    method_name,
                    index: (*position).max(0) as usize,
                })
            }
            _ => None,
        }
    }

    pub(super) fn resolve_reflection_string_arg(&self, arg: &Argument) -> Option<String> {
        match &arg.value.kind {
            ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
            ExprKind::Ident(name) => {
                self.reflection_bindings
                    .get(&self.canon(name))
                    .and_then(|binding| {
                        if let ReflectionBinding::Type(type_name) = binding {
                            Some(type_name.clone())
                        } else {
                            None
                        }
                    })
            }
            _ => None,
        }
    }

    pub(super) fn resolve_reflection_type_arg(&self, expr: &Expression) -> Option<String> {
        match self.resolve_reflection_binding_expr(expr)? {
            ReflectionBinding::Type(type_name) => Some(type_name),
            _ => None,
        }
    }

    pub(super) fn resolve_reflection_type_array_expr(
        &self,
        expr: &Expression,
    ) -> Option<Vec<String>> {
        match &expr.kind {
            ExprKind::Array(items) => items
                .iter()
                .map(|item| self.resolve_reflection_type_arg(&item.value))
                .collect(),
            ExprKind::Lit(Literal::Null) => Some(Vec::new()),
            _ => None,
        }
    }

    pub(super) fn resolve_reflection_invoke_args(
        &self,
        expr: &Expression,
    ) -> Option<Vec<Argument>> {
        match &expr.kind {
            ExprKind::Lit(Literal::Null) => Some(Vec::new()),
            ExprKind::Array(items) => Some(
                items
                    .iter()
                    .map(|item| Argument::positional(item.value.clone()))
                    .collect(),
            ),
            _ => None,
        }
    }

    pub(super) fn resolve_reflection_string_member_expr(
        &self,
        expr: &Expression,
    ) -> Option<String> {
        let ExprKind::Member { object, field, .. } = &expr.kind else {
            return None;
        };
        match (
            self.resolve_reflection_binding_expr(object)?,
            strip_generic_suffix(field.as_str()),
        ) {
            (ReflectionBinding::Type(type_name), "Name") => {
                Some(self.reflection_type_short_name(&type_name))
            }
            (ReflectionBinding::Type(type_name), "FullName") => {
                Some(self.reflection_type_full_name(&type_name))
            }
            (ReflectionBinding::AssemblyName, "Name") => Some("main".to_string()),
            _ => None,
        }
    }

    pub(super) fn reflection_class_expr(&self, type_name: &str) -> Expression {
        let trimmed = type_name.trim().trim_end_matches('?').trim();
        let without_system = trimmed.strip_prefix("System.").unwrap_or(trimmed);
        let mut parts = without_system.split('.').filter(|part| !part.is_empty());
        let first = parts.next().unwrap_or(without_system);
        let mut expr = Expression::ident(first);
        for part in parts {
            expr = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: part.to_string(),
                null_safe: false,
            });
        }
        expr
    }

    pub(crate) fn compile_reflection_type_value(&mut self, type_name: &str) -> Result<(), String> {
        let short_name = self.reflection_type_short_name(type_name);
        let full_name = self.reflection_type_full_name(type_name);
        let is_enum = self.reflection_is_enum_type(type_name);
        let is_value_type = self.reflection_is_value_type(type_name);
        let is_sealed = self.reflection_is_sealed_type(type_name);
        let declaring_type = self
            .reflection_declaring_type_name(type_name)
            .map(|parent| self.reflection_class_expr(&parent))
            .unwrap_or_else(Expression::null);
        self.compile_expr(&Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("Name"),
                value: Expression::string(&short_name),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("FullName"),
                value: Expression::string(&full_name),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("IsEnum"),
                value: Expression::bool(is_enum),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("IsValueType"),
                value: Expression::bool(is_value_type),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("IsSealed"),
                value: Expression::bool(is_sealed),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("DeclaringType"),
                value: declaring_type,
            },
        ])))
    }

    pub(super) fn compile_reflection_type_array(
        &mut self,
        type_names: &[String],
    ) -> Result<(), String> {
        let line = self.line;
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        for type_name in type_names {
            inst!(self, core_wasm::dup);
            self.compile_reflection_type_value(type_name)?;
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    pub(super) fn reflection_attributes_for_binding(
        &self,
        binding: &ReflectionBinding,
        attribute_type: Option<&str>,
        inherit: bool,
    ) -> Vec<Expression> {
        match binding {
            ReflectionBinding::Type(type_name) => {
                self.reflection_attributes_for_type(type_name, attribute_type, inherit)
            }
            ReflectionBinding::Constructor { .. } => Vec::new(),
            ReflectionBinding::Assembly | ReflectionBinding::AssemblyName => Vec::new(),
            ReflectionBinding::Method {
                type_name,
                method_name,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.methods.get(method_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Property {
                type_name,
                property_name,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.properties.get(property_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Field {
                type_name,
                field_name,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.fields.get(field_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Parameter {
                type_name,
                method_name,
                index,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.methods.get(method_name))
                .and_then(|meta| meta.params.get(*index))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
        }
    }

    pub(super) fn reflection_attributes_for_type(
        &self,
        type_name: &str,
        attribute_type: Option<&str>,
        inherit: bool,
    ) -> Vec<Expression> {
        let mut attrs = Vec::new();
        let mut current = Some(type_name.to_string());

        while let Some(current_type) = current {
            let Some(meta) = self.reflection_types.get(&current_type) else {
                break;
            };
            let matching = self.filter_reflection_attributes(&meta.decorators, attribute_type);
            if !matching.is_empty() {
                if let Some(attribute_type) = attribute_type {
                    let usage = self
                        .attribute_usage
                        .get(attribute_type)
                        .copied()
                        .unwrap_or_default();
                    if usage.allow_multiple {
                        attrs.extend(matching);
                    } else {
                        attrs.push(matching[0].clone());
                        break;
                    }
                } else {
                    attrs.extend(matching);
                }
            }

            if !inherit {
                break;
            }
            let should_inherit = attribute_type
                .and_then(|name| self.attribute_usage.get(name))
                .copied()
                .unwrap_or_default()
                .inherited;
            if !should_inherit {
                break;
            }
            current = meta.parents.first().cloned();
        }

        attrs
    }

    pub(super) fn filter_reflection_attributes(
        &self,
        decorators: &[Expression],
        attribute_type: Option<&str>,
    ) -> Vec<Expression> {
        decorators
            .iter()
            .filter(|decorator| {
                attribute_type.is_none_or(|wanted| {
                    self.reflection_attribute_type_name(decorator)
                        .is_some_and(|actual| actual.eq_ignore_ascii_case(wanted))
                })
            })
            .cloned()
            .collect()
    }

    pub(super) fn compile_reflection_attribute_instance(
        &mut self,
        attr: &Expression,
    ) -> Result<(), String> {
        let ExprKind::New { class, args } = &attr.kind else {
            return self.compile_expr(attr);
        };

        if let Some(type_name) = self.reflection_attribute_type_name(attr) {
            let has_runtime_shape = self.reflection_types.get(&type_name).is_some_and(|meta| {
                !meta.constructors.is_empty()
                    || !meta.fields.is_empty()
                    || !meta.properties.is_empty()
            });
            if type_name.starts_with("System.") && !has_runtime_shape {
                let short_name = type_name.rsplit('.').next().unwrap_or(&type_name);
                return self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("__typename"),
                        value: Expression::string(&type_name),
                    },
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(short_name),
                    },
                ])));
            }
        }

        let positional_args: Vec<Argument> = args
            .iter()
            .filter(|arg| arg.name.is_none())
            .cloned()
            .collect();
        let named_args: Vec<&Argument> = args.iter().filter(|arg| arg.name.is_some()).collect();
        if named_args.is_empty() {
            return self.compile_expr(attr);
        }

        self.compile_expr(&Expression::new(ExprKind::New {
            class: class.clone(),
            args: positional_args,
        }))?;
        let slot = self.define_local("__reflection_attr");
        self.emit_u16(Op::LOCAL_SET, slot);

        for arg in named_args {
            self.compile_expr(&arg.value)?;
            self.compile_assign_target(&Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__reflection_attr")),
                field: arg.name.clone().unwrap_or_default(),
                null_safe: false,
            }))?;
        }

        self.emit_u16(Op::LOCAL_GET, slot);
        Ok(())
    }

    pub(super) fn compile_reflection_attribute_array(
        &mut self,
        attrs: &[Expression],
    ) -> Result<(), String> {
        let line = self.line;
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        for attr in attrs {
            inst!(self, core_wasm::dup);
            self.compile_reflection_attribute_instance(attr)?;
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    pub(super) fn compile_reflection_binding_value(
        &mut self,
        binding: &ReflectionBinding,
    ) -> Result<(), String> {
        match binding {
            ReflectionBinding::Type(type_name) => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(type_name),
                    },
                ])))?;
            }
            ReflectionBinding::Constructor { type_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(type_name),
                    },
                ])))?;
            }
            ReflectionBinding::Method { method_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(method_name),
                    },
                ])))?;
            }
            ReflectionBinding::Property { property_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(property_name),
                    },
                ])))?;
            }

            ReflectionBinding::Field { field_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(field_name),
                    },
                ])))?;
            }
            ReflectionBinding::Assembly => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string("main"),
                    },
                ])))?;
            }
            ReflectionBinding::AssemblyName => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string("main"),
                    },
                ])))?;
            }
            ReflectionBinding::Parameter {
                type_name,
                method_name,
                index,
            } => {
                let name = self
                    .reflection_types
                    .get(type_name)
                    .and_then(|meta| meta.methods.get(method_name))
                    .and_then(|meta| meta.params.get(*index))
                    .map(|param| param.name.clone())
                    .unwrap_or_default();
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(&name),
                    },
                    ObjectProperty::KeyValue {
                        key: Expression::string("Position"),
                        value: Expression::int(*index as i64),
                    },
                ])))?;
            }
        }
        Ok(())
    }

    pub(super) fn try_compile_dotnet_attribute_reflection_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        let field_name = strip_generic_suffix(field);
        let receiver_type = terminal_type_name(object).unwrap_or_default();

        if (receiver_type.eq_ignore_ascii_case("Activator")
            || receiver_type.eq_ignore_ascii_case("System.Activator"))
            && field_name == "CreateInstance"
            && !args.is_empty()
        {
            let Some(type_name) = self.resolve_reflection_type_arg(&args[0].value) else {
                return Ok(false);
            };
            self.compile_expr(&Expression::new(ExprKind::New {
                class: Box::new(self.reflection_class_expr(&type_name)),
                args: Vec::new(),
            }))?;
            return Ok(true);
        }

        if (receiver_type.eq_ignore_ascii_case("Attribute")
            || receiver_type.eq_ignore_ascii_case("System.Attribute"))
            && field_name == "GetCustomAttribute"
            && args.len() >= 2
        {
            let Some(provider) = self.resolve_reflection_binding_expr(&args[0].value) else {
                return Ok(false);
            };
            let Some(attribute_type) = self.resolve_reflection_type_arg(&args[1].value) else {
                return Ok(false);
            };
            let attrs =
                self.reflection_attributes_for_binding(&provider, Some(&attribute_type), true);
            if let Some(attr) = attrs.first() {
                self.compile_reflection_attribute_instance(attr)?;
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }

        if (receiver_type.eq_ignore_ascii_case("Attribute")
            || receiver_type.eq_ignore_ascii_case("System.Attribute"))
            && field_name == "IsDefined"
            && args.len() >= 2
        {
            let Some(provider) = self.resolve_reflection_binding_expr(&args[0].value) else {
                return Ok(false);
            };
            let Some(attribute_type) = self.resolve_reflection_type_arg(&args[1].value) else {
                return Ok(false);
            };
            let attrs =
                self.reflection_attributes_for_binding(&provider, Some(&attribute_type), true);
            inst!(self, core_wasm::bool_const, !attrs.is_empty());
            return Ok(true);
        }

        let Some(provider) = self.resolve_reflection_binding_expr(object) else {
            return Ok(false);
        };
        match field_name {
            "GetMethod" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetProperty" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetField" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetConstructor" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetNestedType" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    self.emit(Op::NULL);
                    return Ok(true);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetGenericArguments" if args.is_empty() => {
                let ReflectionBinding::Type(type_name) = provider else {
                    return Ok(false);
                };
                let args = self.reflection_generic_argument_types(&type_name);
                self.compile_reflection_type_array(&args)?;
                Ok(true)
            }
            "GetGenericTypeDefinition" if args.is_empty() => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetInterfaces" if args.is_empty() => {
                let ReflectionBinding::Type(type_name) = provider else {
                    return Ok(false);
                };
                let interfaces = self.reflection_interfaces(&type_name);
                self.compile_reflection_type_array(&interfaces)?;
                Ok(true)
            }
            "IsAssignableFrom" if args.len() >= 1 => {
                let ReflectionBinding::Type(type_name) = provider else {
                    return Ok(false);
                };
                let Some(other_type) = self.resolve_reflection_type_arg(&args[0].value) else {
                    return Ok(false);
                };
                let v = self.reflection_is_assignable_from(&type_name, &other_type);
                inst!(self, core_wasm::bool_const, v);
                Ok(true)
            }
            "GetParameters" if args.is_empty() => {
                let ReflectionBinding::Method {
                    type_name,
                    method_name,
                } = provider
                else {
                    return Ok(false);
                };
                let params = self
                    .reflection_types
                    .get(&type_name)
                    .and_then(|meta| meta.methods.get(&method_name))
                    .map(|meta| meta.params.clone())
                    .unwrap_or_default();
                let line = self.line;
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                for (index, param) in params.iter().enumerate() {
                    inst!(self, core_wasm::dup);
                    self.compile_expr(&Expression::new(ExprKind::Object(vec![
                        ObjectProperty::KeyValue {
                            key: Expression::string("Name"),
                            value: Expression::string(&param.name),
                        },
                        ObjectProperty::KeyValue {
                            key: Expression::string("Position"),
                            value: Expression::int(index as i64),
                        },
                    ])))?;
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                }
                Ok(true)
            }
            "GetCustomAttributes" if !args.is_empty() => {
                let (attribute_type, inherit) = if args.len() >= 2 {
                    let Some(attribute_type) = self.resolve_reflection_type_arg(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    (
                        Some(attribute_type),
                        matches!(args[1].value.kind, ExprKind::Lit(Literal::Bool(true))),
                    )
                } else {
                    (
                        None,
                        matches!(args[0].value.kind, ExprKind::Lit(Literal::Bool(true))),
                    )
                };
                let attrs = self.reflection_attributes_for_binding(
                    &provider,
                    attribute_type.as_deref(),
                    inherit,
                );
                self.compile_reflection_attribute_array(&attrs)?;
                Ok(true)
            }
            "Invoke" => match provider {
                ReflectionBinding::Constructor { type_name, .. } => {
                    let ctor_args = args
                        .first()
                        .and_then(|arg| self.resolve_reflection_invoke_args(&arg.value))
                        .unwrap_or_default();
                    self.compile_expr(&Expression::new(ExprKind::New {
                        class: Box::new(self.reflection_class_expr(&type_name)),
                        args: ctor_args,
                    }))?;
                    Ok(true)
                }
                ReflectionBinding::Method { method_name, .. } => {
                    let Some(instance_arg) = args.first() else {
                        return Ok(false);
                    };
                    let invoke_args = args
                        .get(1)
                        .and_then(|arg| self.resolve_reflection_invoke_args(&arg.value))
                        .unwrap_or_default();
                    self.compile_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(instance_arg.value.clone()),
                            field: method_name,
                            null_safe: false,
                        })),
                        args: invoke_args,
                        optional: false,
                    }))?;
                    Ok(true)
                }
                _ => Ok(false),
            },
            "GetValue" if !args.is_empty() => match provider {
                ReflectionBinding::Property { property_name, .. }
                | ReflectionBinding::Field {
                    field_name: property_name,
                    ..
                } => {
                    self.compile_expr(&Expression::new(ExprKind::Member {
                        object: Box::new(args[0].value.clone()),
                        field: property_name,
                        null_safe: false,
                    }))?;
                    Ok(true)
                }
                _ => Ok(false),
            },
            "SetValue" if args.len() >= 2 => match provider {
                ReflectionBinding::Property { property_name, .. } => {
                    self.compile_expr(&args[1].value)?;
                    self.compile_assign_target(&Expression::new(ExprKind::Member {
                        object: Box::new(args[0].value.clone()),
                        field: property_name,
                        null_safe: false,
                    }))?;
                    self.emit(Op::NULL);
                    Ok(true)
                }
                ReflectionBinding::Field { field_name, .. } => {
                    if let ExprKind::Ident(name) = &args[0].value.kind {
                        let value_slot = self.define_local("__reflection_field_value");
                        self.compile_expr(&args[1].value)?;
                        self.emit_u16(Op::LOCAL_SET, value_slot);

                        self.compile_expr(&args[0].value)?;
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        let field_idx = self.str_const(&self.canon(&field_name));
                        self.emit_u16(Op::STRUCT_SET, field_idx);
                        self.emit(Op::DROP);
                        self.emit_var_set(name);
                        self.emit(Op::NULL);
                        Ok(true)
                    } else {
                        self.compile_expr(&args[1].value)?;
                        self.compile_assign_target(&Expression::new(ExprKind::Member {
                            object: Box::new(args[0].value.clone()),
                            field: field_name,
                            null_safe: false,
                        }))?;
                        self.emit(Op::NULL);
                        Ok(true)
                    }
                }
                _ => Ok(false),
            },
            _ => Ok(false),
        }
    }
}

// ── Chunk-level reflection emit ────────────────────────────────────────────
// Free functions over `&mut Chunk`, merged in from the former `emitter::reflection`
// module. The `impl Compiler` walkers above and these primitives are the two
// halves of the same topic and now live in one file.
use crate::primitives::collections;
use std::sync::Arc;
use vybe_ast::{ArrayElement, ExprKind, Expression, Literal};
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
// Shared reflection substrate for language adapters.
//
// This module is intentionally not "JavaScript reflection". Vybe's runtime
// values are ECMA-shaped enough that `ecma:reflect`, `ecma:object`, and
// `ecma:value` are the portable primitive operations, but each source language
// still owns its public API:
//
// - JavaScript maps these helpers to `typeof`, `instanceof`, `Reflect.*`, and
//   `Object.*` with ECMA quirks such as prototypes and property descriptors.
// - PHP maps them to `gettype`, `is_*`, `get_class`, `is_a`, ReflectionClass,
//   attributes, visibility filters, and dynamic properties.
// - Go maps them to `reflect.Type` / `reflect.Value`, Kind, Elem, CanSet,
//   struct tags, and pointer/ref-aware mutation.
//
// The shared contract is the hidden metadata/stamp shape plus bytecode recipes
// for reading and writing live values. Declaration metadata (fields, methods,
// attributes/tags) is carried by language/compiler metadata; this module only
// emits the compatible runtime objects that expose it.

pub const FIELD_TYPE: &str = "__type";
pub const FIELD_TYPES: &str = "__types";
pub const FIELD_TYPE_ID: &str = "__type_id";
pub const FIELD_TYPE_NAME: &str = "__typename";
pub const FIELD_FIELDS: &str = "__fields";
pub const FIELD_FIELDS_PUBLIC: &str = "__fields_public";
pub const FIELD_METHODS: &str = "__methods";
pub const FIELD_METHODS_PUBLIC: &str = "__methods_public";
pub const FIELD_ATTRIBUTES: &str = "__attributes";
pub const FIELD_TAGS: &str = "__tags";
pub const FIELD_KIND: &str = "__kind";
pub const FIELD_VALUE: &str = "__value";
pub const FIELD_REF: &str = "__ref";

pub const MEMBER_KIND_FIELD: &str = "field";
pub const MEMBER_KIND_METHOD: &str = "method";
pub const MEMBER_KIND_CONSTRUCTOR: &str = "constructor";

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MemberToken {
    pub kind: String,
    pub owner: String,
    pub name: String,
    pub param_count: usize,
    pub type_name: Option<String>,
    pub return_type: Option<String>,
    pub param_types: Vec<String>,
    pub modifiers: i64,
}

/// Build the shared compile-time reflection member token used by language
/// walkers when declaration metadata is known. Language surfaces may wrap this
/// in their own public objects, but the slot order is shared:
/// `[kind, owner, name, param_count, type_name, return_type, param_types, modifiers]`.
pub fn member_token_expr(
    kind: &str,
    owner: &str,
    name: &str,
    param_count: usize,
    type_name: Option<String>,
    return_type: Option<String>,
    param_types: Vec<String>,
    modifiers: i64,
) -> Expression {
    Expression::new(ExprKind::Array(vec![
        array_value(Expression::string(kind)),
        array_value(Expression::string(owner)),
        array_value(Expression::string(name)),
        array_value(Expression::int(param_count as i64)),
        array_value(
            type_name
                .map(|name| Expression::string(&name))
                .unwrap_or_else(Expression::null),
        ),
        array_value(
            return_type
                .map(|name| Expression::string(&name))
                .unwrap_or_else(Expression::null),
        ),
        array_value(string_array_expr(param_types)),
        array_value(Expression::int(modifiers)),
    ]))
}

pub fn member_token(expr: &Expression) -> Option<MemberToken> {
    let ExprKind::Array(elems) = &expr.kind else {
        return None;
    };
    Some(MemberToken {
        kind: token_string(elems, 0)?.to_string(),
        owner: token_string(elems, 1)?.to_string(),
        name: token_string(elems, 2)?.to_string(),
        param_count: token_int(elems, 3).and_then(|value| usize::try_from(value).ok())?,
        type_name: token_string(elems, 4).map(str::to_string),
        return_type: token_string(elems, 5).map(str::to_string),
        param_types: token_string_array(elems, 6).unwrap_or_default(),
        modifiers: token_int(elems, 7).unwrap_or_default(),
    })
}

pub fn string_array_expr(values: Vec<String>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| array_value(Expression::string(&value)))
            .collect(),
    ))
}

fn array_value(value: Expression) -> ArrayElement {
    ArrayElement {
        key: None,
        value,
        spread: false,
        by_ref: false,
    }
}

fn token_string(elems: &[ArrayElement], index: usize) -> Option<&str> {
    match elems.get(index).map(|elem| &elem.value.kind) {
        Some(ExprKind::Lit(Literal::Str(value))) => Some(value.as_str()),
        _ => None,
    }
}

fn token_int(elems: &[ArrayElement], index: usize) -> Option<i64> {
    match elems.get(index).map(|elem| &elem.value.kind) {
        Some(ExprKind::Lit(Literal::Int(value))) => Some(*value),
        _ => None,
    }
}

fn token_string_array(elems: &[ArrayElement], index: usize) -> Option<Vec<String>> {
    let Some(ExprKind::Array(values)) = elems.get(index).map(|elem| &elem.value.kind) else {
        return None;
    };
    Some(
        values
            .iter()
            .filter_map(|elem| match &elem.value.kind {
                ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
                _ => None,
            })
            .collect(),
    )
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ReflectKind {
    Undefined,
    Null,
    Bool,
    Number,
    String,
    Symbol,
    Function,
    Object,
    Array,
    Map,
    Set,
    Struct,
    Class,
    Interface,
    Exception,
    Pointer,
    Slice,
    Channel,
}

impl ReflectKind {
    pub fn as_str(self) -> &'static str {
        match self {
            ReflectKind::Undefined => "undefined",
            ReflectKind::Null => "null",
            ReflectKind::Bool => "bool",
            ReflectKind::Number => "number",
            ReflectKind::String => "string",
            ReflectKind::Symbol => "symbol",
            ReflectKind::Function => "function",
            ReflectKind::Object => "object",
            ReflectKind::Array => "array",
            ReflectKind::Map => "map",
            ReflectKind::Set => "set",
            ReflectKind::Struct => "struct",
            ReflectKind::Class => "class",
            ReflectKind::Interface => "interface",
            ReflectKind::Exception => "exception",
            ReflectKind::Pointer => "ptr",
            ReflectKind::Slice => "slice",
            ReflectKind::Channel => "chan",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Visibility {
    Public,
    Protected,
    Private,
}

impl Visibility {
    pub fn as_str(self) -> &'static str {
        match self {
            Visibility::Public => "public",
            Visibility::Protected => "protected",
            Visibility::Private => "private",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ObjectKeysMode {
    Own,
    ForIn,
    Values,
    Entries,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ObjectOp {
    Object,
    Assign,
    Freeze,
    FromEntries,
    Create,
    Seal,
    IsFrozen,
    IsSealed,
    Is,
    GetPrototypeOf,
    GetOwnPropertyNames,
    GetOwnPropertyDescriptor,
    GetOwnPropertyDescriptors,
    GetOwnPropertySymbols,
    DefineProperty,
    DefineProperties,
    PreventExtensions,
    IsExtensible,
    SetPrototypeOf,
    GroupBy,
    Delete,
    Get,
    Set,
    TrackKey,
    PropertyIsEnumerable,
    HasOwnProperty,
    IsPrototypeOf,
}

impl ObjectOp {
    fn host_name(self) -> &'static str {
        match self {
            ObjectOp::Object => "Object",
            ObjectOp::Assign => "assign",
            ObjectOp::Freeze => "freeze",
            ObjectOp::FromEntries => "fromEntries",
            ObjectOp::Create => "create",
            ObjectOp::Seal => "seal",
            ObjectOp::IsFrozen => "isFrozen",
            ObjectOp::IsSealed => "isSealed",
            ObjectOp::Is => "is",
            ObjectOp::GetPrototypeOf => "getPrototypeOf",
            ObjectOp::GetOwnPropertyNames => "getOwnPropertyNames",
            ObjectOp::GetOwnPropertyDescriptor => "getOwnPropertyDescriptor",
            ObjectOp::GetOwnPropertyDescriptors => "getOwnPropertyDescriptors",
            ObjectOp::GetOwnPropertySymbols => "getOwnPropertySymbols",
            ObjectOp::DefineProperty => "defineProperty",
            ObjectOp::DefineProperties => "defineProperties",
            ObjectOp::PreventExtensions => "preventExtensions",
            ObjectOp::IsExtensible => "isExtensible",
            ObjectOp::SetPrototypeOf => "setPrototypeOf",
            ObjectOp::GroupBy => "groupBy",
            ObjectOp::Delete => "delete",
            ObjectOp::Get => "get",
            ObjectOp::Set => "set",
            ObjectOp::TrackKey => "trackKey",
            ObjectOp::PropertyIsEnumerable => "propertyIsEnumerable",
            ObjectOp::HasOwnProperty => "hasOwnProperty",
            ObjectOp::IsPrototypeOf => "isPrototypeOf",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ReflectOp {
    Get,
    Set,
    Apply,
    Construct,
    DeleteProperty,
    DefineProperty,
    GetOwnPropertyDescriptor,
    GetPrototypeOf,
    Has,
    IsExtensible,
    OwnKeys,
    PreventExtensions,
    SetPrototypeOf,
}

impl ReflectOp {
    fn host_name(self) -> &'static str {
        match self {
            ReflectOp::Get => "get",
            ReflectOp::Set => "set",
            ReflectOp::Apply => "apply",
            ReflectOp::Construct => "construct",
            ReflectOp::DeleteProperty => "deleteProperty",
            ReflectOp::DefineProperty => "defineProperty",
            ReflectOp::GetOwnPropertyDescriptor => "getOwnPropertyDescriptor",
            ReflectOp::GetPrototypeOf => "getPrototypeOf",
            ReflectOp::Has => "has",
            ReflectOp::IsExtensible => "isExtensible",
            ReflectOp::OwnKeys => "ownKeys",
            ReflectOp::PreventExtensions => "preventExtensions",
            ReflectOp::SetPrototypeOf => "setPrototypeOf",
        }
    }
}

fn sconst(chunk: &mut Chunk, s: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(s)))
}

/// Stack: `[value] -> [ecma_type_string]`.
pub fn emit_typeof(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:value", "typeof", 1, line);
}

/// Single-chunk variant. Stack: `[value] -> [ecma_type_string]`.
pub fn emit_typeof_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:value", "typeof", 1, line);
}

/// Stack: `[callable] -> [bool]`.
pub fn emit_is_callable(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", "isCallable", 1, line);
}

/// Single-chunk variant. Stack: `[callable] -> [bool]`.
pub fn emit_is_callable_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:reflect", "isCallable", 1, line);
}

/// Stack: `[object, key] -> [value]`.
pub fn emit_get_property(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", "get", 2, line);
}

/// Single-chunk variant. Stack: `[object, key] -> [value]`.
pub fn emit_get_property_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:reflect", "get", 2, line);
}

/// Stack: `[object, key, value] -> [bool]`.
pub fn emit_set_property(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", "set", 3, line);
}

/// Single-chunk variant. Stack: `[object, key, value] -> [bool]`.
pub fn emit_set_property_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:reflect", "set", 3, line);
}

/// Stack: `[object, key] -> [bool]`.
pub fn emit_has_own(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:object", "hasOwn", 2, line);
}

/// Single-chunk variant. Stack: `[object, key] -> [bool]`.
pub fn emit_has_own_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:object", "hasOwn", 2, line);
}

/// Stack: `[object, key] -> [bool]`.
pub fn emit_has_in(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:object", "hasIn", 2, line);
}

/// Single-chunk variant. Stack: `[object, key] -> [bool]`.
pub fn emit_has_in_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:object", "hasIn", 2, line);
}

/// Stack: `[object] -> [array]`, where the chosen mode preserves the language
/// adapter's enumeration rules.
pub fn emit_object_view(chunks: &mut [Chunk], current: usize, mode: ObjectKeysMode, line: u32) {
    let name = match mode {
        ObjectKeysMode::Own => "keys",
        ObjectKeysMode::ForIn => "iterForIn",
        ObjectKeysMode::Values => "values",
        ObjectKeysMode::Entries => "entries",
    };
    emit_import_call(chunks, current, "ecma:object", name, 1, line);
}

/// Generic ECMA object operation routed through the shared reflection substrate.
/// Stack contract is the underlying host operation's contract.
pub fn emit_object_op(chunks: &mut [Chunk], current: usize, op: ObjectOp, argc: u8, line: u32) {
    emit_import_call(chunks, current, "ecma:object", op.host_name(), argc, line);
}

/// Generic ECMA Reflect operation routed through the shared reflection substrate.
/// Stack contract is the underlying host operation's contract.
pub fn emit_reflect_op(chunks: &mut [Chunk], current: usize, op: ReflectOp, argc: u8, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", op.host_name(), argc, line);
}

/// Single-chunk variant. Stack: `[object] -> [array]`.
pub fn emit_object_view_in_chunk(chunk: &mut Chunk, mode: ObjectKeysMode, line: u32) {
    let name = match mode {
        ObjectKeysMode::Own => "keys",
        ObjectKeysMode::ForIn => "iterForIn",
        ObjectKeysMode::Values => "values",
        ObjectKeysMode::Entries => "entries",
    };
    emit_import_call_in_chunk(chunk, "ecma:object", name, 1, line);
}

/// Stack: `[object, class_name] -> [bool]`.
///
/// A reflection primitive: it reads `FIELD_TYPES` / `FIELD_TYPE` off the object,
/// both defined in this module. It lived in `classes.rs` behind a one-line
/// pass-through until classes were centralized in the compiler
/// (flexclassplan.md §4a-bis); this is its home.
pub fn emit_instanceof(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(3);
    let (obj_s, klass_s, types_s) = (base, base + 1, base + 2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, klass_s, line); // [obj]
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_s, line); // []
    // types = obj["__types"]
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_string_const(FIELD_TYPES, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, types_s, line);
    // if types present -> types.includes(class_name); else obj["__type"] == class_name
    chunks[current].emit_op_u16(Op::LOCAL_GET, types_s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line); // 1 when __types is present
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, types_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, klass_s, line);
    crate::primitives::collections::emit_contains(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_string_const(FIELD_TYPE, line);
    crate::primitives::collections::emit_get(chunks, current, line); // obj["__type"]
    chunks[current].emit_op_u16(Op::LOCAL_GET, klass_s, line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// Stack: unchanged. Writes `object.__type = type_name`.
pub fn emit_stamp_type(chunk: &mut Chunk, object_slot: u16, type_name: &str, line: u32) {
    emit_set_slot_string_field(chunk, object_slot, FIELD_TYPE, type_name, line);
}

/// Stack: unchanged. Writes `object.__typename = type_name`.
pub fn emit_stamp_type_name(chunk: &mut Chunk, object_slot: u16, type_name: &str, line: u32) {
    emit_set_slot_string_field(chunk, object_slot, FIELD_TYPE_NAME, type_name, line);
}

/// Stack: unchanged. Writes `object.__kind = kind`.
pub fn emit_stamp_kind(chunk: &mut Chunk, object_slot: u16, kind: ReflectKind, line: u32) {
    emit_set_slot_string_field(chunk, object_slot, FIELD_KIND, kind.as_str(), line);
}

/// Stack: unchanged. Writes `object[field] = string_value`.
pub fn emit_set_slot_string_field(
    chunk: &mut Chunk,
    object_slot: u16,
    field: &str,
    value: &str,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_string_const(value, line);
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stack: unchanged. Writes `object[field] = local_value`.
pub fn emit_set_slot_field_from_local(
    chunk: &mut Chunk,
    object_slot: u16,
    field: &str,
    value_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stack: unchanged. Writes `object[field] = ref_func(function_chunk)`.
pub fn emit_bind_method(
    chunk: &mut Chunk,
    object_slot: u16,
    field: &str,
    function_chunk: usize,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, function_chunk as u16, line);
    chunk.emit(0, line);
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Create a reflection-shaped object, stamp its type, copy local-backed fields,
/// bind method functions, and leave the object on the stack.
///
/// Stack: unchanged before fields/methods; result `[object]`.
pub fn emit_new_reflection_object(
    chunk: &mut Chunk,
    object_slot: u16,
    type_name: &str,
    fields: &[(&str, u16)],
    methods: &[(&str, usize)],
    line: u32,
) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, object_slot, line);
    emit_stamp_type(chunk, object_slot, type_name, line);
    for (field, value_slot) in fields {
        emit_set_slot_field_from_local(chunk, object_slot, field, *value_slot, line);
    }
    for (method, function_chunk) in methods {
        emit_bind_method(chunk, object_slot, method, *function_chunk, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
}

/// Create `{ __type, __typename, __kind, __fields, __methods, __attributes }`
/// from local-backed metadata and leave it on the stack. Language adapters may
/// stamp extra fields afterward for public API quirks.
pub fn emit_type_descriptor(
    chunk: &mut Chunk,
    descriptor_slot: u16,
    type_name_slot: u16,
    kind: ReflectKind,
    fields_slot: u16,
    methods_slot: u16,
    attributes_slot: u16,
    line: u32,
) {
    emit_new_reflection_object(
        chunk,
        descriptor_slot,
        "ReflectionType",
        &[
            (FIELD_TYPE_NAME, type_name_slot),
            (FIELD_FIELDS, fields_slot),
            (FIELD_METHODS, methods_slot),
            (FIELD_ATTRIBUTES, attributes_slot),
        ],
        &[],
        line,
    );
    chunk.emit_op(Op::DROP, line);
    emit_stamp_kind(chunk, descriptor_slot, kind, line);
    chunk.emit_op_u16(Op::LOCAL_GET, descriptor_slot, line);
}

/// Create `{ __type, __value, __typename, __kind, __ref }` and leave it on the
/// stack. `ref_slot` should contain null when the value is not settable.
pub fn emit_value_descriptor(
    chunk: &mut Chunk,
    descriptor_slot: u16,
    value_slot: u16,
    type_name_slot: u16,
    kind: ReflectKind,
    ref_slot: u16,
    line: u32,
) {
    emit_new_reflection_object(
        chunk,
        descriptor_slot,
        "ReflectionValue",
        &[
            (FIELD_VALUE, value_slot),
            (FIELD_TYPE_NAME, type_name_slot),
            (FIELD_REF, ref_slot),
        ],
        &[],
        line,
    );
    chunk.emit_op(Op::DROP, line);
    emit_stamp_kind(chunk, descriptor_slot, kind, line);
    chunk.emit_op_u16(Op::LOCAL_GET, descriptor_slot, line);
}

/// Stack: `[object] -> [object[field]]`.
pub fn emit_descriptor_field(chunk: &mut Chunk, field: &str, line: u32) {
    chunk.emit_string_const(field, line);
    emit_get_property_in_chunk(chunk, line);
}

/// Create a reflection type descriptor from stack metadata and leave it on the
/// stack. Supported stack layouts:
///
/// - `[value]`
/// - `[value, type_name]`
/// - `[value, type_name, fields]`
/// - `[value, type_name, kind_name, fields]`
pub fn emit_type_descriptor_from_stack(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let fields_slot = chunks[current].alloc_scratch(1);
    let kind_slot = chunks[current].alloc_scratch(1);
    let type_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, fields_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, fields_slot, line);
    }
    if argc >= 2 {
        if argc >= 4 {
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        } else {
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    let methods_slot = chunks[current].alloc_scratch(1);
    let attrs_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, methods_slot, line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
    let out_slot = chunks[current].alloc_scratch(1);
    emit_type_descriptor(
        &mut chunks[current],
        out_slot,
        type_slot,
        ReflectKind::Object,
        fields_slot,
        methods_slot,
        attrs_slot,
        line,
    );
    emit_stamp_kind_from_slot(&mut chunks[current], out_slot, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Create a reflection value descriptor from stack metadata and leave it on the
/// stack. Supported stack layouts:
///
/// - `[value]`
/// - `[value, type_name]`
/// - `[value, type_name, kind_name]`
/// - `[value, type_name, kind_name, ref_marker]`
pub fn emit_value_descriptor_from_stack(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let ref_slot = chunks[current].alloc_scratch(1);
    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, ref_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, ref_slot, line);
    }
    let kind_slot = chunks[current].alloc_scratch(1);
    let type_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        if argc >= 3 {
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        } else {
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    let out_slot = chunks[current].alloc_scratch(1);
    emit_value_descriptor(
        &mut chunks[current],
        out_slot,
        value_slot,
        type_slot,
        ReflectKind::Object,
        ref_slot,
        line,
    );
    emit_stamp_kind_from_slot(&mut chunks[current], out_slot, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Stack: `[value] -> [ReflectionValue(value)]`.
pub fn emit_wrap_existing_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_string_const("any", line);
    chunks[current].emit_string_const("any", line);
    emit_value_descriptor_from_stack(chunks, current, 3, line);
}

/// Stack: `[descriptor] -> [len(descriptor.__fields)]`.
pub fn emit_reflect_num_field(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_descriptor_field(&mut chunks[current], FIELD_FIELDS, line);
    collections::emit_len(chunks, current, line);
}

/// Stack: `[descriptor, index] -> [descriptor.__fields[index]]`.
pub fn emit_reflect_field(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let index = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_descriptor_field(&mut chunks[current], FIELD_FIELDS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
}

/// Stack: `[descriptor, name] -> [field_descriptor|null]`.
///
/// Language walkers may statically lower name lookups into direct field/index
/// access when they have declaration metadata. The runtime fallback is null so
/// unknown reflection queries remain non-panicking.
pub fn emit_reflect_field_by_name(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let name = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Stack: `[value_descriptor] -> [len(value_descriptor.__value)]`.
pub fn emit_reflect_len(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    collections::emit_len(chunks, current, line);
}

/// Stack: `[value_descriptor, index] -> [ReflectionValue(value[index])]`.
pub fn emit_reflect_index(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let index = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
    emit_wrap_existing_value(chunks, current, line);
}

/// Stack: `[map_descriptor, key_descriptor] -> [ReflectionValue(map[key])]`.
pub fn emit_reflect_map_index(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let key = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    collections::emit_get(chunks, current, line);
    emit_wrap_existing_value(chunks, current, line);
}

/// Stack: `[value_descriptor] -> [true]`.
pub fn emit_reflect_is_valid(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_bool_const(true, line);
}

/// Stack: `[value_descriptor] -> [value_descriptor.__value == null]`.
pub fn emit_reflect_is_nil(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Stack: `[value_descriptor] -> [value_descriptor.__ref != null]`.
pub fn emit_reflect_can_set(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_descriptor_field(&mut chunks[current], FIELD_REF, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Stack: `[value_descriptor] -> [bool]`.
pub fn emit_reflect_is_zero(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_descriptor_field(&mut chunks[current], FIELD_KIND, line);
    chunks[current].emit_string_const("string", line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    chunks[current].emit_string_const("", line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    chunks[current].emit_i32_const(0, line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// Stack: `[value_descriptor] -> [value_descriptor.__elem ?? value_descriptor]`.
pub fn emit_reflect_elem(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_descriptor_field(&mut chunks[current], "__elem", line);
    let elem = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
    chunks[current].emit_end(line);
}

/// Stack: `[target_descriptor, value_descriptor] -> [null]`.
pub fn emit_reflect_set_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_descriptor_field(&mut chunks[current], FIELD_VALUE, line);
    emit_set_field_from_stack(&mut chunks[current], FIELD_VALUE, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Stack: `[target_descriptor, primitive_value] -> [null]`.
pub fn emit_reflect_set_primitive(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_set_field_from_stack(&mut chunks[current], FIELD_VALUE, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Stack: `[object, value] -> []`. Writes `object[field] = value`.
pub fn emit_set_field_from_stack(chunk: &mut Chunk, field: &str, line: u32) {
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stack: unchanged. Writes `object[field] = value_slot`.
pub fn emit_stamp_kind_from_slot(chunk: &mut Chunk, object_slot: u16, kind_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    let key = sconst(chunk, FIELD_KIND);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_import_call(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_import_call_in_chunk(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

/// Stamp `class_name` into `this.__types` array for cross-language instanceof.
/// If `__types` is null/missing, creates an empty array first, then pushes the name.
/// Called once per class in the inheritance chain (child calls after parent constructor).
///
/// Bytecode stack trace:
/// ```text
/// local_get this          // [this]
/// dup                     // [this, this]
/// struct_get "__types"    // [this, types_or_null]
/// dup                     // [this, types_or_null, types_or_null]
/// ref_is_null             // [this, types_or_null, i32]
/// i32.eqz; br_if 0        // [this, types_or_null]
/// drop                    // [this]
/// array_new 0             // [this, []]
/// skip:                   // [this, array]
/// const "class_name"      // [this, array, "class_name"]
/// array_push              // [this, array_with_name]
/// struct_set "__types"    // [] (stored on this)
/// drop                    // []
/// ```
///
/// Stack: unchanged
pub fn emit_instanceof_chain(
    chunks: &mut [Chunk],
    current: usize,
    this_slot: u16,
    class_name: &str,
    line: u32,
) {
    let types_key = chunks[current].add_constant(Value::String(Arc::from(
        crate::primitives::reflection::FIELD_TYPES,
    )));

    // Stack: []
    chunks[current].emit_op_u16(Op::LOCAL_GET, this_slot, line); // [this]
    chunks[current].emit_dup(line); // [this, this]
    chunks[current].emit_op_u16(Op::STRUCT_GET, types_key, line); // [this, types_or_null]
    chunks[current].emit_dup(line); // [this, tn, tn]
    chunks[current].emit_op(Op::REF_IS_NULL, line); // [this, tn, bool]
    let init_block = chunks[current].emit_block(line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_op(Op::DROP, line); // [this] (drop the null)
    crate::primitives::collections::emit_array_new(chunks, current, 0, line); // [this, []]
    chunks[current].emit_end(line);
    chunks[current].patch_block(init_block); // skip lands here; [this, array]

    // Push class_name onto array while preserving array on stack.
    // ecma:array.push is [arr, val] → [new_length], so DUP the array
    // first: [this, array] → [this, array, array] → push → [this, array, len] → drop.
    chunks[current].emit_dup(line); // [this, array, array]
    chunks[current].emit_string_const(class_name, line); // [this, array, array, name]
    crate::primitives::collections::emit_push(chunks, current, line); // [this, array, len]
    chunks[current].emit_op(Op::DROP, line); // [this, array]
    // struct_set: [this, array] → sets this.__types = array, leaves array on stack.
    chunks[current].emit_op_u16(Op::STRUCT_SET, types_key, line); // [array]
    chunks[current].emit_op(Op::DROP, line); // []
}
