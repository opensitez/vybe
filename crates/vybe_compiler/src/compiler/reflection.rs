//! .NET reflection compilation — Type/Method/Property/Field/Constructor/
//! Parameter reflection, custom attributes, and Invoke. Compile-time
//! resolution against class/attribute metadata. Moved out of the former dotnet_calls.rs.

use super::*;
use crate::compiler::calls::{strip_generic_suffix, terminal_type_name};

impl Compiler {

    pub(crate) fn resolve_reflection_binding_expr(
        &self,
        expr: &Expression,
    ) -> Option<ReflectionBinding> {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(type_name)) if type_name.starts_with("System.") => {
                Some(ReflectionBinding::Type(type_name.clone()))
            }
            ExprKind::Ident(name) => self.reflection_bindings.get(&self.canon(name)).cloned(),
            ExprKind::Member { object, field, .. } => {
                let receiver = self.resolve_reflection_binding_expr(object)?;
                match (receiver, strip_generic_suffix(field.as_str())) {
                    (ReflectionBinding::Type(type_name), "BaseType") => self
                        .reflection_base_type_name(&type_name)
                        .map(ReflectionBinding::Type),
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
            "GetCustomAttributes" if args.len() >= 2 => {
                let Some(attribute_type) = self.resolve_reflection_type_arg(&args[0].value) else {
                    return Ok(false);
                };
                let inherit = matches!(args[1].value.kind, ExprKind::Lit(Literal::Bool(true)));
                let attrs = self.reflection_attributes_for_binding(
                    &provider,
                    Some(&attribute_type),
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
