//! .NET reflection/enum/TryParse/dictionary call compilation.
//!
//! Extracted from `compiler/calls.rs` (`impl Compiler`).

use super::*;
use crate::compiler::calls::{
    extract_generic_type_name, resolve_receiver_type_hint, strip_generic_suffix, terminal_type_name,
};

impl Compiler {
    pub(super) fn try_compile_dotnet_dictionary_try_get_value(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if args.len() != 2 {
            return Ok(false);
        }
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        if !field.eq_ignore_ascii_case("TryGetValue") {
            return Ok(false);
        }
        let is_dictionary = resolve_receiver_type_hint(self, object)
            .as_deref()
            .map(Self::is_dictionary_type_hint)
            .unwrap_or(false);
        if !is_dictionary {
            return Ok(false);
        }

        if let ExprKind::Ident(name) = &args[1].value.kind {
            let unresolved = self.scope().resolve(name).is_none()
                && (!self.case_sensitive && self.scope().resolve_ci(name).is_none()
                    || self.case_sensitive)
                && !self.defined_globals.contains(&self.canon(name));
            if unresolved {
                self.define_local_typed(name, None);
            }
        }

        self.compile_expr(object)?;
        let map_slot = self.define_local("__dict_try_get_map");
        self.emit_u16(Op::LOCAL_SET, map_slot);

        self.compile_expr(&args[0].value)?;
        if self.expr_uses_case_insensitive_string_keys(object) {
            let line = self.line;
            common::strings::emit_to_lower(self.chunk(), line);
        }
        let key_slot = self.define_local("__dict_try_get_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        let has_idx = self.import("ecma:map", "has");
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_host_call(has_idx, 2);
        let has_slot = self.define_local("__dict_try_get_has");
        self.emit_u16(Op::LOCAL_SET, has_slot);

        self.emit_u16(Op::LOCAL_GET, has_slot);
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, map_slot);
        let getter_key = self.str_const("__get___index__");
        self.emit_u16(Op::STRUCT_GET, getter_key);
        let getter_slot = self.define_local("__dict_try_get_getter");
        self.emit_u16(Op::LOCAL_SET, getter_slot);

        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, self.line);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_u8(Op::CALL_REF, 2);
        self.chunk().emit_end(line);

        self.compile_assign_target(&args[1].value)?;
        inst!(self, core_wasm::bool_const, true);

        self.chunk().emit_else(line);
        self.emit(Op::NULL);
        self.compile_assign_target(&args[1].value)?;
        inst!(self, core_wasm::bool_const, false);
        self.chunk().emit_end(line);
        Ok(true)
    }

    pub(super) fn try_compile_dotnet_case_insensitive_collection_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        if !self.expr_uses_case_insensitive_string_keys(object) {
            return Ok(false);
        }

        let receiver_type = resolve_receiver_type_hint(self, object).unwrap_or_default();
        let normalized = Self::normalize_type_hint(&receiver_type);
        let line = self.line;

        if Self::is_dictionary_type_hint(&normalized) {
            match (field.as_str(), args.len()) {
                ("Add", 2) => {
                    let obj_slot = self.define_local("__dict_add_obj");
                    let key_slot = self.define_local("__dict_add_key");
                    let keys_slot = self.define_local("__dict_add_keys");

                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    self.compile_collection_key(object, &args[0].value)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    self.compile_expr(&args[1].value)?;
                    let idx = self.import("ecma:map", "set");
                    self.emit_host_call(idx, 3);
                    self.emit(Op::DROP);

                    let keys_key = self.str_const("__keys");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::STRUCT_GET, keys_key);
                    self.emit_u16(Op::LOCAL_SET, keys_slot);

                    self.emit_u16(Op::LOCAL_GET, keys_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_SET, keys_slot);
                    self.emit_u16(Op::STRUCT_SET, keys_key);
                    self.emit(Op::DROP);

                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_GET, keys_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                    return Ok(true);
                }
                ("ContainsKey", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:map", "has");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                ("Remove", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:map", "delete");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                _ => {}
            }
        }

        if normalized.contains("hashset") || normalized.contains("sortedset") {
            match (field.as_str(), args.len()) {
                ("Add", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    self.emit_common("dotnet.hashset_add", 2, line);
                    return Ok(true);
                }
                ("Contains", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:set", "has");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                ("Remove", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:set", "delete");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                _ => {}
            }
        }

        Ok(false)
    }

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

    pub(super) fn try_compile_js_iterator_from_generator_take_to_array(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if !self.profile.ecma_array_method_dispatch || !args.is_empty() {
            return Ok(false);
        }

        let ExprKind::Member {
            object: take_call,
            field: to_array_field,
            null_safe: false,
        } = &callee.kind
        else {
            return Ok(false);
        };
        if to_array_field != "toArray" {
            return Ok(false);
        }

        if let ExprKind::Call {
            callee: flat_map_callee,
            args: flat_map_args,
            optional: false,
        } = &take_call.kind
        {
            if flat_map_args.len() == 1 && !flat_map_args[0].spread {
                if let ExprKind::Member {
                    object: from_call,
                    field: flat_map_field,
                    null_safe: false,
                } = &flat_map_callee.kind
                {
                    let mapper_is_generator = matches!(
                        &flat_map_args[0].value.kind,
                        ExprKind::FunctionExpr(stmt)
                            if matches!(
                                &stmt.kind,
                                StmtKind::FunctionDecl {
                                    is_generator: true,
                                    ..
                                }
                            )
                    );
                    if flat_map_field == "flatMap" && mapper_is_generator {
                        if let ExprKind::Call {
                            callee: from_callee,
                            args: from_args,
                            optional: false,
                        } = &from_call.kind
                        {
                            if from_args.len() == 1
                                && !from_args[0].spread
                                && matches!(&from_args[0].value.kind, ExprKind::Array(_))
                            {
                                if let ExprKind::Member {
                                    object: iterator_obj,
                                    field: from_field,
                                    null_safe: false,
                                } = &from_callee.kind
                                {
                                    if from_field == "from"
                                        && matches!(&iterator_obj.kind, ExprKind::Ident(name) if name == "Iterator")
                                    {
                                        self.compile_expr(&from_args[0].value)?;
                                        self.compile_expr(&flat_map_args[0].value)?;
                                        let line = self.line;
                                        crate::emitter::generators::emit_flat_map_generator_mapper_into_array(
                                            &mut self.chunks,
                                            self.current,
                                            line,
                                        );
                                        return Ok(true);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if let ExprKind::Call {
            callee: from_callee,
            args: from_args,
            optional: false,
        } = &take_call.kind
        {
            if from_args.len() == 1 && !from_args[0].spread {
                if let ExprKind::Member {
                    object: iterator_obj,
                    field: from_field,
                    null_safe: false,
                } = &from_callee.kind
                {
                    if from_field == "from"
                        && matches!(&iterator_obj.kind, ExprKind::Ident(name) if name == "Iterator")
                    {
                        let source = &from_args[0].value;
                        if self.is_direct_generator_call(source) {
                            self.compile_expr(source)?;
                            let line = self.line;
                            crate::emitter::generators::emit_drain_into_array(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            return Ok(true);
                        }
                    }
                }
            }
        }

        let ExprKind::Call {
            callee: take_callee,
            args: take_args,
            optional: false,
        } = &take_call.kind
        else {
            return Ok(false);
        };
        if take_args.len() != 1 || take_args[0].spread {
            return Ok(false);
        }

        let ExprKind::Member {
            object: from_call,
            field: take_field,
            null_safe: false,
        } = &take_callee.kind
        else {
            return Ok(false);
        };
        if take_field != "take" {
            return Ok(false);
        }

        let ExprKind::Call {
            callee: from_callee,
            args: from_args,
            optional: false,
        } = &from_call.kind
        else {
            return Ok(false);
        };
        if from_args.len() != 1 || from_args[0].spread {
            return Ok(false);
        }

        let ExprKind::Member {
            object: iterator_obj,
            field: from_field,
            null_safe: false,
        } = &from_callee.kind
        else {
            return Ok(false);
        };
        if from_field != "from" {
            return Ok(false);
        }
        if !matches!(&iterator_obj.kind, ExprKind::Ident(name) if name == "Iterator") {
            return Ok(false);
        }

        let source = &from_args[0].value;
        if !self.is_direct_generator_call(source) {
            return Ok(false);
        }

        self.compile_expr(source)?;
        self.compile_expr(&take_args[0].value)?;
        let line = self.line;
        crate::emitter::generators::emit_take_into_array(&mut self.chunks, self.current, line);
        Ok(true)
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

    pub(super) fn canonical_enum_type_from_runtime_type(
        &self,
        expr: &Expression,
    ) -> Option<String> {
        let ExprKind::Lit(Literal::Str(type_name)) = &expr.kind else {
            return None;
        };
        let short = type_name.rsplit('.').next().unwrap_or(type_name).trim();
        self.resolve_known_enum_type(short)
    }

    pub(super) fn canonical_enum_type_from_expr(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .and_then(|hint| self.resolve_known_enum_type(hint))
                .or_else(|| self.resolve_known_enum_type(name)),
            ExprKind::Member { object, .. } => {
                let enum_type = terminal_type_name(object)?;
                self.resolve_known_enum_type(strip_generic_suffix(&enum_type))
            }
            _ => resolve_receiver_type_hint(self, expr)
                .and_then(|hint| self.resolve_known_enum_type(strip_generic_suffix(&hint))),
        }
    }

    pub(super) fn console_enum_type_from_expr(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(_) => self.canonical_enum_type_from_expr(expr),
            ExprKind::Member { object, .. } if !matches!(&object.kind, ExprKind::Ident(_)) => {
                self.canonical_enum_type_from_expr(expr)
            }
            _ => None,
        }
    }

    pub(super) fn resolve_known_enum_type(&self, name: &str) -> Option<String> {
        let canon = self.canon(name);
        if self.enum_value_names.contains_key(&canon) {
            return Some(canon);
        }
        self.enum_value_names
            .keys()
            .find(|known| known.eq_ignore_ascii_case(name) || known.eq_ignore_ascii_case(&canon))
            .cloned()
    }

    pub(super) fn enum_member_ordinal(&self, enum_type: &str, member_name: &str) -> Option<i64> {
        let enum_type = self.resolve_known_enum_type(enum_type)?;
        self.enum_value_names
            .get(&enum_type)?
            .iter()
            .find(|(_, name)| name.eq_ignore_ascii_case(member_name))
            .map(|(value, _)| *value)
    }

    pub(super) fn enum_entries_sorted(&self, enum_type: &str) -> Option<Vec<(i64, String)>> {
        let mut entries: Vec<(i64, String)> = self
            .enum_value_names
            .get(enum_type)?
            .iter()
            .map(|(value, name)| (*value, name.clone()))
            .collect();
        entries.sort_by_key(|(value, _)| *value);
        Some(entries)
    }

    pub(super) fn compile_string_array(&mut self, values: &[String]) -> Result<(), String> {
        let expr = Expression::new(ExprKind::Array(
            values
                .iter()
                .map(|value| ArrayElement {
                    key: None,
                    value: Expression::string(value),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
        self.compile_expr(&expr)
    }

    pub(super) fn emit_enum_name_lookup(
        &mut self,
        enum_type: &str,
        value_expr: &Expression,
        ignore_case: bool,
    ) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            self.emit(Op::NULL);
            return Ok(());
        };

        let line = self.line;
        let to_str_idx = self.import("ecma:string", "String");
        let lower_idx = if ignore_case {
            Some(self.import("ecma:string", "toLowerCase"))
        } else {
            None
        };

        self.compile_expr(value_expr)?;
        self.emit_host_call(to_str_idx, 1);
        if let Some(lower_idx) = lower_idx {
            self.emit_host_call(lower_idx, 1);
        }
        let input_slot = self.define_local("__enum_name_input");
        self.emit_u16(Op::LOCAL_SET, input_slot);

        let result_slot = self.define_local("__enum_name_result");
        let matched_slot = self.define_local("__enum_name_matched");
        self.emit(Op::NULL);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for (_, name) in entries {
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, input_slot);
            let candidate = if ignore_case {
                name.to_ascii_lowercase()
            } else {
                name.clone()
            };
            self.emit_const(Value::String(Arc::from(candidate.as_str())));
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, result_slot);
        let _ = line;
        Ok(())
    }

    pub(super) fn emit_enum_value_to_string(
        &mut self,
        enum_type: &str,
        value_expr: &Expression,
    ) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            self.compile_expr(value_expr)?;
            let to_str_idx = self.import("ecma:string", "String");
            self.emit_host_call(to_str_idx, 1);
            return Ok(());
        };

        let value_slot = self.define_local("__enum_tostring_value");
        self.compile_expr(value_expr)?;
        self.emit_u16(Op::LOCAL_SET, value_slot);

        let result_slot = self.define_local("__enum_tostring_result");
        let matched_slot = self.define_local("__enum_tostring_matched");
        self.emit(Op::NULL);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for (value, name) in &entries {
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_const(Value::F64(*value as f64));
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        if self.enum_flags.contains(enum_type) {
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_const(Value::String(Arc::from("")));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(0));
            self.emit_u16(Op::LOCAL_SET, matched_slot);

            for (value, name) in &entries {
                if *value <= 0 || (value & (value - 1)) != 0 {
                    continue;
                }
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_const(Value::F64(*value as f64));
                self.emit(Op::I32_AND);
                self.emit_const(Value::F64(*value as f64));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, matched_slot);
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_const(Value::String(Arc::from(", ")));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                };
                self.emit_const(Value::String(Arc::from(name.as_str())));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                };
                self.chunk().emit_else(line);
                self.emit_const(Value::String(Arc::from(name.as_str())));
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.chunk().emit_end(line);
            }

            self.emit_u16(Op::LOCAL_GET, matched_slot);
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, result_slot);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, matched_slot);
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        let to_str_idx = self.import("ecma:string", "String");
        self.emit_host_call(to_str_idx, 1);
        self.chunk().emit_end(line);
        Ok(())
    }

    pub(super) fn emit_dotnet_console_arg(&mut self, expr: &Expression) -> Result<(), String> {
        if let Some(enum_type) = self.console_enum_type_from_expr(expr) {
            self.emit_enum_value_to_string(&enum_type, expr)?;
            return Ok(());
        }

        if !self.profile.namespaces.use_dotnet {
            self.compile_expr(expr)?;
            return Ok(());
        }

        self.compile_expr(expr)?;
        let value_slot = self.define_local("__dotnet_console_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("number")));
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        let helper = self.str_const("__vybe_dotnet_numeric_format");
        self.emit_u16(Op::GLOBAL_GET, helper);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::String(Arc::from("F12")));
        self.emit_const(Value::F64(0.0));
        self.emit_u8(Op::CALL_REF, 3);
        let parse_float = self.import("ecma:number", "parseFloat");
        self.emit_host_call(parse_float, 1);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.chunk().emit_end(line);
        Ok(())
    }

    pub(super) fn emit_enum_has_flag(
        &mut self,
        value_expr: &Expression,
        flag_expr: &Expression,
    ) -> Result<(), String> {
        let flag_slot = self.define_local("__enum_flag_value");
        let value_slot = self.define_local("__enum_flag_source");
        self.compile_expr(flag_expr)?;
        self.emit_u16(Op::LOCAL_SET, flag_slot);
        self.compile_expr(value_expr)?;
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_u16(Op::LOCAL_GET, flag_slot);
        self.emit(Op::I32_AND);
        self.emit_u16(Op::LOCAL_GET, flag_slot);
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        Ok(())
    }

    pub(super) fn try_compile_dotnet_enum_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        let mut static_enum_call = false;
        let (field, instance_object) = match &callee.kind {
            ExprKind::Member { object, field, .. } => {
                if terminal_type_name(object)
                    .is_some_and(|type_name| type_name.eq_ignore_ascii_case("Enum"))
                {
                    static_enum_call = true;
                    (field.as_str(), None)
                } else {
                    (field.as_str(), Some(object.as_ref()))
                }
            }
            ExprKind::Ident(name) => {
                let Some((receiver, field)) = name.rsplit_once('.') else {
                    return Ok(false);
                };
                if receiver
                    .rsplit('.')
                    .next()
                    .is_some_and(|type_name| type_name.eq_ignore_ascii_case("Enum"))
                {
                    static_enum_call = true;
                    (field, None)
                } else {
                    return Ok(false);
                }
            }
            _ => return Ok(false),
        };
        let field_name = strip_generic_suffix(field);

        if static_enum_call {
            match field_name {
                "GetNames" if args.len() == 1 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    let Some(entries) = self.enum_entries_sorted(&enum_type) else {
                        return Ok(false);
                    };
                    let names: Vec<String> = entries.into_iter().map(|(_, name)| name).collect();
                    self.compile_string_array(&names)?;
                    return Ok(true);
                }
                "GetValues" if args.len() == 1 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    let Some(entries) = self.enum_entries_sorted(&enum_type) else {
                        return Ok(false);
                    };
                    let names: Vec<String> = entries.into_iter().map(|(_, name)| name).collect();
                    self.compile_string_array(&names)?;
                    return Ok(true);
                }
                "Parse" if args.len() >= 2 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    self.emit_enum_name_lookup(&enum_type, &args[1].value, false)?;
                    return Ok(true);
                }
                "IsDefined" if args.len() >= 2 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    self.emit_enum_name_lookup(&enum_type, &args[1].value, false)?;
                    self.emit(Op::REF_IS_NULL);
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                    };
                    return Ok(true);
                }
                "GetUnderlyingType" if args.len() == 1 => {
                    let expr = Expression::new(ExprKind::Object(vec![
                        ObjectProperty::KeyValue {
                            key: Expression::string("Name"),
                            value: Expression::string("Int32"),
                        },
                        ObjectProperty::KeyValue {
                            key: Expression::string("FullName"),
                            value: Expression::string("System.Int32"),
                        },
                    ]));
                    self.compile_expr(&expr)?;
                    return Ok(true);
                }
                "Format" if args.len() >= 3 => {
                    self.compile_expr(&args[1].value)?;
                    let to_str_idx = self.import("ecma:string", "String");
                    self.emit_host_call(to_str_idx, 1);
                    return Ok(true);
                }
                "TryParse" if matches!(args.len(), 2 | 3 | 4 | 5) => {
                    let visible_args = if args.len() >= 4 {
                        &args[..args.len() - 2]
                    } else {
                        args
                    };
                    let enum_type = extract_generic_type_name(field)
                        .map(|name| self.canon(&name))
                        .filter(|canon| self.enum_value_names.contains_key(canon))
                        .or_else(|| {
                            (args.len() >= 4)
                                .then(|| {
                                    self.canonical_enum_type_from_expr(&args[args.len() - 2].value)
                                })
                                .flatten()
                        });
                    let Some(enum_type) = enum_type else {
                        return Ok(false);
                    };
                    let (value_arg, ignore_case, out_arg) = if visible_args.len() == 3 {
                        (
                            &visible_args[0].value,
                            matches!(
                                visible_args[1].value.kind,
                                ExprKind::Lit(Literal::Bool(true))
                            ),
                            &visible_args[2].value,
                        )
                    } else {
                        (&visible_args[0].value, false, &visible_args[1].value)
                    };
                    self.emit_enum_name_lookup(&enum_type, value_arg, ignore_case)?;
                    let parsed_slot = self.define_local("__enum_try_parse_value");
                    self.emit_u16(Op::LOCAL_SET, parsed_slot);
                    self.emit_u16(Op::LOCAL_GET, parsed_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit(Op::NULL);
                    self.compile_assign_target(out_arg)?;
                    inst!(self, core_wasm::bool_const, false);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, parsed_slot);
                    self.compile_assign_target(out_arg)?;
                    inst!(self, core_wasm::bool_const, true);
                    self.chunk().emit_end(line);
                    return Ok(true);
                }
                _ => {}
            }
        }

        let Some(object) = instance_object else {
            return Ok(false);
        };

        let Some(enum_type) = self.canonical_enum_type_from_expr(object) else {
            return Ok(false);
        };

        match field_name {
            "HasFlag" if args.len() == 1 => {
                self.emit_enum_has_flag(object, &args[0].value)?;
                Ok(true)
            }
            "ToString" if args.is_empty() => {
                self.emit_enum_value_to_string(&enum_type, object)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Lambda compilation
    // ════════════════════════════════════════════════════════════════════════
}
