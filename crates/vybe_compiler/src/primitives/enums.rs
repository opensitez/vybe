//! .NET enum call compilation (GetName/GetNames/GetValues/Parse/TryParse/
//! IsDefined/ToString/HasFlag/console-arg) — moved out of the former dotnet_calls.rs.
//! The runtime value<->name machinery lives in `crate::primitives::r#enum`; this
//! layer is the compile-time enum-type resolution + dispatch glue.

use super::*;
use crate::primitives::calls::{
    extract_generic_type_name, resolve_receiver_type_hint, strip_generic_suffix, terminal_type_name,
};

impl Compiler {
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
                if let Some(path) = Self::member_access_path(object) {
                    if let Some(enum_type) =
                        self.resolve_known_enum_type(strip_generic_suffix(&path))
                    {
                        return Some(enum_type);
                    }
                }
                if let Some(hint) = self.infer_expr_type_hint(expr) {
                    if let Some(enum_type) = self.resolve_known_enum_type(&hint) {
                        return Some(enum_type);
                    }
                }
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
            .or_else(|| {
                let suffix = format!(".{canon}");
                let mut matches = self
                    .enum_value_names
                    .keys()
                    .filter(|known| known.ends_with(&suffix));
                let resolved = matches.next().cloned();
                if matches.next().is_none() {
                    resolved
                } else {
                    None
                }
            })
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

    pub(super) fn qualified_enum_member_expr(&self, expr: &Expression) -> Option<(String, String)> {
        let ExprKind::Member { object, field, .. } = &expr.kind else {
            return None;
        };
        let path = Self::member_access_path(object)?;
        let base_path = strip_generic_suffix(&path);
        if let Some(enum_type) = self.resolve_known_enum_type(base_path) {
            return self
                .enum_member_ordinal(&enum_type, field)
                .is_some()
                .then(|| (enum_type, field.clone()));
        }
        let member_key = self.canon(field);
        let owner = self.enum_members.get(&member_key)?;
        let canon_path = self.canon(base_path);
        (owner == &canon_path || owner.eq_ignore_ascii_case(base_path))
            .then(|| (owner.clone(), field.clone()))
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

        // Case-sensitive: `input` names a member iff a raw read of the enum
        // object (`ecma:object.get`, bypassing the index getter) yields its
        // NUMERIC forward field. A numeric-string input would instead hit a
        // reverse (value→name) field, which is a string and so correctly
        // rejected. Returns the input name (== the canonical name on an exact
        // match) or null — same contract as the old if-chain.
        if !ignore_case {
            self.compile_expr(&Expression::ident(enum_type))?;
            self.compile_expr(value_expr)?;
            let line = self.line;
            crate::primitives::r#enum::emit_name_to_member_or_null(self.chunk(), line);
            return Ok(());
        }

        // Case-insensitive: the reverse map is keyed by exact name, so fall
        // back to the compile-time member table with lowercased comparison.
        let to_str_idx = self.import("ecma:string", "String");
        let lower_idx = self.import("ecma:string", "toLowerCase");

        self.compile_expr(value_expr)?;
        self.emit_host_call(to_str_idx, 1);
        self.emit_host_call(lower_idx, 1);
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
            self.emit_const(Value::String(Arc::from(name.to_ascii_lowercase().as_str())));
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, result_slot);
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

        // Non-flags enums: `value → name` is a runtime read on the TS-shaped
        // enum object's reverse map. Only NUMERIC values map through it — a
        // value that is already a name string (e.g. an `Enum.Parse` result)
        // passes through `String()` unchanged, matching the old if-chain (which
        // compared against numeric entries and so never rewrote a string). This
        // also avoids the reverse lookup accidentally hitting a forward
        // (name→value) field when the input happens to be a member name.
        if !self.enum_flags.contains(enum_type) {
            self.compile_expr(&Expression::ident(enum_type))?;
            self.compile_expr(value_expr)?;
            let line = self.line;
            crate::primitives::r#enum::emit_value_to_name(self.chunk(), line);
            return Ok(());
        }

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
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, matched_slot);
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_const(Value::String(Arc::from(", ")));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                };
                self.emit_const(Value::String(Arc::from(name.as_str())));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
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
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, result_slot);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, matched_slot);
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
        if let Some((_, member_name)) = self.qualified_enum_member_expr(expr) {
            self.emit_const(Value::String(Arc::from(member_name.as_str())));
            return Ok(());
        }
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
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
        self.compile_expr(value_expr)?;
        self.compile_expr(flag_expr)?;
        let line = self.line;
        crate::primitives::r#enum::emit_has_flag(self.chunk(), line);
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
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
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
                if let Some((_, member_name)) = self.qualified_enum_member_expr(object) {
                    self.emit_const(Value::String(Arc::from(member_name.as_str())));
                    return Ok(true);
                }
                self.emit_enum_value_to_string(&enum_type, object)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}
