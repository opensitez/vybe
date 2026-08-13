//! Shared enum machinery — BOTH halves of the topic, in one file.
//!
//! The `impl Compiler` half is the compile-time enum-type resolution +
//! dispatch glue (GetName/GetNames/GetValues/Parse/TryParse/IsDefined/
//! ToString/HasFlag/console-arg), moved out of the former dotnet_calls.rs.
//! The free `&mut Chunk` half at the bottom is the runtime value<->name
//! machinery it emits through; it used to sit in a separate `enum.rs` that
//! only this file called, which is the split `add_vybex_language.md` says a
//! two-halved topic must not have (and which forced a `pub mod r#enum;` raw
//! identifier, since `enum` is a keyword).
//!
//! **The enum object shape.** An enum compiles to a single object carrying
//! BOTH directions: `{ Red: 0, Green: 1, "0": "Red", "1": "Green" }` — forward
//! `name → value` (values stay bare ints, so flags/arithmetic/comparison/casts
//! never break) plus a reverse `value → name` map. The free fns below
//! implement the enum operations as generic RUNTIME reads on that object, so
//! no language has to hand-roll compile-time ordinal tables. Any language
//! whose enums use this shape (C#, VB, …) shares this one emitter.

use super::*;
use crate::primitives::calls::{
    extract_generic_type_name, resolve_receiver_type_hint, strip_generic_suffix, terminal_type_name,
};
use crate::primitives::instructions::host;

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

    /// `name → the member CONSTANT, or null` — `Enum.Parse`, `Enum.TryParse`.
    ///
    /// A parsed value IS the object a member read yields, so it renders through
    /// the same `ToString` role, coerces through the same `Int` role, and
    /// compares equal to `Color.Green` by identity. This replaced a lookup that
    /// answered with the NAME STRING — a difference invisible while a member
    /// read const-folded to an ordinal (both were "not an object", both
    /// stringified to something plausible), and one more representation of an
    /// enum value once it was not.
    ///
    /// Because the answer is an object rather than a string, callers must still
    /// be able to see that it is an ENUM: `infer_expr_type_hint` names the type
    /// from `Enum.Parse`'s runtime-type argument, and without that the console
    /// and `ToString` paths fall through to a generic one that cannot read the
    /// role.
    ///
    /// Compile-time over the declared members rather than a runtime read of the
    /// class object: a static constant compiles to a GLOBAL, not to a property
    /// hanging off the class, so there is nothing to `get` at runtime. The
    /// members are known here anyway.
    pub(super) fn emit_enum_member_lookup(
        &mut self,
        enum_type: &str,
        value_expr: &Expression,
        ignore_case: bool,
    ) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            self.emit_null();
            return Ok(());
        };

        let to_str_idx = self.import("ecma:string", "String");
        self.compile_expr(value_expr)?;
        self.emit_host_call(to_str_idx, 1);
        if ignore_case {
            let lower_idx = self.import("ecma:string", "toLowerCase");
            self.emit_host_call(lower_idx, 1);
        }
        let input_slot = self.define_local("__enum_member_input");
        self.emit_u16(Op::LOCAL_SET, input_slot);

        let result_slot = self.define_local("__enum_member_result");
        let matched_slot = self.define_local("__enum_member_matched");
        self.emit_null();
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for (_, name) in &entries {
            let probe = if ignore_case {
                name.to_ascii_lowercase()
            } else {
                name.clone()
            };
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, input_slot);
            self.emit_const(Value::String(Arc::from(probe.as_str())));
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            // The constant read itself — the same expression the source would
            // have written, so it resolves however that language's member reads
            // resolve.
            let constant = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(enum_type)),
                field: name.clone(),
                null_safe: false,
            });
            self.compile_expr(&constant)?;
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(())
    }

    /// How an enum value RENDERS — its `ToString` role when it is a constant
    /// OBJECT, and its own stringification when it is a bare number (a flags
    /// combination). The same two-faced shape `emit_dotnet_console_arg` uses,
    /// minus the console's numeric formatting.
    pub(super) fn emit_enum_render(&mut self, expr: &Expression) -> Result<(), String> {
        self.compile_expr(expr)?;
        let value_slot = self.define_local("__enum_render_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        // `ref.test` already yields an i32; `Op::IF` consumes it directly.
        self.chunk().emit_if_value(line);
        {
            let line = self.line;
            let current = self.current;
            crate::primitives::expressions::emit_rich_to_string(
                &mut self.chunks[current],
                value_slot,
                line,
            );
        }
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.chunk().emit_end(line);
        Ok(())
    }

    /// `Enum.IsDefined` — does this NAME or this VALUE name a member.
    ///
    /// It takes either (`IsDefined(typeof(Num), 5)` and
    /// `IsDefined(typeof(Phase), "Start")` are both .NET), so it cannot be the
    /// member lookup above with a null test bolted on. Both halves come from
    /// the one declared member table.
    pub(super) fn emit_enum_is_defined(
        &mut self,
        enum_type: &str,
        value_expr: &Expression,
    ) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            inst!(self, core_wasm::bool_const, false);
            return Ok(());
        };

        self.compile_expr(value_expr)?;
        let input_slot = self.define_local("__enum_is_defined_input");
        self.emit_u16(Op::LOCAL_SET, input_slot);
        let found_slot = self.define_local("__enum_is_defined_found");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, found_slot);

        for (value, name) in &entries {
            for probe in [
                Value::String(Arc::from(name.as_str())),
                Value::F64(*value as f64),
            ] {
                self.emit_u16(Op::LOCAL_GET, input_slot);
                self.emit_const(probe);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, found_slot);
                self.chunk().emit_end(line);
            }
        }

        self.emit_u16(Op::LOCAL_GET, found_slot);
        let line = self.line;
        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
        Ok(())
    }

    /// How an enum-TYPED value stringifies, wherever a declared enum type is
    /// what put it on the stack (`Console.WriteLine(d)`, `d.ToString()`).
    ///
    /// Forks on the value's FACE, because which face a value carries is a
    /// runtime fact and nothing static can settle it: a constant is an OBJECT
    /// that fills the `ToString` role and answers for itself, while a flags
    /// combination or an int stored into an enum-typed local arrives as a bare
    /// NUMBER with no role to ask. The tables below are the number face ONLY.
    ///
    /// Before this fork the number face ran unconditionally, so an `Enum.Parse`
    /// / `TryParse` result — an object since the parse started answering with
    /// the member constant — was looked up in the reverse map by object
    /// identity, found nothing, and rendered empty. The fork lives here, at the
    /// one site that owns "how does an enum value stringify", rather than at
    /// each of its callers.
    pub(super) fn emit_enum_value_to_string(
        &mut self,
        enum_type: &str,
        value_expr: &Expression,
    ) -> Result<(), String> {
        if self.enum_entries_sorted(enum_type).is_none() {
            return self.emit_enum_render(value_expr);
        }

        self.compile_expr(value_expr)?;
        let value_slot = self.define_local("__enum_tostring_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        // `ref.test` already yields an i32; `Op::IF` consumes it directly.
        self.chunk().emit_if_value(line);
        {
            let current = self.current;
            crate::primitives::expressions::emit_rich_to_string(
                &mut self.chunks[current],
                value_slot,
                line,
            );
        }
        self.chunk().emit_else(line);
        self.emit_enum_number_to_string(enum_type, value_slot)?;
        self.chunk().emit_end(line);
        Ok(())
    }

    /// The NUMBER face of the above: an underlying integer back to the name(s)
    /// it stands for. Never reached with an object.
    fn emit_enum_number_to_string(
        &mut self,
        enum_type: &str,
        value_slot: u16,
    ) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            self.emit_u16(Op::LOCAL_GET, value_slot);
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
            self.emit_u16(Op::LOCAL_GET, value_slot);
            let line = self.line;
            emit_value_to_name(self.chunk(), line);
            return Ok(());
        }

        let result_slot = self.define_local("__enum_tostring_result");
        let matched_slot = self.define_local("__enum_tostring_matched");
        self.emit_null();
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
        // No enum special-case here any more. There were two — a member
        // expression folded straight to its name STRING, and an enum-typed
        // value routed through the reverse map — and together with the ordinal
        // fold in `expressions.rs` they gave one enum value three different
        // representations depending on which site rendered it. An enum constant
        // is an OBJECT that fills the `ToString` role like any other, so the
        // object arm below renders it, and it renders a user class's own
        // `ToString` at the same time.

        // No language gate: every caller already reached here BECAUSE the tree
        // resolved the call to `dotnet.console_writeline` / `dotnet.console_write`
        // (builtins.rs, lambdas.rs ×2, calls.rs all test the emit name first).
        // The flag could therefore never be false here — it re-asked a question
        // the resolver had already answered.
        self.compile_expr(expr)?;
        let value_slot = self.define_local("__dotnet_console_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        // An OBJECT renders through the `ToString` SLOT — filled by whatever
        // the object's own language spells it, so an enum constant, a C# class
        // overriding `ToString`, and a Python class defining `__str__` all
        // reach their own conversion. Without this the host's generic
        // stringification answers `[object Color]`, because it looks for the
        // literal spelling `toString` and a case-folding language stored
        // `tostring`.
        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, recipes::is_object);
        let obj_line = self.line;
        // `ref.test` already yields an i32; `Op::IF` consumes it directly.
        self.chunk().emit_if_value(obj_line);
        {
            let line = self.line;
            let current = self.current;
            crate::primitives::expressions::emit_rich_to_string(
                &mut self.chunks[current],
                value_slot,
                line,
            );
        }
        self.chunk().emit_else(obj_line);

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

        self.emit_global_read("__vybe_dotnet_numeric_format");
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::String(Arc::from("F12")));
        self.emit_const(Value::F64(0.0));
        self.emit_direct_callable_invoke(3);
        let parse_float = self.import("ecma:number", "parseFloat");
        self.emit_host_call(parse_float, 1);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.chunk().emit_end(line);
        self.chunk().emit_end(obj_line);
        Ok(())
    }

    /// An enum constant's integer VALUE — the `Int` ROLE the enum lowering
    /// declares (`enum_lowering::INT_METHOD`, returning `__value`).
    ///
    /// THE one site that answers "what is the int of this enum". The `(int)e`
    /// cast, the flags operators `| & ^ ~`, and `HasFlag` all coerce through
    /// here: two sites resolving the same question independently is exactly how
    /// an enum value grew four representations in the first place.
    ///
    /// Tolerant of an operand that is ALREADY a number. A flags combination
    /// names no member — `Perm.A | Perm.B` has no constant behind it — so the
    /// result of one operator flows back into the next as a bare integer, and
    /// the `Int` read has nothing to call. The object test is a runtime one
    /// because that is a runtime fact; the DECISION to read the role at all
    /// stays static, resolved from the operand's declared type by the caller.
    pub(super) fn emit_enum_int_value(&mut self, expr: &Expression) -> Result<(), String> {
        self.compile_expr(expr)?;
        let value_slot = self.define_local("__enum_int_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        // `ref.test` already yields an i32; `Op::IF` consumes it directly.
        self.chunk().emit_if_value(line);
        let key = self.str_const(&vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Int));
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, key);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_direct_callable_invoke(1);
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
        // Both operands through the shared coercion: `HasFlag` is a bit test on
        // the underlying values, and either side may be a constant OBJECT or an
        // already-combined integer.
        self.emit_enum_int_value(value_expr)?;
        self.emit_enum_int_value(flag_expr)?;
        let line = self.line;
        emit_has_flag(self.chunk(), line);
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
                    // `Enum.Parse(type, value, ignoreCase)` is the real .NET
                    // overload, and that third argument is the only thing that
                    // separates a case-insensitive parse from a case-sensitive
                    // one. Reading it here keeps `TryParse`'s ignore-case form
                    // — which the walker normalizes onto this very call — on
                    // the SAME lookup instead of a second one.
                    let ignore_case = args.len() >= 3
                        && matches!(&args[2].value.kind, ExprKind::Lit(Literal::Bool(true)));
                    self.emit_enum_member_lookup(&enum_type, &args[1].value, ignore_case)?;
                    return Ok(true);
                }
                "IsDefined" if args.len() >= 2 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    self.emit_enum_is_defined(&enum_type, &args[1].value)?;
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
                    // The format letter picks WHICH of the constant's two faces
                    // renders — `"D"` its declared value, `"G"`/`"F"` its name —
                    // and each is read through the role that owns it rather
                    // than re-derived here. Ignoring the letter and calling
                    // `String()` answered with whichever face the value
                    // happened to be carrying.
                    let decimal = matches!(
                        &args[2].value.kind,
                        ExprKind::Lit(Literal::Str(fmt)) if fmt.eq_ignore_ascii_case("d")
                    );
                    if decimal {
                        self.emit_enum_int_value(&args[1].value)?;
                    } else {
                        self.emit_enum_render(&args[1].value)?;
                    }
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
                    self.emit_enum_member_lookup(&enum_type, value_arg, ignore_case)?;
                    let parsed_slot = self.define_local("__enum_try_parse_value");
                    self.emit_u16(Op::LOCAL_SET, parsed_slot);
                    self.emit_u16(Op::LOCAL_GET, parsed_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_null();
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

        // `HasFlag` asks the ARGUMENT as well as the receiver. A combined flags
        // value is a bare integer with no declared type left on it — that is
        // what `A | B` produces — so `combined.HasFlag(Perm.Read)` is only
        // recognisable from the constant it is handed. Nothing else spells
        // `HasFlag` with an enum constant for an argument.
        let enum_type = self.canonical_enum_type_from_expr(object).or_else(|| {
            (field_name.eq_ignore_ascii_case("HasFlag") && args.len() == 1)
                .then(|| self.canonical_enum_type_from_expr(&args[0].value))
                .flatten()
        });
        let Some(enum_type) = enum_type else {
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

// ════════════════════════════════════════════════════════════════════════════
// Runtime half — free fns over `&mut Chunk`
//
// Reads use `ecma:object.get` (a raw property-bag read) rather than an index
// expression: an enum object carries an index getter that does array-position
// lookup, which only matches sequential values — the raw read hits the
// reverse field directly.
// ════════════════════════════════════════════════════════════════════════════

/// `value → name` (enum `ToString`, `Enum.GetName`, `Enum.Format("G")`).
/// Stack: `[enumObj, value]` → `[string]`.
///
/// Only NUMERIC values map through the reverse field; a value that is already a
/// name string (e.g. an `Enum.Parse` result flowing into `ToString`) passes
/// through `String()` unchanged. Numeric values that aren't defined members
/// fall back to `String(value)` (matches .NET's numeric `ToString`).
pub fn emit_value_to_name(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    let obj = chunk.alloc_scratch(1);
    let name = chunk.alloc_scratch(1);
    // Stack pushed as [enumObj, value]; pop value first.
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    // typeof(value) === "number"
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("number", line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    // name = ecma:object.get(enumObj, "" + value)  (raw reverse-field read)
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    host::emit(chunk, "ecma:object", "get", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name, line);

    // name undefined ? String(value) : name
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:string", "String", 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    // Non-numeric (already a name string): pass through unchanged.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:string", "String", 1, line);
    chunk.emit_end(line);
}

/// Case-sensitive `name → validated name or null` (enum `Parse` / `IsDefined` /
/// `TryParse`). Stack: `[enumObj, input]` → `[string | null]`.
///
/// `input` names a member iff a raw read of the enum object yields its NUMERIC
/// forward field. A numeric-string input would instead hit a reverse
/// (value→name) field — a string — and is correctly rejected. Returns the
/// input (== the canonical name on an exact match) or null.
pub fn emit_name_to_member_or_null(chunk: &mut Chunk, line: u32) {
    let input = chunk.alloc_scratch(1);
    let obj = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, input, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    // Coerce input to a string once.
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    host::emit(chunk, "ecma:string", "String", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input, line);

    // typeof(ecma:object.get(enumObj, input)) === "number" ? input : null
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    host::emit(chunk, "ecma:object", "get", 2, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("number", line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
}

/// `HasFlag` — `(value & flag) === flag`. Stack: `[value, flag]` → `[bool]`.
pub fn emit_has_flag(chunk: &mut Chunk, line: u32) {
    let flag = chunk.alloc_scratch(1);
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, flag, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, flag, line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    // `HasFlag` returns a BOOLEAN, and `emit_dyn_eq` answers with a raw i32 —
    // so the result printed `1`/`0` where .NET prints `True`/`False`. The same
    // conversion `emit_enum_is_defined` ends with, for the same reason: a
    // predicate's answer is a bool, not the integer that computed it.
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}
