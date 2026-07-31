//! Type-hint inference and type-name resolution.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

impl Compiler {
    pub(super) fn lookup_implicit_self_field_type_hint(&self, name: &str) -> Option<&str> {
        if !self.current_class_implicit_self {
            return None;
        }

        let canon_name = self.canon(name);
        let mut current = self.current_class.as_deref();
        while let Some(class_name) = current {
            let pending = self.pending_classes.get(class_name)?;
            if let Some(type_hint) = pending.instance_field_types.get(&canon_name) {
                return Some(type_hint.as_str());
            }
            current = pending.parent.as_deref();
        }
        None
    }

    pub(super) fn prefers_type_qualified_member_lookup(
        &self,
        type_name: &str,
        member_name: &str,
    ) -> bool {
        if self.enum_member_ordinal(type_name, member_name).is_some() {
            return true;
        }

        let type_canon = self.canon(type_name);
        let Some(pending) = self.pending_classes.get(&type_canon).or_else(|| {
            self.pending_classes
                .iter()
                .find(|(name, _)| {
                    name.eq_ignore_ascii_case(type_name) || name.eq_ignore_ascii_case(&type_canon)
                })
                .map(|(_, pending)| pending)
        }) else {
            return false;
        };

        let member_canon = self.canon(member_name);
        pending
            .static_fields
            .iter()
            .any(|name| name == &member_canon)
            || pending
                .static_method_names
                .iter()
                .any(|name| self.canon(name) == member_canon)
            || pending.static_method_overloads.contains_key(&member_canon)
            || pending
                .nested_types
                .iter()
                .any(|name| self.canon(name) == member_canon)
    }

    pub(super) fn expr_terminal_type_name(expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
            ExprKind::Member { field, .. } => Some(field.clone()),
            _ => None,
        }
    }

    pub(super) fn infer_dotnet_factory_return_type(&self, callee: &Expression) -> Option<String> {
        if !self.profile.namespaces.use_dotnet {
            return None;
        }
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return None;
        };
        // `control.CreateGraphics()` returns a Graphics — typing the result
        // lets `g.DrawLine(...)` resolve through the component descriptor
        // (Graphics no longer binds its drawing methods via a ctor thunk).
        // Independent of the receiver's exact control type, so checked before
        // the terminal-type extraction (which bails on a `new X()` receiver).
        if field.eq_ignore_ascii_case("CreateGraphics") {
            return Some("Graphics".into());
        }
        let class_name = Self::expr_terminal_type_name(object)?;
        vybe_runtime::namespaces::lookup_type_member_return(
            &self.profile.namespaces.type_scopes,
            &class_name,
            field,
        )
    }

    pub(super) fn infer_function_return_type(&self, callee: &Expression) -> Option<String> {
        match &callee.kind {
            ExprKind::Ident(name) => {
                // Builtin free-function return types come from profile data
                // (`[builtin_return_types]`), keyed by lowercased name — e.g.
                // VB `Command`/`Environ` → String, `Timer` → Double.
                if let Some(return_type) =
                    self.profile.builtin_return_types.get(&name.to_lowercase())
                {
                    return Some(return_type.clone());
                }
                if let Some(type_hint) = self.lookup_var_type_hint(name) {
                    if Self::is_callable_type_hint(type_hint) {
                        if let Some(return_type) = Self::callable_return_type_hint(type_hint) {
                            return Some(return_type);
                        }
                    }
                }
                self.function_return_types.get(&self.canon(name)).cloned()
            }
            ExprKind::Member { object, field, .. } => {
                if let Some(receiver_type) = self.infer_expr_type_hint(object) {
                    let receiver_trimmed = receiver_type.trim().trim_end_matches('?').trim();
                    let receiver_base = receiver_trimmed
                        .split('<')
                        .next()
                        .unwrap_or(receiver_trimmed)
                        .trim();
                    let receiver_key = self
                        .resolve_pending_class_name_for_type_hint(&receiver_type)
                        .unwrap_or_else(|| self.canon(receiver_base));
                    let qualified = self.canon(&format!("{}.{}", receiver_key, field));
                    if let Some(return_type) = self.function_return_types.get(&qualified) {
                        return Some(return_type.clone());
                    }
                }
                if let ExprKind::Ident(object_name) = &object.kind {
                    let qualified = self.canon(&format!("{}.{}", object_name, field));
                    if let Some(return_type) = self.function_return_types.get(&qualified) {
                        return Some(return_type.clone());
                    }
                }
                self.function_return_types.get(&self.canon(field)).cloned()
            }
            _ => None,
        }
    }

    pub(super) fn infer_array_element_type_hint<'a>(
        &self,
        values: impl IntoIterator<Item = &'a Expression>,
    ) -> String {
        let mut element_type: Option<String> = None;
        for value in values {
            let inferred = self
                .infer_expr_type_hint(value)
                .unwrap_or_else(|| "object".into());
            match &element_type {
                None => element_type = Some(inferred),
                Some(existing)
                    if Self::normalize_type_hint(existing)
                        == Self::normalize_type_hint(&inferred) => {}
                Some(_) => {
                    element_type = Some("object".into());
                    break;
                }
            }
        }
        element_type.unwrap_or_else(|| "object".into())
    }

    pub(super) fn member_access_path(expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            ExprKind::Member { object, field, .. } => {
                let prefix = Self::member_access_path(object)?;
                Some(format!("{prefix}.{field}"))
            }
            _ => None,
        }
    }

    pub(super) fn infer_vb_runtime_member_type_hint(&self, expr: &Expression) -> Option<String> {
        let path = Self::member_access_path(expr)?;
        match self.canon(&path).as_str() {
            "environment.currentdirectory"
            | "environment.newline"
            | "environment.machinename"
            | "environment.username"
            | "environment.osversion"
            | "system.environment.currentdirectory"
            | "system.environment.newline"
            | "system.environment.machinename"
            | "system.environment.username"
            | "system.environment.osversion"
            | "app.path"
            | "app.title" => Some("string".into()),
            "environment.processorcount"
            | "environment.tickcount"
            | "system.environment.processorcount"
            | "system.environment.tickcount"
            | "screen.width"
            | "screen.height" => Some("integer".into()),
            _ => None,
        }
    }

    /// The receiver's declared type is a class that defines an index
    /// operator — so `x[i]` is a call to it rather than a key lookup.
    /// The receiver's declared type defines `operator []=` — so `x[i] = v` is
    /// a call to it rather than a key store.
    pub(super) fn expr_has_user_index_setter(&self, expr: &Expression) -> bool {
        if self.classes_with_index_setter.is_empty() {
            return false;
        }
        self.infer_expr_type_hint(expr)
            .map(|hint| self.canon(hint.trim()))
            .is_some_and(|hint| self.classes_with_index_setter.contains(&hint))
    }

    pub(super) fn expr_has_user_indexer(&self, expr: &Expression) -> bool {
        if self.classes_with_indexer.is_empty() {
            return false;
        }
        // `canon` on both sides — the set is keyed by the class's canonical
        // name, so the hint has to be canonicalised the same way.
        self.infer_expr_type_hint(expr)
            .map(|hint| self.canon(hint.trim()))
            .is_some_and(|hint| self.classes_with_indexer.contains(&hint))
    }

    pub(super) fn infer_expr_type_hint(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
            ExprKind::Lit(Literal::Int(_)) => Some("int".into()),
            ExprKind::Lit(Literal::Float(_)) => Some("double".into()),
            ExprKind::Lit(Literal::BigInt(_)) => Some("bigint".into()),
            ExprKind::Lit(Literal::Str(_)) => Some("string".into()),
            ExprKind::Lit(Literal::Bytes(_)) => Some("bytes".into()),
            ExprKind::Lit(Literal::Bool(_)) => Some("bool".into()),
            ExprKind::Lit(Literal::Char(_)) => Some("char".into()),
            ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
            ExprKind::Unary {
                op: UnaryOp::Neg | UnaryOp::Pos,
                expr,
            } => self.infer_expr_type_hint(expr),
            ExprKind::RefOf(place) => {
                let pointee_type = match place.as_ref() {
                    PlaceExpr::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
                    PlaceExpr::Member {
                        object,
                        field,
                        null_safe,
                    } => self.infer_expr_type_hint(&Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: field.clone(),
                        null_safe: *null_safe,
                    })),
                    PlaceExpr::Index {
                        object,
                        index,
                        null_safe,
                    } => self.infer_expr_type_hint(&Expression::new(ExprKind::Index {
                        object: object.clone(),
                        index: index.clone(),
                        null_safe: *null_safe,
                    })),
                    PlaceExpr::Deref(expr) => self.infer_expr_type_hint(expr).map(|type_hint| {
                        type_hint
                            .trim()
                            .trim_end_matches('?')
                            .trim()
                            .trim_start_matches('*')
                            .trim_start_matches('^')
                            .trim()
                            .to_string()
                    }),
                }?;
                Some(format!("*{}", pointee_type.trim()))
            }
            ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr,
            } => self
                .infer_expr_type_hint(expr)
                .map(|type_hint| format!("*{}", type_hint.trim().trim_end_matches('?').trim())),
            ExprKind::Unary {
                op: UnaryOp::Deref,
                expr,
            }
            | ExprKind::RefLoad(expr) => self.infer_expr_type_hint(expr).map(|type_hint| {
                type_hint
                    .trim()
                    .trim_end_matches('?')
                    .trim()
                    .trim_start_matches('*')
                    .trim_start_matches('^')
                    .trim()
                    .to_string()
            }),
            ExprKind::New { class, .. } => Self::expr_terminal_type_name(class),
            ExprKind::Array(elements) => Some(format!(
                "{}()",
                self.infer_array_element_type_hint(elements.iter().map(|item| &item.value))
            )),
            ExprKind::Call { callee, args, .. } => {
                if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Array"))
                {
                    return Some(format!(
                        "{}()",
                        self.infer_array_element_type_hint(args.iter().map(|arg| &arg.value))
                    ));
                }
                // JS conversion builtins have a known result type — e.g.
                // `BigInt(x)` is a BigInt, so `BigInt(a) % BigInt(b)` routes
                // through the `ecma:bigint` ops instead of f64 arithmetic.
                if self.profile.has_ecma_bigint {
                    if let ExprKind::Ident(name) = &callee.kind {
                        match name.as_str() {
                            "BigInt" => return Some("bigint".into()),
                            "Number" | "parseInt" | "parseFloat" => return Some("double".into()),
                            "String" => return Some("string".into()),
                            "Boolean" => return Some("bool".into()),
                            _ => {}
                        }
                    }
                }
                // `Foo(...)` naming a declared class constructs one, so the
                // call's type is that class. Languages that spell construction
                // without `new` (Dart, Python) arrive here rather than at
                // `ExprKind::New`, and would otherwise have no type at all.
                if let ExprKind::Ident(name) = &callee.kind {
                    if self.defined_classes.contains(&self.canon(name)) {
                        return Some(name.clone());
                    }
                }
                if self.profile.parens_for_index
                    && args.len() == 1
                    && self
                        .infer_expr_type_hint(callee)
                        .as_deref()
                        .map(Self::normalize_type_hint)
                        .is_some_and(|type_hint| {
                            type_hint.ends_with("()") && !Self::is_callable_type_hint(&type_hint)
                        })
                {
                    return self.infer_expr_type_hint(callee).and_then(|type_hint| {
                        type_hint
                            .trim()
                            .trim_end_matches('?')
                            .trim()
                            .strip_suffix("()")
                            .map(str::to_string)
                    });
                }
                if self.profile.namespaces.use_dotnet {
                    if let ExprKind::Member { object, field, .. } = &callee.kind {
                        if let Some(receiver_type) = self.infer_expr_type_hint(object) {
                            if self
                                .resolve_pending_class_name_for_type_hint(&receiver_type)
                                .is_none()
                            {
                                let class_name = Self::normalize_type_hint(&receiver_type);
                                if let Some(return_type) =
                                    vybe_runtime::namespaces::lookup_type_member_return(
                                        &self.profile.namespaces.type_scopes,
                                        &class_name,
                                        field,
                                    )
                                {
                                    return Some(return_type);
                                }
                            }
                        }
                    }
                }
                self.infer_function_return_type(callee)
                    .or_else(|| self.infer_dotnet_factory_return_type(callee))
            }
            ExprKind::Index { object, .. } => {
                self.infer_expr_type_hint(object).and_then(|type_hint| {
                    let trimmed = type_hint.trim().trim_end_matches('?').trim();
                    trimmed
                        .strip_suffix("()")
                        .map(str::to_string)
                        .or_else(|| Self::pascal_indexed_type_hint(trimmed))
                })
            }
            ExprKind::Member { object, field, .. } => {
                if let Some(type_hint) = self.infer_vb_runtime_member_type_hint(expr) {
                    return Some(type_hint);
                }
                if let Some(receiver_type) = self.infer_expr_type_hint(object) {
                    if let Some(class_name) =
                        self.resolve_pending_class_name_for_type_hint(&receiver_type)
                    {
                        if let Some(type_hint) = self
                            .pending_classes
                            .get(class_name.as_str())
                            .and_then(|pending| {
                                pending.instance_field_types.get(&self.canon(field))
                            })
                        {
                            return Some(type_hint.clone());
                        }
                    }
                }
                let enum_type = Self::expr_terminal_type_name(object)?;
                self.enum_value_names
                    .contains_key(&self.canon(&enum_type))
                    .then_some(enum_type)
            }
            ExprKind::Binary { op, left, right }
                if matches!(
                    op,
                    BinOp::Add
                        | BinOp::Sub
                        | BinOp::Mul
                        | BinOp::Div
                        | BinOp::Mod
                        | BinOp::Pow
                        | BinOp::BitAnd
                        | BinOp::BitOr
                        | BinOp::BitXor
                        | BinOp::Shl
                        | BinOp::Shr
                ) =>
            {
                // BigInt is contagious through arithmetic: if EITHER operand
                // is a BigInt, the result is a BigInt (the op-selection in
                // expressions.rs routes to `ecma:bigint`, and a mix with a
                // known Number throws at runtime). Inferring through chains
                // like `(a * b) % c` keeps every step on the bigint path even
                // when intermediate results have no other type evidence.
                let left_bigint = self.infer_expr_type_hint(left).as_deref() == Some("bigint");
                let right_bigint = self.infer_expr_type_hint(right).as_deref() == Some("bigint");
                if left_bigint || right_bigint {
                    Some("bigint".into())
                } else {
                    None
                }
            }
            ExprKind::Binary { op, left, right }
                if matches!(op, BinOp::BitOr | BinOp::BitAnd | BinOp::BitXor) =>
            {
                let left_type = self.infer_expr_type_hint(left)?;
                let right_type = self.infer_expr_type_hint(right)?;
                if left_type.eq_ignore_ascii_case(&right_type)
                    && self.enum_value_names.contains_key(&self.canon(&left_type))
                {
                    Some(left_type)
                } else {
                    None
                }
            }
            _ => None,
        }
    }

    pub(super) fn user_value_type_name_from_hint(&self, type_hint: &str) -> Option<String> {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();
        if trimmed.starts_with('*')
            || trimmed.starts_with('^')
            || trimmed.starts_with("[]")
            || trimmed.starts_with("map[")
            || trimmed.starts_with("chan ")
            || trimmed.starts_with("func(")
        {
            return None;
        }

        if let Some(class_name) = self.resolve_pending_class_name_for_type_hint(type_hint) {
            if self
                .pending_classes
                .get(&class_name)
                .is_some_and(|pending| pending.is_value_type)
            {
                return Some(class_name);
            }
        }

        for candidate in [
            Some(trimmed),
            trimmed
                .rsplit('.')
                .next()
                .filter(|segment| *segment != trimmed),
        ]
        .into_iter()
        .flatten()
        {
            let canonical = self.canon(candidate);
            if self
                .pending_classes
                .get(&canonical)
                .is_some_and(|pending| pending.is_value_type)
            {
                return Some(canonical);
            }
            if let Some((name, _)) = self.pending_classes.iter().find(|(name, pending)| {
                pending.is_value_type && name.eq_ignore_ascii_case(candidate)
            }) {
                return Some(name.clone());
            }
        }
        None
    }

    pub(super) fn pascal_expr_is_integer_like(&self, expr: &Expression) -> bool {
        if self.profile.name != "pascal" {
            return false;
        }
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_)) => return true,
            ExprKind::Lit(Literal::Float(_) | Literal::Bool(_) | Literal::Str(_)) => return false,
            ExprKind::Unary { op, expr } => {
                return matches!(op, UnaryOp::Not | UnaryOp::BitNot)
                    && self.pascal_expr_is_integer_like(expr);
            }
            ExprKind::Binary { op, left, right } => {
                return matches!(
                    op,
                    BinOp::BitAnd
                        | BinOp::BitOr
                        | BinOp::BitXor
                        | BinOp::Shl
                        | BinOp::Shr
                        | BinOp::And
                        | BinOp::Or
                ) && self.pascal_expr_is_integer_like(left)
                    && self.pascal_expr_is_integer_like(right);
            }
            _ => {}
        }
        let Some(type_hint) = self.infer_expr_type_hint(expr) else {
            return false;
        };
        matches!(
            Self::normalize_type_hint(&self.resolve_source_type_alias(&type_hint)).as_str(),
            "integer"
                | "int"
                | "longint"
                | "shortint"
                | "smallint"
                | "byte"
                | "word"
                | "cardinal"
                | "int64"
                | "uint64"
                | "longword"
        )
    }

    pub(super) fn expr_user_value_type_name(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .and_then(|type_hint| self.user_value_type_name_from_hint(type_hint)),
            _ => self
                .infer_expr_type_hint(expr)
                .and_then(|type_hint| self.user_value_type_name_from_hint(&type_hint)),
        }
    }

    pub(super) fn expr_is_array_like(&self, expr: &Expression) -> bool {
        if self
            .infer_expr_type_hint(expr)
            .as_deref()
            .map(Self::normalize_type_hint)
            .is_some_and(|type_hint| {
                type_hint.ends_with("()") && !Self::is_callable_type_hint(&type_hint)
            })
        {
            return true;
        }

        match &expr.kind {
            ExprKind::Array(_) => true,
            ExprKind::Ident(name) => self.lookup_array_binding(name).is_some(),
            ExprKind::Index { object, index, .. } => {
                matches!(index.kind, ExprKind::Slice { .. }) && self.expr_is_array_like(object)
            }
            ExprKind::Call { callee, .. } => {
                matches!(&callee.kind, ExprKind::Ident(name)
                    if matches!(self.canon(name).as_str(), "array" | "str_split" | "str_getcsv"))
            }
            ExprKind::Binary { op, left, right }
                if matches!(
                    op,
                    BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow
                ) =>
            {
                self.expr_is_array_like(left) || self.expr_is_array_like(right)
            }
            _ => false,
        }
    }

    pub(super) fn vb_generic_type_display_name(&self, type_hint: &str) -> Option<String> {
        let trimmed = type_hint.trim().trim_end_matches('?').trim();
        let short_name = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();

        let angle_arity = self.reflection_generic_argument_types(trimmed).len();
        if angle_arity > 0 {
            let base = short_name.split('<').next().unwrap_or(short_name).trim();
            return Some(format!("{base}`{angle_arity}"));
        }

        let lowered = trimmed.to_lowercase();
        let marker = "(of ";
        let start = lowered.find(marker)?;
        let base = trimmed[..start]
            .trim()
            .rsplit('.')
            .next()
            .unwrap_or(trimmed[..start].trim())
            .trim();
        let inner = trimmed.get(start + marker.len()..trimmed.len().saturating_sub(1))?;
        let mut depth = 0usize;
        let mut arity = 1usize;
        for ch in inner.chars() {
            match ch {
                '(' | '<' => depth += 1,
                ')' | '>' => depth = depth.saturating_sub(1),
                ',' if depth == 0 => arity += 1,
                _ => {}
            }
        }
        Some(format!("{base}`{arity}"))
    }

    pub(super) fn vb_reflection_display_type_name(&self, type_name: &str) -> Option<String> {
        let trimmed = type_name.trim().trim_end_matches('?').trim();
        let short_target = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
        self.reflection_types
            .keys()
            .find(|candidate| {
                candidate.eq_ignore_ascii_case(trimmed)
                    || candidate
                        .rsplit('.')
                        .next()
                        .is_some_and(|leaf| leaf.eq_ignore_ascii_case(short_target))
            })
            .map(|candidate| self.reflection_type_short_name(candidate))
    }

    pub(super) fn vb_typename_from_type_hint(&self, type_hint: &str) -> Option<String> {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();

        if let Some(element_type) = trimmed.strip_suffix("()") {
            return self
                .vb_typename_from_type_hint(element_type.trim())
                .map(|name| format!("{name}()"));
        }

        let normalized = Self::normalize_type_hint(trimmed);
        let primitive = match normalized.as_str() {
            "integer" | "int" | "int32" | "longint" | "system.int32" => Some("Integer"),
            "long" | "int64" | "system.int64" => Some("Long"),
            "short" | "int16" | "system.int16" => Some("Short"),
            "ushort" | "uint16" | "system.uint16" => Some("UShort"),
            "uint" | "uint32" | "system.uint32" => Some("UInteger"),
            "ulong" | "uint64" | "system.uint64" => Some("ULong"),
            "byte" | "system.byte" => Some("Byte"),
            "sbyte" | "system.sbyte" => Some("SByte"),
            "single" | "float" | "system.single" => Some("Single"),
            "double" | "real" | "system.double" => Some("Double"),
            "decimal" | "system.decimal" => Some("Decimal"),
            "boolean" | "bool" | "system.boolean" => Some("Boolean"),
            "char" | "system.char" => Some("Char"),
            "string" | "system.string" => Some("String"),
            "datetime" | "date" | "system.datetime" => Some("Date"),
            "object" | "system.object" => Some("Object"),
            _ => None,
        };
        if let Some(name) = primitive {
            return Some(name.into());
        }

        if let Some(name) = self.vb_generic_type_display_name(trimmed) {
            return Some(name);
        }

        if let Some(name) = self.vb_reflection_display_type_name(trimmed) {
            return Some(name);
        }

        let short_target = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
        if let Some((display_name, _)) = self.pending_classes.iter().find(|(candidate, _)| {
            candidate.eq_ignore_ascii_case(trimmed)
                || candidate
                    .rsplit('.')
                    .next()
                    .is_some_and(|leaf| leaf.eq_ignore_ascii_case(short_target))
        }) {
            return Some(
                display_name
                    .rsplit('.')
                    .next()
                    .unwrap_or(display_name)
                    .to_string(),
            );
        }

        if self.reflection_type_metadata(trimmed).is_some() || self.reflection_is_enum_type(trimmed)
        {
            return Some(self.reflection_type_short_name(trimmed));
        }
        None
    }

    pub(super) fn vb_typename_from_expr(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_)) => Some("Integer".into()),
            ExprKind::Lit(Literal::Float(_)) => Some("Double".into()),
            ExprKind::Lit(Literal::Str(_)) => Some("String".into()),
            ExprKind::Lit(Literal::Bool(_)) => Some("Boolean".into()),
            ExprKind::Lit(Literal::Char(_)) => Some("Char".into()),
            ExprKind::Lit(Literal::Null | Literal::Undefined) => Some("Nothing".into()),
            _ => self
                .infer_expr_type_hint(expr)
                .and_then(|type_hint| self.vb_typename_from_type_hint(&type_hint)),
        }
    }

    pub(super) fn vb_is_reference_type_hint(&self, type_hint: &str) -> bool {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();
        if trimmed.ends_with("()") {
            return true;
        }
        if self.reflection_is_enum_type(trimmed) || self.reflection_is_value_type(trimmed) {
            return false;
        }
        match self.vb_typename_from_type_hint(trimmed).as_deref() {
            Some(
                "Integer" | "Long" | "Short" | "UShort" | "UInteger" | "ULong" | "Byte" | "SByte"
                | "Single" | "Double" | "Decimal" | "Boolean" | "Char" | "Date",
            ) => false,
            Some("String" | "Object") => true,
            Some(name) if name.ends_with("()") => true,
            Some(_) => true,
            None => false,
        }
    }

    pub(super) fn vb_is_object_type_hint(&self, type_hint: &str) -> bool {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();
        if trimmed.ends_with("()") {
            return true;
        }
        if self.reflection_is_enum_type(trimmed) || self.reflection_is_value_type(trimmed) {
            return false;
        }
        match self.vb_typename_from_type_hint(trimmed).as_deref() {
            Some("Object") => true,
            Some(
                "String" | "Integer" | "Long" | "Short" | "UShort" | "UInteger" | "ULong" | "Byte"
                | "SByte" | "Single" | "Double" | "Decimal" | "Boolean" | "Char" | "Date",
            ) => false,
            Some(name) if name.ends_with("()") => true,
            Some(_) => true,
            None => false,
        }
    }

    pub(super) fn vb_is_reference_expr(&self, expr: &Expression) -> Option<bool> {
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_))
            | ExprKind::Lit(Literal::Float(_))
            | ExprKind::Lit(Literal::Bool(_))
            | ExprKind::Lit(Literal::Char(_))
            | ExprKind::Lit(Literal::Null | Literal::Undefined) => Some(false),
            ExprKind::Lit(Literal::Str(_)) => Some(true),
            _ => self
                .infer_expr_type_hint(expr)
                .map(|type_hint| self.vb_is_reference_type_hint(&type_hint)),
        }
    }

    pub(super) fn vb_is_object_expr(&self, expr: &Expression) -> Option<bool> {
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_))
            | ExprKind::Lit(Literal::Float(_))
            | ExprKind::Lit(Literal::Bool(_))
            | ExprKind::Lit(Literal::Char(_))
            | ExprKind::Lit(Literal::Str(_))
            | ExprKind::Lit(Literal::Null | Literal::Undefined) => Some(false),
            _ => self
                .infer_expr_type_hint(expr)
                .map(|type_hint| self.vb_is_object_type_hint(&type_hint)),
        }
    }

    pub(super) fn compile_expr_with_value_copy(&mut self, expr: &Expression) -> Result<(), String> {
        self.compile_expr(expr)?;
        let should_clone = matches!(
            expr.kind,
            ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. }
        );
        if should_clone {
            if let Some(type_name) = self.expr_user_value_type_name(expr) {
                self.emit_user_value_type_clone_from_stack(&type_name);
            }
        }
        Ok(())
    }

    pub(super) fn emit_array_clone_from_stack(&mut self) {
        let source_slot = self.define_local("__array_clone_src");
        let len_slot = self.define_local("__array_clone_len");

        self.emit_u16(Op::LOCAL_SET, source_slot);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        common::collections::emit_len(&mut self.chunks, self.current, self.line);
        self.emit_u16(Op::LOCAL_SET, len_slot);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_slice(&mut self.chunks, self.current, self.line);
    }

    pub(super) fn emit_user_value_type_clone_from_stack(&mut self, type_name: &str) {
        let Some((fields, instance_member_names)) =
            self.pending_classes.get(type_name).map(|pending| {
                (
                    pending.fields.clone(),
                    pending.instance_member_names.clone(),
                )
            })
        else {
            return;
        };

        let source_slot = self.define_local("__value_type_src");
        self.emit_u16(Op::LOCAL_SET, source_slot);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.chunk().emit_else(line);

        let clone_slot = self.define_local("__value_type_clone");
        crate::primitives::classes::emit_new_typed_object(
            self.chunk(),
            clone_slot,
            type_name,
            line,
        );

        for member_name in fields.iter().chain(instance_member_names.iter()) {
            let member_key = self.str_const(member_name);
            self.emit_u16(Op::LOCAL_GET, clone_slot);
            self.emit_u16(Op::LOCAL_GET, source_slot);
            self.emit_u16(Op::STRUCT_GET, member_key);
            self.emit_u16(Op::STRUCT_SET, member_key);
            self.emit(Op::DROP);
        }

        self.emit_u16(Op::LOCAL_GET, clone_slot);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
    }

    pub(super) fn expr_is_known_string_receiver(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(_)) | ExprKind::Interpolation(_) => true,
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(Self::is_string_type_hint),
            _ => false,
        }
    }
}
