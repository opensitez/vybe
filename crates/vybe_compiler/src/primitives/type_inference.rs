//! Type-hint inference and type-name resolution.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use crate::primitives::class_slots;
use super::*;

impl Compiler {
    /// Is `field` a declared INSTANCE FIELD of this class (or an ancestor)?
    ///
    /// A control's field is not one of its properties. `TForm1` holding
    /// `btn1: TButton` is a form AND has a field named `btn1`, so once a user
    /// class was recognised as a control, `Self.btn1 := TButton.Create(…)`
    /// started lowering to `setAttribute("btn1", <element>)` — the control
    /// stored as markup and the field left null. 14 tests went red on exactly
    /// that.
    pub(crate) fn is_declared_instance_field(&self, class_name: &str, field: &str) -> bool {
        let canon_field = self.canon(field);
        self.pending_class_ancestry(class_name).any(|pending| {
            pending.instance_field_types.contains_key(&canon_field)
                || pending.fields.iter().any(|f| self.canon(f) == canon_field)
        })
    }

    /// One pending class, by whatever spelling reaches here.
    fn pending_class_entry(&self, class_name: &str) -> Option<&PendingClass> {
        self.pending_classes
            .get(&self.canon(class_name))
            .or_else(|| self.pending_classes.get(class_name))
    }

    /// Every `PendingClass` in a class's ancestry, nearest first.
    ///
    /// ⛔ ONE ORDER, THREE QUESTIONS — AND THE THIRD HAD NO WALK AT ALL.
    /// `is_declared_instance_field` asks *is this a field*,
    /// `lookup_implicit_self_field_type_hint` asks *what is a bare name's
    /// declared type*, and the `ExprKind::Member` arm of `infer_expr_type_hint`
    /// asks that same question through an explicit receiver. The first two each
    /// wrote their own loop up `parent`; the third looked only in the
    /// receiver's OWN map, so
    ///
    ///     class B { public Action H; }
    ///     class D : B { }
    ///     d.H += handler;
    ///
    /// inferred nothing, fell to the arithmetic arm of `+=` and trapped with
    /// `toF64 — not a number`, while the identical line on a `B` instance
    /// worked. Only a lowering that BRANCHES ON THE DECLARED TYPE could see it:
    /// the field's storage, its reads and its plain writes were already
    /// correct, so nothing about the value was wrong — only the compile-time
    /// question *what was this declared as*.
    ///
    /// ⛔ The two existing copies also DISAGREED about canonicalization: one
    /// folded the parent's name at every step, the other did not. `canon` is
    /// the policy the field KEYS already use, so folding is the correct half —
    /// un-folded, an inherited field was unreachable in a case-insensitive
    /// scope whenever a parent was spelled differently from its declaration.
    ///
    /// The walk STOPS at the first ancestor this compilation has no pending
    /// class for — a framework or catalog parent. That is what both copies
    /// already did, and it is deliberate: nothing further up can be answered
    /// from this table, and a `None` that means *stop* must not be confused
    /// with one that means *no such field*.
    fn pending_class_ancestry<'a>(
        &'a self,
        class_name: &str,
    ) -> impl Iterator<Item = &'a PendingClass> + 'a {
        std::iter::successors(self.pending_class_entry(class_name), move |pending| {
            pending
                .parent
                .as_deref()
                .and_then(|parent| self.pending_class_entry(parent))
        })
    }

    /// The declared type of an instance field, looked up through the ancestry.
    pub(super) fn lookup_instance_field_type(
        &self,
        class_name: &str,
        field: &str,
    ) -> Option<&FieldType> {
        let canon_field = self.canon(field);
        self.pending_class_ancestry(class_name)
            .find_map(|pending| pending.instance_field_types.get(&canon_field))
    }

    /// The declared parent of a user class, by whatever spelling reaches here.
    pub(crate) fn pending_class_parent(&self, class_name: &str) -> Option<String> {
        let canon = self.canon(class_name);
        self.pending_classes
            .get(&canon)
            .or_else(|| self.pending_classes.get(class_name))
            .and_then(|pending| pending.parent.clone())
    }

    pub(super) fn lookup_implicit_self_field_type_hint(&self, name: &str) -> Option<&str> {
        if !self.current_class_implicit_self {
            return None;
        }

        let class_name = self.current_class.as_deref()?;
        self.lookup_instance_field_type(class_name, name)
            .map(|field_type| field_type.hint.as_str())
    }

    /// Unused since case folding became a scope policy: its only caller
    /// refined the fold-only half of a local-vs-type decision, which a folding
    /// scope cannot distinguish. Kept because the predicate itself is sound.
    #[allow(dead_code)]
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

    pub(super) fn infer_namespace_tree_factory_return_type(
        &self,
        callee: &Expression,
    ) -> Option<String> {
        if self.profile.namespaces.type_scopes.is_empty() {
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
        self.tree_member_return(&class_name, field)
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
            // The type of `self` is the class being compiled. Without this the
            // Member arm below has no receiver type for `self.field`, so a
            // field's declared type was known at top level (`var lbl: TLabel`)
            // and unknown one line into a method — the same expression taking
            // two different paths.
            // `self` is an ordinary BINDING, so its type is read off its slot
            // like any other name — `compile_class` declares it where the
            // class is known (see `define_local_typed(&self_kw, …)`).
            //
            // Reading the slot rather than asserting `current_class` is what
            // makes this safe where `this` is dynamically rebound: JS's
            // ambient-`this` local is loaded from `__js_this`, whose value is
            // decided at CALL time, so that site declares no type and
            // inference correctly answers None instead of naming the
            // enclosing class.
            //
            // Measured neutral name-for-name over 164 class tests across
            // csharp/java/vb/js/python, including the JS lexical-`this`
            // categories.
            ExprKind::This => {
                let self_kw = self.profile.self_keyword.clone();
                self.lookup_var_type_hint(&self_kw).map(str::to_string)
            }
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
            ExprKind::New { class, .. } => Self::expr_terminal_type_name(class)
                .map(|name| self.resolve_source_type_alias(&name)),
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
                // `Enum.Parse(typeof(T), s)` answers with a `T`, and saying so
                // is what keeps a parsed value an ENUM to everything
                // downstream. A parse result is a member OBJECT now, so a site
                // that asks "is this enum-typed?" and hears nothing falls to a
                // generic path that cannot read the `ToString` role: C#
                // `p.ToString()` trapped in a numeric coercion and VB's
                // `Console.WriteLine(p)` printed `[object Status]`. Both are
                // one missing answer, not two rendering bugs — fixing either
                // print path alone would have left the other broken.
                //
                // The type is already in the first argument: the same
                // runtime-type string `Enum.Parse` itself reads to find the
                // members, read here through the same helper.
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if field.eq_ignore_ascii_case("Parse")
                        && args.len() >= 2
                        && Self::member_access_path(object).is_some_and(|path| {
                            path.eq_ignore_ascii_case("Enum")
                                || path.eq_ignore_ascii_case("System.Enum")
                        })
                    {
                        if let Some(enum_type) =
                            self.canonical_enum_type_from_runtime_type(&args[0].value)
                        {
                            return Some(enum_type);
                        }
                    }
                }
                // JS conversion builtins have a known result type — e.g.
                // `BigInt(x)` is a BigInt, so `BigInt(a) % BigInt(b)` routes
                // through the `ecma:bigint` ops instead of f64 arithmetic.
                if self.bigint_semantics() {
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
                // A call `compile_call` lowers to a DOM element is worth that
                // control type. Same predicate, asked from the other end — see
                // `constructed_control_type_name`. Without this the value has
                // NO type, so `Window.Forms.Button()` assigned to a widened
                // declaration stayed the top type and every `.Text`/`.Left`
                // write on it missed the DOM path and landed on a plain object
                // property: a control with no caption at the origin.
                if let Some(control_type) = self.constructed_control_type_name(callee) {
                    return Some(control_type);
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
                if !self.profile.namespaces.type_scopes.is_empty() {
                    if let ExprKind::Member { object, field, .. } = &callee.kind {
                        if let Some(receiver_type) = self.infer_expr_type_hint(object) {
                            if self
                                .resolve_pending_class_name_for_type_hint(&receiver_type)
                                .is_none()
                            {
                                let class_name = Self::tree_type_key(&receiver_type);
                                if let Some(return_type) =
                                    self.tree_member_return(&class_name, field)
                                {
                                    return Some(return_type);
                                }
                            }
                        }
                    }
                }
                self.infer_function_return_type(callee)
                    .or_else(|| self.infer_namespace_tree_factory_return_type(callee))
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
                        // Through the ANCESTRY, not the receiver's own map —
                        // see `pending_class_ancestry` for what a single-level
                        // lookup here cost.
                        if let Some(field_type) =
                            self.lookup_instance_field_type(class_name.as_str(), field)
                        {
                            return Some(field_type.hint.clone());
                        }
                    }
                }
                // A PROPERTY read carries the same tree-declared return
                // type as a member call — the Call arm below already asks
                // `lookup_type_member_return`; without the symmetric ask
                // here, chains through properties (`date.dayOfWeek.value`)
                // lost their type at the property hop.
                if !self.profile.namespaces.type_scopes.is_empty() {
                    if let Some(receiver_type) = self.infer_expr_type_hint(object) {
                        if self
                            .resolve_pending_class_name_for_type_hint(&receiver_type)
                            .is_none()
                        {
                            let class_name = Self::tree_type_key(&receiver_type);
                            if let Some(return_type) =
                                self.tree_member_return(&class_name, field)
                            {
                                return Some(return_type);
                            }
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
                let left_bigint = self.hint_is_bigint(self.infer_expr_type_hint(left).as_deref());
                let right_bigint = self.hint_is_bigint(self.infer_expr_type_hint(right).as_deref());
                if left_bigint || right_bigint {
                    Some("bigint".into())
                } else {
                    None
                }
            }
            // UNREACHABLE, and load-bearing that way. The arm above matches the
            // same three ops with no further guard, so a bitwise `Binary`
            // always lands there and answers `None` unless an operand is a
            // BigInt. Rust warns about neither, because both arms are guarded.
            //
            // What depends on it: `expressions.rs` lowers `a | b` on enum
            // operands to the underlying INTEGER, so `var c = A | B;` must
            // infer as not-an-enum. If this arm were ever reordered ahead of
            // the BigInt one, `(int)c` would resolve `c` as enum-typed and read
            // the `Int` role off a raw number. Delete it or leave it dead —
            // do not promote it without changing the flags lowering with it.
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

    /// Does this spelling denote a type STORED BY VALUE, or one that merely
    /// refers to one?
    ///
    /// `^TNode`, `*Point`, `[]Point`, `map[string]Point`, `chan Point` and
    /// `func(Point)` all name a value type without storing it, so copying the
    /// owner must not copy through them. The test has to survive alias
    /// resolution: Pascal's `type PNode = ^TNode` resolves to the pointee, and
    /// treating the result as a stored value deep-copies a pointer field —
    /// which is exactly what `test_pointers_advanced::typed_pointer_in_record_field`
    /// catches.
    ///
    /// Shared by the declaration-pass resolution and by the emit-time lookup so
    /// the two cannot drift; it is the only spelling-level judgement either of
    /// them makes.
    pub(super) fn type_hint_stores_by_value(resolved: &str) -> Option<&str> {
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
        Some(trimmed)
    }

    pub(super) fn user_value_type_name_from_hint(&self, type_hint: &str) -> Option<String> {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = Self::type_hint_stores_by_value(&resolved)?;

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

    /// Is `expr` statically an INTEGER, for languages whose `and`/`or`/`not`
    /// are bitwise on integers and logical on booleans?
    ///
    /// Gated on `logical_ops_bitwise_for_integers`, so the six shared call
    /// sites need no per-language condition of their own.
    pub(super) fn expr_is_integer_like(&self, expr: &Expression) -> bool {
        if !self.profile.logical_ops_bitwise_for_integers {
            return false;
        }
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_)) => return true,
            ExprKind::Lit(Literal::Float(_) | Literal::Bool(_) | Literal::Str(_)) => return false,
            ExprKind::Unary { op, expr } => {
                return matches!(op, UnaryOp::Not | UnaryOp::BitNot)
                    && self.expr_is_integer_like(expr);
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
                ) && self.expr_is_integer_like(left)
                    && self.expr_is_integer_like(right);
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

    /// Is `expr` already EXACTLY representable in `normalized_hint`, so that the
    /// width coercion in `coerce_c_value_for_type_hint` is a mathematical no-op?
    ///
    /// This is not an optimization heuristic — it answers "would the coercion
    /// change the value?", and only `false` is ever the safe-but-slow answer.
    /// Every arm below must be a value the coercion provably maps to itself.
    ///
    /// Why it matters: that coercion is a modular-arithmetic chain in f64
    /// (`trunc`, three `%`, an add, and for signed types a compare + subtract).
    /// Emitting it for `int k = 0` costs ~20 opcodes to compute `0`. Measured,
    /// a C loop body ran **302 executed opcodes per iteration** for work that is
    /// ~11 opcodes of WASM — see `statictypelowering.md`.
    ///
    /// Deliberately conservative in this first cut:
    /// - **Literals** are checked against the target's exact range.
    /// - **Bitwise ops** already leave an exact i32 (the compiler emits
    ///   `I32_AND`/`I32_OR`/… for them), so they are exact for `int`.
    ///   `UShr` is EXCLUDED — under ECMA coercion it yields a u32, which can
    ///   exceed `i32::MAX`.
    /// - **Idents are NOT trusted.** A variable is coerced at its own
    ///   declaration and assignments, but a typed *parameter* is not coerced on
    ///   entry, so `int y = x;` could legitimately need the wrap.
    fn value_is_exact_in_type_hint(&self, expr: &Expression, normalized_hint: &str) -> bool {
        // Same ONE spelling table the narrowing emitter resolves through, so
        // the two can never disagree about what a width is.
        let Some(width) = vybe_ast::builtin_types::int_width_of(normalized_hint) else {
            return false;
        };
        let range = width.range();
        let is_i32 = width == vybe_ast::builtin_types::IntWidth::I32;

        match &expr.kind {
            ExprKind::Lit(Literal::Int(n)) => *n >= range.0 && *n <= range.1,

            ExprKind::Binary { op, left, right } => {
                // A bitwise op only leaves an exact i32 when it really is the
                // i32 opcode. If either operand is a user value type, the
                // operator may be OVERLOADED (C# `operator &`, Python
                // `__and__`) and can return anything — so claim nothing.
                if self.expr_user_value_type_name(left).is_some()
                    || self.expr_user_value_type_name(right).is_some()
                {
                    return false;
                }
                match op {
                    // `x & MASK` is bounded by the mask whenever the mask is a
                    // non-negative literal: every result bit is a mask bit.
                    // This is what makes `(unsigned char)(i & 0xFF)` free.
                    BinOp::BitAnd => {
                        let mask_fits = |e: &Expression| {
                            matches!(&e.kind, ExprKind::Lit(Literal::Int(m))
                                if *m >= 0 && *m <= range.1)
                        };
                        mask_fits(left) || mask_fits(right)
                    }
                    // The remaining bitwise ops leave an exact i32, which is
                    // exact for `int` but says nothing about narrower widths.
                    BinOp::BitOr | BinOp::BitXor | BinOp::Shl | BinOp::Shr => is_i32,
                    _ => false,
                }
            }

            _ => false,
        }
    }

    /// Would `coerce_c_value_for_type_hint` be a no-op for this value?
    pub(super) fn coercion_is_redundant(&self, expr: &Expression, type_hint: Option<&str>) -> bool {
        let Some(type_hint) = type_hint else {
            return false;
        };
        let normalized = Self::normalize_type_hint(type_hint);
        // A `char` that holds a CHARACTER is not a numeric width at all; the
        // coercion arm skips it too, so never claim to have narrowed one.
        if normalized == "char" && self.hint_is_builtin_string(&normalized) {
            return false;
        }
        self.value_is_exact_in_type_hint(expr, &normalized)
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

    /// Does this expression's STATIC type name a declared value type — one
    /// whose language said `b = a` hands back an independent value?
    ///
    /// The `__value_copy` instance stamp is what carries the semantics across a
    /// language boundary, but it is only written where an instance is
    /// CONSTRUCTED. Pascal's `var a, b: TR` default-initialises its records
    /// from a synthesized literal and never runs a constructor, so nothing
    /// stamps them — which is why Go's `P{X: 1}` and C#'s `new S()` gained
    /// value semantics from the shared path and Pascal's declaration form did
    /// not.
    ///
    /// So the static type answers it when it can, and the stamp answers it when
    /// the type is unknown — a value that arrived from another language.
    pub(super) fn expr_is_declared_value_type(&self, expr: &Expression) -> bool {
        let Some(hint) = self.infer_expr_type_hint(expr) else {
            return false;
        };
        // `normalized_classes` is keyed by the CANONICAL name the declaration
        // pass inserted (`link.rs`), so ask through `canon` — a raw or
        // normalized spelling misses, silently, and the copy never happens.
        let bare = hint.split('<').next().unwrap_or(&hint).trim().to_string();
        let canon = self.canon(&bare);
        self.normalized_classes
            .get(&canon)
            .or_else(|| self.normalized_classes.get(&bare))
            .or_else(|| {
                self.normalized_classes
                    .get(&Self::normalize_type_hint(&hint))
            })
            .is_some_and(|nc| nc.is_value_type)
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
        let mut in_progress = Vec::new();
        self.emit_user_value_type_clone_inner(type_name, &mut in_progress);
    }

    /// A field whose own declared type is a value type has to be copied too,
    /// otherwise `b := a` hands back an outer copy pointing at the SAME inner
    /// record — `b.I.V := 99` then mutates `a.I.V`. Three lines of Pascal
    /// demonstrate it, and the flat case passed for a long time while this one
    /// silently aliased.
    ///
    /// The recursion is at COMPILE time, driven by the declared field type, so
    /// no instance stamp is involved and the copy keeps its rtt at every level.
    /// That matters because the runtime alternative cannot: a generic walk over
    /// an object's own keys has nothing to allocate the copy *as*, so it
    /// produces a shape-alike that has lost its type identity.
    ///
    /// `in_progress` breaks a declaration cycle. A record cannot contain itself
    /// BY VALUE in any language that has records — the size would be infinite —
    /// but a malformed or mutually-recursive declaration must not expand
    /// forever at compile time, and inlining is what makes that a hang rather
    /// than a stack overflow at runtime.
    fn emit_user_value_type_clone_inner(&mut self, type_name: &str, in_progress: &mut Vec<String>) {
        if in_progress.iter().any(|n| n == type_name) {
            return;
        }
        let Some((fields, instance_member_names, field_types)) =
            self.pending_classes.get(type_name).map(|pending| {
                (
                    pending.fields.clone(),
                    pending.instance_member_names.clone(),
                    pending.instance_field_types.clone(),
                )
            })
        else {
            return;
        };
        in_progress.push(type_name.to_string());

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
        // The clone must carry the SAME rtt as its source, so it allocates
        // through `struct.new_default $T` like every other instance. The type
        // is already registered by the time a value-type copy is emitted, so
        // its 1-based table index is just its position.
        // `type_name` arrives CANONICAL (it came from a `pending_classes` key)
        // while `types` holds the name as DECLARED, so an exact compare misses
        // for every case-folding language and silently yields slot 0 — a clone
        // with the wrong rtt, which is the one thing this path exists to get
        // right. Fold through `canon` rather than an ASCII compare, so the
        // language's own case policy decides.
        let type_slot = self.chunks[0]
            .types
            .iter()
            .position(|t| t.name == type_name || self.canon(&t.name) == type_name)
            .map(|i| i as u16 + 1)
            .unwrap_or(0);
        crate::primitives::classes::emit_new_typed_object(
            self.chunk(),
            clone_slot,
            type_name,
            type_slot,
            line,
        );

        for member_name in fields.iter().chain(instance_member_names.iter()) {
            let member_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal(member_name));
            // Resolved in the declaration pass, keyed by the same storage name
            // this loop iterates. Nothing is derived from a spelling here — a
            // method member simply has no entry.
            let nested = field_types
                .get(member_name)
                .and_then(|field_type| field_type.value_type.clone());

            self.emit_u16(Op::LOCAL_GET, source_slot);
            self.class_get_resolved(class_slots::ObjSource::Stack, &member_key);
            if let Some(nested_type) = nested {
                self.emit_user_value_type_clone_inner(&nested_type, in_progress);
            }
            // Sink the field value before pushing the destination. The nested
            // clone emits its own `if`/`else` blocks and local traffic, and
            // leaving `clone_slot` on the stack underneath them is the dirty-
            // stack shape that already cost a day on `CALL_REF` in `clone.rs`.
            let field_slot = self.define_local("__value_field_copy");
            self.emit_u16(Op::LOCAL_SET, field_slot);
            self.emit_u16(Op::LOCAL_GET, clone_slot);
            self.emit_u16(Op::LOCAL_GET, field_slot);
            self.class_set_resolved(
                class_slots::ObjSource::Stack,
                &member_key,
                class_slots::ValueSource::Stack,
            );
        }

        self.emit_u16(Op::LOCAL_GET, clone_slot);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        in_progress.pop();
    }

    pub(super) fn expr_is_known_string_receiver(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(_)) | ExprKind::Interpolation(_) => true,
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(Self::is_string_type_hint),
            // Anything else the compiler can already TYPE. A string is a string
            // however it was reached, and `infer_expr_type_hint` answers for the
            // shapes the two arms above cannot: an array element, a field, a
            // call's return. Without this, `Labels[I][1]` read null while the
            // identical `S := Labels[I]; S[1]` read the character — the same
            // expression taking two answers depending on whether a temporary
            // happened to be introduced. Deferring to the one resolver is what
            // keeps them the same.
            _ => self
                .infer_expr_type_hint(expr)
                .is_some_and(|type_hint| Self::is_string_type_hint(&type_hint)),
        }
    }
}
