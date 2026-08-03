//! Kotlin `ClassDecl` → `NormalClass` normalization pass.

use vybe_ast::class_normalize::{NormalMembers, build_normal_method, from_method_stmt, types::*};
use vybe_ast::{
    Argument, BinOp, ClassKind, ClassMember, ClassModifiers, ConstructorInitializerTarget,
    ExprKind, Expression, Modifiers, Param, PassBy, PropertySetter, Span, Statement, StmtKind };

/// `this.<name>`.
/// A component read from inside a derived member.
///
/// Explicitly `this.<name>`, never the bare identifier. `copy`'s parameters are
/// named after the components they replace, so a bare `n` inside `copy(n = …)`
/// binds to the PARAMETER and every default became its own self-reference.
fn this_field(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: name.to_string(),
        null_safe: false })
}

fn this_component(index: usize) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(Expression::new(ExprKind::This)),
        index: Box::new(Expression::int(index as i64)),
        null_safe: false })
}

/// Render a value the way Kotlin does — `emitter/tostring.rs` dispatches on the
/// VALUE, so a nested collection or record renders as Kotlin spells it.
fn kt_render(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__kt_tostring")),
        args: vec![Argument::positional(expr)],
        optional: false })
}

fn ret(expr: Expression) -> Vec<Statement> {
    vec![Statement::new(StmtKind::Return(Some(expr)))]
}

fn no_arg_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: true }
}

/// `ClassKind::Record` — derive the members a `data class` gets for free.
///
/// This runs in NORMALIZATION, not in the walker: the shape is a function of
/// the primary constructor's components, which is the same computation in every
/// language that has records, and the results are bound to `ProtocolSlot`s
/// rather than published as spellings (flexclassplan §2a). The walker's only
/// job was to say `kind: Record`.
fn derive_record_members(class_name: &str, out: &mut NormalMembers) {
    let Some(primary) = out.constructor.clone() else {
        return;
    };
    let components: Vec<String> = primary.params.iter().map(|p| p.name.clone()).collect();
    if components.is_empty() {
        return;
    }

    // A derived member is named exactly like a hand-written one: SOURCE spelling
    // for the call site, CANONICAL name for the slot. Naming it after its
    // spelling twice is what left `toString` unbound to the ToString slot —
    // `special_methods` keys on the canonical name (classes.rs `class_slots`).
    let method = |name: &str, params: Vec<Param>, body: Vec<Statement>| {
        let (canonical, _) = crate::protocol::canonical_method(name);
        build_normal_method(
            primary.span.clone(),
            &canonical,
            name,
            params,
            None,
            body,
            Access::Public,
            false,
            false,
            false,
            Modifiers::default(),
        )
    };

    // `component1()` … — what destructuring binds, in declaration order.
    for (idx, comp) in components.iter().enumerate() {
        let name = format!("component{}", idx + 1);
        out.instance_methods
            .push(method(&name, vec![], ret(this_field(comp))));
    }

    // `copy(a = this.a, …)` — every component defaults to the current value.
    let copy_params: Vec<Param> = components
        .iter()
        .enumerate()
        .map(|(idx, comp)| Param {
            name: comp.clone(),
            type_hint: None,
            default: Some(this_component(idx)),
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: true,
            is_nullable: false })
        .collect();
    let copy_call = Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(class_name)),
        args: components
            .iter()
            .map(|comp| Argument::positional(Expression::ident(comp)))
            .collect() });
    out.instance_methods
        .push(method("copy", copy_params, ret(copy_call)));

    // `toString()` → `Name(a=1, b=2)`, each part through the value renderer so
    // a nested record prints as its own `toString`.
    let mut text = Expression::string(&format!("{}({}=", class_name, components[0]));
    let concat = |left: Expression, right: Expression| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Concat,
            left: Box::new(left),
            right: Box::new(right) })
    };
    text = concat(text, kt_render(this_field(&components[0])));
    for comp in components.iter().skip(1) {
        text = concat(text, Expression::string(&format!(", {}=", comp)));
        text = concat(text, kt_render(this_field(comp)));
    }
    text = concat(text, Expression::string(")"));
    out.instance_methods
        .push(method("toString", vec![], ret(text)));

    // `equals(other)` — structural over the components.
    let other = "__kt_other";
    let mut eq = Expression::new(ExprKind::IsType {
        expr: Box::new(Expression::ident(other)),
        type_name: class_name.to_string() });
    for comp in &components {
        eq = Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(eq),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(this_field(comp)),
                right: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(other)),
                    field: comp.clone(),
                    null_safe: false })) })) });
    }
    out.instance_methods
        .push(method("equals", vec![no_arg_param(other)], ret(eq)));

    // `hashCode()` — equal values render alike, so they hash alike.
    let hash = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__kt_hash")),
        args: vec![Argument::positional(kt_render(Expression::new(
            ExprKind::This,
        )))],
        optional: false });
    out.instance_methods
        .push(method("hashCode", vec![], ret(hash)));

    // Bind the SLOTS. `==`, Set membership and Map keys resolve the slot, never
    // the spelling — which is what makes the derived members reachable from
    // any language, not just from Kotlin source that spells `equals`.
    for (spelling, slot) in [
        ("toString", SpecialMethodKind::ToString),
        ("equals", SpecialMethodKind::Eq),
        ("hashCode", SpecialMethodKind::Hash),
    ] {
        let (canonical, _) = crate::protocol::canonical_method(spelling);
        out.special_methods.push(SpecialMethod {
            kind: slot,
            canonical_name: canonical,
            source_name: spelling.to_string() });
    }
}

/// Kotlin's `by` delegation, stated once.
///
/// `AfterOwn` — a member the class declares itself overrides the delegated one,
/// which is Kotlin's rule (`override fun` beside `by`). `FirstWins` — a class
/// may delegate several interfaces, and the first one listed answers, since
/// Kotlin requires an explicit override for a genuine clash rather than
/// rejecting the declaration. `OwnParent` — `super` in a delegating class still
/// means its own superclass; the delegate is not in the chain.
/// A plain `: I` — the interface's DEFAULT implementations become the class's.
///
/// `Copy`, not `Chain`: Kotlin resolves a default at compile time and a class
/// implementing two interfaces with the same default must override it, so there
/// is no lookup order to preserve. `AfterOwn` — the class's own `override fun`
/// always wins. `Ambiguous` — two interfaces supplying the same default IS the
/// error Kotlin reports, and silently picking one would compile a program
/// `kotlinc` rejects.
fn kotlin_interface_defaults() -> AugmentationPolicy {
    AugmentationPolicy {
        mode: AugmentationMode::Copy,
        position: AugmentationPosition::AfterOwn,
        conflict: AugmentationConflict::Ambiguous,
        super_target: AugmentationSuper::OwnParent,
        contributes: AugmentationContributes {
            methods: true,
            // An interface holds no state, so it has no fields to give.
            fields: false,
            statics: false,
            constructors: false,
            // An abstract declaration is a REQUIREMENT, not an implementation;
            // copying the bodiless stub in would shadow whatever supplies it.
            abstract_members: false } }
}

fn kotlin_delegation() -> AugmentationPolicy {
    AugmentationPolicy {
        mode: AugmentationMode::Promote,
        position: AugmentationPosition::AfterOwn,
        conflict: AugmentationConflict::FirstWins,
        super_target: AugmentationSuper::OwnParent,
        contributes: AugmentationContributes {
            methods: true,
            // The delegate keeps its own state — copying its fields onto the
            // outer class would give it separate storage and silently
            // desynchronise the two.
            fields: false,
            statics: false,
            constructors: false,
            // The augmenting type is almost always an INTERFACE, whose members
            // are all abstract. Those declarations are the only record of WHAT
            // to forward — the forwarder's body is generated to call the
            // delegate — so excluding them leaves nothing to promote at all.
            abstract_members: true } }
}

/// A copy of `stmt` whose `FunctionDecl` name is `name`.
///
/// Cheap and only ever hit for the handful of members whose source spelling the
/// walker had to annotate; every other member is renamed to itself.
fn rename_method(stmt: &vybe_ast::Statement, name: &str) -> vybe_ast::Statement {
    let mut out = stmt.clone();
    if let StmtKind::FunctionDecl { name: n, .. } = &mut out.kind {
        *n = name.to_string();
    }
    out
}

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    _interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut out = NormalMembers::default();

    for member in members {
        match member {
            ClassMember::Field {
                name: fname,
                type_hint,
                init,
                modifiers: m,
                array_bounds,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: Access::from(m.visibility),
                    readonly: m.is_readonly };
                out.push_field(m.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: src_name,
                    modifiers: m,
                    ..
                } = &stmt.kind
                else {
                    continue;
                };

                let (canonical, special_kind) = crate::protocol::canonical_method(src_name);
                let access = Access::from(m.visibility);
                // The walker marks a member operator by prefixing its name with
                // `"operator "` so `protocol::canonical_method` can tell it from
                // a plain method of the same name. That marker must NOT reach
                // the class machinery: the bound member name comes from
                // `source_name`, so leaving it on stored `plus` as
                // `operator plus` — and `a.plus(b)` / `obj.invoke()`, which
                // Kotlin allows, found nothing. Only the SLOT was published, so
                // `a + b` worked while the named call did not.
                // ONLY strips the marker. Renaming every special method would
                // also move `toString` to `tostring`, `equals` to `eq` and so
                // on — the class machinery binds members by `source_name`, and
                // those names are what Kotlin code calls.
                let renamed;
                let stmt = if src_name.starts_with("operator ") {
                    renamed = rename_method(stmt, &canonical);
                    &renamed
                } else {
                    stmt
                };
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    out.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical.clone(),
                        source_name: src_name.clone() });
                }
                out.push_method(m.is_static, method);
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                initializer_target,
                ..
            } => {
                let normalized = NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => match initializer_target {
                            ConstructorInitializerTarget::Base => BaseCall::Explicit(
                                args.iter()
                                    .map(|e| vybe_ast::Argument::positional(e.clone()))
                                    .collect(),
                            ),
                            ConstructorInitializerTarget::This => BaseCall::This(
                                args.iter()
                                    .map(|e| vybe_ast::Argument::positional(e.clone()))
                                    .collect(),
                            ) },
                        None => {
                            if parents.is_empty() {
                                BaseCall::None
                            } else {
                                BaseCall::Auto
                            }
                        }
                    },
                    named_name: None };
                out.push_constructor(normalized);
            }
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers: m,
                ..
            } => {
                // Kotlin is case-sensitive and a property is reached by member
                // access, which resolves `__get_<Name>` / `__set_<Name>` by the
                // EXACT source spelling — so the canonical name keeps its case.
                let access = Access::from(m.visibility);
                let synthetic_backing_getter = !*is_auto && getter.is_none() && setter.is_some();
                let getter_body = getter.clone().or_else(|| {
                    synthetic_backing_getter
                        .then(|| ret(this_field(&format!("__kt_field_{}", pname))))
                });
                let getter_method = getter_body.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        pname,
                        pname,
                        vec![],
                        None,
                        body.clone(),
                        access,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                let setter_method = setter.as_ref().map(|s: &PropertySetter| {
                    build_normal_method(
                        span.clone(),
                        pname,
                        pname,
                        vec![s.param.clone()],
                        None,
                        s.body.clone(),
                        access,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                out.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: pname.clone(),
                    source_name: pname.clone(),
                    is_static: m.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None } });
            }
            // `class C(d: I) : I by d` — the members of `I` become available on
            // `C`, running on the DELEGATE. That is exactly `Promote`, the mode
            // Go embedding uses: the receiver rebinds to the field named by
            // `via_field`. Declared here as DATA; the shared
            // `class_augmentation` pass applies it once, for every language.
            ClassMember::Augment(decl) => {
                // Two different clauses reach the same AST node. `: I by d` names
                // a delegate FIELD and forwards to it; a bare `: I` names no
                // field and copies the interface's default implementations in.
                // `via_field` is what tells them apart.
                let policy = if decl.via_field.is_some() {
                    kotlin_delegation()
                } else {
                    kotlin_interface_defaults()
                };
                out.push_augment_decl(decl, policy);
            }
            other => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    // `data class` — the derived members are a function of the primary
    // constructor's components, produced HERE in normalization from the kind
    // the walker declared. The walker synthesizes nothing.
    if modifiers.kind == ClassKind::Record {
        derive_record_members(name, &mut out);
    }

    NormalClass {
        implicit_self_fields: true,
        ..Default::default()
    }
    .with_members(out)
}
