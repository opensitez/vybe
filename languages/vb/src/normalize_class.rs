//! VB `ClassDecl` → `NormalClass` walker pass.
//!
//! The VB walker (`walker.rs::parse_class_decl`) has already done the
//! heavy lifting:
//!   - `Sub New` → `ClassMember::Constructor`.
//!   - `Handles ctrl.Event` clauses on methods → `AddHandler`
//!     statements injected into the constructor body.
//!   - `Inherits ClassName` → `parents[0]`; implicit
//!     `MyBase.New()` injected at ctor body start when a parent exists.
//!   - `Implements I1, I2` → `interfaces`.
//!   - `Partial Class` merging is handled at the compiler level
//!     (`merge_partial_classes` runs before `compile_stmt`), so by
//!     the time we see a `ClassDecl`, all parts are fused.
//!   - `Property Foo` with `Get`/`Set` blocks → `ClassMember::Property`;
//!     auto-implemented `Property Foo As T` → `ClassMember::Field`.
//!   - `MustInherit` → `is_abstract`; `NotInheritable` → `is_sealed`.
//!
//! This pass only needs to:
//!   - Canonicalise method names via the `ClassLang::Vb` table
//!     (VB is case-insensitive — canonical names land lowercase:
//!     `ToString` → `tostring`, `GetEnumerator` → `iterator`, etc.).
//!   - Stamp `SpecialMethodKind` for operator overloads / protocol
//!     methods.
//!   - Carry visibility + is_override / is_virtual / is_abstract.
//!   - Detect `InitializeComponent` (if the class defines a method by
//!     that name, the compiler's legacy path auto-invokes it in every
//!     ctor). We populate `auto_init_methods` so the shim preserves
//!     that semantic when the direct `emit_class` path lands.

use vybe_ast::{ClassMember, ClassModifiers, ExprKind, Literal, PropertySetter, Span, StmtKind};
use vybe_ast::class_normalize::{
    NormalMembers, build_normal_method, from_method_stmt, types::* };

pub fn normalize_class(
    span: Span,
    _name: &str,
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
                    init: init.clone().or_else(|| vb_default_field_init(type_hint)),
                    array_bounds: array_bounds.clone(),
                    access: Access::from(m.visibility),
                    readonly: m.is_readonly };
                out.push_field(m.is_static || m.is_shared, field);
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

                // InitializeComponent auto-call detection: VB / C# WinForms
                // convention is that ctors implicitly call this method if
                // defined. Flag it here; the direct-emit path in
                // `emit_class` (Phase 2b.2) will emit the call. Today the
                // legacy `compile_class` handles it via the
                // `auto_init_methods` profile flag, so populating here is
                // redundant but forward-compatible.
                if src_name.eq_ignore_ascii_case("InitializeComponent")
                    && !out
                        .auto_init_methods
                        .iter()
                        .any(|n| n.eq_ignore_ascii_case("InitializeComponent"))
                {
                    out.auto_init_methods.push(src_name.clone());
                }

                let (canonical, name_kind) = crate::protocol::canonical_method(src_name);
                let special_kind = m.protocol_slot.or(name_kind);
                let access = Access::from(m.visibility);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                // `Finalize` is VB's destructor. It used to be left as an
                // ordinary override, which is why VB was the one language with
                // a real destructor concept and no `destructor` on its
                // NormalClass.
                if special_kind == Some(SpecialMethodKind::Destructor) {
                    out.destructor = Some(method);
                    continue;
                }
                if let Some(kind) = special_kind {
                    out.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical,
                        source_name: src_name.clone() });
                }
                // VB treats `Shared` as "static" and `Static` is
                // separately for locals — both compile paths land on
                // `is_static` via the walker.
                out.push_method(m.is_static || m.is_shared, method);
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
                            vybe_ast::ConstructorInitializerTarget::Base => BaseCall::Explicit(
                                args.iter()
                                    .map(|e| vybe_ast::Argument::positional(e.clone()))
                                    .collect(),
                            ),
                            vybe_ast::ConstructorInitializerTarget::This => BaseCall::This(
                                args.iter()
                                    .map(|e| vybe_ast::Argument::positional(e.clone()))
                                    .collect(),
                            ) },
                        // The VB walker injects `MyBase.New()` at the start
                        // of the body when the class DECLARATION itself
                        // has an `Inherits` clause. For partial classes,
                        // the walker runs per-file — the declaration that
                        // owns `Sub New` may not be the one that carries
                        // `Inherits`, so the body arrives here without an
                        // explicit super call even though the merged
                        // class has a parent. Emit `BaseCall::Auto` so
                        // `compile_class` auto-injects when needed.
                        // `body_has_super_call` is consulted downstream,
                        // so this is a no-op when the body already starts
                        // with `MyBase.New(...)`.
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
                let (canonical, _) = crate::protocol::canonical_method(pname);
                let access = Access::from(m.visibility);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        vec![],
                        None,
                        body.clone(),
                        access,
                        false,
                        false,
                        false,
                        m.clone(),
                    )
                });
                let setter_method = setter.as_ref().map(|s: &PropertySetter| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        vec![s.param.clone()],
                        None,
                        s.body.clone(),
                        access,
                        false,
                        false,
                        false,
                        m.clone(),
                    )
                });
                out.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: m.is_static || m.is_shared,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None } });
            }
            // VB.NET has single inheritance plus interfaces and no trait/mixin
            // mechanism, so the walker never produces this.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        is_partial: modifiers.is_partial, // informational; merging already done
        explicit_self_param: false,       // VB: Me is implicit
        implicit_self_fields: true,       // VB: bare field names resolve to Me.field
        // No first-class destructor: VB `Finalize` is a regular override. No
        // event bindings: the walker already turned `Handles` into AddHandler
        // statements. Both stay at their neutral default.
        ..Default::default()
    }
    .with_members(out)
}

fn vb_default_field_init(type_hint: &Option<String>) -> Option<vybe_ast::Expression> {
    let ty = type_hint.as_deref()?.trim();
    if matches!(
        ty.to_ascii_lowercase().as_str(),
        "integer"
            | "int32"
            | "short"
            | "int16"
            | "long"
            | "int64"
            | "byte"
            | "sbyte"
            | "ushort"
            | "uint16"
            | "uinteger"
            | "uint32"
            | "ulong"
            | "uint64"
            | "single"
            | "double"
            | "decimal"
    ) {
        return Some(vybe_ast::Expression::new(ExprKind::Lit(Literal::Int(0))));
    }
    if ty.eq_ignore_ascii_case("Boolean") {
        return Some(vybe_ast::Expression::new(ExprKind::Lit(Literal::Bool(
            false,
        ))));
    }
    if ty.eq_ignore_ascii_case("Char") {
        return Some(vybe_ast::Expression::new(ExprKind::Lit(Literal::Str(
            "\0".into(),
        ))));
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{Modifiers, ProtocolSlot};

    fn dummy_span() -> Span {
        Span::default()
    }

    fn make_method(src_name: &str) -> ClassMember {
        make_method_with_modifiers(src_name, Modifiers::default())
    }

    fn make_method_with_modifiers(src_name: &str, modifiers: Modifiers) -> ClassMember {
        ClassMember::Method(Box::new(vybe_ast::Statement::new(StmtKind::FunctionDecl {
            name: src_name.into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers,
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false })))
    }

    #[test]
    fn tostring_canonicalises_case_insensitive() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("ToString")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn getenumerator_maps_to_iterator() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("GetEnumerator")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "iterator");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Iterator);
    }

    #[test]
    fn initializecomponent_populates_auto_init_methods() {
        let nc = normalize_class(
            dummy_span(),
            "MyForm",
            &["Form".into()],
            &[],
            &[make_method("InitializeComponent")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.auto_init_methods.len(), 1);
        assert!(nc.auto_init_methods[0].eq_ignore_ascii_case("InitializeComponent"));
    }

    #[test]
    fn walker_declared_protocol_slot_wins_over_name_guessing() {
        let mut modifiers = Modifiers::default();
        modifiers.protocol_slot = Some(ProtocolSlot::Add);
        let nc = normalize_class(
            dummy_span(),
            "Vector",
            &[],
            &[],
            &[make_method_with_modifiers("operator", modifiers)],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }
}
