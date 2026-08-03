//! PHP `ClassDecl` → `NormalClass` walker pass.
//!
//! Per-language responsibilities:
//!   - `__construct` → `NormalConstructor`.
//!   - `__destruct` → `destructor`.
//!   - `__toString` → canonical `tostring` (SpecialMethodKind::ToString).
//!   - `__invoke` / `__call` → canonical `call` (SpecialMethodKind::Call).
//!   - `__get` / `__set` / `__unset` → attribute-access interceptors
//!     (SpecialMethodKind::GetAttr/SetAttr/DelAttr).
//!   - Visibility keywords `public` / `protected` / `private` from the
//!     method/field `modifiers.visibility`.
//!   - `readonly` fields carried through.
//!   - Traits (`use Trait;`) are flattened by the walker BEFORE reaching
//!     here — the member list contains the class's own members plus
//!     every trait method copied in (PHP's semantics).
//!   - `abstract class` / `final class` → `is_abstract` / `is_sealed`.
//!   - `class Foo implements I1, I2` → `interfaces`.

use vybe_ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind};
use vybe_ast::class_normalize::{
    NormalMembers, build_normal_method,
    from_method_stmt,
    types::* };

/// PHP `use SomeTrait;` — the language's rules, stated once.
///
/// `Copy`: a trait's members are duplicated into the using class and report as
/// the class's own (`get_class_methods` lists them). `AfterOwn`: the class's own
/// declaration always beats the trait's. `RequireExplicit`: two traits supplying
/// the same name is a FATAL error in PHP unless the class resolves it with
/// `insteadof` — not a silent last-one-wins pick. `OwnParent`: `parent::` inside
/// a trait method means the USING CLASS's parent; the trait is not in the
/// inheritance chain at all, which is why PHP flattens rather than linking.
/// `statics: true` because a trait's static property gives each using class its
/// OWN copy, not a shared one.
const PHP_TRAIT: AugmentationPolicy = AugmentationPolicy {
    mode: AugmentationMode::Copy,
    position: AugmentationPosition::AfterOwn,
    conflict: AugmentationConflict::RequireExplicit,
    super_target: AugmentationSuper::OwnParent,
    contributes: AugmentationContributes {
        methods: true,
        fields: true,
        statics: true,
        constructors: false,
        abstract_members: true } };

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
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
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                // The destructor is a lifecycle member, not a method — routed
                // by KIND, so PHP states `__destruct` once (in the shared
                // canonical table) instead of testing for it here.
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
                out.push_method(m.is_static, method);
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                ..
            } => {
                out.push_constructor(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter()
                                .map(|e| vybe_ast::Argument::positional(e.clone()))
                                .collect(),
                        ),
                        // PHP: `parent::__construct(...)` is optional;
                        // when missing no auto-call is emitted.
                        None => BaseCall::None },
                    named_name: None });
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
                        Modifiers::default(),
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
                        Modifiers::default(),
                    )
                });
                out.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: m.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None } });
            }
            ClassMember::Const {
                name: cname,
                type_hint,
                value,
                ..
            } => {
                // Class constants are stamped on the class object as
                // static fields so PHP `Class::CONST` static access
                // (struct_get on the constructor object) resolves to
                // the value. Mirrors how `self::CONST` works inside
                // class methods.
                out.push_field(
                    true,
                    NormalField {
                        span: span.clone(),
                        name: cname.clone(),
                        type_hint: type_hint.clone(),
                        init: Some(value.clone()),
                        array_bounds: None,
                        access: Access::Public,
                        readonly: true },
                );
                // Keep the raw entry too so the legacy `Class.Const`
                // global path is still emitted for any caller that
                // resolves consts that way.
                out.raw_extra_members.push(member.clone());
            }
            ClassMember::Augment(decl) => out.push_augment_decl(decl, PHP_TRAIT),
            other @ (ClassMember::Event { .. } | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        implicit_self_fields: false, // PHP requires `$this->field`
        ..Default::default()
    }
    .with_members(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::Modifiers;

    fn dummy_span() -> Span {
        Span::default()
    }

    fn make_method(src_name: &str) -> ClassMember {
        ClassMember::Method(Box::new(vybe_ast::Statement::new(StmtKind::FunctionDecl {
            name: src_name.into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false })))
    }

    #[test]
    fn tostring_magic_maps_to_canonical() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("__toString")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn destruct_goes_to_destructor_slot() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("__destruct")],
            &ClassModifiers::default(),
        );
        assert!(nc.destructor.is_some());
        assert!(nc.instance_methods.is_empty());
    }

    /// `__invoke` (the object is callable) and `__call` (a method was not
    /// found) are DIFFERENT roles. They shared the `Call` slot until
    /// 2026-07-28, so a class defining both published one under the other's
    /// slot and the second install evicted the first.
    #[test]
    fn invoke_and_call_map_to_distinct_slots() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("__invoke"), make_method("__call")],
            &ClassModifiers::default(),
        );
        let kinds: Vec<_> = nc.special_methods.iter().map(|s| s.kind).collect();
        assert!(kinds.contains(&SpecialMethodKind::Call));
        assert!(kinds.contains(&SpecialMethodKind::CallMissing));
    }

    #[test]
    fn get_and_set_map_to_getattr_setattr() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("__get"), make_method("__set")],
            &ClassModifiers::default(),
        );
        let kinds: Vec<_> = nc.special_methods.iter().map(|s| s.kind).collect();
        assert!(kinds.contains(&SpecialMethodKind::GetAttr));
        assert!(kinds.contains(&SpecialMethodKind::SetAttr));
    }

    #[test]
    fn interfaces_passed_through() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &["Countable".into(), "IteratorAggregate".into()],
            &[],
            &ClassModifiers::default(),
        );
        assert_eq!(
            nc.interfaces,
            vec!["Countable".to_string(), "IteratorAggregate".to_string()]
        );
    }
}
