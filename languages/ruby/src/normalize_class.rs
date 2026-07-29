//! Ruby `ClassDecl` → `NormalClass` walker pass.
//!
//! The Ruby walker (`walker.rs::walk_class_def` + `walk_class_body`)
//! has already turned `def initialize` into `ClassMember::Constructor`,
//! `attr_accessor/reader/writer` into fields/properties, and
//! normal `def` into `ClassMember::Method` with visibility from the
//! current `private`/`protected`/`public` section.
//!
//! This pass normalises cross-language naming and identity:
//!   - `to_s` → canonical `tostring` (+ `SpecialMethodKind::ToString`).
//!   - `inspect` → `repr`.
//!   - `<=>` → `compare`; `==` → `eq`; `<` / `<=` / `>` / `>=` mapped
//!     to `lt`/`le`/`gt`/`ge`.
//!   - Binary operators `+` / `-` / `*` / `/` / `%` / `**` → arithmetic
//!     special kinds.
//!   - Unary `-@` → `neg`.
//!   - `[]` / `[]=` → `getitem` / `setitem`.
//!   - `each` → `iterator`.
//!   - `size` / `length` → `len`.
//!   - `hash` → `hash`.
//!   - `include?` → `contains`.
//!   - `call` → callable protocol.

use vybe_ast::{
    ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind, Visibility,
};
use vybe_ast::class_normalize::{
    NormalMembers, build_normal_method,
    from_method_stmt,
    types::*,
};

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
                    readonly: false,
                };
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
                if let Some(kind) = special_kind {
                    out.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical,
                        source_name: src_name.clone(),
                    });
                }
                out.push_method(m.is_static, method);
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                visibility,
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
                        // Ruby: `super` without args passes through the
                        // current method's args; bare `super()` with
                        // empty parens means explicit-no-args. Walker
                        // today doesn't distinguish — treat missing as
                        // `None`; `emit_class` won't auto-inject for Ruby.
                        None => BaseCall::None,
                    },
                    named_name: None,
                });
                // Constructor visibility suppressed here — Ruby's
                // `initialize` is always effectively private, user
                // calls go through `new` which the stdlib handles.
                let _ = visibility;
            }
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers: m,
                ..
            } => {
                // Ruby attr_accessor/reader/writer produces these via
                // the walker's `walk_attr_decl`.
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
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            // Ruby `include` / `prepend` / `extend` ARE augmentations, and they
            // are the first real `Chain` user — a module is inserted into the
            // ancestor chain, not copied. Migrating them is flexclassplan.md
            // §4c-R step R6, which needs `Chain` (R5) first; until then the
            // walker keeps folding and never produces this.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        implicit_self_fields: true, // Ruby: bare ivars resolve to @self
        ..Default::default() // Ruby has no destructor; GC-finalised
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
            is_sub: false,
        })))
    }

    #[test]
    fn to_s_maps_to_canonical_tostring() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("to_s")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn spaceship_operator_maps_to_compare() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("<=>")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "compare");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Compare);
    }

    #[test]
    fn each_maps_to_iterator() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("each")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "iterator");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Iterator);
    }

    #[test]
    fn index_operators_map_to_getitem_setitem() {
        let get_nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("[]")],
            &ClassModifiers::default(),
        );
        assert_eq!(get_nc.instance_methods[0].canonical_name, "getitem");
        assert_eq!(get_nc.special_methods[0].kind, SpecialMethodKind::GetItem);

        let set_nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("[]=")],
            &ClassModifiers::default(),
        );
        assert_eq!(set_nc.instance_methods[0].canonical_name, "setitem");
        assert_eq!(set_nc.special_methods[0].kind, SpecialMethodKind::SetItem);
    }

    #[test]
    fn size_and_length_both_map_to_len() {
        let nc_size = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("size")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc_size.instance_methods[0].canonical_name, "len");
        assert_eq!(nc_size.special_methods[0].kind, SpecialMethodKind::Len);

        let nc_length = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("length")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc_length.instance_methods[0].canonical_name, "len");
    }
}
