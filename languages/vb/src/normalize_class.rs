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

use vybe_ast::{ClassMember, ClassModifiers, PropertySetter, Span, StmtKind, Visibility};
use vybe_plugin::class_normalize::{
    build_normal_method,
    canonical::{ClassLang, canonicalize_method},
    from_method_stmt,
    types::*,
};

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut raw_extra_members: Vec<ClassMember> = Vec::new();
    let mut instance_fields: Vec<NormalField> = Vec::new();
    let mut static_fields: Vec<NormalField> = Vec::new();
    let mut instance_methods: Vec<NormalMethod> = Vec::new();
    let mut static_methods: Vec<NormalMethod> = Vec::new();
    let mut properties: Vec<NormalProperty> = Vec::new();
    let mut constructors: Vec<NormalConstructor> = Vec::new();
    let mut constructor: Option<NormalConstructor> = None;
    let mut special_methods: Vec<SpecialMethod> = Vec::new();
    let mut auto_init_methods: Vec<String> = Vec::new();

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
                    access: access_from_visibility(m.visibility),
                    readonly: m.is_readonly,
                };
                if m.is_static || m.is_shared {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
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
                    && !auto_init_methods
                        .iter()
                        .any(|n| n.eq_ignore_ascii_case("InitializeComponent"))
                {
                    auto_init_methods.push(src_name.clone());
                }

                let (canonical, special_kind) = match src_name.as_str() {
                    "__add__" | "__sub__" | "__mul__" | "__truediv__" | "__mod__" | "__eq__"
                    | "__lt__" | "__le__" | "__gt__" | "__ge__" => {
                        canonicalize_method(ClassLang::Python, src_name)
                    }
                    "__getitem__" => ("getitem".to_string(), Some(SpecialMethodKind::GetItem)),
                    "__setitem__" => ("setitem".to_string(), Some(SpecialMethodKind::SetItem)),
                    "__call__" => ("call".to_string(), Some(SpecialMethodKind::Call)),
                    "operator+" => ("add".to_string(), Some(SpecialMethodKind::Add)),
                    "operator-" => ("sub".to_string(), Some(SpecialMethodKind::Sub)),
                    "operator*" => ("mul".to_string(), Some(SpecialMethodKind::Mul)),
                    "operator/" => ("div".to_string(), Some(SpecialMethodKind::Div)),
                    "operator\\" => ("div".to_string(), Some(SpecialMethodKind::Div)),
                    "operatorMod" | "operatormod" => {
                        ("mod".to_string(), Some(SpecialMethodKind::Mod))
                    }
                    "operator=" => ("eq".to_string(), Some(SpecialMethodKind::Eq)),
                    "operator<>" => ("eq".to_string(), Some(SpecialMethodKind::Eq)),
                    "operator<" => ("lt".to_string(), Some(SpecialMethodKind::Lt)),
                    "operator<=" => ("le".to_string(), Some(SpecialMethodKind::Le)),
                    "operator>" => ("gt".to_string(), Some(SpecialMethodKind::Gt)),
                    "operator>=" => ("ge".to_string(), Some(SpecialMethodKind::Ge)),
                    _ => canonicalize_method(ClassLang::Vb, src_name),
                };
                let access = access_from_visibility(m.visibility);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical,
                        source_name: src_name.clone(),
                    });
                }
                // VB treats `Shared` as "static" and `Static` is
                // separately for locals — both compile paths land on
                // `is_static` via the walker.
                if m.is_static || m.is_shared {
                    static_methods.push(method);
                } else {
                    instance_methods.push(method);
                }
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
                            ),
                        },
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
                    named_name: None,
                };
                constructor = Some(normalized.clone());
                constructors.push(normalized);
            }
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers: m,
                ..
            } => {
                let (canonical, _) = canonicalize_method(ClassLang::Vb, pname);
                let access = access_from_visibility(m.visibility);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        Vec::new(),
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
                        Vec::new(),
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
                properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: m.is_static || m.is_shared,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        span,
        name: name.to_string(),
        parent: parents.first().cloned(),
        bases: Vec::new(),
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: modifiers.is_partial, // informational; merging already done
        is_value_type: false,
        explicit_self_param: false, // VB: Me is implicit
        implicit_self_fields: true, // VB: bare field names resolve to Me.field
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructors,
        constructor,
        destructor: None, // VB `Finalize` is conventionally handled as a regular override; no first-class destructor here
        auto_init_methods,
        special_methods,
        event_bindings: Vec::new(), // walker already turned `Handles` into AddHandler statements
        raw_extra_members,
    }
}

fn access_from_visibility(v: Visibility) -> Access {
    match v {
        Visibility::Public => Access::Public,
        Visibility::Protected => Access::Protected,
        Visibility::Private => Access::Private,
        Visibility::Internal => Access::Internal,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::Modifiers;

    fn dummy_span() -> Span {
        Span::default()
    }

    fn make_method(src_name: &str) -> ClassMember {
        ClassMember::Method(Box::new(vybe_ast::Statement::new(
            StmtKind::FunctionDecl {
                name: src_name.into(),
                params: vec![],
                return_type: None,
                body: vec![],
                modifiers: Modifiers::default(),
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        )))
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
}
