//! C# `ClassDecl` → `NormalClass` walker pass.
//!
//! The C# walker has already done:
//!   - `public ClassName(args) { ... }` (constructor by name match) →
//!     `ClassMember::Constructor`.
//!   - `: base(args)` on the constructor → `base_args`.
//!   - `public int Foo { get; set; }` (auto-property) → `ClassMember::Field`
//!     with backing flag.
//!   - `public int Foo { get { ... } set { ... } }` → `ClassMember::Property`.
//!   - `public static T operator +(...)` → static method with name
//!     "operator+" (or similar).
//!   - `partial class` merging handled at compiler level before
//!     `compile_stmt` sees this.
//!   - `abstract class` / `sealed class` → `is_abstract` / `is_sealed`.
//!   - `override` / `virtual` / `abstract` / `static` keywords → `Modifiers`.
//!
//! This pass applies cross-language canonical naming:
//!   - `ToString` → `tostring`; `GetHashCode` → `hash`; `Equals` → `eq`.
//!   - `GetEnumerator` → `iterator` (for `foreach` interop).
//!   - `operator +` / `operator ==` / `operator <` etc. → arithmetic /
//!     comparison special kinds.
//!   - `CompareTo` → `compare`.
//!   - Populates `auto_init_methods` for `InitializeComponent` (the
//!     WinForms convention — C# ctors implicitly call it).

use crate::ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind, Visibility};
use crate::common::classes::{
    build_normal_method,
    canonical::{canonicalize_method, ClassLang},
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
    let mut constructor: Option<NormalConstructor> = None;
    let mut destructor: Option<NormalMethod> = None;
    let mut special_methods: Vec<SpecialMethod> = Vec::new();
    let mut auto_init_methods: Vec<String> = Vec::new();

    for member in members {
        match member {
            ClassMember::Field { name: fname, init, modifiers: m, .. } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    init: init.clone(),
                    access: access_from_visibility(m.visibility),
                    readonly: m.is_readonly,
                };
                if m.is_static {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { name: src_name, modifiers: m, .. } = &stmt.kind else {
                    continue;
                };

                // Finalizer (`~ClassName`) — walker emits these with
                // the source name starting with `~`. Route to destructor.
                if src_name.starts_with('~') {
                    if let Some(d) = from_method_stmt(
                        span.clone(), stmt, "destructor",
                        access_from_visibility(m.visibility),
                    ) {
                        destructor = Some(d);
                    }
                    continue;
                }

                // InitializeComponent auto-call: WinForms convention.
                if src_name == "InitializeComponent"
                    && !auto_init_methods.iter().any(|n| n == "InitializeComponent")
                {
                    auto_init_methods.push(src_name.clone());
                }

                let (canonical, special_kind) = canonicalize_method(ClassLang::CSharp, src_name);
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
                if m.is_static {
                    static_methods.push(method);
                } else {
                    instance_methods.push(method);
                }
            }
            ClassMember::Constructor { params, body, base_args, .. } => {
                constructor = Some(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter().map(|e| crate::ast::Argument::positional(e.clone())).collect(),
                        ),
                        // C#: if no `: base(...)` clause and there IS a
                        // parent class, C# auto-invokes the parameterless
                        // parent ctor. Mirror with Auto.
                        None => if parents.is_empty() { BaseCall::None } else { BaseCall::Auto },
                    },
                    named_name: None,
                });
            }
            ClassMember::Property { name: pname, getter, setter, is_auto, modifiers: m, .. } => {
                let (canonical, _) = canonicalize_method(ClassLang::CSharp, pname);
                let access = access_from_visibility(m.visibility);
                let getter_method = getter.as_ref().map(|body| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![], None, body.clone(),
                    access, false, false, false, Modifiers::default(),
                ));
                let setter_method = setter.as_ref().map(|s: &PropertySetter| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![s.param.clone()], None, s.body.clone(),
                    access, false, false, false, Modifiers::default(),
                ));
                properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            other @ (ClassMember::Event { .. } | ClassMember::Const { .. } | ClassMember::NestedType(_)) => {
                raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        span,
        name: name.to_string(),
        parent: parents.first().cloned(),
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: modifiers.is_partial,
        explicit_self_param: false,
        implicit_self_fields: true, // C#: bare field names resolve to this.field
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructor,
        destructor,
        auto_init_methods,
        special_methods,
        event_bindings: Vec::new(),
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

    fn dummy_span() -> Span { Span::default() }

    fn make_method(src_name: &str) -> ClassMember {
        ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
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
    fn tostring_maps_to_canonical() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("ToString")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn operator_plus_canonicalises_to_add() {
        let nc = normalize_class(
            dummy_span(), "Vec", &[], &[], &[make_method("operator+")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }

    #[test]
    fn tilde_finalizer_routes_to_destructor() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("~Foo")],
            &ClassModifiers::default(),
        );
        assert!(nc.destructor.is_some());
        assert!(nc.instance_methods.is_empty());
    }

    #[test]
    fn initializecomponent_flagged_for_auto_init() {
        let nc = normalize_class(
            dummy_span(), "MyForm", &["Form".into()], &[],
            &[make_method("InitializeComponent")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.auto_init_methods, vec!["InitializeComponent".to_string()]);
    }
}
