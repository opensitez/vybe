//! Dart `ClassDecl` → `NormalClass` walker pass.
//!
//! Dart specifics:
//!   - Primary constructor: `ClassName(args) { ... }` or short-form
//!     `ClassName(this.x, this.y);`. Walker produces
//!     `ClassMember::Constructor`.
//!   - Named constructors: `ClassName.named(args)` — walker should
//!     wrap these as methods tagged "constructor.<name>" or similar;
//!     we don't yet distinguish them here (future improvement).
//!   - Factory constructors: `factory ClassName(args) => ...` — static
//!     method that returns an instance; walker represents as a static
//!     method.
//!   - `operator +` / `operator ==` / `operator []` / `operator []=`
//!     → SpecialMethodKind::{Add, Eq, GetItem, SetItem}.
//!   - `get foo => expr` / `set foo(val) { ... }` → NormalProperty.
//!   - `extends A` → parent; `implements I1, I2` → interfaces;
//!     `with M1, M2` (mixins) — walker flattens mixin methods into
//!     member list before we see it.
//!   - `abstract class` → `is_abstract`.
//!   - Dart doesn't have destructors (uses Finalizer from dart:ffi).

use vybe_ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind};
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
    let mut constructor: Option<NormalConstructor> = None;
    let mut constructors: Vec<NormalConstructor> = Vec::new();
    let mut special_methods: Vec<SpecialMethod> = Vec::new();

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
                    access: Access::Public, // Dart's `_name` convention isn't enforced
                    readonly: m.is_readonly,
                };
                if m.is_static {
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
                let (canonical, special_kind) = canonicalize_method(ClassLang::Dart, src_name);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, Access::Public)
                else {
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
            ClassMember::Constructor {
                name: ctor_name,
                params,
                body,
                base_args,
                ..
            } => {
                let normal = NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter()
                                .map(|e| vybe_ast::Argument::positional(e.clone()))
                                .collect(),
                        ),
                        // A Dart subclass ctor without an explicit
                        // `: super(...)` implicitly calls the parent's no-arg
                        // ctor — which is exactly `BaseCall::Auto`, the same
                        // preamble C# and VB already use. A class with no
                        // parent has nothing to call.
                        None if !parents.is_empty() => BaseCall::Auto,
                        None => BaseCall::None,
                    },
                    named_name: ctor_name.clone(),
                };
                // A class can declare an unnamed ctor AND several named ones
                // (`Point(this.x)` + `Point.origin()`); they are distinct
                // constructors, not overloads, and often share an arity. Every
                // one is a variant — assigning to a single slot would keep
                // only the last one walked.
                if ctor_name.is_none() {
                    constructor = Some(normal.clone());
                }
                constructors.push(normal);
            }
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers: m,
                ..
            } => {
                let (canonical, _) = canonicalize_method(ClassLang::Dart, pname);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        Vec::new(),
                        vec![],
                        None,
                        body.clone(),
                        Access::Public,
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
                        Vec::new(),
                        vec![s.param.clone()],
                        None,
                        s.body.clone(),
                        Access::Public,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: m.is_static,
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
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: false,
        is_value_type: false,
        explicit_self_param: false,
        // Dart resolves bare identifiers in instance methods to `this.field`
        // (the `this.` is only needed when a local shadows). Setting this true
        // lets `String toString() => '($x, $y)'` and `_balance += v` reach
        // the instance fields without explicit `this.` qualification.
        implicit_self_fields: true,
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructors,
        constructor,
        destructor: None,
        auto_init_methods: Vec::new(),
        special_methods,
        event_bindings: Vec::new(),
        raw_extra_members,
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
    fn to_string_canonicalises_to_tostring() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("toString")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn operator_plus_maps_to_add() {
        let nc = normalize_class(
            dummy_span(),
            "Vec",
            &[],
            &[],
            &[make_method("operator+")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }

    #[test]
    fn hash_code_maps_to_hash() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("hashCode")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "hash");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Hash);
    }
}
