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

use vybe_ast::class_normalize::{NormalMembers, build_normal_method, from_method_stmt, types::*};
use vybe_ast::{
    ClassMember, ClassModifiers, ConstructorInitializerTarget, Modifiers, PropertySetter, Span,
    StmtKind, Visibility,
};

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut out = NormalMembers::default();

    // Slots this class DECLARES, collected before the member loop so a
    // declaration anywhere in the body outranks a guess anywhere else.
    //
    // C# spells one slot two ways and means two different things by them:
    // `operator ==` defines `a == b`, while `Equals` is the virtual
    // object-equality method — `Equals` is mapped to `Eq` by name so that a
    // type with no `operator ==` still compares sensibly. When BOTH are
    // present they collide on one slot key and the loser is silently
    // overwritten, which made `a == b` run `Equals`. A declaration is
    // evidence; a name is a guess, so the guess yields.
    let declared_slots: Vec<ProtocolSlot> = members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl { modifiers, .. } => modifiers.protocol_slot,
                _ => None,
            },
            _ => None,
        })
        .collect();

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
                    readonly: m.is_readonly,
                    value_type: None,
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

                // InitializeComponent auto-call: WinForms convention.
                if src_name == "InitializeComponent"
                    && !out
                        .auto_init_methods
                        .iter()
                        .any(|n| n == "InitializeComponent")
                {
                    out.auto_init_methods.push(src_name.clone());
                }

                let (canonical, name_kind) = crate::protocol::canonical_method(src_name);
                // A slot the WALKER stated wins over one guessed from the
                // spelling: `operator ==` knows it fills `Eq` from the
                // declaration form, while the name it carries (`op_Equality`)
                // is the CLR ABI spelling and means nothing to the name table.
                let special_kind = m.protocol_slot.or_else(|| {
                    name_kind.filter(|guessed| {
                        !declared_slots.iter().any(|declared| declared == guessed)
                    })
                });
                let access = Access::from(m.visibility);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                // Finalizer (`~ClassName`) — a lifecycle member, not a method.
                // The `~` sigil is declared in the shared canonical table,
                // which is the only spelling that is a PATTERN rather than a
                // fixed name.
                if special_kind == Some(SpecialMethodKind::Destructor) {
                    out.destructor = Some(method);
                    continue;
                }
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
                            ),
                        },
                        // C#: if no `: base(...)` clause and there IS a
                        // parent class, C# auto-invokes the parameterless
                        // parent ctor. Mirror with Auto.
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
                // C# is case-sensitive, and a property is read/written via member
                // access (`obj.Prop`), which the VM's STRUCT_GET/SET resolves to a
                // `__get_<Prop>` / `__set_<Prop>` accessor by the EXACT source name.
                // Methods can lowercase their canonical because their call sites
                // lowercase too, but property accessors are keyed by the raw field
                // name — lowercasing here would bind `__get_prop` while access looks
                // up `__get_Prop`, so the getter never fires. Preserve case (like JS).
                let canonical = pname.clone();
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
            // C# has no augmentation in this sense: `partial` merges
            // declarations of the SAME type, and extension methods do not enter
            // the type at all. Default interface members would qualify; the
            // walker does not produce them yet.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        implicit_self_fields: true, // C#: bare field names resolve to this.field
        ..Default::default()
    }
    .with_members(out)
}

#[cfg(test)]
mod tests {
    use super::*;

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
    fn tostring_maps_to_canonical() {
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
    fn operator_plus_canonicalises_to_add() {
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
    fn tilde_finalizer_routes_to_destructor() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("~Foo")],
            &ClassModifiers::default(),
        );
        assert!(nc.destructor.is_some());
        assert!(nc.instance_methods.is_empty());
    }

    #[test]
    fn initializecomponent_flagged_for_auto_init() {
        let nc = normalize_class(
            dummy_span(),
            "MyForm",
            &["Form".into()],
            &[],
            &[make_method("InitializeComponent")],
            &ClassModifiers::default(),
        );
        assert_eq!(
            nc.auto_init_methods,
            vec!["InitializeComponent".to_string()]
        );
    }
}
