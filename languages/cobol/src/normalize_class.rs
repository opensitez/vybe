//! COBOL `ClassDecl` -> `NormalClass` shim.
//!
//! COBOL classes walk directly into the common AST as fields and methods.
//! This shim preserves that shape so the shared class compiler can consume
//! OO COBOL without a dedicated semantic lowering pass.
//!
//! ⛔ THIS FILE WAS DEAD CODE UNTIL 2026-08-29. It is registered on
//! `LanguageDef::normalize_class` and looks live, but `normalize_class_from_ast`
//! only calls the registry when `profile.uses_normalize_class` is set — and the
//! COBOL profile had no such row. Every OO COBOL class went through
//! `normalize_from_ast_legacy` instead, which answers three COBOL questions
//! wrongly (`explicit_self_param: true`, `implicit_self_fields: false`, statics
//! routed on `is_shared` where the COBOL walker sets `is_static`) and has no
//! constructor concept at all. Symptom: a bare `WS-X` inside a method compiled
//! to `global.get`, so instance state was written by the constructor and never
//! read back. `tests/cobol/oo_class_state/` is 0/4 with that row removed.
//!
//! ⇒ A normalizer being WRITTEN is not the same as a normalizer being CALLED.
//! The registration is not the gate; the profile row is.

use vybe_ast::class_normalize::{
    Access, BaseCall, NormalClass, NormalConstructor, NormalField, NormalMembers, from_method_stmt,
};
use vybe_ast::{ClassMember, ClassModifiers, Span, StmtKind};

pub fn normalize_class(
    span: Span,
    _name: &str,
    parents: &[String],
    _interfaces: &[String],
    members: &[ClassMember],
    _modifiers: &ClassModifiers,
) -> NormalClass {
    let mut m = NormalMembers::default();

    for member in members {
        match member {
            ClassMember::Field {
                name: field_name,
                type_hint,
                init,
                modifiers: field_modifiers,
                array_bounds,
                storage,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: field_name.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: Access::Public,
                    readonly: field_modifiers.is_readonly,
                    value_type: None,
                    // CARRIED, never re-defaulted — the same rule
                    // `normalize_from_ast_legacy` states for itself. A COBOL
                    // class field IS its `PIC`: the width, the decimal places
                    // and the sign are the declaration, so `None` here is a
                    // frontend throwing away what it just parsed.
                    //
                    // ⚠ MEASURED 2026-08-29: carrying it changes nothing YET.
                    // `MOVE "ZZZZZ" TO WS-S` truncates to `ZZZ` for a PIC X(3)
                    // at PROGRAM scope and does NOT truncate for the identical
                    // PIC on a class field — with `*storage` and with `None`
                    // alike. The class field path does not consult
                    // `NormalField.storage`. That is a gap in the consumer, not
                    // a reason to stop declaring: the value is now available to
                    // whoever closes it.
                    storage: *storage,
                };
                m.push_field(field_modifiers.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: source_name,
                    modifiers: method_modifiers,
                    params,
                    body,
                    ..
                } = &stmt.kind
                else {
                    continue;
                };

                let is_constructor = source_name.eq_ignore_ascii_case("new");
                if is_constructor {
                    m.push_constructor(NormalConstructor {
                        span: span.clone(),
                        params: params.clone(),
                        body: body.clone(),
                        base_call: if parents.is_empty() {
                            BaseCall::None
                        } else {
                            BaseCall::Auto
                        },
                        named_name: None,
                    });
                    continue;
                }

                let canonical_name = source_name.to_ascii_lowercase();
                if let Some(method) =
                    from_method_stmt(span.clone(), stmt, &canonical_name, Access::Public)
                {
                    m.push_method(method_modifiers.is_static, method);
                }
            }
            // COBOL has no trait/mixin/module mechanism — `INHERITS` is plain
            // single inheritance — so the walker never produces this.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Constructor { .. }
            | ClassMember::Property { .. }
            | ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                m.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        implicit_self_fields: true,
        ..Default::default()
    }
    .with_members(m)
}
