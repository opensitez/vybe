//! THE enum lowering — one member model for every language.
//!
//! Promoted from `platforms/jvm/src/lang_enum.rs`, which was the most complete
//! enum lowering in the tree and the only one that got IDENTITY right. See
//! `documentation/enumunificationplan.md`.
//!
//! ## An enum constant is an INSTANCE
//!
//! Not an ordinal. `StmtKind::EnumDecl` used to lower to `Color.Red = 0` plus a
//! reverse map, which meant a member read const-folded to a number — so `==`
//! compared ordinals, `->name` had nothing to read, and a member could not
//! survive crossing frontends. That is exactly why the JVM frontends refused
//! the common node and built their enums privately.
//!
//! Every constant is therefore built through the enum's own constructor, which
//! carries three implicit leading parameters (JLS §8.9.2 gives Java the first
//! two; the third is the declared value that backed enums and `(int)e` need):
//! `__name`, `__ordinal`, `__value`. A body that declares its own constructor
//! keeps its parameters after them, so `RED("r", 1)` still binds `s`/`n`.
//!
//! ## Why the constants go in a static-init block
//!
//! JLS §8.9.2: constants are created by the class's static initializer, before
//! any other. A static FIELD initializer runs while the class is still being
//! compiled, and for a NESTED enum the class is renamed while the initializer
//! still names the leaf — measured to kill every member-declared enum with
//! `undefined is not callable`. The block runs after every class is defined.
//!
//! The JVM frontends call the block via their own `inject_static_init_calls`.
//! The other languages have no such injector, so `compile_shared_enum_decl`
//! emits the call itself, immediately after the class is defined.

use vybe_ast::{
    Argument, ArrayElement, BinOp, ClassMember, ConstructorInitializerTarget, EnumMember, ExprKind,
    Expression, Literal, Modifiers, Param, PassBy, Statement, StmtKind, Visibility,
};

/// The name the static initializer is published under.
pub const STATIC_INIT: &str = "__static_init_block__";
/// A constant's declared name — what `name`/`toString` read.
pub const NAME_FIELD: &str = "__name";
/// A constant's declaration index — what `ordinal`/`index` read.
pub const ORDINAL_FIELD: &str = "__ordinal";
/// A constant's declared value — what `value` and `(int)e` read. Equal to the
/// ordinal unless the source declared otherwise (`Red = 4`, a backed enum's
/// `case Red = 'r'`, a flags enum's `1 << 0`).
pub const VALUE_FIELD: &str = "__value";
/// The declaring type's name — identity across frontends, since a tree type has
/// no rtt of its own.
// ⛔ ONE OWNER. A THIRD constant for `"__type"`, under a third name
// (`TYPE_FIELD` vs `FIELD_TYPE`) — the word order alone defeated the grep
// that found the other two.
pub use crate::primitives::reflection::FIELD_TYPE as TYPE_FIELD;
/// The integer-coercion method. Its NAME is private to this lowering — every
/// language spells the coercion (`(int)e`, `Ord(e)`, `ordinal()`), never the
/// method, so it is reached through the `Int` slot and not by this spelling.
const INT_METHOD: &str = "__enum_to_int";

/// How the source language SPELLS one piece of the implicit enum surface.
///
/// Pure surface. A name can be installed as a method or as data, never both —
/// they share one key on the object — so each spelling has to be declared.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Spelling {
    /// A call: `x.name()`, `Color.values()`.
    Method,
    /// Data: `x.name`, `Color.values`.
    Property,
}

/// The declared spellings of an enum's implicit surface.
///
/// Two axes, because they vary INDEPENDENTLY — Kotlin is the proof: it reads
/// `c.name` as a property (like Dart and PHP) but calls `Color.values()` as a
/// method (like Java). One axis cannot describe it.
#[derive(Clone, Copy)]
pub struct Surface {
    /// `name` / `ordinal` on a constant.
    pub accessors: Spelling,
    /// `values` on the type.
    pub values: Spelling,
}

impl Surface {
    /// Java: `c.name()`, `Color.values()`.
    pub const JAVA: Self = Self {
        accessors: Spelling::Method,
        values: Spelling::Method,
    };
    /// Kotlin: `c.name`, `Color.values()`.
    pub const KOTLIN: Self = Self {
        accessors: Spelling::Property,
        values: Spelling::Method,
    };
    /// Dart, PHP, Python, C#, and every language on the shared path:
    /// `c.name`, `Color.values`.
    pub const PROPERTIES: Self = Self {
        accessors: Spelling::Property,
        values: Spelling::Property,
    };
}

fn param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

fn this_field(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: field.to_string(),
        null_safe: false,
    })
}

fn assign(target: Expression, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })
}

fn method(name: &str, params: Vec<Param>, body: Vec<Statement>, is_static: bool) -> ClassMember {
    method_filling(name, params, body, is_static, None)
}

/// A member that also fills a protocol ROLE.
///
/// Installing the member is not enough on its own: a role is resolved from the
/// declared slot, never from the spelling, and these members are spelled in the
/// lowering's own vocabulary (`toString`) rather than any language's. C# is the
/// proof — its name table maps `ToString`, so the installed `toString` filled
/// nothing and `WriteLine(c)` rendered `[object Color]`.
fn method_filling(
    name: &str,
    params: Vec<Param>,
    body: Vec<Statement>,
    is_static: bool,
    protocol_slot: Option<vybe_ast::ProtocolSlot>,
) -> ClassMember {
    let mut modifiers = Modifiers::default();
    modifiers.is_static = is_static;
    modifiers.protocol_slot = protocol_slot;
    ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name: name.to_string(),
        params,
        return_type: None,
        body,
        modifiers,
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    })))
}

fn declared_methods(members: &[ClassMember]) -> Vec<String> {
    members
        .iter()
        .filter_map(|m| match m {
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl { name, .. } => Some(name.clone()),
                _ => None,
            },
            _ => None,
        })
        .collect()
}

/// `EnumType.MEMBER`.
fn constant_read(class_name: &str, member: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(class_name)),
        field: member.to_string(),
        null_safe: false,
    })
}

/// The value each constant carries, and whether it is a compile-time integer.
///
/// Auto-increment follows the source rule every language shares: an explicit
/// integer resets the counter, the next implicit member continues from it.
/// A non-literal (`1 << 0`, `'r'`) is carried as written but cannot key the
/// compile-time reverse map.
pub fn member_values(members: &[EnumMember]) -> Vec<(Expression, Option<i64>)> {
    let mut next = 0i64;
    let mut out = Vec::with_capacity(members.len());
    for member in members {
        let entry = match &member.value {
            Some(value) => {
                if let ExprKind::Lit(Literal::Int(n)) = &value.kind {
                    next = *n;
                    (value.clone(), Some(*n))
                } else {
                    (value.clone(), None)
                }
            }
            None => (
                Expression::new(ExprKind::Lit(Literal::Int(next))),
                Some(next),
            ),
        };
        out.push(entry);
        next += 1;
    }
    out
}

/// Install the whole enum surface onto a lowered enum class body.
///
/// Every synthesized member yields to a user declaration of the same name — an
/// enum may override `toString`, and a body may declare its own `values`.
pub fn install(
    class_name: &str,
    members_decl: &[EnumMember],
    values: &[(Expression, Option<i64>)],
    members: &mut Vec<ClassMember>,
    surface: Surface,
) {
    install_constructor(class_name, members, surface.accessors);
    let declared = declared_methods(members);
    install_instance_methods(members, &declared, surface.accessors);
    install_statics(class_name, members_decl, members, &declared, surface.values);
    install_constants(class_name, members_decl, values, members);
}

/// Thread `__name`/`__ordinal`/`__value` through the constructor and stamp them
/// first, so a body's own constructor statements see them already set.
fn install_constructor(class_name: &str, members: &mut Vec<ClassMember>, accessors: Spelling) {
    // Stamped as LITERALS/params, never as a read of the class's own statics:
    // the constructions are PREPENDED to the static-init block, so a read would
    // depend on an ordering that one reorder silently breaks.
    let mut stamps = vec![
        assign(this_field(NAME_FIELD), Expression::ident(NAME_FIELD)),
        assign(this_field(ORDINAL_FIELD), Expression::ident(ORDINAL_FIELD)),
        assign(this_field(VALUE_FIELD), Expression::ident(VALUE_FIELD)),
        assign(this_field(TYPE_FIELD), Expression::string(class_name)),
    ];
    if accessors == Spelling::Property {
        // The read-side spellings, as plain data. Java must NOT get these:
        // there they collide with the `name()`/`ordinal()` methods, which live
        // in the same property namespace on the object. `index` is dart's
        // spelling of the ordinal; `value` is php's and C#'s of the value.
        for (property, source) in [
            ("name", NAME_FIELD),
            ("ordinal", ORDINAL_FIELD),
            ("index", ORDINAL_FIELD),
            ("value", VALUE_FIELD),
        ] {
            stamps.push(assign(this_field(property), Expression::ident(source)));
        }
    }

    let mut found = false;
    for member in members.iter_mut() {
        if let ClassMember::Constructor { params, body, .. } = member {
            found = true;
            // Reversed, because each insert(0) pushes the previous one right —
            // the body's own params end up after `__name, __ordinal, __value`.
            for name in [VALUE_FIELD, ORDINAL_FIELD, NAME_FIELD] {
                params.insert(0, param(name));
            }
            body.splice(0..0, stamps.iter().cloned());
        }
    }
    if !found {
        members.push(ClassMember::Constructor {
            name: None,
            params: vec![param(NAME_FIELD), param(ORDINAL_FIELD), param(VALUE_FIELD)],
            body: stamps,
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: Visibility::Private,
        });
    }
}

/// `name()`, `ordinal()`, `toString()`.
fn install_instance_methods(
    members: &mut Vec<ClassMember>,
    declared: &[String],
    accessors: Spelling,
) {
    let accessor_methods: &[(&str, &str)] = match accessors {
        Spelling::Method => &[("name", NAME_FIELD), ("ordinal", ORDINAL_FIELD)],
        Spelling::Property => &[],
    };
    for (name, field) in accessor_methods
        .iter()
        .copied()
        .chain(std::iter::once(("toString", NAME_FIELD)))
    {
        if declared.iter().any(|d| d == name) {
            continue;
        }
        // How an enum constant RENDERS is the `ToString` role, declared here
        // once rather than left to each language's name table to recognise a
        // spelling this lowering chose.
        let slot = (name == "toString").then_some(vybe_ast::ProtocolSlot::ToString);
        members.push(method_filling(
            name,
            vec![],
            vec![Statement::new(StmtKind::Return(Some(this_field(field))))],
            false,
            slot,
        ));
    }
    // The other half of the same table: coercing a constant to an integer —
    // C# `(int)Color.Blue`, Pascal `Ord(c)`, Java `ordinal()` — is the `Int`
    // role, and it reads the declared VALUE. Nameless on purpose: no language
    // spells this method, they all spell the coercion, and the coercion
    // resolves through the slot.
    if !declared.iter().any(|d| d == INT_METHOD) {
        members.push(method_filling(
            INT_METHOD,
            vec![],
            vec![Statement::new(StmtKind::Return(Some(this_field(
                VALUE_FIELD,
            ))))],
            false,
            Some(vybe_ast::ProtocolSlot::Int),
        ));
    }
}

/// `values` — the constants in declaration order — and `valueOf`.
fn install_statics(
    class_name: &str,
    members_decl: &[EnumMember],
    members: &mut Vec<ClassMember>,
    declared: &[String],
    values_spelling: Spelling,
) {
    let declares_values = declared.iter().any(|d| d == "values")
        || members.iter().any(|m| {
            matches!(m,
            ClassMember::Field { name, .. } if name == "values")
        });
    if !declares_values {
        let array = Expression::new(ExprKind::Array(
            members_decl
                .iter()
                .map(|m| ArrayElement {
                    key: None,
                    value: constant_read(class_name, &m.name),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
        match values_spelling {
            Spelling::Method => members.push(method(
                "values",
                vec![],
                vec![Statement::new(StmtKind::Return(Some(array)))],
                true,
            )),
            // As DATA it cannot be a field initializer: the array names every
            // constant, and the constants do not exist until the static-init
            // block has run. So the field is empty and the block fills it —
            // appended here, and `install_constants` PREPENDS the constructions,
            // which is what puts the two in the only order that works.
            Spelling::Property => {
                let mut modifiers = Modifiers::default();
                modifiers.is_static = true;
                members.push(ClassMember::Field {
                    name: "values".to_string(),
                    type_hint: None,
                    init: None,
                    modifiers,
                    with_events: false,
                    array_bounds: None,
                    storage: None,
                });
                append_to_static_init(
                    members,
                    vec![assign(constant_read(class_name, "values"), array)],
                );
            }
        }
    }

    if declared.iter().any(|d| d == "valueOf") {
        return;
    }
    let mut body: Vec<Statement> = members_decl
        .iter()
        .map(|m| {
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(Expression::ident("__s")),
                    right: Box::new(Expression::string(&m.name)),
                }),
                then_body: vec![Statement::new(StmtKind::Return(Some(constant_read(
                    class_name, &m.name,
                ))))],
                elifs: vec![],
                else_body: None,
            })
        })
        .collect();
    // An unmatched name is null here, not a throw: the JVM frontends keep their
    // own `valueOf` (JLS §8.9.3 requires IllegalArgumentException), and the
    // languages reaching THIS one spell the failure as a null/None result
    // (php `tryFrom`, C# `Enum.TryParse`).
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Lit(Literal::Null),
    )))));
    members.push(method("valueOf", vec![param("__s")], body, true));
}

/// The constants: a static field per constant, assigned in the static
/// initializer. The field carries NO initializer — see the module header.
fn install_constants(
    class_name: &str,
    members_decl: &[EnumMember],
    values: &[(Expression, Option<i64>)],
    members: &mut Vec<ClassMember>,
) {
    let mut init: Vec<Statement> = Vec::with_capacity(members_decl.len());
    for (ordinal, member) in members_decl.iter().enumerate() {
        let value_expr = values
            .get(ordinal)
            .map(|(expr, _)| expr.clone())
            .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(ordinal as i64))));
        let mut args = vec![
            Argument::positional(Expression::string(&member.name)),
            Argument::positional(Expression::int(ordinal as i64)),
            Argument::positional(value_expr),
        ];
        args.extend(
            member
                .constructor_args
                .iter()
                .cloned()
                .map(Argument::positional),
        );

        let mut modifiers = Modifiers::default();
        modifiers.is_static = true;
        members.push(ClassMember::Field {
            name: member.name.clone(),
            type_hint: Some(class_name.to_string()),
            init: None,
            modifiers,
            with_events: false,
            array_bounds: None,
            storage: None,
        });

        init.push(assign(
            constant_read(class_name, &member.name),
            Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(class_name)),
                args,
            }),
        ));
    }
    if init.is_empty() {
        return;
    }
    splice_static_init(members, init, 0);
}

/// Run `statements` LAST in the static initializer — after the constants exist.
fn append_to_static_init(members: &mut Vec<ClassMember>, statements: Vec<Statement>) {
    let at = static_init_body(members).map_or(0, |body| body.len());
    splice_static_init(members, statements, at);
}

/// Splice into the static initializer at `at`, creating the block if absent.
///
/// A body may already declare one (a `static { … }` block in the source), and
/// the constants must precede everything in it — hence `at`, rather than always
/// appending.
fn splice_static_init(members: &mut Vec<ClassMember>, statements: Vec<Statement>, at: usize) {
    if let Some(body) = static_init_body(members) {
        let at = at.min(body.len());
        body.splice(at..at, statements);
        return;
    }
    members.push(method(STATIC_INIT, vec![], statements, true));
}

fn static_init_body(members: &mut [ClassMember]) -> Option<&mut Vec<Statement>> {
    members.iter_mut().find_map(|member| match member {
        ClassMember::Method(stmt) => match &mut stmt.kind {
            StmtKind::FunctionDecl { name, body, .. } if name == STATIC_INIT => Some(body),
            _ => None,
        },
        _ => None,
    })
}
