//! `java.lang.Enum` — the JLS §8.9 enum surface, declared ONCE for the JVM.
//!
//! `name()`, `ordinal()`, `toString()`, `values()` and `valueOf(String)` are
//! members of `java.lang.Enum`, the same class in Java, Kotlin, Scala and
//! Groovy. They belong here for the same reason `java.util.List` does: a
//! second JVM frontend must not have to re-derive them.
//!
//! This used to live in `languages/java/src/walker.rs::walk_enum_decl`, which
//! made a JDK class the property of one frontend.
//!
//! ## The lowering
//!
//! JLS §8.9 says an enum IS a class and each constant IS an instance of it, so
//! the output is a plain `ClassDecl` — deliberately NOT `StmtKind::EnumDecl`,
//! whose ordinal tables constant-fold `Mode.ON` to an F64 and destroy the
//! instance identity `==` depends on. Each constant carries `__name` and
//! `__ordinal`, stamped by the constructor.
//!
//! ## Why the constants go in the static-init block
//!
//! JLS §8.9.2: the constants are created by the class's static initializer,
//! *before* any other static initializer in the enum body. Emitting them as
//! static FIELD initializers instead is both less faithful and actively
//! broken: a field initializer runs while the class is still being compiled,
//! and for a NESTED enum the class is renamed to `Owner.Day` while the
//! initializer still says `new Day(…)`, whose leaf alias is not installed
//! until the nested declaration finishes. Measured: every member-declared
//! enum died with `undefined is not callable` before a line of `main` ran,
//! and so did a plain `static Inner A = new Inner(7)` in C# — the shape, not
//! the enum, is what breaks. A static-init block runs after every class is
//! defined, so the self-reference resolves.

use vybe_ast::{
    Argument, ArrayElement, BinOp, ClassMember, ConstructorInitializerTarget, ExprKind, Expression,
    Modifiers, Param, PassBy, Statement, StmtKind, Visibility,
};

/// The name the static initializer is published under. Matches what a Java
/// `static { … }` block lowers to, so an enum that declares one shares the
/// method and the constants simply come first.
pub const STATIC_INIT: &str = "__static_init_block__";

/// `valueOf`'s storage name.
///
/// NOT `valueOf`: that spelling is intercepted by shared compiler paths ahead
/// of user-class static dispatch, so the static would never be reached. The
/// frontend rewrites `EnumType.valueOf(x)` call sites to this name.
pub const VALUE_OF: &str = "__j_enum_value_of";

/// The instance field holding a constant's declared name (`name()`/`toString()`).
pub const NAME_FIELD: &str = "__name";
/// The instance field holding a constant's declaration index (`ordinal()`).
pub const ORDINAL_FIELD: &str = "__ordinal";
/// The declaration-ordered constant names, carried by every constant.
///
/// `EnumSet` is a bit vector over an enum's constants, so every operation
/// needs that list. Reading it off the VALUE is what lets `EnumSet.of(x, y)`
/// be an ordinary tree leaf: the leaf gets only its source arguments, and one
/// of them is a constant that already knows.
pub const NAMES_FIELD: &str = "__java_enum_names";
/// The declaring class's simple name, carried by every constant.
pub const CLASS_FIELD: &str = "__java_class_name";

/// One enum constant as written in source: `PENNY(1)` is `("PENNY", [1])`.
pub struct EnumConstant {
    pub name: String,
    pub ctor_args: Vec<Argument>,
}

/// How the source language spells a constant's name and ordinal.
///
/// The only difference between a Java enum and a Kotlin one, and it is pure
/// surface: `java.lang.Enum` declares `name()`/`ordinal()` as METHODS, and
/// Kotlin re-exposes the same two as PROPERTIES (`Color.RED.name`). Everything
/// else — the class shape, the constants, `values`, `valueOf`, `toString` — is
/// the same JDK class, which is why both frontends install from here.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Accessors {
    /// `x.name()` / `x.ordinal()` — Java.
    Methods,
    /// `x.name` / `x.ordinal` — Kotlin.
    Properties,
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
    let mut modifiers = Modifiers::default();
    modifiers.is_static = is_static;
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

/// `EnumType.MEMBER`, the read every synthesized member uses to reach a constant.
fn constant_read(class_name: &str, member: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(class_name)),
        field: member.to_string(),
        null_safe: false,
    })
}

/// Install the `java.lang.Enum` surface onto a lowered enum class body.
///
/// Every synthesized member yields to a user declaration of the same name —
/// JLS §8.9.3 lets an enum override `toString`, and a body may declare its own
/// `values`/`valueOf`.
pub fn install(
    class_name: &str,
    constants: &[EnumConstant],
    members: &mut Vec<ClassMember>,
    accessors: Accessors,
) {
    install_constructor(class_name, constants, members, accessors);
    let declared = declared_methods(members);
    install_instance_methods(members, &declared, accessors);
    install_statics(class_name, constants, members, &declared);
    install_constants(class_name, constants, members);
}

/// `["RED", "GREEN", …]` in declaration order.
fn names_array(constants: &[EnumConstant]) -> Expression {
    Expression::new(ExprKind::Array(
        constants
            .iter()
            .map(|c| ArrayElement {
                key: None,
                value: Expression::string(&c.name),
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

/// `java.lang.Enum.__vybe_declare("Color", ["RED", …])`.
///
/// A fully-qualified JDK path, resolved by the common resolver through the
/// `jvm.java` tree mount every JVM frontend already declares — which is the
/// point: the metadata is published by the platform's own surface, not by a
/// per-language profile entry.
fn declare_call(class_name: &str, constants: &[EnumConstant]) -> Statement {
    let mut path = Expression::ident("java");
    for segment in ["lang", "Enum", "__vybe_declare"] {
        path = Expression::new(ExprKind::Member {
            object: Box::new(path),
            field: segment.to_string(),
            null_safe: false,
        });
    }
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(path),
        args: vec![
            Argument::positional(Expression::string(class_name)),
            Argument::positional(names_array(constants)),
        ],
        optional: false,
    })))
}

/// Thread `__name`/`__ordinal` through the constructor and stamp them first.
///
/// JLS §8.9.2 gives every enum constructor these two implicit leading
/// parameters; a body that declares its own constructor keeps its parameters
/// after them.
fn install_constructor(
    class_name: &str,
    constants: &[EnumConstant],
    members: &mut Vec<ClassMember>,
    accessors: Accessors,
) {
    // Stamped as LITERALS, deliberately not as a read of the class's own
    // static field: that would make every constant depend on the static
    // running first, and `install_constants` prepends the constructions to the
    // static-init block — one reorder and the field is silently `undefined`.
    let mut stamps = vec![
        assign(this_field(NAME_FIELD), Expression::ident(NAME_FIELD)),
        assign(this_field(ORDINAL_FIELD), Expression::ident(ORDINAL_FIELD)),
        assign(this_field(NAMES_FIELD), names_array(constants)),
        assign(this_field(CLASS_FIELD), Expression::string(class_name)),
    ];
    if accessors == Accessors::Properties {
        stamps.extend(property_stamps());
    }

    let mut found = false;
    for member in members.iter_mut() {
        if let ClassMember::Constructor { params, body, .. } = member {
            found = true;
            params.insert(0, param(ORDINAL_FIELD));
            params.insert(0, param(NAME_FIELD));
            body.splice(0..0, stamps.iter().cloned());
        }
    }
    if !found {
        members.push(ClassMember::Constructor {
            name: None,
            params: vec![param(NAME_FIELD), param(ORDINAL_FIELD)],
            body: stamps,
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: Visibility::Private,
        });
    }
}

/// `this.name = __name; this.ordinal = __ordinal` — the Kotlin spelling, as
/// plain data properties. Java must NOT get these: there they would collide
/// with the `name()` / `ordinal()` methods, which live in the same property
/// namespace on the object.
fn property_stamps() -> Vec<Statement> {
    vec![
        assign(this_field("name"), Expression::ident(NAME_FIELD)),
        assign(this_field("ordinal"), Expression::ident(ORDINAL_FIELD)),
    ]
}

/// `name()`, `ordinal()`, `toString()` — JLS §8.9.3.
fn install_instance_methods(
    members: &mut Vec<ClassMember>,
    declared: &[String],
    accessors: Accessors,
) {
    // `toString` is a method in BOTH languages; `name`/`ordinal` are methods
    // only where the source calls them as such.
    let accessor_methods: &[(&str, &str)] = match accessors {
        Accessors::Methods => &[("name", NAME_FIELD), ("ordinal", ORDINAL_FIELD)],
        Accessors::Properties => &[],
    };
    for (name, field) in accessor_methods
        .iter()
        .copied()
        .chain(std::iter::once(("toString", NAME_FIELD)))
    {
        if declared.iter().any(|d| d == name) {
            continue;
        }
        members.push(method(
            name,
            vec![],
            vec![Statement::new(StmtKind::Return(Some(this_field(field))))],
            false,
        ));
    }
}

/// `values()` and `valueOf(String)` — JLS §8.9.3's implicitly declared statics.
fn install_statics(
    class_name: &str,
    constants: &[EnumConstant],
    members: &mut Vec<ClassMember>,
    declared: &[String],
) {
    if !declared.iter().any(|d| d == "values") {
        let array = Expression::new(ExprKind::Array(
            constants
                .iter()
                .map(|c| ArrayElement {
                    key: None,
                    value: constant_read(class_name, &c.name),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
        members.push(method(
            "values",
            vec![],
            vec![Statement::new(StmtKind::Return(Some(array)))],
            true,
        ));
    }

    if declared.iter().any(|d| d == "valueOf") {
        return;
    }
    let mut body: Vec<Statement> = constants
        .iter()
        .map(|c| {
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(Expression::ident("__s")),
                    right: Box::new(Expression::string(&c.name)),
                }),
                then_body: vec![Statement::new(StmtKind::Return(Some(constant_read(
                    class_name, &c.name,
                ))))],
                elifs: vec![],
                else_body: None,
            })
        })
        .collect();
    // JLS §8.9.3: an unmatched name is an IllegalArgumentException, not null.
    body.push(Statement::new(StmtKind::Throw {
        expr: Some(Expression::new(ExprKind::New {
            class: Box::new(Expression::ident("IllegalArgumentException")),
            args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::string(&format!(
                    "No enum constant {class_name}."
                ))),
                right: Box::new(Expression::ident("__s")),
            }))],
        })),
        cause: None,
    }));
    members.push(method(VALUE_OF, vec![param("__s")], body, true));
    // Published under the SOURCE spelling as well, delegating to the storage
    // name. A frontend whose call sites are not intercepted (Kotlin) then
    // reaches it with no rewrite of its own; one whose are (Java) keeps
    // rewriting to `VALUE_OF` and this alias is simply never consulted.
    members.push(method(
        "valueOf",
        vec![param("__s")],
        vec![Statement::new(StmtKind::Return(Some(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(class_name)),
                    field: VALUE_OF.to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(Expression::ident("__s"))],
                optional: false,
            },
        ))))],
        true,
    ));
}

/// The constants themselves: a static field per constant, assigned in the
/// static initializer.
///
/// The field carries NO initializer — see the module header. The assignments
/// are PREPENDED to any static-init block the body already declared, because
/// JLS §8.9.2 creates the constants before the body's own static initializers
/// run (`static { COUNT = values().length; }` must see them).
fn install_constants(class_name: &str, constants: &[EnumConstant], members: &mut Vec<ClassMember>) {
    let mut init: Vec<Statement> = Vec::with_capacity(constants.len() + 1);
    // `Class.getEnumConstants()` for this class, published before the
    // constants exist so `EnumSet.allOf(Color.class)` — which is handed only
    // the NAME — can reach the list. See `emitter::enum_adapter`.
    init.push(declare_call(class_name, constants));
    for (ordinal, constant) in constants.iter().enumerate() {
        let mut args = vec![
            Argument::positional(Expression::string(&constant.name)),
            Argument::positional(Expression::int(ordinal as i64)),
        ];
        args.extend(constant.ctor_args.iter().cloned());

        let mut modifiers = Modifiers::default();
        modifiers.is_static = true;
        members.push(ClassMember::Field {
            name: constant.name.clone(),
            type_hint: Some(class_name.to_string()),
            init: None,
            modifiers,
            with_events: false,
            array_bounds: None,
        });

        init.push(assign(
            constant_read(class_name, &constant.name),
            Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(class_name)),
                args,
            }),
        ));
    }
    if init.is_empty() {
        return;
    }

    for member in members.iter_mut() {
        if let ClassMember::Method(stmt) = member {
            if let StmtKind::FunctionDecl { name, body, .. } = &mut stmt.kind {
                if name == STATIC_INIT {
                    body.splice(0..0, init);
                    return;
                }
            }
        }
    }
    members.push(method(STATIC_INIT, vec![], init, true));
}

/// Emit the top-level calls that RUN each class's static initializer.
///
/// The constants are created by `__static_init_block__` (see the module
/// header), and a static initializer only runs because something calls it.
/// Java's frontend already had its own injector for user `static { … }`
/// blocks; Kotlin had none, so its enum classes compiled with the block
/// present and never executed — every constant read `undefined`.
///
/// Appended AFTER every declaration in `body`, which is what makes a constant
/// that names its own class (`new Day("MON", 0)`) resolve: by then the class
/// global exists. Recurses into nested types, since a member enum's block is
/// its own.
pub fn inject_static_init_calls(body: &mut Vec<Statement>) {
    let mut classes = Vec::new();
    collect_static_init_classes(body, &mut classes);
    body.extend(classes.into_iter().map(|class_name| {
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&class_name)),
                field: STATIC_INIT.to_string(),
                null_safe: false,
            })),
            args: vec![],
            optional: false,
        })))
    }));
}

fn collect_static_init_classes(body: &[Statement], out: &mut Vec<String>) {
    for stmt in body {
        let StmtKind::ClassDecl { name, members, .. } = &stmt.kind else {
            if let StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } = &stmt.kind
            {
                collect_static_init_classes(stmts, out);
            }
            continue;
        };
        if members.iter().any(|m| {
            matches!(m, ClassMember::Method(method)
                if matches!(&method.kind, StmtKind::FunctionDecl { name, .. } if name == STATIC_INIT))
        }) {
            out.push(name.clone());
        }
        for member in members {
            if let ClassMember::NestedType(nested) = member {
                collect_static_init_classes(std::slice::from_ref(nested), out);
            }
        }
    }
}
