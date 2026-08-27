//! AST-construction sugar for the Flutter adapter classes.
//!
//! Deliberately small: only the shapes [`super::foundation`] needs, built
//! straight from `vybe_ast`'s own constructors. `languages/dart` has a larger
//! vocabulary of the same kind for its `dart:core` classes; that one is
//! `pub(super)` to its module, and neither can depend on the other (a platform
//! must not depend on a language). Sharing them would mean lifting the
//! vocabulary into `vybe_ast`, which is a shared-crate change and needs the
//! user's approval — so this stays a local, minimal helper set rather than a
//! copy of that file.

use vybe_ast::{
    Argument, BindingPattern, ClassMember, Expression, ExprKind, Literal, Modifiers, Param, PassBy,
    PropertySetter, Span, Statement, StmtKind, VarDeclKind, VarDeclarator, Visibility,
};

pub(super) fn span() -> Span {
    Span::default()
}

pub(super) fn ident(name: &str) -> Expression {
    Expression::ident(name)
}

pub(super) fn bool_lit(value: bool) -> Expression {
    Expression::bool(value)
}

pub(super) fn null_lit() -> Expression {
    Expression::null()
}

/// `this.<name>` — the instance field read. Dart resolves a bare identifier in
/// an instance method to the field, but these bodies are synthesized rather
/// than parsed, so the receiver is written explicitly.
pub(super) fn this_field(name: &str) -> Expression {
    field_of(this_ref(), name)
}

/// The receiver INSIDE A PROPERTY ACCESSOR body.
///
/// ⛔A getter/setter is compiled to a chunk that takes the receiver as its
/// FIRST PARAMETER, bound as a local named `this`, so only the IDENTIFIER form
/// reaches it. A method or constructor body wants [`this_ref`] instead.
/// MEASURED both ways: with the method form in accessors, `notifyListeners()`
/// saw a listener while `hasListeners` on the SAME object answered false, and
/// `ValueNotifier.value` always read null; with the identifier form in method
/// bodies, methods stop reading their own fields (14/20 → 3/20). Two spellings
/// for one concept, split by MEMBER KIND — pick by which body you are in.
pub(super) fn accessor_this() -> Expression {
    Expression::ident("this")
}

/// `this.<name>` as spelled inside a property accessor. See [`accessor_this`].
pub(super) fn accessor_field(name: &str) -> Expression {
    field_of(accessor_this(), name)
}

/// The receiver.
///
/// ⛔MEASURED, and it is NOT one form. `ExprKind::This` is what a METHOD body
/// needs (with it, 14/20 of the ChangeNotifier corpus passes; with
/// `Expression::ident("this")` instead, 3/20 — methods stop reading their own
/// fields). But a PROPERTY ACCESSOR is compiled as a chunk taking the receiver
/// as its first parameter, bound as a local named `this`, and only the
/// IDENTIFIER form reaches that — which is why `notifyListeners()` sees a
/// listener while `hasListeners` on the same object answers false.
///
/// Two receiver spellings for one concept, split by member kind, is exactly
/// the knowledge that must live in ONE shared class-building vocabulary rather
/// than be rediscovered by every producer of synthesized class AST. Left as
/// the method form here because that is the larger, working half.
pub(super) fn this_ref() -> Expression {
    Expression::new(ExprKind::This)
}

pub(super) fn field_of(object: Expression, name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: name.to_string(),
        null_safe: false,
    })
}

pub(super) fn call_member(object: Expression, method: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(field_of(object, method)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

pub(super) fn call_value(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

pub(super) fn binary(op: vybe_ast::BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

pub(super) fn not(value: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: vybe_ast::UnaryOp::Not,
        expr: Box::new(value),
    })
}

/// `<target> = <value>;`
///
/// ⛔The dedicated `StmtKind::Assign`, NOT `StmtKind::Expr(ExprKind::Assign)`.
/// Written as an expression statement the write does not reach the instance's
/// field storage: methods still read and mutate the field in place, so a list
/// field looked fine while every `this.x = …` silently did nothing —
/// `dispose()` never armed its guard and `ValueNotifier.value` always read
/// null. `languages/dart`'s core-class builders use this same node.
pub(super) fn assign(target: Expression, value: Expression) -> Statement {
    Statement::with_span(
        StmtKind::Assign {
            targets: vec![target],
            value,
            by_ref: false,
        },
        span(),
    )
}

pub(super) fn expr_stmt(value: Expression) -> Statement {
    Statement::with_span(StmtKind::Expr(value), span())
}

pub(super) fn ret(value: Expression) -> Statement {
    Statement::with_span(StmtKind::Return(Some(value)), span())
}

pub(super) fn ret_void() -> Statement {
    Statement::with_span(StmtKind::Return(None), span())
}

pub(super) fn local(name: &str, value: Expression) -> Statement {
    Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(name.to_string()),
                type_hint: None,
                init: Some(value),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        },
        span(),
    )
}

pub(super) fn if_stmt(cond: Expression, then_body: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            elifs: Vec::new(),
            else_body: None,
        },
        span(),
    )
}

pub(super) fn for_in(var: &str, iter: Expression, body: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::ForIn {
            var: var.to_string(),
            key: None,
            iter,
            body,
            of: true,
            else_body: None,
            is_async: false,
        },
        span(),
    )
}

/// `throw <value>;`
pub(super) fn throw(value: Expression) -> Statement {
    Statement::with_span(
        StmtKind::Throw {
            expr: Some(value),
            cause: None,
        },
        span(),
    )
}

pub(super) fn empty_list() -> Expression {
    Expression::new(ExprKind::Array(Vec::new()))
}

pub(super) fn param(name: &str, type_hint: Option<&str>) -> Param {
    Param {
        name: name.to_string(),
        type_hint: type_hint.map(|hint| hint.to_string().into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

pub(super) fn field(name: &str, type_hint: &str, init: Expression) -> ClassMember {
    ClassMember::Field {
        name: name.to_string(),
        type_hint: Some(type_hint.to_string()),
        init: Some(init),
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
        storage: None,
    }
}

/// A NAMED constructor — Dart `EdgeInsets.all(v)`.
pub(super) fn named_constructor(
    name: &str,
    params: Vec<Param>,
    body: Vec<Statement>,
) -> ClassMember {
    ClassMember::Constructor {
        name: Some(name.to_string()),
        params,
        body,
        base_args: None,
        initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    }
}

/// `static T get name { … }` — Dart's named constants (`Alignment.center`).
pub(super) fn static_getter(name: &str, type_hint: &str, body: Vec<Statement>) -> ClassMember {
    ClassMember::Property {
        name: name.to_string(),
        type_hint: Some(type_hint.to_string()),
        getter: Some(body),
        setter: None,
        is_auto: false,
        modifiers: Modifiers {
            is_static: true,
            ..Modifiers::default()
        },
    }
}

/// `this.<method>(<args>)`
pub(super) fn this_call(method: &str) -> Expression {
    call_member(this_ref(), method, vec![])
}

/// `<a> && <b>`
pub(super) fn and(a: Expression, b: Expression) -> Expression {
    binary(vybe_ast::BinOp::And, a, b)
}

/// A double literal. Dart distinguishes `0` from `0.0`, and these value types
/// are declared `double` in Flutter.
pub(super) fn float_lit(value: f64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Float(value)))
}

/// `__dart_to_string(<expr>)` — the FINAL spelling of Dart's `"$x"`.
///
/// A spliced body never sees walker normalisation, so an `Interpolation` node
/// would not be lowered here; this is the emitted form directly.
pub(super) fn stringify(expr: Expression) -> Expression {
    call_value(ident("__dart_to_string"), vec![expr])
}

/// A construction of a synthesized class — `Alignment(-1.0, 0.0)`.
pub(super) fn construct(class: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::New {
        class: Box::new(ident(class)),
        args: args.into_iter().map(Argument::positional).collect(),
    })
}

pub(super) fn constructor(params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    }
}

pub(super) fn method(
    name: &str,
    params: Vec<Param>,
    return_type: Option<&str>,
    body: Vec<Statement>,
) -> ClassMember {
    ClassMember::Method(Box::new(Statement::with_span(
        StmtKind::FunctionDecl {
            name: name.to_string(),
            params,
            return_type: return_type.map(str::to_string),
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span(),
    )))
}

/// A Dart `T get name => …`.
///
/// A GETTER, never a method: `normalize_class` maps a property name through
/// the protocol table, so `hasListeners`-style reads publish the right slot and
/// — critically — a `Property` never enters the flat, class-less
/// `defined_class_methods` set that every untyped receiver consults. Declaring
/// these as methods is what took the dart slices to zero when `StringBuffer`
/// did it (see `languages/dart/src/core_classes/string_buffer.rs`).
pub(super) fn getter(name: &str, type_hint: &str, body: Vec<Statement>) -> ClassMember {
    ClassMember::Property {
        name: name.to_string(),
        type_hint: Some(type_hint.to_string()),
        getter: Some(body),
        setter: None,
        is_auto: false,
        modifiers: Modifiers::default(),
    }
}

/// A Dart `T get name` / `set name(T v)` pair on one property.
pub(super) fn property(
    name: &str,
    type_hint: &str,
    getter_body: Vec<Statement>,
    setter_param: &str,
    setter_body: Vec<Statement>,
) -> ClassMember {
    ClassMember::Property {
        name: name.to_string(),
        type_hint: Some(type_hint.to_string()),
        getter: Some(getter_body),
        setter: Some(PropertySetter {
            param: param(setter_param, None),
            body: setter_body,
        }),
        is_auto: false,
        modifiers: Modifiers::default(),
    }
}

pub(super) fn class_decl(name: &str, members: Vec<ClassMember>) -> Statement {
    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.to_string(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: vybe_ast::ClassModifiers::default(),
            decorators: Vec::new(),
        },
        span(),
    )
}

/// `Literal::Str` without going through `Expression::string`'s borrow.
pub(super) fn str_lit(value: &str) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Str(value.to_string())))
}
