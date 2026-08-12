//! The AST-construction vocabulary every core class is written in.
//!
//! Nothing here is Dart-specific or class-specific: these are the shapes a
//! walker produces for source-declared members, spelled directly. Keeping them
//! in one place is what lets a class file read as the class rather than as
//! `Expression::with_span(ExprKind::…)` noise.

use vybe_ast::{
    ClassMember, ExprKind, Expression, Literal, Modifiers, Param, PassBy, PropertySetter, Span,
    Statement, StmtKind,
};

pub(super) fn span() -> Span {
    Span::default()
}

pub(super) fn ident(name: &str) -> Expression {
    Expression::with_span(ExprKind::Ident(name.to_string()), span())
}

pub(super) fn str_lit(value: &str) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Str(value.to_string())), span())
}

/// `<left> <op> <right>`
pub(super) fn binary(op: vybe_ast::BinOp, left: Expression, right: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        },
        span(),
    )
}

/// `<left> + <right>`
pub(super) fn concat(left: Expression, right: Expression) -> Expression {
    binary(vybe_ast::BinOp::Add, left, right)
}

/// `this.<name>`
pub(super) fn this_field(name: &str) -> Expression {
    Expression::with_span(
        ExprKind::Member {
            object: Box::new(Expression::with_span(ExprKind::This, span())),
            field: name.to_string(),
            null_safe: false,
        },
        span(),
    )
}

/// `<object>.<field>`
pub(super) fn field_of(object: Expression, name: &str) -> Expression {
    Expression::with_span(
        ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false,
        },
        span(),
    )
}

/// `<target> = <value>;`
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

/// `this.<name> = <value>;`
pub(super) fn set_this(name: &str, value: Expression) -> Statement {
    assign(this_field(name), value)
}

/// `return <value>;`
pub(super) fn ret(value: Expression) -> Statement {
    Statement::with_span(StmtKind::Return(Some(value)), span())
}

pub(super) fn num_lit(value: f64) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Float(value)), span())
}

pub(super) fn int_lit(value: i64) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Int(value)), span())
}

/// `<cond> ? <then> : <else_>`
pub(super) fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then),
            else_: Box::new(else_),
        },
        span(),
    )
}

/// `'$expr'` — stringification through the shared `to_string` slot, which is
/// what interpolation lowers to. A `.toString()` call would instead re-enter
/// member dispatch on an arbitrary receiver.
pub(super) fn stringify(expr: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(ident("__dart_to_string")),
            args: vec![vybe_ast::Argument::positional(expr)],
            optional: false,
        },
        span(),
    )
}

/// `<object>.<method>(<args>)`
pub(super) fn call_member(object: Expression, method: &str, args: Vec<Expression>) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(object),
                    field: method.to_string(),
                    null_safe: false,
                },
                span(),
            )),
            args: args
                .into_iter()
                .map(vybe_ast::Argument::positional)
                .collect(),
            optional: false,
        },
        span(),
    )
}

/// A local `var <name> = <value>;`
pub(super) fn local(name: &str, value: Expression) -> Statement {
    Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![vybe_ast::VarDeclarator {
                pattern: vybe_ast::BindingPattern::Ident(name.to_string()),
                type_hint: None,
                init: Some(value),
                array_bounds: None,
                with_events: false,
            }],
            kind: vybe_ast::VarDeclKind::Var,
        },
        span(),
    )
}

/// `this.<method>()` — a CALL, not the member read `this_field` gives.
///
/// The difference is not cosmetic: reading a zero-arg method yields the
/// function object, and interpolating that renders `[function authority]`.
pub(super) fn this_call(method: &str) -> Expression {
    call_member(
        Expression::with_span(ExprKind::This, span()),
        method,
        Vec::new(),
    )
}

/// `<expr>;` — an expression evaluated for its effect.
pub(super) fn expr_stmt(value: Expression) -> Statement {
    Statement::with_span(StmtKind::Expr(value), span())
}

/// `if (<cond>) { <then> }`
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

/// `if (<cond>) { <then> } else { <else_> }`
pub(super) fn if_else(
    cond: Expression,
    then_body: Vec<Statement>,
    else_body: Vec<Statement>,
) -> Statement {
    Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            elifs: Vec::new(),
            else_body: Some(else_body),
        },
        span(),
    )
}

/// `for (var <var> in <iter>) { <body> }` — `of: true`, the value-iteration
/// form every language but JS's `for…in` means.
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

/// `<object>[<index>] = <value>;`
pub(super) fn index_set(object: Expression, index: Expression, value: Expression) -> Statement {
    assign(
        Expression::with_span(
            ExprKind::Index {
                object: Box::new(object),
                index: Box::new(index),
                null_safe: false,
            },
            span(),
        ),
        value,
    )
}

/// `[]`
pub(super) fn empty_list() -> Expression {
    Expression::with_span(ExprKind::Array(Vec::new()), span())
}

/// `{}` — an empty Dart map.
///
/// **`ExprKind::Object`, not `ExprKind::Map`.** Verified against `--dump-ast`:
/// the dart walker lowers a `{}` literal to `Object([])`, and that is what the
/// map primitives are built around. Using `Map` here produced a value that
/// indexed correctly (`m['k']` read back) but that `containsKey` answered
/// `false` for — two backings, one of which the `__keys` sidecar
/// ([[project_dict_primitives_need_keys_sidecar]]) never sees. Build the node
/// the walker builds.
pub(super) fn empty_map() -> Expression {
    Expression::with_span(ExprKind::Object(Vec::new()), span())
}

/// `<a> || <b>`
pub(super) fn or(a: Expression, b: Expression) -> Expression {
    binary(vybe_ast::BinOp::Or, a, b)
}

/// `<a> && <b>`
#[allow(dead_code)]
pub(super) fn and(a: Expression, b: Expression) -> Expression {
    binary(vybe_ast::BinOp::And, a, b)
}

pub(super) fn bool_lit(value: bool) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Bool(value)), span())
}

pub(super) fn param(name: &str, type_hint: Option<&str>, default: Option<Expression>) -> Param {
    Param {
        name: name.to_string(),
        type_hint: type_hint.map(|t| t.to_string().into()),
        is_optional: default.is_some(),
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_nullable: false,
    }
}

/// A plain field with an initializer.
pub(super) fn field(name: &str, type_hint: &str, init: Expression) -> ClassMember {
    ClassMember::Field {
        name: name.to_string(),
        type_hint: Some(type_hint.to_string()),
        init: Some(init),
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
    }
}

/// An instance method, as the `FunctionDecl` statement `ClassMember::Method`
/// wraps — the same shape a walker produces for a source-declared method.
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
            body,
            return_type: return_type.map(str::to_string),
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span(),
    )))
}

/// A read-only property — Dart's `int get length => …`.
///
/// A name in `ZERO_ARG_GETTER_NAMES` (see the module doc) does NOT work as a
/// getter — but see the StringBuffer members for why the method form can be
/// worse.
pub(super) fn getter(name: &str, type_hint: &str, value: Expression) -> ClassMember {
    getter_body(
        name,
        type_hint,
        vec![Statement::with_span(StmtKind::Return(Some(value)), span())],
    )
}

/// A read-only property whose body is more than one statement.
pub(super) fn getter_body(name: &str, type_hint: &str, body: Vec<Statement>) -> ClassMember {
    ClassMember::Property {
        name: name.to_string(),
        type_hint: Some(type_hint.to_string()),
        getter: Some(body),
        setter: None::<PropertySetter>,
        is_auto: false,
        modifiers: Modifiers::default(),
    }
}

/// The unnamed constructor.
pub(super) fn constructor(params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
        visibility: vybe_ast::Visibility::Public,
    }
}

/// `class <name> { <members> }` — no parents, no interfaces, no decorators.
pub(super) fn class(name: &str, members: Vec<ClassMember>) -> Statement {
    class_extending(name, &[], members)
}

/// `class <name> extends <parent> { <members> }`.
///
/// The parent is what puts the ancestor into the `__types` chain
/// `compile_class` stamps (`classes.rs:3931` walks the MRO), and that chain is
/// half of what `reflection::emit_is_instance_of` unions with the rtt to answer
/// `catch`/`is`. So a declared parent is a CATCHABILITY statement, not just
/// documentation.
pub(super) fn class_extending(
    name: &str,
    parents: &[&str],
    members: Vec<ClassMember>,
) -> Statement {
    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.to_string(),
            parents: parents.iter().map(|p| p.to_string()).collect(),
            interfaces: Vec::new(),
            members,
            modifiers: vybe_ast::ClassModifiers::default(),
            decorators: vec![],
        },
        span(),
    )
}
