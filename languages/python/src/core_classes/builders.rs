//! The AST-construction vocabulary every core class is written in.
//!
//! Nothing here is Python-specific or class-specific: these are the shapes the
//! walker produces for source-declared members, spelled directly. Keeping them
//! in one place is what lets a class file read as the class rather than as
//! `Expression::with_span(ExprKind::…)` noise. Ported from
//! `languages/dart/src/core_classes/builders.rs`, which is the same vocabulary
//! for the same reason.

use vybe_ast::{
    Argument, BinOp, ClassMember, ExprKind, Expression, Literal, Modifiers, Param, PassBy,
    PropertySetter, Span, Statement, StmtKind,
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

pub(super) fn num(value: f64) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Float(value)), span())
}

pub(super) fn bool_lit(value: bool) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Bool(value)), span())
}

pub(super) fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        },
        span(),
    )
}

pub(super) fn add(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Add, left, right)
}

/// `self.<name>` as a READ.
///
/// ⛔⛔ NOT `ExprKind::Member`. Python's walker desugars every attribute read
/// into `__py_obj_get__(obj, "name")` and every attribute write into an
/// `Index` on the instance — the parallel attribute system recorded in
/// `project_python_attributes_bypass_shared_classes`. A declared class whose
/// fields are written as `Member` lands them in the GC struct's named fields,
/// where python's own read path cannot see them: construction succeeds and
/// EVERY subsequent `a.x` throws. Measured on this module, 11/20 → 1/20.
///
/// So the builders speak the shapes the walker produces. When `classes.rs`
/// learns per-instance storage and python attribute reads become ordinary
/// `Member` nodes, these two helpers are the only places that change.
pub(super) fn this_field(name: &str) -> Expression {
    read_attr(ident("self"), name)
}

/// `<object>.<name>` as a READ — python's desugared form.
pub(super) fn read_attr(object: Expression, name: &str) -> Expression {
    call(ident("__py_obj_get__"), vec![object, str_lit(name)])
}

/// `self.<name>` as an ASSIGNMENT TARGET — the instance-dict subscript.
pub(super) fn this_slot(name: &str) -> Expression {
    index(ident("self"), str_lit(name))
}

/// `<object>.<name>` — a read on someone else's object, same desugaring.
pub(super) fn field_of(object: Expression, name: &str) -> Expression {
    read_attr(object, name)
}

/// `<object>.<name>` as a genuine MEMBER node — for a METHOD call, which does
/// dispatch through the class, unlike a field read.
pub(super) fn member(object: Expression, name: &str) -> Expression {
    Expression::with_span(
        ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false,
        },
        span(),
    )
}

/// `f(*args)` — a call whose single argument is SPREAD.
pub(super) fn call_spread(callee: Expression, arg: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(callee),
            args: vec![Argument { value: arg, name: None, by_ref: false, spread: true }],
            optional: false,
        },
        span(),
    )
}

/// `<Class>(*args)` — construction with a spread.
pub(super) fn new_spread(class_name: &str, arg: Expression) -> Expression {
    Expression::with_span(
        ExprKind::New {
            class: Box::new(ident(class_name)),
            args: vec![Argument { value: arg, name: None, by_ref: false, spread: true }],
        },
        span(),
    )
}

pub(super) fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(callee),
            args: args.into_iter().map(Argument::positional).collect(),
            optional: false,
        },
        span(),
    )
}

/// A call of a global by name — a builtin (`str`, `int`, `len`) or one of the
/// module-level functions declared alongside the classes.
pub(super) fn call_global(name: &str, args: Vec<Expression>) -> Expression {
    call(ident(name), args)
}

/// `<Class>(<args>)` — construction. The walker normalizes Python's
/// call-a-class into `ExprKind::New`, and a core class must produce the same
/// node so `compile_class`'s constructor path runs.
pub(super) fn new(class_name: &str, args: Vec<Expression>) -> Expression {
    Expression::with_span(
        ExprKind::New {
            class: Box::new(ident(class_name)),
            args: args.into_iter().map(Argument::positional).collect(),
        },
        span(),
    )
}

/// `s[start:]` — ⛔ a `Slice` is the INDEX of an `Index` node, not a node with
/// an object of its own.
pub(super) fn slice_from(object: Expression, start: Expression) -> Expression {
    index(
        object,
        Expression::with_span(
            ExprKind::Slice { lower: Some(Box::new(start)), upper: None, step: None },
            span(),
        ),
    )
}

/// `s[start:end]`
pub(super) fn slice_range(object: Expression, start: Expression, end: Expression) -> Expression {
    index(
        object,
        Expression::with_span(
            ExprKind::Slice {
                lower: Some(Box::new(start)),
                upper: Some(Box::new(end)),
                step: None,
            },
            span(),
        ),
    )
}

pub(super) fn index(object: Expression, at: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Index {
            object: Box::new(object),
            index: Box::new(at),
            null_safe: false,
        },
        span(),
    )
}

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

/// `not <expr>`
pub(super) fn unary_not(expr: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Unary { op: vybe_ast::UnaryOp::Not, expr: Box::new(expr) },
        span(),
    )
}

/// `raise StopIteration()` — how an iterator declares exhaustion.
/// `raise <Name>(args…)` — the exception classes are not in
/// `py_defined_classes`, so this is a CALL, never a `New`.
pub(super) fn raise_call(name: &str, args: Vec<Expression>) -> Statement {
    Statement::with_span(
        StmtKind::Throw { expr: Some(call_global(name, args)), cause: None },
        span(),
    )
}

pub(super) fn raise_stop_iteration() -> Statement {
    Statement::with_span(
        StmtKind::Throw {
            expr: Some(call_global("StopIteration", vec![])),
            cause: None,
        },
        span(),
    )
}

pub(super) fn set_this(name: &str, value: Expression) -> Statement {
    Statement::with_span(
        StmtKind::Assign {
            targets: vec![this_slot(name)],
            value,
            by_ref: false,
        },
        span(),
    )
}

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

pub(super) fn ret(value: Expression) -> Statement {
    Statement::with_span(StmtKind::Return(Some(value)), span())
}

pub(super) fn expr_stmt(value: Expression) -> Statement {
    Statement::with_span(StmtKind::Expr(value), span())
}

pub(super) fn if_stmt(cond: Expression, then: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::If {
            cond,
            then_body: then,
            elifs: vec![],
            else_body: None,
        },
        span(),
    )
}

pub(super) fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::While {
            cond,
            body,
            else_body: None,
        },
        span(),
    )
}

pub(super) fn param(name: &str, default: Option<Expression>) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        is_optional: default.is_some(),
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_nullable: false,
    }
}

/// `*args`.
pub(super) fn rest_param(name: &str) -> Param {
    Param { is_rest: true, ..param(name, None) }
}

/// `**kwargs`.
pub(super) fn kwargs_param(name: &str) -> Param {
    Param { is_kwargs: true, ..param(name, None) }
}

/// The `(*a, **k)` tail every stub in these modules takes.
pub(super) fn any_args() -> Vec<Param> {
    vec![rest_param("a"), kwargs_param("k")]
}

/// A method that accepts anything and answers `value` — the shape most of the
/// logging/traceback surface is: present, callable, inert.
pub(super) fn stub(name: &str, value: Expression) -> ClassMember {
    method(name, any_args(), vec![ret(value)])
}

/// A module-level `name = value`.
pub(super) fn global_assign(name: &str, value: Expression) -> Statement {
    assign(ident(name), value)
}

/// A module-level function that accepts anything and answers `value`.
pub(super) fn stub_fn(name: &str, value: Expression) -> Statement {
    function(name, any_args(), vec![ret(value)])
}

/// `for <var> in <iter>: <body>`
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

/// `try: <body> except <exc>: <handler>`
pub(super) fn try_except(body: Vec<Statement>, exc: &str, handler: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::Try {
            body,
            catches: vec![vybe_ast::CatchClause {
                types: vec![exc.to_string()],
                var_name: None,
                stack_var: None,
                body: handler,
                when_clause: None,
            }],
            else_body: None,
            finally: None,
        },
        span(),
    )
}

/// `x is None`.
///
/// ⛔ NOT `x == None`. `==` routes through `__py_value_eq`, which TRAPS on a
/// `None` operand — `Thread(None, None, "MainThread")` died in its own
/// constructor on `daemon == True`. The walker lowers `is` / `is not` to
/// `__py_is__` / `__py_is_not__`, so a declared class says the same thing.
pub(super) fn is_none(expr: Expression) -> Expression {
    call_global("__py_is__", vec![expr, null()])
}

/// `x is not None`.
pub(super) fn is_not_none(expr: Expression) -> Expression {
    call_global("__py_is_not__", vec![expr, null()])
}

/// `x is True` — what a `daemon=` flag is actually asking.
pub(super) fn is_true(expr: Expression) -> Expression {
    call_global("__py_is__", vec![expr, bool_lit(true)])
}

pub(super) fn null() -> Expression {
    Expression::null()
}

pub(super) fn tuple_of(items: Vec<Expression>) -> Expression {
    Expression::with_span(ExprKind::Tuple(items), span())
}

pub(super) fn list_of(items: Vec<Expression>) -> Expression {
    Expression::with_span(
        ExprKind::Array(
            items
                .into_iter()
                .map(|value| vybe_ast::ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ),
        span(),
    )
}

/// A CLASS-level constant — `ssl.TLSVersion.TLSv1_2`. `is_static` is what puts
/// it on the class rather than the instance.
pub(super) fn static_field(name: &str, init: Expression) -> ClassMember {
    ClassMember::Field {
        name: name.to_string(),
        type_hint: None,
        init: Some(init),
        modifiers: Modifiers { is_static: true, ..Modifiers::default() },
        with_events: false,
        array_bounds: None,
        storage: None,
    }
}

pub(super) fn field(name: &str, init: Expression) -> ClassMember {
    ClassMember::Field {
        name: name.to_string(),
        type_hint: None,
        init: Some(init),
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
        storage: None,
    }
}

/// An instance method. ⛔ Python's receiver is EXPLICIT — `self` must be the
/// first parameter, exactly as the walker produces for source, or
/// `normalize_class` (which sets `explicit_self_param`) binds the wrong slot.
pub(super) fn method(name: &str, params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    let mut all = vec![param("self", None)];
    all.extend(params);
    ClassMember::Method(Box::new(Statement::with_span(
        StmtKind::FunctionDecl {
            name: name.to_string(),
            params: all,
            body,
            return_type: None,
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span(),
    )))
}

/// A read-only property — what `@property` produces. Used for the members that
/// are NOT pure functions of a stored field; anything that is stays a plain
/// field set in the constructor, which costs no accessor at all.
pub(super) fn getter(name: &str, body: Vec<Statement>) -> ClassMember {
    ClassMember::Property {
        name: name.to_string(),
        type_hint: None,
        getter: Some(body),
        setter: None::<PropertySetter>,
        is_auto: false,
        modifiers: Modifiers::default(),
    }
}

/// `def __init__(self, …)`.
///
/// ⛔⛔ A `ClassMember::Constructor`, NOT a method named `__init__`. The walker
/// converts a source `__init__` into this node before `normalize_class` ever
/// sees the class — a dump of a working source class contains no `__init__`
/// method at all. Declaring it as a method instead means it is never run as the
/// constructor: `IPv4Address(7)` builds an EMPTY object, construction appears to
/// succeed, and every field reads `None`. That cost four wrong diagnoses.
///
/// `self` is still the first parameter, because python is
/// `explicit_self_param` and the body says `self`.
pub(super) fn init(params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    let mut all = vec![param("self", None)];
    all.extend(params);
    ClassMember::Constructor {
        name: None,
        params: all,
        body,
        base_args: None,
        initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
        visibility: vybe_ast::Visibility::Public,
    }
}

/// A module-level function, declared alongside the classes.
pub(super) fn function(name: &str, params: Vec<Param>, body: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::FunctionDecl {
            name: name.to_string(),
            params,
            body,
            return_type: None,
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span(),
    )
}

pub(super) fn class(name: &str, members: Vec<ClassMember>) -> Statement {
    class_extending(name, &[], members)
}

/// The parent is what puts the ancestor into the `__types` chain
/// `compile_class` stamps, and that chain is half of what
/// `reflection::emit_is_instance_of` unions with the rtt — so a declared parent
/// is a catchability statement, not documentation.
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
