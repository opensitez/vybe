use vybe_ast::{
    BinOp, ClassKind, ClassMember, ClassModifiers, ConstructorInitializerTarget, ExprKind,
    Expression, Literal, Modifiers, Param, PassBy, Span, Statement, StmtKind, Visibility,
};

fn param(name: &str, default: Option<Expression>) -> Param {
    let is_optional = default.is_some();
    Param {
        name: name.to_string(),
        type_hint: None,
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional,
        is_nullable: false,
    }
}

fn this_field(field: &str, span: &Span) -> Expression {
    Expression::with_span(
        ExprKind::Member {
            object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
            field: field.to_string(),
            null_safe: false,
        },
        span.clone(),
    )
}

fn exception_field(name: &str, init: Expression) -> ClassMember {
    ClassMember::Field {
        name: name.to_string(),
        type_hint: None,
        init: Some(init),
        modifiers: Modifiers {
            visibility: Visibility::Protected,
            ..Modifiers::default()
        },
        with_events: false,
        array_bounds: None,
        storage: None,
    }
}

fn lsb_class_field(name: &str, span: &Span) -> ClassMember {
    ClassMember::Field {
        name: "__php_lsb_class".to_string(),
        type_hint: Some("string".to_string()),
        init: Some(Expression::with_span(
            ExprKind::Lit(Literal::Str(name.replace('.', "\\"))),
            span.clone(),
        )),
        modifiers: Modifiers {
            is_static: true,
            ..Modifiers::default()
        },
        with_events: false,
        array_bounds: None,
        storage: None,
    }
}

fn assign_this_field(field: &str, value: Expression, span: &Span) -> Statement {
    Statement::with_span(
        StmtKind::Expr(Expression::with_span(
            ExprKind::Assign {
                target: Box::new(this_field(field, span)),
                value: Box::new(value),
            },
            span.clone(),
        )),
        span.clone(),
    )
}

fn return_expr(expr: Expression, span: &Span) -> Statement {
    Statement::with_span(StmtKind::Return(Some(expr)), span.clone())
}

fn exception_method(name: &str, expr: Expression, span: &Span) -> ClassMember {
    ClassMember::Method(Box::new(Statement::with_span(
        StmtKind::FunctionDecl {
            name: name.to_string(),
            params: Vec::new(),
            return_type: None,
            body: vec![return_expr(expr, span)],
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span.clone(),
    )))
}

fn exception_to_string_expr(class_name: &str, span: &Span) -> Expression {
    let class_name = Expression::with_span(
        ExprKind::Lit(Literal::Str(class_name.to_string())),
        span.clone(),
    );
    let sep_and_message = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Concat,
            left: Box::new(Expression::with_span(
                ExprKind::Lit(Literal::Str(": ".to_string())),
                span.clone(),
            )),
            right: Box::new(this_field("message", span)),
        },
        span.clone(),
    );
    Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Concat,
            left: Box::new(class_name),
            right: Box::new(sep_and_message),
        },
        span.clone(),
    )
}

fn empty_array(span: &Span) -> Expression {
    Expression::with_span(ExprKind::Array(Vec::new()), span.clone())
}

fn exception_constructor(span: &Span) -> ClassMember {
    let message_value = Expression::with_span(
        ExprKind::NullCoalesce {
            left: Box::new(Expression::with_span(
                ExprKind::Ident("message".to_string()),
                span.clone(),
            )),
            right: Box::new(Expression::with_span(
                ExprKind::Lit(Literal::Str(String::new())),
                span.clone(),
            )),
        },
        span.clone(),
    );
    let code_value = Expression::with_span(
        ExprKind::NullCoalesce {
            left: Box::new(Expression::with_span(
                ExprKind::Ident("code".to_string()),
                span.clone(),
            )),
            right: Box::new(Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone())),
        },
        span.clone(),
    );
    let previous_value = Expression::with_span(
        ExprKind::NullCoalesce {
            left: Box::new(Expression::with_span(
                ExprKind::Ident("previous".to_string()),
                span.clone(),
            )),
            right: Box::new(Expression::with_span(ExprKind::Lit(Literal::Null), span.clone())),
        },
        span.clone(),
    );
    ClassMember::Constructor {
        name: None,
        params: vec![
            param(
                "message",
                Some(Expression::with_span(
                    ExprKind::Lit(Literal::Str(String::new())),
                    span.clone(),
                )),
            ),
            param(
                "code",
                Some(Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone())),
            ),
            param(
                "previous",
                Some(Expression::with_span(ExprKind::Lit(Literal::Null), span.clone())),
            ),
        ],
        body: vec![
            assign_this_field("message", message_value, span),
            assign_this_field("code", code_value, span),
            assign_this_field("previous", previous_value.clone(), span),
            assign_this_field("cause", previous_value, span),
        ],
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    }
}

fn exception_base_members(class_name: &str, span: &Span) -> Vec<ClassMember> {
    let mut members = vec![
        exception_field(
            "message",
            Expression::with_span(ExprKind::Lit(Literal::Str(String::new())), span.clone()),
        ),
        exception_field(
            "code",
            Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone()),
        ),
        exception_field(
            "previous",
            Expression::with_span(ExprKind::Lit(Literal::Null), span.clone()),
        ),
        exception_field(
            "cause",
            Expression::with_span(ExprKind::Lit(Literal::Null), span.clone()),
        ),
        exception_constructor(span),
        exception_method("getMessage", this_field("message", span), span),
        exception_method("getCode", this_field("code", span), span),
        exception_method("getPrevious", this_field("previous", span), span),
        exception_method(
            "getLine",
            Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone()),
            span,
        ),
        exception_method(
            "getFile",
            Expression::with_span(ExprKind::Lit(Literal::Str(String::new())), span.clone()),
            span,
        ),
        exception_method("getTrace", empty_array(span), span),
        exception_method(
            "getTraceAsString",
            Expression::with_span(
                ExprKind::Lit(Literal::Str("#0 {main}".to_string())),
                span.clone(),
            ),
            span,
        ),
        exception_method("__toString", exception_to_string_expr(class_name, span), span),
    ];
    members.push(lsb_class_field(class_name, span));
    members
}

fn core_exception_decl(
    name: &str,
    parent: Option<&str>,
    interfaces: Vec<String>,
    members: Vec<ClassMember>,
    kind: ClassKind,
    span: &Span,
) -> Statement {
    let parents = parent.into_iter().map(|p| p.to_string()).collect::<Vec<_>>();
    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.to_string(),
            parents,
            interfaces,
            members,
            modifiers: ClassModifiers {
                kind,
                ..ClassModifiers::default()
            },
            decorators: Vec::new(),
        },
        span.clone(),
    )
}

/// SPL/core throwable surfaces as normal PHP class/interface declarations.
///
/// These declarations travel through the same class normalizer, class slots,
/// reflection metadata, and shared throw/catch machinery as user PHP classes.
/// They are not parsed source preludes.
pub(crate) fn declarations() -> Vec<Statement> {
    let span = Span::default();
    let mut stmts = vec![
        core_exception_decl(
            "Throwable",
            None,
            Vec::new(),
            Vec::new(),
            ClassKind::Interface,
            &span,
        ),
        core_exception_decl(
            "Exception",
            None,
            vec!["Throwable".to_string()],
            exception_base_members("Exception", &span),
            ClassKind::Class,
            &span,
        ),
        core_exception_decl(
            "Error",
            None,
            vec!["Throwable".to_string()],
            exception_base_members("Error", &span),
            ClassKind::Class,
            &span,
        ),
    ];
    for (name, parent) in [
        ("ErrorException", "Exception"),
        ("TypeError", "Error"),
        ("ValueError", "Error"),
        ("ArithmeticError", "Error"),
        ("DivisionByZeroError", "ArithmeticError"),
        ("ArgumentCountError", "TypeError"),
        ("CompileError", "Error"),
        ("ParseError", "CompileError"),
        ("AssertionError", "Error"),
        ("UnhandledMatchError", "Error"),
        ("FiberError", "Error"),
        ("RuntimeException", "Exception"),
        ("LogicException", "Exception"),
        ("InvalidArgumentException", "LogicException"),
        ("DomainException", "LogicException"),
        ("LengthException", "LogicException"),
        ("OutOfRangeException", "LogicException"),
        ("BadFunctionCallException", "LogicException"),
        ("BadMethodCallException", "BadFunctionCallException"),
        ("OutOfBoundsException", "RuntimeException"),
        ("RangeException", "RuntimeException"),
        ("OverflowException", "RuntimeException"),
        ("UnderflowException", "RuntimeException"),
        ("UnexpectedValueException", "RuntimeException"),
        ("JsonException", "Exception"),
    ] {
        stmts.push(core_exception_decl(
            name,
            Some(parent),
            Vec::new(),
            exception_base_members(name, &span),
            ClassKind::Class,
            &span,
        ));
    }
    stmts.push(core_exception_decl(
        "__PHP_Incomplete_Class",
        None,
        Vec::new(),
        vec![lsb_class_field("__PHP_Incomplete_Class", &span)],
        ClassKind::Class,
        &span,
    ));
    stmts
}
