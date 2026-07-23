//! Shared .NET exception hierarchy normalization.
//!
//! Frontends still decide when to inject these declarations, but the hierarchy
//! and constructor field semantics are .NET surface, not C# surface.

use vybe_ast::{
    ClassMember, ClassModifiers, ConstructorInitializerTarget, ExprKind, Expression, Literal,
    Modifiers, Param, PassBy, Span, Statement, StmtKind, Visibility,
};

pub const EXCEPTION_HIERARCHY: &[(&str, &str)] = &[
    ("Exception", ""),
    ("SystemException", "Exception"),
    ("ApplicationException", "Exception"),
    ("AggregateException", "Exception"),
    ("ArithmeticException", "SystemException"),
    ("ArgumentException", "SystemException"),
    ("InvalidOperationException", "SystemException"),
    ("FormatException", "SystemException"),
    ("InvalidCastException", "SystemException"),
    ("NotImplementedException", "SystemException"),
    ("NotSupportedException", "SystemException"),
    ("NullReferenceException", "SystemException"),
    ("IndexOutOfRangeException", "SystemException"),
    ("KeyNotFoundException", "SystemException"),
    ("TimeoutException", "SystemException"),
    ("DivideByZeroException", "ArithmeticException"),
    ("OverflowException", "ArithmeticException"),
    ("ArgumentNullException", "ArgumentException"),
    ("ArgumentOutOfRangeException", "ArgumentException"),
    ("ObjectDisposedException", "InvalidOperationException"),
];

pub fn synthesize_exception_classes() -> Vec<Statement> {
    EXCEPTION_HIERARCHY
        .iter()
        .map(|(name, parent)| synthesize_exception_class(name, parent))
        .collect()
}

fn synthesize_exception_class(name: &str, parent: &str) -> Statement {
    let span = Span::default();
    let needs_param_name = matches!(
        name,
        "ArgumentNullException" | "ArgumentOutOfRangeException" | "ArgumentException"
    );

    let assign = |field: &str, ident: &str| {
        Statement::with_span(
            StmtKind::Assign {
                targets: vec![Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
                        field: field.into(),
                        null_safe: false,
                    },
                    span.clone(),
                )],
                value: Expression::with_span(ExprKind::Ident(ident.into()), span.clone()),
            },
            span.clone(),
        )
    };

    let canon = vybe_emitter::errors::canonical_exception_name(name).to_string();
    let assign_extype = Statement::with_span(
        StmtKind::Assign {
            targets: vec![Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
                    field: "__exception_type".into(),
                    null_safe: false,
                },
                span.clone(),
            )],
            value: Expression::with_span(ExprKind::Lit(Literal::Str(canon)), span.clone()),
        },
        span.clone(),
    );

    let mk_param = |pname: &str| Param {
        name: pname.into(),
        type_hint: Some("string".into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };

    let (params, body) = if needs_param_name {
        match name {
            "ArgumentException" => (
                vec![mk_param("msg"), mk_param("paramName")],
                vec![
                    assign("Message", "msg"),
                    assign("ParamName", "paramName"),
                    assign_extype.clone(),
                ],
            ),
            "ArgumentNullException" => (
                vec![mk_param("paramName")],
                vec![assign("ParamName", "paramName"), assign_extype.clone()],
            ),
            "ArgumentOutOfRangeException" => (
                vec![mk_param("paramName"), mk_param("msg")],
                vec![
                    assign("ParamName", "paramName"),
                    assign("Message", "msg"),
                    assign_extype.clone(),
                ],
            ),
            _ => unreachable!(),
        }
    } else {
        (
            vec![mk_param("msg")],
            vec![assign("Message", "msg"), assign_extype.clone()],
        )
    };

    let mk_field = |fname: &str| ClassMember::Field {
        name: fname.into(),
        type_hint: None,
        init: None,
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
    };

    let mut members = vec![mk_field("Message"), mk_field("InnerException")];
    if needs_param_name {
        members.push(ClassMember::Field {
            name: "ParamName".into(),
            type_hint: Some("string".into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        });
    }

    members.push(ClassMember::Constructor {
        name: None,
        params: Vec::new(),
        body: vec![assign_extype.clone()],
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    });
    members.push(ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    });

    if name == "ArgumentException" {
        members.push(ClassMember::Constructor {
            name: None,
            params: vec![mk_param("msg")],
            body: vec![assign("Message", "msg"), assign_extype.clone()],
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
    }

    if !needs_param_name {
        members.push(ClassMember::Constructor {
            name: None,
            params: vec![mk_param("msg"), mk_param("inner")],
            body: vec![
                assign("Message", "msg"),
                assign("InnerException", "inner"),
                assign_extype,
            ],
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
    }

    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.into(),
            parents: if parent.is_empty() {
                Vec::new()
            } else {
                vec![parent.into()]
            },
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: Vec::new(),
        },
        span,
    )
}
