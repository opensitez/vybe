//! Shared .NET exception hierarchy normalization.
//!
//! Frontends still decide when to inject these declarations, but the hierarchy
//! and constructor field semantics are .NET surface, not C# surface.

use vybe_ast::{
    ArrayElement, ClassMember, ClassModifiers, ConstructorInitializerTarget, ExprKind, Expression,
    Literal, Modifiers, Param, PassBy, Span, Statement, StmtKind, Visibility,
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
    ("OperationCanceledException", "SystemException"),
    ("DivideByZeroException", "ArithmeticException"),
    ("OverflowException", "ArithmeticException"),
    ("ArgumentNullException", "ArgumentException"),
    ("ArgumentOutOfRangeException", "ArgumentException"),
    ("ObjectDisposedException", "InvalidOperationException"),
    ("UriFormatException", "FormatException"),
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
    let assign_expr = |field: &str, value: Expression| {
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
                value,
            },
            span.clone(),
        )
    };

    let canon = vybe_compiler::compiler::errors::canonical_exception_name(name).to_string();
    let assign_name = assign_expr("name", Expression::string(name));
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
    let assign_types = assign_expr("__types", dotnet_exception_types_expr(name, parent));

    let mk_param = |pname: &str| Param {
        name: pname.into(),
        type_hint: Some("string".into()),
        default: Some(Expression::string("")),
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: false,
    };
    let mk_obj_param = |pname: &str| Param {
        name: pname.into(),
        type_hint: None,
        default: Some(Expression::null()),
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: true,
    };

    let (params, body) = if needs_param_name {
        match name {
            "ArgumentException" => (
                vec![
                    mk_param("msg"),
                    mk_param("paramName"),
                    mk_obj_param("inner"),
                ],
                vec![
                    assign("Message", "msg"),
                    assign("ParamName", "paramName"),
                    assign("InnerException", "inner"),
                    assign_name.clone(),
                    assign_extype.clone(),
                    assign_types.clone(),
                ],
            ),
            "ArgumentNullException" => (
                vec![mk_param("paramName"), mk_param("msg")],
                vec![
                    assign("ParamName", "paramName"),
                    assign("Message", "msg"),
                    assign_name.clone(),
                    assign_extype.clone(),
                    assign_types.clone(),
                ],
            ),
            "ArgumentOutOfRangeException" => (
                vec![
                    mk_param("paramName"),
                    mk_obj_param("actualValue"),
                    mk_param("msg"),
                ],
                vec![
                    assign("ParamName", "paramName"),
                    assign("ActualValue", "actualValue"),
                    assign("Message", "msg"),
                    assign_name.clone(),
                    assign_extype.clone(),
                    assign_types.clone(),
                ],
            ),
            _ => unreachable!(),
        }
    } else {
        let params = if name == "ObjectDisposedException" {
            vec![
                mk_param("objectName"),
                mk_param("msg"),
                mk_obj_param("inner"),
            ]
        } else {
            vec![mk_param("msg"), mk_obj_param("inner")]
        };
        let mut body = Vec::new();
        if name == "ObjectDisposedException" {
            body.push(assign("ObjectName", "objectName"));
        }
        body.extend([
            assign("Message", "msg"),
            assign("InnerException", "inner"),
            assign_name.clone(),
            assign_extype.clone(),
            assign_types.clone(),
        ]);
        (params, body)
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
        members.push(mk_field("ParamName"));
    }
    if name == "ArgumentOutOfRangeException" {
        members.push(mk_field("ActualValue"));
    }
    if name == "ObjectDisposedException" {
        members.push(mk_field("ObjectName"));
    }

    members.push(ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    });

    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.into(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: Vec::new(),
        },
        span,
    )
}

fn dotnet_exception_types_expr(name: &str, parent: &str) -> Expression {
    let mut names = Vec::new();
    push_unique(&mut names, name);
    let mut current = parent;
    while !current.is_empty() {
        push_unique(&mut names, current);
        current = EXCEPTION_HIERARCHY
            .iter()
            .find_map(|(child, next)| (*child == current).then_some(*next))
            .unwrap_or("");
    }
    for canon in vybe_compiler::compiler::errors::exception_ancestors(name) {
        push_unique(&mut names, canon);
    }
    let canonical = vybe_compiler::compiler::errors::canonical_exception_name(name);
    push_unique(&mut names, canonical);
    push_unique(&mut names, "Exception");
    let elements = names
        .into_iter()
        .map(|name| ArrayElement {
            key: None,
            value: Expression::string(&name),
            spread: false,
            by_ref: false,
        })
        .collect();
    Expression::new(ExprKind::Array(elements))
}

fn push_unique(names: &mut Vec<String>, name: &str) {
    let name = name.trim();
    if !name.is_empty() && !names.iter().any(|item| item == name) {
        names.push(name.to_string());
    }
}
