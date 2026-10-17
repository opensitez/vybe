// Source-linked pass checks also allow compile verification against the AST
// artifact while the shared Cargo build is occupied.
#[path = "../src/normalize_affine_bytes.rs"]
mod normalize_affine_bytes;
#[path = "../src/normalize_loop_backing.rs"]
mod normalize_loop_backing;
#[path = "../src/normalize_storage_access.rs"]
mod normalize_storage_access;

#[cfg(test)]
pub use vybe_language_c::{frontend, parse, parse_with_peg};

#[cfg(not(test))]
fn main() {
    use vybe_ast::{
        Argument, BinOp, ExprKind, Expression, Param, PassBy, Statement, StmtKind, StorageAccess,
        TypeHint, UnaryOp,
    };
    let param = |name: &str, spelling: &str| {
        let mut hint = TypeHint::descriptive(spelling);
        hint.access = Some(StorageAccess::Ordinary);
        Param {
            name: name.into(),
            type_hint: Some(hint),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }
    };
    let call = |name: &str, args: Vec<Expression>| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args: args.into_iter().map(Argument::positional).collect(),
            optional: false,
        })
    };
    let loop_body = vec![
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Ternary {
            cond: Box::new(Expression::new(ExprKind::IsType {
                expr: Box::new(Expression::ident("d")),
                type_name: "number".into(),
            })),
            then: Box::new(call(
                "__c_ptr_byte_store",
                vec![
                    Expression::ident("d"),
                    call("__c_ptr_i32_load8_u", vec![Expression::ident("s")]),
                ],
            )),
            else_: Box::new(call("managed_store", vec![])),
        }))),
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("d")),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident("d")),
                right: Box::new(Expression::int(3)),
            })),
        }))),
    ];
    let mut function = Statement::new(StmtKind::FunctionDecl {
        name: "arbitrary".into(),
        params: vec![
            param("d", "unsigned char *"),
            param("s", "unsigned char *"),
            param("n", "int"),
        ],
        return_type: None,
        body: vec![Statement::new(StmtKind::DoWhile {
            body: loop_body,
            cond: Expression::new(ExprKind::Unary {
                op: UnaryOp::PostDec,
                expr: Box::new(Expression::ident("n")),
            }),
            until: false,
        })],
        modifiers: Default::default(),
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    });
    let mut body = vec![function.clone()];
    normalize_loop_backing::normalize(&mut body);
    let StmtKind::FunctionDecl {
        body: statements, ..
    } = &body[0].kind
    else {
        unreachable!()
    };
    assert!(matches!(statements[0].kind, StmtKind::If { .. }));
    let before = format!("{body:?}");
    normalize_loop_backing::normalize(&mut body);
    assert_eq!(before, format!("{body:?}"));
    normalize_storage_access::finish(&mut body);
    let StmtKind::FunctionDecl { params, .. } = &body[0].kind else {
        unreachable!()
    };
    assert!(params.iter().all(|param| param.type_hint.is_none()));
    if let StmtKind::FunctionDecl { params, .. } = &mut function.kind {
        params[0].type_hint.as_mut().unwrap().access = Some(StorageAccess::Volatile);
    }
    let mut body = vec![function];
    normalize_loop_backing::normalize(&mut body);
    let StmtKind::FunctionDecl { body, .. } = &body[0].kind else {
        unreachable!()
    };
    assert!(matches!(body[0].kind, StmtKind::DoWhile { .. }));
}
