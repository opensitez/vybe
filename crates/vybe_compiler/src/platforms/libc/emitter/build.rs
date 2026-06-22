//! Shared AST builder helpers for the libc platform.
//!
//! These are the small constructors used to emit the libc runtime as common
//! AST (the same JS-shaped nodes every front-end produces). They live here so
//! both the C walker and the `platforms/libc/*` adapters build the runtime from
//! one set of helpers instead of each duplicating them.

use crate::ast::{
    Argument, BindingPattern, ExprKind, Expression, Literal, Modifiers, Param, PassBy, Statement,
    StmtKind, VarDeclKind, VarDeclarator,
};

// ── C FILE record slot indices (the array layout behind a FILE* handle) ──────
pub const CFILE_PATH: i64 = 0;
pub const CFILE_CONTENT: i64 = 2;
pub const CFILE_POS: i64 = 3;
pub const CFILE_EOF: i64 = 4;
pub const CFILE_UNGOT: i64 = 5;
pub const CFILE_DIRTY: i64 = 6;
pub const CFILE_SPECIAL: i64 = 8;

pub fn stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}

pub fn expr(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

pub fn ident(name: &str) -> Expression {
    expr(ExprKind::Ident(name.to_string()))
}

pub fn str_lit(value: &str) -> Expression {
    expr(ExprKind::Lit(Literal::Str(value.to_string())))
}

pub fn int_lit(value: i64) -> Expression {
    expr(ExprKind::Lit(Literal::Int(value)))
}

pub fn null_lit() -> Expression {
    expr(ExprKind::Lit(Literal::Null))
}

pub fn member(object: Expression, field: &str) -> Expression {
    expr(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

pub fn index_expr(object: Expression, index: Expression) -> Expression {
    expr(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

pub fn file_slot(file: Expression, idx: i64) -> Expression {
    index_expr(file, int_lit(idx))
}

pub fn call_expr(callee: Expression, args: Vec<Expression>) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

pub fn call_member(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    call_expr(member(object, field), args)
}

pub fn assign_expr(target: Expression, value: Expression) -> Expression {
    expr(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })
}

pub fn var_decl_stmt(name: &str, init: Expression) -> Statement {
    stmt(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    })
}

pub fn if_stmt(
    cond: Expression,
    then_body: Vec<Statement>,
    else_body: Option<Vec<Statement>>,
) -> Statement {
    stmt(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

pub fn function_stmt(name: &str, params: Vec<&str>, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::FunctionDecl {
        name: name.to_string(),
        params: params
            .into_iter()
            .map(|param| Param {
                name: param.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect(),
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}
