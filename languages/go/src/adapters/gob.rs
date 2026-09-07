use vybe_ast::{Argument, BinOp, ExprKind, Expression, PlaceExpr, Statement, StmtKind, UnaryOp};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("encoding.gob.NewEncoder", "go.encoding.gob.NewEncoder"),
        ("encoding.gob.NewDecoder", "go.encoding.gob.NewDecoder"),
        ("encoding.gob.Register", "go.encoding.gob.Register"),
        ("encoding.gob.RegisterName", "go.encoding.gob.RegisterName"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    register_type(root, "encoding.gob.Encoder", &["Encode", "EncodeValue"]);
    register_type(root, "encoding.gob.Decoder", &["Decode", "DecodeValue"]);
    register_type(root, "encoding.gob.GobEncoder", &[]);
    register_type(root, "encoding.gob.GobDecoder", &[]);
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "gob.Encoder" | "encoding/gob.Encoder" | "encoding.gob.Encoder" => Some("__goGobEncoder"),
        "gob.Decoder" | "encoding/gob.Decoder" | "encoding.gob.Decoder" => Some("__goGobDecoder"),
        "gob.GobEncoder"
        | "gob.GobDecoder"
        | "encoding/gob.GobEncoder"
        | "encoding/gob.GobDecoder"
        | "encoding.gob.GobEncoder"
        | "encoding.gob.GobDecoder" => Some("any"),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "gob.NewEncoder"
        | "b.NewEncoder"
        | "encoding.gob.NewEncoder"
        | "go.encoding.gob.NewEncoder" => Some(pointer_arg(typed_object(
            "__goGobEncoder",
            vec![("w", writer_value(arg(args, 0)))],
        ))),
        "gob.NewDecoder"
        | "b.NewDecoder"
        | "encoding.gob.NewDecoder"
        | "go.encoding.gob.NewDecoder" => Some(pointer_arg(typed_object(
            "__goGobDecoder",
            vec![
                ("r", writer_value(arg(args, 0))),
                ("pos", Expression::int(0)),
            ],
        ))),
        "gob.Register" | "b.Register" | "encoding.gob.Register" | "go.encoding.gob.Register" => {
            Some(Expression::null())
        }
        "gob.RegisterName"
        | "b.RegisterName"
        | "encoding.gob.RegisterName"
        | "go.encoding.gob.RegisterName" => Some(Expression::null()),
        _ => None,
    }
}

pub(crate) fn encode_expr(encoder: Expression, value: Expression) -> Expression {
    let e = Expression::ident("__go_gob_encoder");
    let w = member(e.clone(), "w");
    let len = member(w.clone(), "gob_len");
    lambda_call(vec![
        var_decl("__go_gob_encoder", value_arg(encoder)),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Or,
                binary(BinOp::Eq, e.clone(), Expression::null()),
                binary(BinOp::Eq, w.clone(), Expression::null()),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::null())))],
            elifs: Vec::new(),
            else_body: None,
        }),
        gob_slot_assign(len.clone(), w.clone(), value),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(w.clone(), "gob_len")),
            value: Box::new(binary(BinOp::Add, len, Expression::int(1))),
        })),
        expr_stmt(crate::adapters::bytes_io::buffer_write_string_expr(
            w,
            Expression::string("g"),
        )),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

pub(crate) fn next_expr(decoder: Expression) -> Expression {
    let d = Expression::ident("__go_gob_decoder");
    let r = member(d.clone(), "r");
    lambda_call(vec![
        var_decl("__go_gob_decoder", value_arg(decoder)),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Or,
                binary(BinOp::Eq, d.clone(), Expression::null()),
                binary(
                    BinOp::Or,
                    binary(BinOp::Eq, r.clone(), Expression::null()),
                    binary(
                        BinOp::GtEq,
                        member(d.clone(), "pos"),
                        member(r.clone(), "gob_len"),
                    ),
                ),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::null())))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl("__go_gob_val", Expression::null()),
        gob_slot_load(member(d.clone(), "pos"), r.clone()),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(d.clone(), "pos")),
            value: Box::new(binary(BinOp::Add, member(d, "pos"), Expression::int(1))),
        })),
        Statement::new(StmtKind::Return(Some(Expression::ident("__go_gob_val")))),
    ])
}

fn register_type(root: &mut Subtree, name: &str, methods: &[&str]) {
    let mut method_tree = Subtree::new();
    for method in methods {
        method_tree.insert(
            (*method).to_string(),
            NamespaceNode::CommonEmit(format!("go.{name}.{method}")),
        );
    }
    insert_path(
        root,
        name,
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: method_tree,
            member_returns: Default::default(),
        },
    );
}

fn arg(args: &[Argument], idx: usize) -> Expression {
    args.get(idx)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(Expression::null)
}

fn gob_slot_assign(pos: Expression, receiver: Expression, value: Expression) -> Statement {
    let mut fallback = Some(vec![expr_stmt(Expression::new(ExprKind::Assign {
        target: Box::new(member(receiver.clone(), "gob7")),
        value: Box::new(value.clone()),
    }))]);
    for idx in (0..7).rev() {
        fallback = Some(vec![Statement::new(StmtKind::If {
            cond: binary(BinOp::Eq, pos.clone(), Expression::int(idx)),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(member(receiver.clone(), &format!("gob{idx}"))),
                value: Box::new(value.clone()),
            }))],
            elifs: Vec::new(),
            else_body: fallback,
        })]);
    }
    fallback
        .and_then(|mut stmts| stmts.pop())
        .unwrap_or_else(|| expr_stmt(Expression::null()))
}

fn gob_slot_load(pos: Expression, receiver: Expression) -> Statement {
    let mut fallback = Some(vec![expr_stmt(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident("__go_gob_val")),
        value: Box::new(member(receiver.clone(), "gob7")),
    }))]);
    for idx in (0..7).rev() {
        fallback = Some(vec![Statement::new(StmtKind::If {
            cond: binary(BinOp::Eq, pos.clone(), Expression::int(idx)),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_gob_val")),
                value: Box::new(member(receiver.clone(), &format!("gob{idx}"))),
            }))],
            elifs: Vec::new(),
            else_body: fallback,
        })]);
    }
    fallback
        .and_then(|mut stmts| stmts.pop())
        .unwrap_or_else(|| expr_stmt(Expression::null()))
}

fn writer_value(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::RefOf(place) => place_expr(&place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        _ => expr,
    }
}

fn place_expr(place: &PlaceExpr) -> Expression {
    match place {
        PlaceExpr::Ident(name) => Expression::ident(name),
        PlaceExpr::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: object.clone(),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: object.clone(),
            index: index.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Deref(expr) => Expression::new(ExprKind::RefLoad(expr.clone())),
    }
}

fn typed_object(type_name: &str, fields: Vec<(&str, Expression)>) -> Expression {
    crate::adapters::bytes_io::typed_object(type_name, fields)
}

fn pointer_arg(expr: Expression) -> Expression {
    crate::adapters::bytes_io::pointer_arg(expr)
}

fn value_arg(expr: Expression) -> Expression {
    crate::adapters::bytes_io::value_arg(expr)
}

fn member(object: Expression, field: &str) -> Expression {
    crate::adapters::bytes_io::member(object, field)
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    crate::adapters::bytes_io::binary(op, left, right)
}

fn var_decl(name: &str, init: Expression) -> Statement {
    crate::adapters::bytes_io::var_decl(name, init)
}

fn expr_stmt(expr: Expression) -> Statement {
    crate::adapters::bytes_io::expr_stmt(expr)
}

fn lambda_call(body: Vec<Statement>) -> Expression {
    crate::adapters::bytes_io::lambda_call(body)
}

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.insert(leaf.to_string(), node);
}
