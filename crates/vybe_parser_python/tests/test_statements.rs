use vybe_parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybe_parser_python::parse(src).expect("parse failed")
}

// Assignment

#[test]
fn simple_assign() {
    let m = parse("x = 42\n");
    match &m.body[0] {
        Statement::Assign { targets, value } => {
            assert_eq!(targets.len(), 1);
            assert!(matches!(&targets[0], Expression::Name(n) if n == "x"));
            assert!(matches!(value, Expression::Int(42)));
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn multiple_assign() {
    let m = parse("a = b = 1\n");
    match &m.body[0] {
        Statement::Assign { targets, value } => {
            assert_eq!(targets.len(), 2);
            assert!(matches!(value, Expression::Int(1)));
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn tuple_unpacking() {
    let m = parse("a, b, c = 1, 2, 3\n");
    match &m.body[0] {
        Statement::Assign { targets, value } => {
            assert!(matches!(&targets[0], Expression::Tuple(t) if t.len() == 3));
            assert!(matches!(value, Expression::Tuple(t) if t.len() == 3));
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn augmented_assign_all_ops() {
    let ops = vec![
        ("x += 1\n", AugOp::Add),
        ("x -= 1\n", AugOp::Sub),
        ("x *= 1\n", AugOp::Mul),
        ("x /= 1\n", AugOp::Div),
        ("x //= 1\n", AugOp::FloorDiv),
        ("x %= 1\n", AugOp::Mod),
        ("x **= 1\n", AugOp::Pow),
        ("x <<= 1\n", AugOp::LShift),
        ("x >>= 1\n", AugOp::RShift),
        ("x |= 1\n", AugOp::BitOr),
        ("x &= 1\n", AugOp::BitAnd),
        ("x ^= 1\n", AugOp::BitXor),
    ];
    for (src, expected_op) in ops {
        let m = parse(src);
        match &m.body[0] {
            Statement::AugAssign { op, .. } => assert_eq!(*op, expected_op, "failed for: {}", src),
            other => panic!("expected AugAssign for {}, got: {:?}", src, other),
        }
    }
}

#[test]
fn type_annotation() {
    let m = parse("x: int = 5\n");
    assert!(matches!(&m.body[0], Statement::AnnAssign { .. }));
}

// Control flow

#[test]
fn if_simple() {
    let m = parse("if True:\n    pass\n");
    match &m.body[0] {
        Statement::If { body, elif_clauses, else_body, .. } => {
            assert_eq!(body.len(), 1);
            assert!(elif_clauses.is_empty());
            assert!(else_body.is_none());
        }
        other => panic!("expected If, got: {:?}", other),
    }
}

#[test]
fn if_else() {
    let m = parse("if True:\n    x = 1\nelse:\n    x = 2\n");
    match &m.body[0] {
        Statement::If { else_body, .. } => {
            assert!(else_body.is_some());
        }
        other => panic!("expected If, got: {:?}", other),
    }
}

#[test]
fn if_elif_else() {
    let m = parse("if a:\n    pass\nelif b:\n    pass\nelif c:\n    pass\nelse:\n    pass\n");
    match &m.body[0] {
        Statement::If { elif_clauses, else_body, .. } => {
            assert_eq!(elif_clauses.len(), 2);
            assert!(else_body.is_some());
        }
        other => panic!("expected If, got: {:?}", other),
    }
}

#[test]
fn single_line_if() {
    let m = parse("if True: pass\n");
    match &m.body[0] {
        Statement::If { body, .. } => {
            assert_eq!(body.len(), 1);
            assert!(matches!(&body[0], Statement::Pass));
        }
        other => panic!("expected If, got: {:?}", other),
    }
}

#[test]
fn while_loop() {
    let m = parse("while True:\n    break\n");
    match &m.body[0] {
        Statement::While { body, .. } => {
            assert!(matches!(&body[0], Statement::Break));
        }
        other => panic!("expected While, got: {:?}", other),
    }
}

#[test]
fn while_else() {
    let m = parse("while x > 0:\n    x -= 1\nelse:\n    print(x)\n");
    match &m.body[0] {
        Statement::While { else_body, .. } => {
            assert!(else_body.is_some());
        }
        other => panic!("expected While, got: {:?}", other),
    }
}

#[test]
fn for_loop() {
    let m = parse("for i in items:\n    print(i)\n");
    match &m.body[0] {
        Statement::For { target, body, .. } => {
            assert!(matches!(target, Expression::Name(n) if n == "i"));
            assert_eq!(body.len(), 1);
        }
        other => panic!("expected For, got: {:?}", other),
    }
}

#[test]
fn for_tuple_unpacking() {
    let m = parse("for k, v in items:\n    pass\n");
    match &m.body[0] {
        Statement::For { target, .. } => {
            assert!(matches!(target, Expression::Tuple(t) if t.len() == 2));
        }
        other => panic!("expected For, got: {:?}", other),
    }
}

#[test]
fn for_else() {
    let m = parse("for x in lst:\n    pass\nelse:\n    print('done')\n");
    match &m.body[0] {
        Statement::For { else_body, .. } => {
            assert!(else_body.is_some());
        }
        other => panic!("expected For, got: {:?}", other),
    }
}

#[test]
fn break_continue_pass() {
    let m = parse("while True:\n    break\n    continue\n    pass\n");
    match &m.body[0] {
        Statement::While { body, .. } => {
            assert!(matches!(&body[0], Statement::Break));
            assert!(matches!(&body[1], Statement::Continue));
            assert!(matches!(&body[2], Statement::Pass));
        }
        other => panic!("expected While, got: {:?}", other),
    }
}

// Return

#[test]
fn return_value() {
    let m = parse("def f():\n    return 42\n");
    match &m.body[0] {
        Statement::FunctionDef { body, .. } => {
            assert!(matches!(&body[0], Statement::Return(Some(Expression::Int(42)))));
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn return_none() {
    let m = parse("def f():\n    return\n");
    match &m.body[0] {
        Statement::FunctionDef { body, .. } => {
            assert!(matches!(&body[0], Statement::Return(None)));
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

// Delete

#[test]
fn delete_statement() {
    let m = parse("del x\n");
    assert!(matches!(&m.body[0], Statement::Delete(targets) if targets.len() == 1));
}

// Assert

#[test]
fn assert_simple() {
    let m = parse("assert True\n");
    assert!(matches!(&m.body[0], Statement::Assert { msg: None, .. }));
}

#[test]
fn assert_with_message() {
    let m = parse(r#"assert x > 0, "must be positive""#);
    assert!(matches!(&m.body[0], Statement::Assert { msg: Some(_), .. }));
}

// Global / Nonlocal

#[test]
fn global_statement() {
    let m = parse("global x, y\n");
    match &m.body[0] {
        Statement::Global(names) => assert_eq!(names, &["x", "y"]),
        other => panic!("expected Global, got: {:?}", other),
    }
}

#[test]
fn nonlocal_statement() {
    let m = parse("nonlocal x\n");
    match &m.body[0] {
        Statement::Nonlocal(names) => assert_eq!(names, &["x"]),
        other => panic!("expected Nonlocal, got: {:?}", other),
    }
}

// Comments

#[test]
fn comments_ignored() {
    let m = parse("# this is a comment\nx = 1  # inline comment\n");
    assert_eq!(m.body.len(), 1);
}

// Multiline with implicit continuation

#[test]
fn multiline_parens() {
    let m = parse("x = (1 +\n    2 +\n    3)\n");
    match &m.body[0] {
        Statement::Assign { value, .. } => {
            assert!(matches!(value, Expression::BinOp { .. }));
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn multiline_brackets() {
    let m = parse("x = [\n    1,\n    2,\n    3\n]\n");
    match &m.body[0] {
        Statement::Assign { value, .. } => {
            assert!(matches!(value, Expression::List(v) if v.len() == 3));
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}
