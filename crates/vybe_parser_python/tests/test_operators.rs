use vybe_parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybe_parser_python::parse(src).expect("parse failed")
}

fn first_value(m: &Module) -> &Expression {
    match &m.body[0] {
        Statement::Assign { value, .. } => value,
        Statement::Expression(e) => e,
        other => panic!("expected assign/expr, got: {:?}", other),
    }
}

// Arithmetic

#[test]
fn add() {
    let m = parse("x = 1 + 2\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::Add, .. }));
}

#[test]
fn sub() {
    let m = parse("x = 5 - 3\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::Sub, .. }));
}

#[test]
fn mul() {
    let m = parse("x = 2 * 3\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::Mul, .. }));
}

#[test]
fn div() {
    let m = parse("x = 10 / 3\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::Div, .. }));
}

#[test]
fn floor_div() {
    let m = parse("x = 10 // 3\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::FloorDiv, .. }));
}

#[test]
fn modulo() {
    let m = parse("x = 10 % 3\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::Mod, .. }));
}

#[test]
fn power() {
    let m = parse("x = 2 ** 10\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::Pow, .. }));
}

// Precedence

#[test]
fn mul_before_add() {
    let m = parse("x = 1 + 2 * 3\n");
    // Should be Add(1, Mul(2, 3))
    match first_value(&m) {
        Expression::BinOp { op: BinOp::Add, left, right } => {
            assert!(matches!(left.as_ref(), Expression::Int(1)));
            assert!(matches!(right.as_ref(), Expression::BinOp { op: BinOp::Mul, .. }));
        }
        other => panic!("expected Add, got: {:?}", other),
    }
}

#[test]
fn power_before_mul() {
    let m = parse("x = 2 * 3 ** 4\n");
    match first_value(&m) {
        Expression::BinOp { op: BinOp::Mul, right, .. } => {
            assert!(matches!(right.as_ref(), Expression::BinOp { op: BinOp::Pow, .. }));
        }
        other => panic!("expected Mul, got: {:?}", other),
    }
}

#[test]
fn parens_override_precedence() {
    let m = parse("x = (1 + 2) * 3\n");
    match first_value(&m) {
        Expression::BinOp { op: BinOp::Mul, left, .. } => {
            assert!(matches!(left.as_ref(), Expression::BinOp { op: BinOp::Add, .. }));
        }
        other => panic!("expected Mul, got: {:?}", other),
    }
}

// Comparison

#[test]
fn eq() {
    let m = parse("x = a == b\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => assert_eq!(ops, &[CmpOp::Eq]),
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn not_eq() {
    let m = parse("x = a != b\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => assert_eq!(ops, &[CmpOp::NotEq]),
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn less_than() {
    let m = parse("x = a < b\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => assert_eq!(ops, &[CmpOp::Lt]),
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn chained_comparison() {
    let m = parse("x = a < b < c\n");
    match first_value(&m) {
        Expression::Compare { ops, comparators, .. } => {
            assert_eq!(ops.len(), 2);
            assert_eq!(comparators.len(), 2);
            assert_eq!(ops, &[CmpOp::Lt, CmpOp::Lt]);
        }
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn mixed_chained_comparison() {
    let m = parse("x = 0 <= val < 100\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => {
            assert_eq!(ops, &[CmpOp::LtE, CmpOp::Lt]);
        }
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn in_operator() {
    let m = parse("x = a in b\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => assert_eq!(ops, &[CmpOp::In]),
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn not_in_operator() {
    let m = parse("x = a not in b\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => assert_eq!(ops, &[CmpOp::NotIn]),
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn is_operator() {
    let m = parse("x = a is b\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => assert_eq!(ops, &[CmpOp::Is]),
        other => panic!("expected Compare, got: {:?}", other),
    }
}

#[test]
fn is_not_operator() {
    let m = parse("x = a is not b\n");
    match first_value(&m) {
        Expression::Compare { ops, .. } => assert_eq!(ops, &[CmpOp::IsNot]),
        other => panic!("expected Compare, got: {:?}", other),
    }
}

// Boolean operators

#[test]
fn bool_and() {
    let m = parse("x = a and b\n");
    assert!(matches!(first_value(&m), Expression::BoolOp { op: BoolOp::And, .. }));
}

#[test]
fn bool_or() {
    let m = parse("x = a or b\n");
    assert!(matches!(first_value(&m), Expression::BoolOp { op: BoolOp::Or, .. }));
}

#[test]
fn bool_not() {
    let m = parse("x = not a\n");
    assert!(matches!(first_value(&m), Expression::UnaryOp { op: UnaryOp::Not, .. }));
}

#[test]
fn chained_and() {
    let m = parse("x = a and b and c\n");
    match first_value(&m) {
        Expression::BoolOp { op: BoolOp::And, values } => assert_eq!(values.len(), 3),
        other => panic!("expected BoolOp And, got: {:?}", other),
    }
}

// Bitwise

#[test]
fn bitwise_and() {
    let m = parse("x = a & b\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::BitAnd, .. }));
}

#[test]
fn bitwise_or() {
    let m = parse("x = a | b\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::BitOr, .. }));
}

#[test]
fn bitwise_xor() {
    let m = parse("x = a ^ b\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::BitXor, .. }));
}

#[test]
fn bitwise_invert() {
    let m = parse("x = ~a\n");
    assert!(matches!(first_value(&m), Expression::UnaryOp { op: UnaryOp::Invert, .. }));
}

#[test]
fn left_shift() {
    let m = parse("x = a << 2\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::LShift, .. }));
}

#[test]
fn right_shift() {
    let m = parse("x = a >> 2\n");
    assert!(matches!(first_value(&m), Expression::BinOp { op: BinOp::RShift, .. }));
}

// Ternary

#[test]
fn ternary_expression() {
    let m = parse(r#"x = "yes" if cond else "no""#);
    assert!(matches!(first_value(&m), Expression::IfExp { .. }));
}

// Walrus

#[test]
fn walrus_operator() {
    let m = parse("x = (n := 10)\n");
    assert!(matches!(first_value(&m), Expression::NamedExpr { .. }));
}

// Unary

#[test]
fn unary_positive() {
    let m = parse("x = +a\n");
    assert!(matches!(first_value(&m), Expression::UnaryOp { op: UnaryOp::UAdd, .. }));
}

#[test]
fn unary_negative() {
    let m = parse("x = -a\n");
    assert!(matches!(first_value(&m), Expression::UnaryOp { op: UnaryOp::USub, .. }));
}
