use vybec::parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybec::parser_python::parse(src).expect("parse failed")
}

fn first_expr(m: &Module) -> &Expression {
    match &m.body[0] {
        Statement::Expression(e) => e,
        Statement::Assign { value, .. } => value,
        other => panic!("expected expression or assign, got: {:?}", other),
    }
}

#[test]
fn int_literal() {
    let m = parse("42\n");
    assert!(matches!(first_expr(&m), Expression::Int(42)));
}

#[test]
fn negative_int() {
    let m = parse("x = -5\n");
    match first_expr(&m) {
        Expression::UnaryOp { op: UnaryOp::USub, operand } => {
            assert!(matches!(operand.as_ref(), Expression::Int(5)));
        }
        other => panic!("expected unary neg, got: {:?}", other),
    }
}

#[test]
fn float_literal() {
    let m = parse("3.14\n");
    match first_expr(&m) {
        Expression::Float(f) => assert!((*f - 3.14).abs() < 1e-10),
        other => panic!("expected float, got: {:?}", other),
    }
}

#[test]
fn float_exponent() {
    let m = parse("1e10\n");
    match first_expr(&m) {
        Expression::Float(f) => assert!((*f - 1e10).abs() < 1.0),
        other => panic!("expected float, got: {:?}", other),
    }
}

#[test]
fn hex_int() {
    let m = parse("0xFF\n");
    assert!(matches!(first_expr(&m), Expression::Int(255)));
}

#[test]
fn octal_int() {
    let m = parse("0o77\n");
    assert!(matches!(first_expr(&m), Expression::Int(63)));
}

#[test]
fn binary_int() {
    let m = parse("0b1010\n");
    assert!(matches!(first_expr(&m), Expression::Int(10)));
}

#[test]
fn underscore_in_number() {
    let m = parse("1_000_000\n");
    assert!(matches!(first_expr(&m), Expression::Int(1_000_000)));
}

#[test]
fn bool_true() {
    let m = parse("True\n");
    assert!(matches!(first_expr(&m), Expression::Bool(true)));
}

#[test]
fn bool_false() {
    let m = parse("False\n");
    assert!(matches!(first_expr(&m), Expression::Bool(false)));
}

#[test]
fn none_literal() {
    let m = parse("None\n");
    assert!(matches!(first_expr(&m), Expression::None));
}

#[test]
fn ellipsis_literal() {
    let m = parse("...\n");
    assert!(matches!(first_expr(&m), Expression::Ellipsis));
}

#[test]
fn string_double_quote() {
    let m = parse(r#"x = "hello""#);
    match first_expr(&m) {
        Expression::Str(s) => assert_eq!(s, "hello"),
        other => panic!("expected str, got: {:?}", other),
    }
}

#[test]
fn string_single_quote() {
    let m = parse("x = 'world'\n");
    match first_expr(&m) {
        Expression::Str(s) => assert_eq!(s, "world"),
        other => panic!("expected str, got: {:?}", other),
    }
}

#[test]
fn string_escape_sequences() {
    let m = parse(r#"x = "a\nb\tc""#);
    match first_expr(&m) {
        Expression::Str(s) => assert_eq!(s, "a\nb\tc"),
        other => panic!("expected str, got: {:?}", other),
    }
}

#[test]
fn triple_quoted_string() {
    let m = parse("x = '''multi\nline'''\n");
    match first_expr(&m) {
        Expression::Str(s) => assert!(s.contains('\n')),
        other => panic!("expected str, got: {:?}", other),
    }
}

#[test]
fn raw_string() {
    let m = parse(r#"x = r"no\nescape""#);
    match first_expr(&m) {
        Expression::Str(s) => assert_eq!(s, r"no\nescape"),
        other => panic!("expected str, got: {:?}", other),
    }
}

#[test]
fn byte_string() {
    let m = parse(r#"x = b"bytes""#);
    // byte strings are parsed as Str in our AST
    match first_expr(&m) {
        Expression::Str(s) => assert_eq!(s, "bytes"),
        other => panic!("expected str, got: {:?}", other),
    }
}

#[test]
fn string_concatenation() {
    let m = parse(r#"x = "hello" " " "world""#);
    match first_expr(&m) {
        Expression::Str(s) => assert_eq!(s, "hello world"),
        other => panic!("expected str, got: {:?}", other),
    }
}
