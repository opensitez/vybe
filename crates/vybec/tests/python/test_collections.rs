use vybec::parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybec::parser_python::parse(src).expect("parse failed")
}

fn first_value(m: &Module) -> &Expression {
    match &m.body[0] {
        Statement::Assign { value, .. } => value,
        Statement::Expression(e) => e,
        other => panic!("expected assign/expr, got: {:?}", other),
    }
}

// Lists

#[test]
fn empty_list() {
    let m = parse("x = []\n");
    assert!(matches!(first_value(&m), Expression::List(v) if v.is_empty()));
}

#[test]
fn list_of_ints() {
    let m = parse("x = [1, 2, 3]\n");
    match first_value(&m) {
        Expression::List(v) => assert_eq!(v.len(), 3),
        other => panic!("expected list, got: {:?}", other),
    }
}

#[test]
fn list_mixed_types() {
    let m = parse(r#"x = [1, "two", True, None]"#);
    match first_value(&m) {
        Expression::List(v) => assert_eq!(v.len(), 4),
        other => panic!("expected list, got: {:?}", other),
    }
}

#[test]
fn nested_list() {
    let m = parse("x = [[1, 2], [3, 4]]\n");
    match first_value(&m) {
        Expression::List(v) => {
            assert_eq!(v.len(), 2);
            assert!(matches!(&v[0], Expression::List(_)));
        }
        other => panic!("expected list, got: {:?}", other),
    }
}

#[test]
fn list_trailing_comma() {
    let m = parse("x = [1, 2, 3,]\n");
    match first_value(&m) {
        Expression::List(v) => assert_eq!(v.len(), 3),
        other => panic!("expected list, got: {:?}", other),
    }
}

// Tuples

#[test]
fn empty_tuple() {
    let m = parse("x = ()\n");
    assert!(matches!(first_value(&m), Expression::Tuple(v) if v.is_empty()));
}

#[test]
fn tuple_of_ints() {
    let m = parse("x = (1, 2, 3)\n");
    match first_value(&m) {
        Expression::Tuple(v) => assert_eq!(v.len(), 3),
        other => panic!("expected tuple, got: {:?}", other),
    }
}

#[test]
fn single_element_tuple() {
    let m = parse("x = (1,)\n");
    match first_value(&m) {
        Expression::Tuple(v) => assert_eq!(v.len(), 1),
        other => panic!("expected tuple, got: {:?}", other),
    }
}

#[test]
fn parenthesized_expr_not_tuple() {
    let m = parse("x = (42)\n");
    // No trailing comma → not a tuple
    assert!(matches!(first_value(&m), Expression::Int(42)));
}

// Dicts

#[test]
fn empty_dict() {
    let m = parse("x = {}\n");
    match first_value(&m) {
        Expression::Dict { keys, .. } => assert!(keys.is_empty()),
        other => panic!("expected dict, got: {:?}", other),
    }
}

#[test]
fn dict_str_keys() {
    let m = parse(r#"x = {"a": 1, "b": 2}"#);
    match first_value(&m) {
        Expression::Dict { keys, values } => {
            assert_eq!(keys.len(), 2);
            assert_eq!(values.len(), 2);
        }
        other => panic!("expected dict, got: {:?}", other),
    }
}

#[test]
fn dict_int_keys() {
    let m = parse("x = {1: 'a', 2: 'b'}\n");
    match first_value(&m) {
        Expression::Dict { keys, .. } => assert_eq!(keys.len(), 2),
        other => panic!("expected dict, got: {:?}", other),
    }
}

// Sets

#[test]
fn set_literal() {
    let m = parse("x = {1, 2, 3}\n");
    match first_value(&m) {
        Expression::Set(v) => assert_eq!(v.len(), 3),
        other => panic!("expected set, got: {:?}", other),
    }
}

// F-strings

#[test]
fn fstring_basic() {
    let m = parse(r#"x = f"hello {name}""#);
    match first_value(&m) {
        Expression::FString { parts } => {
            assert_eq!(parts.len(), 2); // "hello " + expr
        }
        other => panic!("expected fstring, got: {:?}", other),
    }
}

#[test]
fn fstring_multiple_exprs() {
    let m = parse(r#"x = f"{a} + {b} = {c}""#);
    match first_value(&m) {
        Expression::FString { parts } => {
            // parts: "" + a + " + " + b + " = " + c + ""
            // The exact count depends on how empty strings are handled
            assert!(parts.len() >= 3);
        }
        other => panic!("expected fstring, got: {:?}", other),
    }
}

#[test]
fn fstring_expression() {
    let m = parse(r#"x = f"result: {1 + 2}""#);
    assert!(matches!(first_value(&m), Expression::FString { .. }));
}
