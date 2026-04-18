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

// Attribute access

#[test]
fn attribute_access() {
    let m = parse("x = obj.attr\n");
    match first_value(&m) {
        Expression::Attribute { attr, .. } => assert_eq!(attr, "attr"),
        other => panic!("expected Attribute, got: {:?}", other),
    }
}

#[test]
fn chained_attribute_access() {
    let m = parse("x = a.b.c\n");
    match first_value(&m) {
        Expression::Attribute { value, attr, .. } => {
            assert_eq!(attr, "c");
            assert!(matches!(value.as_ref(), Expression::Attribute { attr, .. } if attr == "b"));
        }
        other => panic!("expected Attribute, got: {:?}", other),
    }
}

// Method calls

#[test]
fn method_call() {
    let m = parse("x = obj.method(1, 2)\n");
    match first_value(&m) {
        Expression::Call { func, args, .. } => {
            assert!(matches!(func.as_ref(), Expression::Attribute { attr, .. } if attr == "method"));
            assert_eq!(args.len(), 2);
        }
        other => panic!("expected Call, got: {:?}", other),
    }
}

#[test]
fn chained_method_calls() {
    let m = parse(r#"x = "hello".upper().strip()"#);
    match first_value(&m) {
        Expression::Call { func, .. } => {
            // strip() is called on the result of upper()
            assert!(matches!(func.as_ref(), Expression::Attribute { attr, .. } if attr == "strip"));
        }
        other => panic!("expected Call, got: {:?}", other),
    }
}

// Subscript / indexing

#[test]
fn index_access() {
    let m = parse("x = lst[0]\n");
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            assert!(matches!(slice.as_ref(), Expression::Int(0)));
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

#[test]
fn string_key_access() {
    let m = parse(r#"x = d["key"]"#);
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            assert!(matches!(slice.as_ref(), Expression::Str(s) if s == "key"));
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

#[test]
fn negative_index() {
    let m = parse("x = lst[-1]\n");
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            assert!(matches!(slice.as_ref(), Expression::UnaryOp { op: UnaryOp::USub, .. }));
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

// Slicing

#[test]
fn slice_basic() {
    let m = parse("x = lst[1:3]\n");
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            match slice.as_ref() {
                Expression::Slice { lower, upper, step } => {
                    assert!(lower.is_some());
                    assert!(upper.is_some());
                    assert!(step.is_none());
                }
                other => panic!("expected Slice, got: {:?}", other),
            }
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

#[test]
fn slice_from_start() {
    let m = parse("x = lst[:3]\n");
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            match slice.as_ref() {
                Expression::Slice { lower, upper, .. } => {
                    assert!(lower.is_none());
                    assert!(upper.is_some());
                }
                other => panic!("expected Slice, got: {:?}", other),
            }
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

#[test]
fn slice_to_end() {
    let m = parse("x = lst[1:]\n");
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            match slice.as_ref() {
                Expression::Slice { lower, upper, .. } => {
                    assert!(lower.is_some());
                    assert!(upper.is_none());
                }
                other => panic!("expected Slice, got: {:?}", other),
            }
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

#[test]
fn slice_with_step() {
    let m = parse("x = lst[::2]\n");
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            match slice.as_ref() {
                Expression::Slice { lower, upper, step } => {
                    assert!(lower.is_none());
                    assert!(upper.is_none());
                    assert!(step.is_some());
                }
                other => panic!("expected Slice, got: {:?}", other),
            }
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

#[test]
fn slice_full() {
    let m = parse("x = lst[1:10:2]\n");
    match first_value(&m) {
        Expression::Subscript { slice, .. } => {
            match slice.as_ref() {
                Expression::Slice { lower, upper, step } => {
                    assert!(lower.is_some());
                    assert!(upper.is_some());
                    assert!(step.is_some());
                }
                other => panic!("expected Slice, got: {:?}", other),
            }
        }
        other => panic!("expected Subscript, got: {:?}", other),
    }
}

// Function calls

#[test]
fn call_no_args() {
    let m = parse("f()\n");
    match first_value(&m) {
        Expression::Call { args, keywords, .. } => {
            assert!(args.is_empty());
            assert!(keywords.is_empty());
        }
        other => panic!("expected Call, got: {:?}", other),
    }
}

#[test]
fn call_positional_args() {
    let m = parse("f(1, 2, 3)\n");
    match first_value(&m) {
        Expression::Call { args, .. } => assert_eq!(args.len(), 3),
        other => panic!("expected Call, got: {:?}", other),
    }
}

#[test]
fn call_keyword_args() {
    let m = parse("f(x=1, y=2)\n");
    match first_value(&m) {
        Expression::Call { keywords, .. } => {
            assert_eq!(keywords.len(), 2);
            assert_eq!(keywords[0].name.as_deref(), Some("x"));
        }
        other => panic!("expected Call, got: {:?}", other),
    }
}

#[test]
fn call_star_args() {
    let m = parse("f(*args)\n");
    match first_value(&m) {
        Expression::Call { args, .. } => {
            assert_eq!(args.len(), 1);
            assert!(matches!(&args[0], Expression::Starred(_)));
        }
        other => panic!("expected Call, got: {:?}", other),
    }
}

#[test]
fn call_double_star_kwargs() {
    let m = parse("f(**kwargs)\n");
    match first_value(&m) {
        Expression::Call { keywords, .. } => {
            assert_eq!(keywords.len(), 1);
            assert!(keywords[0].name.is_none()); // **kwargs has name=None
        }
        other => panic!("expected Call, got: {:?}", other),
    }
}

// With statement

#[test]
fn with_statement() {
    let m = parse("with open('f') as f:\n    pass\n");
    match &m.body[0] {
        Statement::With { items, .. } => {
            assert_eq!(items.len(), 1);
            assert!(items[0].optional_vars.is_some());
        }
        other => panic!("expected With, got: {:?}", other),
    }
}

#[test]
fn with_no_as() {
    let m = parse("with open('f'):\n    pass\n");
    match &m.body[0] {
        Statement::With { items, .. } => {
            assert_eq!(items.len(), 1);
            assert!(items[0].optional_vars.is_none());
        }
        other => panic!("expected With, got: {:?}", other),
    }
}
