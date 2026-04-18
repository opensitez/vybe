use vybec::parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybec::parser_python::parse(src).expect("parse failed")
}

#[test]
fn simple_function() {
    let m = parse("def greet():\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { name, params, body, .. } => {
            assert_eq!(name, "greet");
            assert!(params.args.is_empty());
            assert_eq!(body.len(), 1);
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn function_with_args() {
    let m = parse("def add(a, b):\n    return a + b\n");
    match &m.body[0] {
        Statement::FunctionDef { params, .. } => {
            assert_eq!(params.args.len(), 2);
            assert_eq!(params.args[0].name, "a");
            assert_eq!(params.args[1].name, "b");
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn function_with_defaults() {
    let m = parse("def greet(name, greeting=\"hello\"):\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { params, .. } => {
            assert_eq!(params.args.len(), 2);
            assert_eq!(params.defaults.len(), 1);
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn function_varargs() {
    let m = parse("def f(*args):\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { params, .. } => {
            assert!(params.vararg.is_some());
            assert_eq!(params.vararg.as_ref().unwrap().name, "args");
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn function_kwargs() {
    let m = parse("def f(**kwargs):\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { params, .. } => {
            assert!(params.kwarg.is_some());
            assert_eq!(params.kwarg.as_ref().unwrap().name, "kwargs");
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn function_all_param_types() {
    let m = parse("def f(a, b=1, *args, c, d=2, **kwargs):\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { params, .. } => {
            assert_eq!(params.args.len(), 2);
            assert_eq!(params.defaults.len(), 1);
            assert!(params.vararg.is_some());
            assert_eq!(params.kwonly_args.len(), 2);
            assert_eq!(params.kw_defaults.len(), 2);
            assert!(params.kwarg.is_some());
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn function_return_annotation() {
    let m = parse("def f(x: int) -> str:\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { params, returns, .. } => {
            assert!(params.args[0].annotation.is_some());
            assert!(returns.is_some());
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn nested_function() {
    let m = parse("def outer():\n    def inner():\n        pass\n    inner()\n");
    match &m.body[0] {
        Statement::FunctionDef { body, .. } => {
            assert_eq!(body.len(), 2);
            assert!(matches!(&body[0], Statement::FunctionDef { .. }));
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn lambda_no_args() {
    let m = parse("f = lambda: 42\n");
    match &m.body[0] {
        Statement::Assign { value, .. } => {
            match value {
                Expression::Lambda { params, .. } => assert!(params.args.is_empty()),
                other => panic!("expected Lambda, got: {:?}", other),
            }
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn lambda_with_args() {
    let m = parse("f = lambda x, y: x + y\n");
    match &m.body[0] {
        Statement::Assign { value, .. } => {
            match value {
                Expression::Lambda { params, body } => {
                    assert_eq!(params.args.len(), 2);
                    assert!(matches!(body.as_ref(), Expression::BinOp { op: BinOp::Add, .. }));
                }
                other => panic!("expected Lambda, got: {:?}", other),
            }
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn lambda_with_default() {
    let m = parse("f = lambda x, y=0: x + y\n");
    match &m.body[0] {
        Statement::Assign { value, .. } => {
            match value {
                Expression::Lambda { params, .. } => {
                    assert_eq!(params.args.len(), 2);
                    assert_eq!(params.defaults.len(), 1);
                }
                other => panic!("expected Lambda, got: {:?}", other),
            }
        }
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn decorator() {
    let m = parse("@staticmethod\ndef foo():\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { decorators, .. } => {
            assert_eq!(decorators.len(), 1);
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn multiple_decorators() {
    let m = parse("@dec1\n@dec2\ndef foo():\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { decorators, .. } => {
            assert_eq!(decorators.len(), 2);
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}

#[test]
fn decorator_with_args() {
    let m = parse("@app.route(\"/\")\ndef index():\n    pass\n");
    match &m.body[0] {
        Statement::FunctionDef { decorators, .. } => {
            assert_eq!(decorators.len(), 1);
            assert!(matches!(&decorators[0], Expression::Call { .. }));
        }
        other => panic!("expected FunctionDef, got: {:?}", other),
    }
}
