use vybe_parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybe_parser_python::parse(src).expect("parse failed")
}

#[test]
fn try_except_basic() {
    let m = parse("try:\n    x = 1\nexcept:\n    pass\n");
    match &m.body[0] {
        Statement::Try { body, handlers, else_body, finally_body } => {
            assert_eq!(body.len(), 1);
            assert_eq!(handlers.len(), 1);
            assert!(handlers[0].exc_type.is_none());
            assert!(handlers[0].name.is_none());
            assert!(else_body.is_none());
            assert!(finally_body.is_none());
        }
        other => panic!("expected Try, got: {:?}", other),
    }
}

#[test]
fn try_except_typed() {
    let m = parse("try:\n    pass\nexcept ValueError:\n    pass\n");
    match &m.body[0] {
        Statement::Try { handlers, .. } => {
            assert!(handlers[0].exc_type.is_some());
            assert!(handlers[0].name.is_none());
        }
        other => panic!("expected Try, got: {:?}", other),
    }
}

#[test]
fn try_except_as() {
    let m = parse("try:\n    pass\nexcept ValueError as e:\n    print(e)\n");
    match &m.body[0] {
        Statement::Try { handlers, .. } => {
            assert!(handlers[0].exc_type.is_some());
            assert_eq!(handlers[0].name.as_deref(), Some("e"));
        }
        other => panic!("expected Try, got: {:?}", other),
    }
}

#[test]
fn multiple_except_handlers() {
    let m = parse("try:\n    pass\nexcept TypeError:\n    pass\nexcept ValueError:\n    pass\n");
    match &m.body[0] {
        Statement::Try { handlers, .. } => {
            assert_eq!(handlers.len(), 2);
        }
        other => panic!("expected Try, got: {:?}", other),
    }
}

#[test]
fn try_except_else() {
    let m = parse("try:\n    pass\nexcept:\n    pass\nelse:\n    print('ok')\n");
    match &m.body[0] {
        Statement::Try { else_body, .. } => {
            assert!(else_body.is_some());
        }
        other => panic!("expected Try, got: {:?}", other),
    }
}

#[test]
fn try_finally() {
    let m = parse("try:\n    pass\nfinally:\n    cleanup()\n");
    match &m.body[0] {
        Statement::Try { handlers, finally_body, .. } => {
            assert!(handlers.is_empty());
            assert!(finally_body.is_some());
        }
        other => panic!("expected Try, got: {:?}", other),
    }
}

#[test]
fn try_except_finally() {
    let m = parse("try:\n    pass\nexcept:\n    pass\nfinally:\n    cleanup()\n");
    match &m.body[0] {
        Statement::Try { handlers, finally_body, .. } => {
            assert_eq!(handlers.len(), 1);
            assert!(finally_body.is_some());
        }
        other => panic!("expected Try, got: {:?}", other),
    }
}

#[test]
fn raise_bare() {
    let m = parse("raise\n");
    match &m.body[0] {
        Statement::Raise { exc, cause } => {
            assert!(exc.is_none());
            assert!(cause.is_none());
        }
        other => panic!("expected Raise, got: {:?}", other),
    }
}

#[test]
fn raise_exception() {
    let m = parse("raise ValueError()\n");
    match &m.body[0] {
        Statement::Raise { exc, cause } => {
            assert!(exc.is_some());
            assert!(cause.is_none());
        }
        other => panic!("expected Raise, got: {:?}", other),
    }
}

#[test]
fn raise_from() {
    let m = parse("raise ValueError() from e\n");
    match &m.body[0] {
        Statement::Raise { exc, cause } => {
            assert!(exc.is_some());
            assert!(cause.is_some());
        }
        other => panic!("expected Raise, got: {:?}", other),
    }
}
