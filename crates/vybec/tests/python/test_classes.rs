use vybec::parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybec::parser_python::parse(src).expect("parse failed")
}

#[test]
fn empty_class() {
    let m = parse("class Foo:\n    pass\n");
    match &m.body[0] {
        Statement::ClassDef { name, body, bases, .. } => {
            assert_eq!(name, "Foo");
            assert!(bases.is_empty());
            assert_eq!(body.len(), 1);
            assert!(matches!(&body[0], Statement::Pass));
        }
        other => panic!("expected ClassDef, got: {:?}", other),
    }
}

#[test]
fn class_with_methods() {
    let m = parse("class Dog:\n    def bark(self):\n        pass\n    def sit(self):\n        pass\n");
    match &m.body[0] {
        Statement::ClassDef { body, .. } => {
            assert_eq!(body.len(), 2);
            assert!(matches!(&body[0], Statement::FunctionDef { .. }));
            assert!(matches!(&body[1], Statement::FunctionDef { .. }));
        }
        other => panic!("expected ClassDef, got: {:?}", other),
    }
}

#[test]
fn class_with_init() {
    let src = "class Dog:\n    def __init__(self, name):\n        self.name = name\n";
    let m = parse(src);
    match &m.body[0] {
        Statement::ClassDef { body, .. } => {
            match &body[0] {
                Statement::FunctionDef { name, params, .. } => {
                    assert_eq!(name, "__init__");
                    assert_eq!(params.args.len(), 2); // self, name
                }
                other => panic!("expected FunctionDef, got: {:?}", other),
            }
        }
        other => panic!("expected ClassDef, got: {:?}", other),
    }
}

#[test]
fn class_with_inheritance() {
    let m = parse("class Dog(Animal):\n    pass\n");
    match &m.body[0] {
        Statement::ClassDef { bases, .. } => {
            assert_eq!(bases.len(), 1);
            assert!(matches!(&bases[0], Expression::Name(n) if n == "Animal"));
        }
        other => panic!("expected ClassDef, got: {:?}", other),
    }
}

#[test]
fn class_multiple_bases() {
    let m = parse("class C(A, B):\n    pass\n");
    match &m.body[0] {
        Statement::ClassDef { bases, .. } => {
            assert_eq!(bases.len(), 2);
        }
        other => panic!("expected ClassDef, got: {:?}", other),
    }
}

#[test]
fn class_with_metaclass() {
    let m = parse("class Foo(metaclass=ABCMeta):\n    pass\n");
    match &m.body[0] {
        Statement::ClassDef { keywords, .. } => {
            assert_eq!(keywords.len(), 1);
            assert_eq!(keywords[0].name.as_deref(), Some("metaclass"));
        }
        other => panic!("expected ClassDef, got: {:?}", other),
    }
}

#[test]
fn decorated_class() {
    let m = parse("@dataclass\nclass Point:\n    x: int\n    y: int\n");
    match &m.body[0] {
        Statement::ClassDef { decorators, .. } => {
            assert_eq!(decorators.len(), 1);
        }
        other => panic!("expected ClassDef, got: {:?}", other),
    }
}
