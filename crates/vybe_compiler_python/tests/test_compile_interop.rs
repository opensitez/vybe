use vybe_parser_python::parse;
use vybe_compiler_python::Compiler;

fn compile(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    c.compile(&module).expect("compile failed")
}

#[test]
fn class_emits_type_entry() {
    let chunks = compile(r#"
class Dog:
    def bark(self):
        print("woof")
"#);
    let types = &chunks[0].types;
    assert!(!types.is_empty(), "should have type entries");
    let dog = types.iter().find(|t| t.name == "dog").expect("Dog type not found");
    assert!(!dog.methods.is_empty(), "Dog should have methods");
    assert!(dog.methods.iter().any(|(n, _)| n == "bark"));
}

#[test]
fn class_with_init_has_constructor() {
    let chunks = compile(r#"
class Dog:
    def __init__(self, name):
        self.name = name
    def bark(self):
        return self.name
"#);
    let dog = chunks[0].types.iter().find(|t| t.name == "dog").expect("Dog type not found");
    assert!(dog.constructor_chunk.is_some(), "Dog should have constructor chunk");
    assert!(dog.methods.iter().any(|(n, _)| n == "__init__"));
    assert!(dog.methods.iter().any(|(n, _)| n == "bark"));
}

#[test]
fn class_with_inheritance() {
    let chunks = compile(r#"
class Animal:
    def speak(self):
        pass

class Dog(Animal):
    def bark(self):
        pass
"#);
    let dog = chunks[0].types.iter().find(|t| t.name == "dog").expect("Dog type not found");
    assert_eq!(dog.parent, "animal");
}

#[test]
fn multiple_classes() {
    let chunks = compile(r#"
class Cat:
    def meow(self):
        pass

class Dog:
    def bark(self):
        pass
"#);
    assert!(chunks[0].types.len() >= 2);
    assert!(chunks[0].types.iter().any(|t| t.name == "cat"));
    assert!(chunks[0].types.iter().any(|t| t.name == "dog"));
}

#[test]
fn class_not_interface() {
    let chunks = compile(r#"
class Foo:
    def bar(self):
        pass
"#);
    let foo = chunks[0].types.iter().find(|t| t.name == "foo").unwrap();
    assert!(!foo.is_interface);
}
