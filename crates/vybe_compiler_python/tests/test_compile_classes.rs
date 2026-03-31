use vybe_parser_python::parse;
use vybe_compiler_python::Compiler;
use vybe_bytecode::chunk::TypeEntry;

fn compile(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    c.compile(&module).expect("compile failed")
}

#[test]
fn class_produces_constructor_chunk() {
    let chunks = compile(r#"
class Dog:
    def __init__(self, name):
        self.name = name
    def bark(self):
        return self.name
"#);
    // Should have: script(0) + __init__(N) + bark(N+1) + dog constructor(N+2) + stdlib
    assert!(chunks.len() > 3);
    // The constructor chunk should be named "dog"
    let ctor = chunks.iter().find(|c| c.name == "dog");
    assert!(ctor.is_some(), "should have a 'dog' constructor chunk");
}

#[test]
fn class_has_type_entry() {
    let chunks = compile(r#"
class Cat:
    def meow(self):
        pass
"#);
    let types = &chunks[0].types;
    let cat = types.iter().find(|t| t.name == "cat");
    assert!(cat.is_some(), "should have Cat type entry");
    let cat = cat.unwrap();
    assert!(cat.methods.iter().any(|(n, _)| n == "meow"));
    assert!(cat.constructor_chunk.is_some());
}

#[test]
fn class_with_inheritance_has_parent() {
    let chunks = compile(r#"
class Animal:
    def speak(self):
        pass
class Dog(Animal):
    def bark(self):
        pass
"#);
    let types = &chunks[0].types;
    let dog = types.iter().find(|t| t.name == "dog").unwrap();
    assert_eq!(dog.parent, "animal");
}

#[test]
fn constructor_arity_excludes_self() {
    let chunks = compile(r#"
class Dog:
    def __init__(self, name, breed):
        self.name = name
        self.breed = breed
"#);
    let ctor = chunks.iter().find(|c| c.name == "dog").unwrap();
    // __init__ takes (self, name, breed) but constructor arity should be 2 (name, breed)
    assert_eq!(ctor.arity, 2, "constructor arity should exclude self");
}

#[test]
fn class_constructor_stamps_type() {
    // Verify the constructor bytecode contains set_type_id
    let chunks = compile(r#"
class Foo:
    pass
"#);
    let ctor = chunks.iter().find(|c| c.name == "foo").unwrap();
    // The bytecode should contain the set_type_id opcode
    let has_set_type = ctor.code.iter().any(|&b| b == vybe_bytecode::opcode::Op::set_type_id.encode().0
        || (b == 0xFE)); // extended opcode prefix
    // set_type_id is an extended opcode, check for 0xFE prefix
    let set_type_id_encoded = vybe_bytecode::opcode::Op::set_type_id.encode();
    let has_it = if let (prefix, Some(ext)) = set_type_id_encoded {
        ctor.code.windows(2).any(|w| w[0] == prefix && w[1] == ext)
    } else {
        ctor.code.contains(&set_type_id_encoded.0)
    };
    assert!(has_it, "constructor should stamp type_id");
}

#[test]
fn no_args_class() {
    let chunks = compile(r#"
class Empty:
    pass
"#);
    let ctor = chunks.iter().find(|c| c.name == "empty").unwrap();
    assert_eq!(ctor.arity, 0);
}

#[test]
fn class_with_only_methods() {
    let chunks = compile(r#"
class Calculator:
    def add(self, a, b):
        return a + b
    def sub(self, a, b):
        return a - b
"#);
    let types = &chunks[0].types;
    let calc = types.iter().find(|t| t.name == "calculator").unwrap();
    assert!(calc.methods.iter().any(|(n, _)| n == "add"));
    assert!(calc.methods.iter().any(|(n, _)| n == "sub"));
    // No __init__ means constructor arity = 0
    let ctor = chunks.iter().find(|c| c.name == "calculator").unwrap();
    assert_eq!(ctor.arity, 0);
}
