use vybe_parser_python::parse;
use vybe_compiler_python::Compiler;
use vybe_bytecode::{VM, Value};
use std::rc::Rc;
use std::cell::RefCell;

fn compile(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    c.compile(&module).expect("compile failed")
}

fn run(src: &str) -> Value {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    vm.run(chunks).unwrap()
}

fn run_prints(src: &str) -> Vec<String> {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    // Override log AFTER register_all to capture output
    let output = Rc::new(RefCell::new(Vec::<String>::new()));
    let out = output.clone();
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
        out.borrow_mut().push(s.join(" "));
        Value::Null
    }));
    vm.run(chunks).unwrap();
    output.borrow().clone()
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
    assert_eq!(dog.parent, "Animal");
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

// ── Runtime class tests (compile + execute) ─────────────────

#[test]
fn class_instance_has_properties() {
    let out = run_prints(r#"
class Dog:
    def __init__(self, name, age):
        self.name = name
        self.age = age
d = Dog("Rex", 3)
print(d.name)
print(d.age)
"#);
    assert_eq!(out[0], "Rex");
    assert_eq!(out[1], "3");
}

#[test]
fn class_method_returns_value() {
    let out = run_prints(r#"
class Dog:
    def __init__(self, name):
        self.name = name
    def bark(self):
        return self.name
d = Dog("Rex")
print(d.bark())
"#);
    assert_eq!(out[0], "Rex");
}

#[test]
fn class_attribute() {
    let out = run_prints(r#"
class Dog:
    species = "canine"
    def __init__(self, name):
        self.name = name
d = Dog("Rex")
print(d.species)
"#);
    assert_eq!(out[0], "canine");
}

#[test]
fn class_multiple_instances() {
    let out = run_prints(r#"
class Counter:
    def __init__(self):
        self.count = 0
    def inc(self):
        self.count = self.count + 1
    def get(self):
        return self.count
a = Counter()
b = Counter()
a.inc()
a.inc()
b.inc()
print(a.get())
print(b.get())
"#);
    assert_eq!(out[0], "2");
    assert_eq!(out[1], "1");
}

#[test]
fn class_method_modifies_state() {
    let out = run_prints(r#"
class Stack:
    def __init__(self):
        self.items = []
    def push(self, item):
        self.items.append(item)
    def size(self):
        return len(self.items)
s = Stack()
s.push(1)
s.push(2)
s.push(3)
print(s.size())
"#);
    assert_eq!(out[0], "3");
}

#[test]
fn class_property_decorator() {
    let chunks = compile(r#"
class Circle:
    def __init__(self, r):
        self.r = r
    @property
    def radius(self):
        return self.r
"#);
    let types = &chunks[0].types;
    let circle = types.iter().find(|t| t.name == "circle").unwrap();
    assert!(circle.methods.iter().any(|(n, _)| n == "__get_radius"));
}

#[test]
fn class_dunder_str_alias() {
    let chunks = compile(r#"
class Dog:
    def __str__(self):
        return "dog"
"#);
    // Constructor should set both __str__ and toString on the object
    let ctor = chunks.iter().find(|c| c.name == "dog").unwrap();
    let has_tostring = ctor.constants.iter().any(|c| {
        if let Value::String(s) = c { s.as_ref() == "toString" } else { false }
    });
    assert!(has_tostring, "constructor should alias __str__ as toString");
}

#[test]
fn class_exception_types_registered() {
    let chunks = compile("x = 1\n");
    let types = &chunks[0].types;
    assert!(types.iter().any(|t| t.name == "valueerror"), "ValueError should be registered");
    assert!(types.iter().any(|t| t.name == "typeerror"), "TypeError should be registered");
    assert!(types.iter().any(|t| t.name == "keyerror"), "KeyError should be registered");
}

#[test]
#[ignore] // Known: typed except handler (try_table) doesn't preserve thrown object properties
fn class_exception_constructor() {
    let out = run_prints(r#"
try:
    raise ValueError("bad input")
except ValueError as e:
    print(e.message)
"#);
    assert_eq!(out[0], "bad input");
}

#[test]
fn class_method_self_access() {
    let out = run_prints(r#"
class Person:
    def __init__(self, first, last):
        self.first = first
        self.last = last
    def full_name(self):
        return self.first + " " + self.last
p = Person("John", "Doe")
print(p.full_name())
"#);
    assert_eq!(out[0], "John Doe");
}

#[test]
fn class_chained_method_calls() {
    let out = run_prints(r#"
class Builder:
    def __init__(self):
        self.parts = []
    def add(self, part):
        self.parts.append(part)
        return self
    def build(self):
        return len(self.parts)
b = Builder()
b.add("a")
b.add("b")
print(b.build())
"#);
    assert_eq!(out[0], "2");
}
