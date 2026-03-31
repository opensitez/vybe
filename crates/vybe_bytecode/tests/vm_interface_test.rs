/// Tests for cross-language interface sharing and constructor chaining.

use vybe_bytecode::{VM, Value, Chunk, Op, TypeDef};
use vybe_bytecode::typedef::Method;
use vybe_bytecode::chunk::TypeEntry;
use std::rc::Rc;

// ── Interface Registration ──────────────────────────────────

#[test]
fn register_interface() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let iface_id = reg.register_interface("IAnimal", &[("speak", 1), ("move", 1)]);
    assert!(iface_id > 0);
    let td = reg.get(iface_id).unwrap();
    assert!(td.is_interface);
    assert_eq!(td.required_methods.len(), 2);
}

#[test]
fn type_implements_interface() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let iface_id = reg.register_interface("IAnimal", &[("speak", 1)]);

    let mut dog = TypeDef::new("Dog");
    dog.methods.insert("speak".into(), Method::ChunkFn(0));
    let dog_id = reg.register(dog);

    reg.add_implements(dog_id, iface_id);

    // is_subtype should return true for interface implementation
    assert!(reg.is_subtype(dog_id, iface_id));
}

#[test]
fn type_does_not_implement_interface() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let iface_id = reg.register_interface("IFlyable", &[("fly", 1)]);

    let dog_id = reg.register(TypeDef::new("Dog"));

    // Dog doesn't implement IFlyable
    assert!(!reg.is_subtype(dog_id, iface_id));
}

#[test]
fn satisfies_interface_check() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let iface_id = reg.register_interface("IDrawable", &[("draw", 1), ("resize", 2)]);

    let mut widget = TypeDef::new("Widget");
    widget.methods.insert("draw".into(), Method::ChunkFn(0));
    widget.methods.insert("resize".into(), Method::ChunkFn(1));
    let widget_id = reg.register(widget);

    assert!(reg.satisfies_interface(widget_id, iface_id));
}

#[test]
fn does_not_satisfy_missing_method() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let iface_id = reg.register_interface("ISerializable", &[("serialize", 1), ("deserialize", 1)]);

    let mut partial = TypeDef::new("Partial");
    partial.methods.insert("serialize".into(), Method::ChunkFn(0));
    // Missing "deserialize"
    let partial_id = reg.register(partial);

    assert!(!reg.satisfies_interface(partial_id, iface_id));
}

#[test]
fn inherited_interface_implementation() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let iface_id = reg.register_interface("IAnimal", &[("speak", 1)]);

    let mut animal = TypeDef::new("Animal");
    animal.methods.insert("speak".into(), Method::ChunkFn(0));
    let animal_id = reg.register(animal);
    reg.add_implements(animal_id, iface_id);

    // Dog inherits from Animal
    let mut dog = TypeDef::new("Dog");
    dog.parent = Some(animal_id);
    let dog_id = reg.register(dog);

    // Dog should satisfy IAnimal through inheritance
    assert!(reg.is_subtype(dog_id, iface_id));
}

#[test]
fn multiple_interfaces() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let iface_a = reg.register_interface("IReadable", &[("read", 1)]);
    let iface_b = reg.register_interface("IWritable", &[("write", 1)]);

    let mut stream = TypeDef::new("Stream");
    stream.methods.insert("read".into(), Method::ChunkFn(0));
    stream.methods.insert("write".into(), Method::ChunkFn(1));
    let stream_id = reg.register(stream);
    reg.add_implements(stream_id, iface_a);
    reg.add_implements(stream_id, iface_b);

    assert!(reg.is_subtype(stream_id, iface_a));
    assert!(reg.is_subtype(stream_id, iface_b));
}

// ── Constructor Chaining ────────────────────────────────────

#[test]
fn constructor_registered_in_type() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let mut td = TypeDef::new("Foo");
    td.constructor = Some(Method::ChunkFn(5));
    let id = reg.register(td);

    assert!(matches!(reg.get_constructor(id), Some(Method::ChunkFn(5))));
}

#[test]
fn resolve_constructor_walks_parent() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();

    // Parent has constructor
    let mut parent = TypeDef::new("Base");
    parent.constructor = Some(Method::ChunkFn(10));
    let parent_id = reg.register(parent);

    // Child has no constructor
    let mut child = TypeDef::new("Child");
    child.parent = Some(parent_id);
    let child_id = reg.register(child);

    // resolve_constructor should find parent's constructor
    assert!(matches!(reg.resolve_constructor(child_id), Some(Method::ChunkFn(10))));
}

#[test]
fn child_constructor_overrides_parent() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();

    let mut parent = TypeDef::new("Base");
    parent.constructor = Some(Method::ChunkFn(10));
    let parent_id = reg.register(parent);

    let mut child = TypeDef::new("Child");
    child.parent = Some(parent_id);
    child.constructor = Some(Method::ChunkFn(20));
    let child_id = reg.register(child);

    // Child's own constructor takes precedence
    assert!(matches!(reg.resolve_constructor(child_id), Some(Method::ChunkFn(20))));
}

// ── load_type_table with interfaces ─────────────────────────

#[test]
fn load_type_table_interface() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();

    let entries = vec![
        TypeEntry {
            name: "ianimal".into(),
            parent: String::new(),
            fields: Vec::new(),
            methods: vec![("speak".into(), 0)],
            is_interface: true,
            implements: Vec::new(),
            constructor_chunk: None,
        },
        TypeEntry {
            name: "dog".into(),
            parent: String::new(),
            fields: vec!["name".into()],
            methods: vec![("speak".into(), 1), ("fetch".into(), 2)],
            is_interface: false,
            implements: vec!["ianimal".into()],
            constructor_chunk: Some(3),
        },
    ];

    reg.load_type_table(&entries);

    let iface_id = reg.get_id("ianimal").expect("interface not registered");
    let dog_id = reg.get_id("dog").expect("dog not registered");

    assert!(reg.get(iface_id).unwrap().is_interface);
    assert!(reg.is_subtype(dog_id, iface_id));
    assert!(matches!(reg.get_constructor(dog_id), Some(Method::ChunkFn(3))));
}

#[test]
fn load_type_table_cross_language_inheritance() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();

    // Simulate: VB defines Animal, C# defines Dog : Animal
    let vb_types = vec![TypeEntry {
        name: "animal".into(),
        parent: String::new(),
        fields: vec!["name".into(), "species".into()],
        methods: vec![("speak".into(), 0)],
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: Some(1),
    }];

    let cs_types = vec![TypeEntry {
        name: "dog".into(),
        parent: "animal".into(),
        fields: vec!["breed".into()],
        methods: vec![("fetch".into(), 5), ("speak".into(), 6)],
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: Some(7),
    }];

    // Load VB types first, then C# types
    reg.load_type_table(&vb_types);
    reg.load_type_table(&cs_types);

    let animal_id = reg.get_id("animal").unwrap();
    let dog_id = reg.get_id("dog").unwrap();

    // Dog inherits from Animal
    assert!(reg.is_subtype(dog_id, animal_id));

    // Dog has its own fields + inherits Animal's
    let dog_td = reg.get(dog_id).unwrap();
    assert!(dog_td.field_index("breed").is_some());
    assert_eq!(dog_td.parent, Some(animal_id));

    // Dog overrides speak
    assert!(matches!(reg.resolve_method(dog_id, "speak"), Some(Method::ChunkFn(6))));

    // Dog inherits... but Animal's speak is overridden, so check a non-overridden case
    let animal_td = reg.get(animal_id).unwrap();
    assert!(matches!(animal_td.constructor.as_ref(), Some(Method::ChunkFn(1))));
}

// ── ref_test with interfaces in VM ──────────────────────────

#[test]
fn ref_test_interface_in_vm() {
    let mut vm = VM::new();

    // Register interface
    let iface_id = vm.type_registry.register_interface("IAnimal", &[("speak", 1)]);

    // Register Dog implementing IAnimal
    let mut dog_td = TypeDef::new("Dog");
    dog_td.methods.insert("speak".into(), Method::HostFn(0));
    let dog_id = vm.type_registry.register(dog_td);
    vm.type_registry.add_implements(dog_id, iface_id);

    // Create a Dog object and test if it's an IAnimal
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create typed Dog object
    let tid = chunk.add_constant(Value::I32(dog_id as i32));
    chunk.emit_op_u16(Op::r#const, tid, 0);
    chunk.emit_op(Op::shared_new, 0);

    // ref_test against "ianimal"
    let type_name = chunk.add_constant(Value::String(Rc::from("ianimal")));
    chunk.emit_op_u16(Op::ref_test, type_name, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Bool(true)), "Dog should pass ref_test for IAnimal");
}
