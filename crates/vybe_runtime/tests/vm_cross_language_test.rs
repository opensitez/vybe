use std::sync::Arc;
use vybe_runtime::value::Object;
/// Cross-language compatibility tests.
/// Verify that objects created by different language compilers are interoperable.
use vybe_runtime::*;

fn assert_wasm_true(value: &Value, message: &str) {
    assert!(
        matches!(value, Value::Bool(true) | Value::I32(1)),
        "{message}"
    );
}

/// Simulate what Python `{"name": "Rex", "age": 3}` compiles to
fn make_python_dict() -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("name".into(), Value::String(Arc::from("Rex")));
    obj.properties.insert("age".into(), Value::I32(3));
    Value::Object(Arc::new(std::sync::Mutex::new(obj)))
}

/// Simulate what JS `{name: "Rex", age: 3}` compiles to
fn make_js_object() -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("name".into(), Value::String(Arc::from("Rex")));
    obj.properties.insert("age".into(), Value::I32(3));
    Value::Object(Arc::new(std::sync::Mutex::new(obj)))
}

/// Simulate what VB `New With {.Name = "Rex", .Age = 3}` compiles to
#[allow(dead_code)]
fn make_vb_object() -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("name".into(), Value::String(Arc::from("Rex")));
    obj.properties.insert("age".into(), Value::I32(3));
    Value::Object(Arc::new(std::sync::Mutex::new(obj)))
}

#[test]
fn python_dict_equals_js_object() {
    let py = make_python_dict();
    let js = make_js_object();
    // Both should have identical structure
    if let (Value::Object(a), Value::Object(b)) = (&py, &js) {
        let a = a.lock().unwrap();
        let b = b.lock().unwrap();
        assert_eq!(
            a.properties.get("name").unwrap().to_string(),
            b.properties.get("name").unwrap().to_string()
        );
        assert_eq!(
            a.properties.get("age").unwrap().as_i32(),
            b.properties.get("age").unwrap().as_i32()
        );
        assert_eq!(a.properties.len(), b.properties.len());
    }
}

#[test]
fn js_can_read_python_dict_fields() {
    // Python creates: {"name": "Rex", "age": 3}
    // JS reads: obj.name, obj.age (via struct_get)
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 3;

    // Create dict like Python does (struct_new + struct_set)
    chunk.emit_struct_new(0, 0, 0);
    chunk.emit_dup(0);
    chunk.emit_string_const("Rex", 0);
    let name_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, name_key, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // Read like JS does (struct_get)
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let get_name = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, get_name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.to_string(), "Rex");
}

#[test]
fn python_can_read_js_object_fields() {
    // JS creates: {name: "Rex"} — same bytecode as Python
    // Python reads: d["name"] — same struct_get
    // This is the same test — they produce identical bytecode
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 3;

    // JS-style object creation
    chunk.emit_struct_new(0, 0, 0);
    chunk.emit_dup(0);
    chunk.emit_string_const("hello", 0);
    let key = chunk.add_constant(Value::String(Arc::from("msg")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // Python-style dict access
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let get_key = chunk.add_constant(Value::String(Arc::from("msg")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, get_key, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.to_string(), "hello");
}

#[test]
fn numeric_equality_across_types() {
    // Python I32(42) == JS F64(42.0) == VB I32(42)
    assert!(Value::I32(42).eq(&Value::F64(42.0)));
    assert!(Value::F64(42.0).eq(&Value::I32(42)));
    assert!(Value::I32(0).eq(&Value::F64(0.0)));
    assert!(Value::I64(100).eq(&Value::F64(100.0)));
    assert!(Value::I32(100).eq(&Value::I64(100)));
}

// ── CLS Case Resolution ─────────────────────────────────────

#[test]
fn cls_case_resolution_at_link_time() {
    use std::collections::HashMap;
    use vybe_runtime::chunk::TypeEntry;
    use vybe_runtime::component::*;

    // C# component defines Dog with PascalCase field "Name"
    let mut cs_script = Chunk::new("<script>");
    cs_script.local_count = 1;
    cs_script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    cs_script.types.push(TypeEntry {
        name: "Dog".to_string(),
        kind: vybe_runtime::chunk::CompositeKind::Struct,
        parent_index: 0,
        fields: vec!["Name".to_string(), "Breed".to_string()],
        methods: vec![("Bark".to_string(), 0)],
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new(),
    });
    // Add "Name" as a constant (simulates C# accessing obj.Name)
    cs_script.add_constant(Value::String(Arc::from("Name")));

    let cs_comp = Component {
        name: "csharp-app".into(),
        language: Language::CSharp,
        chunks: vec![cs_script],
        imports: vec![],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![],
    };

    // VB component accesses Dog with lowercase "name" (VB convention)
    let mut vb_script = Chunk::new("<script>");
    vb_script.local_count = 1;
    vb_script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    // VB's constant pool has lowercased "name" from VB compiler
    let _name_idx = vb_script.add_constant(Value::String(Arc::from("name")));
    let _breed_idx = vb_script.add_constant(Value::String(Arc::from("breed")));
    let _bark_idx = vb_script.add_constant(Value::String(Arc::from("bark")));

    let vb_comp = Component {
        name: "vb-app".into(),
        language: Language::VB,
        chunks: vec![vb_script],
        imports: vec![],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![],
    };

    let mut linker = Linker::new();
    linker.add_component(cs_comp);
    linker.add_component(vb_comp);

    let result = linker.link().expect("link failed");

    // The VB chunk's constants should be rewritten to match C# casing
    let vb_chunk = &result.chunks[1]; // VB is second component
    let name_val = &vb_chunk.constants[0];
    let breed_val = &vb_chunk.constants[1];
    let bark_val = &vb_chunk.constants[2];

    assert_eq!(
        name_val.to_string(),
        "Name",
        "VB 'name' should be rewritten to 'Name'"
    );
    assert_eq!(
        breed_val.to_string(),
        "Breed",
        "VB 'breed' should be rewritten to 'Breed'"
    );
    assert_eq!(
        bark_val.to_string(),
        "Bark",
        "VB 'bark' should be rewritten to 'Bark'"
    );
}

#[test]
fn cls_case_preserves_case_sensitive_languages() {
    use std::collections::HashMap;
    use vybe_runtime::component::*;

    // JS component with camelCase "firstName"
    let mut js_script = Chunk::new("<script>");
    js_script.local_count = 1;
    js_script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    js_script.add_constant(Value::String(Arc::from("firstName")));

    let js_comp = Component {
        name: "js-app".into(),
        language: Language::JS,
        chunks: vec![js_script],
        imports: vec![],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![],
    };

    // Python component with snake_case "first_name"
    let mut py_script = Chunk::new("<script>");
    py_script.local_count = 1;
    py_script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    py_script.add_constant(Value::String(Arc::from("first_name")));

    let py_comp = Component {
        name: "py-app".into(),
        language: Language::Python,
        chunks: vec![py_script],
        imports: vec![],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![],
    };

    let mut linker = Linker::new();
    linker.add_component(js_comp);
    linker.add_component(py_comp);

    let result = linker.link().expect("link failed");

    // Case-sensitive languages should NOT be rewritten
    assert_eq!(result.chunks[0].constants[0].to_string(), "firstName");
    assert_eq!(result.chunks[1].constants[0].to_string(), "first_name");
}

#[test]
fn string_identity_is_shared() {
    // Same Arc<str> across components = pointer equality
    let s = Value::String(Arc::from("shared"));
    let s2 = s.clone(); // same Rc
    match (&s, &s2) {
        (Value::String(a), Value::String(b)) => assert!(Arc::ptr_eq(a, b)),
        _ => panic!(),
    }
}
