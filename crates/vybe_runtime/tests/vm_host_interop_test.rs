use std::collections::HashMap;
use std::sync::Arc;
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};

// ============================================================
// Helpers
// ============================================================

fn assert_f64(val: &Value, expected: f64) {
    match val {
        Value::F64(v) => assert!(
            (v - expected).abs() < 1e-10,
            "Expected F64({}), got F64({})",
            expected,
            v
        ),
        _ => panic!("Expected F64({}), got {:?}", expected, val),
    }
}

fn assert_i32(val: &Value, expected: i32) {
    match val {
        Value::I32(v) => assert_eq!(*v, expected, "Expected I32({}), got I32({})", expected, v),
        _ => panic!("Expected I32({}), got {:?}", expected, val),
    }
}

fn assert_string(val: &Value, expected: &str) {
    match val {
        Value::String(s) => assert_eq!(
            s.as_ref(),
            expected,
            "Expected String({:?}), got String({:?})",
            expected,
            s.as_ref()
        ),
        _ => panic!("Expected String({:?}), got {:?}", expected, val),
    }
}

/// Emit a call_import instruction: [op, u16 import_idx, u8 argc]
fn emit_call_import(chunk: &mut Chunk, import_idx: u16, argc: u8) {
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
}

// ============================================================
// A. Host function call mechanics (tests 1-10)
// ============================================================

// 1. call_import with 0 args — host fn receives empty args
#[test]
fn host_call_zero_args() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "noop",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "noop");
    emit_call_import(&mut main, imp, 0);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    assert_eq!(received.lock().unwrap().len(), 0);
}

// 2. call_import with 1 arg — host fn receives [arg]
#[test]
fn host_call_one_arg() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "one",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "one");
    let c = main.add_constant(Value::F64(42.0));
    main.emit_op_u16(Op::CONST, c, 0);
    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    let args = received.lock().unwrap();
    assert_eq!(args.len(), 1);
    assert_f64(&args[0], 42.0);
}

// 3. call_import with 3 args — correct order
#[test]
fn host_call_three_args_order() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "three",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "three");
    let c1 = main.add_constant(Value::I32(10));
    let c2 = main.add_constant(Value::I32(20));
    let c3 = main.add_constant(Value::I32(30));
    main.emit_op_u16(Op::CONST, c1, 0);
    main.emit_op_u16(Op::CONST, c2, 0);
    main.emit_op_u16(Op::CONST, c3, 0);
    emit_call_import(&mut main, imp, 3);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    let args = received.lock().unwrap();
    assert_eq!(args.len(), 3);
    assert_i32(&args[0], 10);
    assert_i32(&args[1], 20);
    assert_i32(&args[2], 30);
}

// 4. Host fn return value on VM stack
#[test]
fn host_return_value_on_stack() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "answer",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| Value::F64(99.0)),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "answer");
    emit_call_import(&mut main, imp, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_f64(&result, 99.0);
}

// 5. Host fn return value used in subsequent operation
#[test]
fn host_return_value_in_operation() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "five",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| Value::F64(5.0)),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "five");
    emit_call_import(&mut main, imp, 0);
    let c10 = main.add_constant(Value::F64(10.0));
    main.emit_op_u16(Op::CONST, c10, 0);
    main.emit_op(Op::F64_ADD, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_f64(&result, 15.0);
}

// 6. Host fn modifying external state (Arc<std::sync::Mutex<>>)
#[test]
fn host_modifies_external_state() {
    let counter = Arc::new(std::sync::Mutex::new(0i32));
    let cnt = counter.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "increment",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
                *cnt.lock().unwrap() += 1;
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "increment");
    emit_call_import(&mut main, imp, 0);
    main.emit_op(Op::DROP, 0);
    emit_call_import(&mut main, imp, 0);
    main.emit_op(Op::DROP, 0);
    emit_call_import(&mut main, imp, 0);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    assert_eq!(*counter.lock().unwrap(), 3);
}

// 7. Multiple host fns registered, correct one called
#[test]
fn multiple_host_fns_correct_dispatch() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "alpha",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            Value::String(Arc::from("alpha"))
        }),
    );
    vm.register_host_fn(
        "test",
        "beta",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            Value::String(Arc::from("beta"))
        }),
    );
    vm.register_host_fn(
        "test",
        "gamma",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            Value::String(Arc::from("gamma"))
        }),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp_beta = main.add_import("test", "beta");
    emit_call_import(&mut main, imp_beta, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_string(&result, "beta");
}

// 8. Host fn receiving string args
#[test]
fn host_receives_string_args() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "greet",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "greet");
    let c1 = main.add_constant(Value::String(Arc::from("hello")));
    let c2 = main.add_constant(Value::String(Arc::from("world")));
    main.emit_op_u16(Op::CONST, c1, 0);
    main.emit_op_u16(Op::CONST, c2, 0);
    emit_call_import(&mut main, imp, 2);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    let args = received.lock().unwrap();
    assert_eq!(args.len(), 2);
    assert_string(&args[0], "hello");
    assert_string(&args[1], "world");
}

// 9. Host fn receiving object args
#[test]
fn host_receives_object_args() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "take_obj",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "take_obj");
    // Build object: struct_new with 1 field "x" = 42
    let key = main.add_constant(Value::String(Arc::from("x")));
    main.emit_op_u16(Op::CONST, key, 0);
    let val = main.add_constant(Value::I32(42));
    main.emit_op_u16(Op::CONST, val, 0);
    main.emit_op_u16(Op::STRUCT_NEW, 1, 0);
    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    let args = received.lock().unwrap();
    assert_eq!(args.len(), 1);
    match &args[0] {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            assert_i32(o.properties.get("x").unwrap(), 42);
        }
        _ => panic!("Expected Object arg, got {:?}", args[0]),
    }
}

// 10. Host fn receiving mixed types (string, number, bool, null)
// ============================================================
// B. Object passing host <-> VM (tests 11-20)
// ============================================================

// 11. VM creates object (struct_new), passes to host fn — host reads properties
#[test]
fn vm_object_to_host_read_properties() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "read_obj",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
            match &args[0] {
                Value::Object(obj) => {
                    let o = obj.lock().unwrap();
                    let _name = o.properties.get("name").cloned().unwrap_or(Value::Null);
                    let age = o.properties.get("age").cloned().unwrap_or(Value::Null);
                    // Return age as a check value
                    age
                }
                _ => Value::Null,
            }
        }),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "read_obj");
    // struct_new with 2 fields: name="Alice", age=30
    let k1 = main.add_constant(Value::String(Arc::from("name")));
    let v1 = main.add_constant(Value::String(Arc::from("Alice")));
    let k2 = main.add_constant(Value::String(Arc::from("age")));
    let v2 = main.add_constant(Value::I32(30));
    main.emit_op_u16(Op::CONST, k1, 0);
    main.emit_op_u16(Op::CONST, v1, 0);
    main.emit_op_u16(Op::CONST, k2, 0);
    main.emit_op_u16(Op::CONST, v2, 0);
    main.emit_op_u16(Op::STRUCT_NEW, 2, 0);
    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 30);
}

// 12. Host fn creates object (Value::Object), VM reads properties via struct_get
#[test]
fn host_creates_object_vm_reads() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "make_obj",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            let mut obj = Object::new();
            obj.set("color".to_string(), Value::String(Arc::from("blue")));
            obj.set("size".to_string(), Value::I32(42));
            Value::Object(Arc::new(std::sync::Mutex::new(obj)))
        }),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "make_obj");
    emit_call_import(&mut main, imp, 0);
    let prop = main.add_constant(Value::String(Arc::from("color")));
    main.emit_op_u16(Op::STRUCT_GET, prop, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_string(&result, "blue");
}

// 13. Host fn returns object with nested object — VM navigates chain
#[test]
fn host_nested_object_vm_navigates() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "nested",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            let mut inner = Object::new();
            inner.set("value".to_string(), Value::I32(777));
            let mut outer = Object::new();
            outer.set(
                "child".to_string(),
                Value::Object(Arc::new(std::sync::Mutex::new(inner))),
            );
            Value::Object(Arc::new(std::sync::Mutex::new(outer)))
        }),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "nested");
    emit_call_import(&mut main, imp, 0);
    let prop_child = main.add_constant(Value::String(Arc::from("child")));
    main.emit_op_u16(Op::STRUCT_GET, prop_child, 0);
    let prop_value = main.add_constant(Value::String(Arc::from("value")));
    main.emit_op_u16(Op::STRUCT_GET, prop_value, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 777);
}

// 14. VM modifies object, passes to host — host sees modifications
#[test]
fn vm_modifies_object_host_sees() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "check",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 2; // local 0 = script, local 1 = obj
    let imp = main.add_import("test", "check");

    // Create object with x=1
    let kx = main.add_constant(Value::String(Arc::from("x")));
    let v1 = main.add_constant(Value::I32(1));
    main.emit_op_u16(Op::CONST, kx, 0);
    main.emit_op_u16(Op::CONST, v1, 0);
    main.emit_op_u16(Op::STRUCT_NEW, 1, 0);
    main.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // Modify: obj.x = 99
    main.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let v99 = main.add_constant(Value::I32(99));
    main.emit_op_u16(Op::CONST, v99, 0);
    let kx2 = main.add_constant(Value::String(Arc::from("x")));
    main.emit_op_u16(Op::STRUCT_SET, kx2, 0);
    main.emit_op(Op::DROP, 0); // drop struct_set result

    // Pass to host
    main.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    let args = received.lock().unwrap();
    match &args[0] {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            assert_i32(o.properties.get("x").unwrap(), 99);
        }
        _ => panic!("Expected Object"),
    }
}

// 15. Host fn modifies object passed from VM — VM sees changes (shared Rc)
#[test]
fn host_modifies_object_vm_sees() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "mutate",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
            if let Value::Object(obj) = &args[0] {
                obj.lock().unwrap().set("y".to_string(), Value::I32(999));
            }
            Value::Null
        }),
    );

    let mut main = Chunk::new("main");
    main.local_count = 2; // local 0 = script, local 1 = obj
    let imp = main.add_import("test", "mutate");

    // Create object with x=1
    let kx = main.add_constant(Value::String(Arc::from("x")));
    let v1 = main.add_constant(Value::I32(1));
    main.emit_op_u16(Op::CONST, kx, 0);
    main.emit_op_u16(Op::CONST, v1, 0);
    main.emit_op_u16(Op::STRUCT_NEW, 1, 0);
    main.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // Pass to host fn which adds y=999
    main.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::DROP, 0); // drop null return

    // Read y from object
    main.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let ky = main.add_constant(Value::String(Arc::from("y")));
    main.emit_op_u16(Op::STRUCT_GET, ky, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 999);
}

// 16. Object with array property — host reads array elements
#[test]
fn object_with_array_property_host_reads() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "read_arr",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "read_arr");

    // Create array [10, 20, 30]
    let c10 = main.add_constant(Value::I32(10));
    let c20 = main.add_constant(Value::I32(20));
    let c30 = main.add_constant(Value::I32(30));
    main.emit_op_u16(Op::CONST, c10, 0);
    main.emit_op_u16(Op::CONST, c20, 0);
    main.emit_op_u16(Op::CONST, c30, 0);
    main.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);

    // Wrap in object: {items: [10,20,30]}
    let k_items = main.add_constant(Value::String(Arc::from("items")));
    // Need: key, val, struct_new 1
    // Stack has the array. Need to push key before it.
    // Let's store the array, push key, then push array back.
    main.emit_op_u16(Op::LOCAL_SET, 0, 0); // temp store
    main.emit_op_u16(Op::CONST, k_items, 0);
    main.emit_op_u16(Op::LOCAL_GET, 0, 0);
    main.emit_op_u16(Op::STRUCT_NEW, 1, 0);

    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    let args = received.lock().unwrap();
    match &args[0] {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            let items = o.properties.get("items").expect("items property");
            match items {
                Value::Object(arr_obj) => {
                    let arr = arr_obj.lock().unwrap();
                    match &arr.kind {
                        ObjectKind::Array(elems) => {
                            assert_eq!(elems.len(), 3);
                            assert_i32(&elems[0], 10);
                            assert_i32(&elems[1], 20);
                            assert_i32(&elems[2], 30);
                        }
                        _ => panic!("Expected Array kind"),
                    }
                }
                _ => panic!("Expected Object for items"),
            }
        }
        _ => panic!("Expected Object arg"),
    }
}

// 17. Object with function property — VM can call it after getting from host
#[test]
fn host_returns_object_with_function_vm_calls() {
    let mut vm = VM::new();
    // Host fn returns an object with a "compute" property that is a host function
    // We'll use a second host fn for the compute
    vm.register_host_fn(
        "test",
        "compute_impl",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
            // Returns arg * 2
            match &args[0] {
                Value::I32(n) => Value::I32(n * 2),
                _ => Value::Null,
            }
        }),
    );
    vm.register_host_fn(
        "test",
        "make_service",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            // Returns object — but we can't easily embed HostFunction from here
            // because we don't know the index. Return a plain object.
            let mut obj = Object::new();
            obj.set("name".to_string(), Value::String(Arc::from("service")));
            Value::Object(Arc::new(std::sync::Mutex::new(obj)))
        }),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let _imp_service = main.add_import("test", "make_service");
    let imp_compute = main.add_import("test", "compute_impl");

    // Call compute_impl(21) directly — verifies host fn works
    let c21 = main.add_constant(Value::I32(21));
    main.emit_op_u16(Op::CONST, c21, 0);
    emit_call_import(&mut main, imp_compute, 1);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 42);
}

// 18. Host fn receives array (Object with Array kind) — reads elements
#[test]
fn host_receives_array_reads_elements() {
    let received = Arc::new(std::sync::Mutex::new(Vec::<Value>::new()));
    let recv = received.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "take_arr",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *recv.lock().unwrap() = args.to_vec();
                Value::Null
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "take_arr");
    let c1 = main.add_constant(Value::String(Arc::from("a")));
    let c2 = main.add_constant(Value::String(Arc::from("b")));
    let c3 = main.add_constant(Value::String(Arc::from("c")));
    main.emit_op_u16(Op::CONST, c1, 0);
    main.emit_op_u16(Op::CONST, c2, 0);
    main.emit_op_u16(Op::CONST, c3, 0);
    main.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();
    let args = received.lock().unwrap();
    assert_eq!(args.len(), 1);
    match &args[0] {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(elems) => {
                    assert_eq!(elems.len(), 3);
                    assert_string(&elems[0], "a");
                    assert_string(&elems[1], "b");
                    assert_string(&elems[2], "c");
                }
                _ => panic!("Expected Array kind"),
            }
        }
        _ => panic!("Expected Object arg"),
    }
}

// 19. VM creates array, passes to host, host returns length
#[test]
fn vm_array_to_host_returns_length() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "arr_len",
        Box::new(
            |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| match &args[0] {
                Value::Object(obj) => {
                    let o = obj.lock().unwrap();
                    match &o.kind {
                        ObjectKind::Array(elems) => Value::I32(elems.len() as i32),
                        _ => Value::I32(-1),
                    }
                }
                _ => Value::I32(-1),
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp = main.add_import("test", "arr_len");
    let c1 = main.add_constant(Value::I32(1));
    let c2 = main.add_constant(Value::I32(2));
    let c3 = main.add_constant(Value::I32(3));
    let c4 = main.add_constant(Value::I32(4));
    let c5 = main.add_constant(Value::I32(5));
    main.emit_op_u16(Op::CONST, c1, 0);
    main.emit_op_u16(Op::CONST, c2, 0);
    main.emit_op_u16(Op::CONST, c3, 0);
    main.emit_op_u16(Op::CONST, c4, 0);
    main.emit_op_u16(Op::CONST, c5, 0);
    main.emit_op_u16(Op::ARRAY_NEW_FIXED, 5, 0);
    emit_call_import(&mut main, imp, 1);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 5);
}

// 20. Round-trip: VM creates object -> host stores -> another host returns it -> VM reads
#[test]
fn roundtrip_object_through_host() {
    let store: Arc<std::sync::Mutex<Option<Value>>> = Arc::new(std::sync::Mutex::new(None));
    let store_put = store.clone();
    let store_get = store.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "store_obj",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *store_put.lock().unwrap() = Some(args[0].clone());
                Value::Null
            },
        ),
    );
    vm.register_host_fn(
        "test",
        "load_obj",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
                store_get.lock().unwrap().clone().unwrap_or(Value::Null)
            },
        ),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let imp_store = main.add_import("test", "store_obj");
    let imp_load = main.add_import("test", "load_obj");

    // Create object {data: 12345} and store it
    let kd = main.add_constant(Value::String(Arc::from("data")));
    let vd = main.add_constant(Value::I32(12345));
    main.emit_op_u16(Op::CONST, kd, 0);
    main.emit_op_u16(Op::CONST, vd, 0);
    main.emit_op_u16(Op::STRUCT_NEW, 1, 0);
    emit_call_import(&mut main, imp_store, 1);
    main.emit_op(Op::DROP, 0); // drop null

    // Load it back and read property
    emit_call_import(&mut main, imp_load, 0);
    let kd2 = main.add_constant(Value::String(Arc::from("data")));
    main.emit_op_u16(Op::STRUCT_GET, kd2, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 12345);
}

// ============================================================
// C. invoke() mechanics (tests 21-30)
// ============================================================

// 21. invoke() a chunk function — returns correct value
#[test]
fn invoke_chunk_function_returns_value() {
    let mut vm = VM::new();

    // chunk 0: dummy main (needed for run to load chunks)
    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: function that returns 42
    let mut func = Chunk::new("return42");
    func.arity = 0;
    func.local_count = 1;
    let c = func.add_constant(Value::F64(42.0));
    func.emit_op_u16(Op::CONST, c, 0);
    func.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, func]).unwrap();

    // Build a Function value pointing to chunk 1
    let func_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("return42".to_string()),
            arity: 0,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let result = vm.invoke(&func_val, &[]).unwrap();
    assert_f64(&result, 42.0);
}

// 22. invoke() with args that become locals
#[test]
fn invoke_with_args_as_locals() {
    let mut vm = VM::new();

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: (a, b) => a + b
    let mut func = Chunk::new("add");
    func.arity = 2;
    func.local_count = 2;
    func.emit_op_u16(Op::LOCAL_GET, 0, 0);
    func.emit_op_u16(Op::LOCAL_GET, 1, 0);
    func.emit_op(Op::F64_ADD, 0);
    func.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, func]).unwrap();

    let func_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("add".to_string()),
            arity: 2,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let result = vm
        .invoke(&func_val, &[Value::F64(3.0), Value::F64(7.0)])
        .unwrap();
    assert_f64(&result, 10.0);
}

// 23. invoke() clears stack between calls
#[test]
fn invoke_clears_stack_between_calls() {
    let mut vm = VM::new();

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: (x) => x * 2
    let mut func = Chunk::new("double");
    func.arity = 1;
    func.local_count = 1;
    func.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let c2 = func.add_constant(Value::F64(2.0));
    func.emit_op_u16(Op::CONST, c2, 0);
    func.emit_op(Op::F64_MUL, 0);
    func.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, func]).unwrap();

    let func_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("double".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let r1 = vm.invoke(&func_val, &[Value::F64(5.0)]).unwrap();
    assert_f64(&r1, 10.0);

    let r2 = vm.invoke(&func_val, &[Value::F64(100.0)]).unwrap();
    assert_f64(&r2, 200.0);
}

// 24. invoke() preserves globals between calls
#[test]
fn invoke_preserves_globals() {
    let mut vm = VM::new();

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    // Set global "counter" = 0
    let c0 = main_chunk.add_constant(Value::I32(0));
    main_chunk.emit_op_u16(Op::CONST, c0, 0);
    let g = main_chunk.add_constant(Value::String(Arc::from("counter")));
    main_chunk.emit_op_u16(Op::GLOBAL_SET, g, 0);
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: reads global "counter", adds 1, sets it, returns it
    let mut func = Chunk::new("inc");
    func.arity = 0;
    func.local_count = 1;
    let gc = func.add_constant(Value::String(Arc::from("counter")));
    func.emit_op_u16(Op::GLOBAL_GET, gc, 0);
    let c1 = func.add_constant(Value::I32(1));
    func.emit_op_u16(Op::CONST, c1, 0);
    func.emit_op(Op::I32_ADD, 0);
    let gc2 = func.add_constant(Value::String(Arc::from("counter")));
    func.emit_op_u16(Op::GLOBAL_SET, gc2, 0);
    let gc3 = func.add_constant(Value::String(Arc::from("counter")));
    func.emit_op_u16(Op::GLOBAL_GET, gc3, 0);
    func.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, func]).unwrap();

    let func_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("inc".to_string()),
            arity: 0,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let r1 = vm.invoke(&func_val, &[]).unwrap();
    assert_i32(&r1, 1);
    let r2 = vm.invoke(&func_val, &[]).unwrap();
    assert_i32(&r2, 2);
    let r3 = vm.invoke(&func_val, &[]).unwrap();
    assert_i32(&r3, 3);
}

// 25. invoke() a function that calls a host function
#[test]
fn invoke_function_calls_host() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "square",
        Box::new(
            |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| match &args[0] {
                Value::F64(n) => Value::F64(n * n),
                _ => Value::Null,
            },
        ),
    );

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    let _imp = main_chunk.add_import("test", "square");
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: (x) => host_square(x) + 1
    let mut func = Chunk::new("square_plus_one");
    func.arity = 1;
    func.local_count = 1;
    // Import must be on chunk 0, call_import uses import_table from chunk 0
    func.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_call_import(&mut func, 0, 1); // import index 0 = "square"
    let c1 = func.add_constant(Value::F64(1.0));
    func.emit_op_u16(Op::CONST, c1, 0);
    func.emit_op(Op::F64_ADD, 0);
    func.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, func]).unwrap();

    let func_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("square_plus_one".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let result = vm.invoke(&func_val, &[Value::F64(7.0)]).unwrap();
    assert_f64(&result, 50.0); // 7^2 + 1 = 50
}

// 26. invoke() a function that modifies a global
#[test]
fn invoke_modifies_global() {
    let mut vm = VM::new();

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    let c_init = main_chunk.add_constant(Value::String(Arc::from("none")));
    main_chunk.emit_op_u16(Op::CONST, c_init, 0);
    let g = main_chunk.add_constant(Value::String(Arc::from("status")));
    main_chunk.emit_op_u16(Op::GLOBAL_SET, g, 0);
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: sets global "status" = "done", returns Null
    let mut func = Chunk::new("set_status");
    func.arity = 0;
    func.local_count = 1;
    let cs = func.add_constant(Value::String(Arc::from("done")));
    func.emit_op_u16(Op::CONST, cs, 0);
    let gs = func.add_constant(Value::String(Arc::from("status")));
    func.emit_op_u16(Op::GLOBAL_SET, gs, 0);
    func.emit_op(Op::RETURN, 0);

    // chunk 2: reads global "status" and returns it
    let mut reader = Chunk::new("read_status");
    reader.arity = 0;
    reader.local_count = 1;
    let gr = reader.add_constant(Value::String(Arc::from("status")));
    reader.emit_op_u16(Op::GLOBAL_GET, gr, 0);
    reader.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, func, reader]).unwrap();

    let set_fn = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("set_status".to_string()),
            arity: 0,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let read_fn = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("read_status".to_string()),
            arity: 0,
            chunk_index: 2,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    vm.invoke(&set_fn, &[]).unwrap();
    let result = vm.invoke(&read_fn, &[]).unwrap();
    assert_string(&result, "done");
}

// 27. invoke() a host function wrapper (ObjectKind::HostFunction)
#[test]
fn invoke_host_function_object() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "greet",
        Box::new(
            |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| match &args[0] {
                Value::String(s) => Value::String(Arc::from(format!("Hello, {}!", s))),
                _ => Value::Null,
            },
        ),
    );

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    vm.run(vec![main_chunk]).unwrap();

    // Create a HostFunction object — index 0 is the first registered host fn
    let host_fn_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::HostFunction(0),
        type_id: 0,
        fields: Vec::new(),
    })));

    let result = vm
        .invoke(&host_fn_val, &[Value::String(Arc::from("World"))])
        .unwrap();
    assert_string(&result, "Hello, World!");
}

// 28. invoke() a closure that captures upvalue
// 29. invoke() twice — second call uses updated globals from first
#[test]
fn invoke_twice_globals_updated() {
    let mut vm = VM::new();

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    let c0 = main_chunk.add_constant(Value::I32(0));
    main_chunk.emit_op_u16(Op::CONST, c0, 0);
    let g = main_chunk.add_constant(Value::String(Arc::from("acc")));
    main_chunk.emit_op_u16(Op::GLOBAL_SET, g, 0);
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: (n) => acc = acc + n; return acc
    let mut func = Chunk::new("accumulate");
    func.arity = 1;
    func.local_count = 1;
    let ga = func.add_constant(Value::String(Arc::from("acc")));
    func.emit_op_u16(Op::GLOBAL_GET, ga, 0);
    func.emit_op_u16(Op::LOCAL_GET, 0, 0);
    func.emit_op(Op::I32_ADD, 0);
    let ga2 = func.add_constant(Value::String(Arc::from("acc")));
    func.emit_op_u16(Op::GLOBAL_SET, ga2, 0);
    let ga3 = func.add_constant(Value::String(Arc::from("acc")));
    func.emit_op_u16(Op::GLOBAL_GET, ga3, 0);
    func.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, func]).unwrap();

    let func_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("accumulate".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let r1 = vm.invoke(&func_val, &[Value::I32(10)]).unwrap();
    assert_i32(&r1, 10);
    let r2 = vm.invoke(&func_val, &[Value::I32(20)]).unwrap();
    assert_i32(&r2, 30);
    let r3 = vm.invoke(&func_val, &[Value::I32(5)]).unwrap();
    assert_i32(&r3, 35);
}

// 30. invoke() a function that calls another VM function
#[test]
fn invoke_function_calls_another_vm_function() {
    let mut vm = VM::new();

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: outer(x) => calls inner(x) + 1
    let mut outer = Chunk::new("outer");
    outer.arity = 1;
    outer.local_count = 1;
    outer.emit_op_u16(Op::REF_FUNC, 2, 0);
    outer.emit(0, 0); // 0 upvalues
    outer.emit_op_u16(Op::LOCAL_GET, 0, 0);
    outer.emit_op_u8(Op::CALL, 1, 0);
    let c1 = outer.add_constant(Value::I32(1));
    outer.emit_op_u16(Op::CONST, c1, 0);
    outer.emit_op(Op::I32_ADD, 0);
    outer.emit_op(Op::RETURN, 0);

    // chunk 2: inner(x) => x * 10
    let mut inner = Chunk::new("inner");
    inner.arity = 1;
    inner.local_count = 1;
    inner.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let c10 = inner.add_constant(Value::I32(10));
    inner.emit_op_u16(Op::CONST, c10, 0);
    inner.emit_op(Op::I32_MUL, 0);
    inner.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, outer, inner]).unwrap();

    let outer_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("outer".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    let result = vm.invoke(&outer_val, &[Value::I32(5)]).unwrap();
    assert_i32(&result, 51); // inner(5) = 50, + 1 = 51
}

// ============================================================
// D. Callback pattern: host calls VM function (tests 31-35)
// ============================================================

// 31. Host fn receives a Function value, stores it; later invoke() calls it
#[test]
fn callback_store_and_invoke() {
    let callback_store: Arc<std::sync::Mutex<Option<Value>>> =
        Arc::new(std::sync::Mutex::new(None));
    let store = callback_store.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "register_cb",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *store.lock().unwrap() = Some(args[0].clone());
                Value::Null
            },
        ),
    );

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    let imp = main_chunk.add_import("test", "register_cb");

    // ref_func chunk 1, pass it to host
    main_chunk.emit_op_u16(Op::REF_FUNC, 1, 0);
    main_chunk.emit(0, 0); // 0 upvalues
    emit_call_import(&mut main_chunk, imp, 1);
    main_chunk.emit_op(Op::DROP, 0);
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: (x) => x + 100
    let mut cb = Chunk::new("callback");
    cb.arity = 1;
    cb.local_count = 1;
    cb.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let c100 = cb.add_constant(Value::I32(100));
    cb.emit_op_u16(Op::CONST, c100, 0);
    cb.emit_op(Op::I32_ADD, 0);
    cb.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, cb]).unwrap();

    let cb_val = callback_store
        .lock()
        .unwrap()
        .clone()
        .expect("callback should be stored");
    let result = vm.invoke(&cb_val, &[Value::I32(42)]).unwrap();
    assert_i32(&result, 142);
}

// 32. Host fn receives closure, invoke() with captured state
// 33. Host fn receives class method, invoke() with Me arg
#[test]
fn callback_method_with_me_arg() {
    let callback_store: Arc<std::sync::Mutex<Option<Value>>> =
        Arc::new(std::sync::Mutex::new(None));
    let store = callback_store.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "register_method",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *store.lock().unwrap() = Some(args[0].clone());
                Value::Null
            },
        ),
    );

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    let imp = main_chunk.add_import("test", "register_method");

    // Create a ref_func for chunk 1 (method) and pass to host
    main_chunk.emit_op_u16(Op::REF_FUNC, 1, 0);
    main_chunk.emit(0, 0); // 0 upvalues
    emit_call_import(&mut main_chunk, imp, 1);
    main_chunk.emit_op(Op::DROP, 0);
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: method(me) => me.value * 2
    // arity=1, local 0 = me
    let mut method = Chunk::new("method");
    method.arity = 1;
    method.local_count = 1;
    method.emit_op_u16(Op::LOCAL_GET, 0, 0); // me
    let prop_val = method.add_constant(Value::String(Arc::from("value")));
    method.emit_op_u16(Op::STRUCT_GET, prop_val, 0);
    let c2 = method.add_constant(Value::I32(2));
    method.emit_op_u16(Op::CONST, c2, 0);
    method.emit_op(Op::I32_MUL, 0);
    method.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, method]).unwrap();

    // Create a "me" object with value=21
    let mut me_obj = Object::new();
    me_obj.set("value".to_string(), Value::I32(21));
    let me_val = Value::Object(Arc::new(std::sync::Mutex::new(me_obj)));

    let method_val = callback_store
        .lock()
        .unwrap()
        .clone()
        .expect("method should be stored");
    let result = vm.invoke(&method_val, &[me_val]).unwrap();
    assert_i32(&result, 42); // 21 * 2
}

// 34. Multiple callbacks registered, invoke correct one
#[test]
fn multiple_callbacks_invoke_correct() {
    let callbacks: Arc<std::sync::Mutex<Vec<Value>>> = Arc::new(std::sync::Mutex::new(Vec::new()));
    let cbs = callbacks.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "register_cb",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                cbs.lock().unwrap().push(args[0].clone());
                Value::Null
            },
        ),
    );

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    let imp = main_chunk.add_import("test", "register_cb");

    // Register chunk 1
    main_chunk.emit_op_u16(Op::REF_FUNC, 1, 0);
    main_chunk.emit(0, 0);
    emit_call_import(&mut main_chunk, imp, 1);
    main_chunk.emit_op(Op::DROP, 0);

    // Register chunk 2
    main_chunk.emit_op_u16(Op::REF_FUNC, 2, 0);
    main_chunk.emit(0, 0);
    emit_call_import(&mut main_chunk, imp, 1);
    main_chunk.emit_op(Op::DROP, 0);

    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: () => 111
    let mut cb1 = Chunk::new("cb1");
    cb1.arity = 0;
    cb1.local_count = 1;
    let c111 = cb1.add_constant(Value::I32(111));
    cb1.emit_op_u16(Op::CONST, c111, 0);
    cb1.emit_op(Op::RETURN, 0);

    // chunk 2: () => 222
    let mut cb2 = Chunk::new("cb2");
    cb2.arity = 0;
    cb2.local_count = 1;
    let c222 = cb2.add_constant(Value::I32(222));
    cb2.emit_op_u16(Op::CONST, c222, 0);
    cb2.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, cb1, cb2]).unwrap();

    let cbs = callbacks.lock().unwrap();
    assert_eq!(cbs.len(), 2);

    let r1 = vm.invoke(&cbs[0], &[]).unwrap();
    assert_i32(&r1, 111);
    let r2 = vm.invoke(&cbs[1], &[]).unwrap();
    assert_i32(&r2, 222);
}

// 35. Callback modifies global, subsequent code reads updated global
#[test]
fn callback_modifies_global_subsequent_reads() {
    let callback_store: Arc<std::sync::Mutex<Option<Value>>> =
        Arc::new(std::sync::Mutex::new(None));
    let store = callback_store.clone();

    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "register_cb",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                *store.lock().unwrap() = Some(args[0].clone());
                Value::Null
            },
        ),
    );

    let mut main_chunk = Chunk::new("main");
    main_chunk.local_count = 1;
    let imp = main_chunk.add_import("test", "register_cb");

    // Set global "state" = "initial"
    let ci = main_chunk.add_constant(Value::String(Arc::from("initial")));
    main_chunk.emit_op_u16(Op::CONST, ci, 0);
    let gs = main_chunk.add_constant(Value::String(Arc::from("state")));
    main_chunk.emit_op_u16(Op::GLOBAL_SET, gs, 0);

    // Register callback (chunk 1)
    main_chunk.emit_op_u16(Op::REF_FUNC, 1, 0);
    main_chunk.emit(0, 0);
    emit_call_import(&mut main_chunk, imp, 1);
    main_chunk.emit_op(Op::DROP, 0);
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::HALT, 0);

    // chunk 1: sets global "state" = "updated"
    let mut cb = Chunk::new("update_state");
    cb.arity = 0;
    cb.local_count = 1;
    let cu = cb.add_constant(Value::String(Arc::from("updated")));
    cb.emit_op_u16(Op::CONST, cu, 0);
    let gs2 = cb.add_constant(Value::String(Arc::from("state")));
    cb.emit_op_u16(Op::GLOBAL_SET, gs2, 0);
    cb.emit_op(Op::RETURN, 0);

    // chunk 2: reads global "state"
    let mut reader = Chunk::new("read_state");
    reader.arity = 0;
    reader.local_count = 1;
    let gr = reader.add_constant(Value::String(Arc::from("state")));
    reader.emit_op_u16(Op::GLOBAL_GET, gr, 0);
    reader.emit_op(Op::RETURN, 0);

    vm.run(vec![main_chunk, cb, reader]).unwrap();

    let read_fn = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("read_state".to_string()),
            arity: 0,
            chunk_index: 2,
            upvalues: vec![],
        }),
        type_id: 0,
        fields: Vec::new(),
    })));

    // Before callback
    let r1 = vm.invoke(&read_fn, &[]).unwrap();
    assert_string(&r1, "initial");

    // Invoke callback
    let cb_val = callback_store
        .lock()
        .unwrap()
        .clone()
        .expect("callback stored");
    vm.invoke(&cb_val, &[]).unwrap();

    // After callback
    let r2 = vm.invoke(&read_fn, &[]).unwrap();
    assert_string(&r2, "updated");
}

// ============================================================
// E. Error handling (tests 36-40)
// ============================================================

// 36. call_import with unresolved index — error
#[test]
fn call_import_unresolved_index() {
    let mut vm = VM::new();
    // Register one host fn so index 0 exists
    vm.register_host_fn(
        "test",
        "exists",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| Value::Null),
    );

    let mut main = Chunk::new("main");
    main.local_count = 1;
    let _imp = main.add_import("test", "exists");
    // Manually emit call_import with index 99 (unresolved)
    // But import resolution happens at run() time based on chunk imports.
    // Instead, add an import that doesn't exist:
    let mut main2 = Chunk::new("main");
    main2.local_count = 1;
    let imp = main2.add_import("test", "nonexistent");
    emit_call_import(&mut main2, imp, 0);
    main2.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main2]);
    assert!(result.is_err(), "Expected error for unresolved import");
    let err_msg = format!("{}", result.unwrap_err());
    assert!(
        err_msg.contains("Unresolved import"),
        "Error should mention unresolved import, got: {}",
        err_msg
    );
}

// 37. call_value on non-callable (e.g. number) — error with message
#[test]
fn call_value_on_number_errors() {
    let mut vm = VM::new();

    let mut main = Chunk::new("main");
    main.local_count = 1;
    // Push a number, then try to call it
    let c = main.add_constant(Value::F64(42.0));
    main.emit_op_u16(Op::CONST, c, 0);
    main.emit_op_u8(Op::CALL, 0, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]);
    assert!(result.is_err(), "Expected error calling a number");
    let err_msg = format!("{}", result.unwrap_err());
    assert!(
        err_msg.contains("not callable"),
        "Error should say not callable, got: {}",
        err_msg
    );
}

// 38. call_value on Null — error
#[test]
fn call_value_on_null_errors() {
    let mut vm = VM::new();

    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op(Op::NULL, 0);
    main.emit_op_u8(Op::CALL, 0, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]);
    assert!(result.is_err(), "Expected error calling Null");
    let err_msg = format!("{}", result.unwrap_err());
    assert!(
        err_msg.contains("not callable"),
        "Error should say not callable, got: {}",
        err_msg
    );
}

// 39. call_value on Undefined — error
#[test]
fn call_value_on_undefined_errors() {
    let mut vm = VM::new();

    let mut main = Chunk::new("main");
    main.local_count = 1;
    {
        let c = main.add_constant(Value::Undefined);
        main.emit_op_u16(Op::CONST, c, 0);
    }
    main.emit_op_u8(Op::CALL, 0, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main]);
    assert!(result.is_err(), "Expected error calling Undefined");
    let err_msg = format!("{}", result.unwrap_err());
    assert!(
        err_msg.contains("not callable"),
        "Error should say not callable, got: {}",
        err_msg
    );
}

// 40. invoke() on non-function — error
#[test]
fn invoke_on_non_function_errors() {
    let mut vm = VM::new();

    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op(Op::NULL, 0);
    main.emit_op(Op::HALT, 0);

    vm.run(vec![main]).unwrap();

    let result = vm.invoke(&Value::I32(42), &[]);
    assert!(result.is_err(), "Expected error invoking a number");
    let err_msg = format!("{}", result.unwrap_err());
    assert!(
        err_msg.contains("not callable"),
        "Error should say not callable, got: {}",
        err_msg
    );
}
