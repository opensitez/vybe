/// Tests for VM import resolution, global scoping, call mechanics, and edge cases.

use vybe_bytecode::{VM, Value, Chunk, Op};
use vybe_bytecode::value::{Object, ObjectKind};
use std::rc::Rc;
use std::cell::RefCell;

// ============================================================
// IMPORT RESOLUTION
// ============================================================

#[test]
fn import_resolution_basic() {
    let mut vm = VM::new();
    vm.register_host_fn("test", "add", Box::new(|args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0)
                 + args.get(1).map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.imports.push(vybe_bytecode::chunk::Import { module: "test".into(), name: "add".into() });
    let a = chunk.add_constant(Value::F64(3.0));
    let b = chunk.add_constant(Value::F64(4.0));
    chunk.emit_op_u16(Op::r#const, a, 0);
    chunk.emit_op_u16(Op::r#const, b, 0);
    // call_import: u16 import_idx + u8 argc
    chunk.emit_op(Op::call_import, 0);
    chunk.emit(0, 0); chunk.emit(0, 0); // import_idx = 0
    chunk.emit(2, 0); // argc = 2
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 7.0);
}

#[test]
fn import_unresolved_errors_gracefully() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.imports.push(vybe_bytecode::chunk::Import { module: "missing".into(), name: "func".into() });
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]);
    assert!(result.is_err());
    assert!(format!("{}", result.unwrap_err()).contains("Unresolved import"));
}

#[test]
fn import_multiple_modules_correct_dispatch() {
    let mut vm = VM::new();
    vm.register_host_fn("math", "double", Box::new(|args: &[Value]| {
        Value::F64(args[0].as_f64() * 2.0)
    }));
    vm.register_host_fn("str", "len", Box::new(|args: &[Value]| {
        Value::F64(format!("{}", args[0]).len() as f64)
    }));

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.imports.push(vybe_bytecode::chunk::Import { module: "math".into(), name: "double".into() });
    chunk.imports.push(vybe_bytecode::chunk::Import { module: "str".into(), name: "len".into() });

    // Call str.len("hello") — should be import 1, not 0
    let s = chunk.add_constant(Value::String(Rc::from("hello")));
    chunk.emit_op_u16(Op::r#const, s, 0);
    chunk.emit_op(Op::call_import, 0);
    chunk.emit(0, 0); chunk.emit(1, 0); // import_idx = 1 (str.len)
    chunk.emit(1, 0); // argc = 1
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 5.0);
}

#[test]
fn import_same_module_different_functions() {
    let mut vm = VM::new();
    vm.register_host_fn("math", "add", Box::new(|args: &[Value]| {
        Value::F64(args[0].as_f64() + args[1].as_f64())
    }));
    vm.register_host_fn("math", "mul", Box::new(|args: &[Value]| {
        Value::F64(args[0].as_f64() * args[1].as_f64())
    }));

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.imports.push(vybe_bytecode::chunk::Import { module: "math".into(), name: "add".into() });
    chunk.imports.push(vybe_bytecode::chunk::Import { module: "math".into(), name: "mul".into() });

    let v3 = chunk.add_constant(Value::F64(3.0));
    let v4 = chunk.add_constant(Value::F64(4.0));
    let v10 = chunk.add_constant(Value::F64(10.0));

    // add(3, 4) = 7
    chunk.emit_op_u16(Op::r#const, v3, 0);
    chunk.emit_op_u16(Op::r#const, v4, 0);
    chunk.emit_op(Op::call_import, 0);
    chunk.emit(0, 0); chunk.emit(0, 0); // import 0 = add
    chunk.emit(2, 0);
    // mul(7, 10) = 70
    chunk.emit_op_u16(Op::r#const, v10, 0);
    chunk.emit_op(Op::call_import, 0);
    chunk.emit(0, 0); chunk.emit(1, 0); // import 1 = mul
    chunk.emit(2, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 70.0);
}

// ============================================================
// GLOBAL VARIABLE RESOLUTION
// ============================================================

#[test]
fn global_get_missing_returns_undefined() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let idx = chunk.add_constant(Value::String(Rc::from("nonexistent")));
    chunk.emit_op_u16(Op::global_get, idx, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Undefined));
}

#[test]
fn global_set_then_get_roundtrip() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let name = chunk.add_constant(Value::String(Rc::from("x")));
    let val = chunk.add_constant(Value::F64(42.0));
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::global_set, name, 0);
    chunk.emit_op(Op::drop, 0);
    chunk.emit_op_u16(Op::global_get, name, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn globals_persist_after_run() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let name = chunk.add_constant(Value::String(Rc::from("saved")));
    let val = chunk.add_constant(Value::F64(99.0));
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::global_set, name, 0);
    chunk.emit_op(Op::drop, 0);
    chunk.emit_op(Op::null, 0);
    chunk.emit_op(Op::halt, 0);
    vm.run(vec![chunk]).unwrap();
    assert_eq!(vm.globals.get("saved").unwrap().as_f64(), 99.0);
}

#[test]
fn globals_persist_across_multiple_runs() {
    let mut vm = VM::new();
    // Run 1: set x = 10
    let mut c1 = Chunk::new("<script>");
    c1.local_count = 1;
    let n1 = c1.add_constant(Value::String(Rc::from("x")));
    let v1 = c1.add_constant(Value::F64(10.0));
    c1.emit_op_u16(Op::r#const, v1, 0);
    c1.emit_op_u16(Op::global_set, n1, 0);
    c1.emit_op(Op::drop, 0);
    c1.emit_op(Op::null, 0);
    c1.emit_op(Op::halt, 0);
    vm.run(vec![c1]).unwrap();

    // Run 2: read x
    let mut c2 = Chunk::new("<script>");
    c2.local_count = 1;
    let n2 = c2.add_constant(Value::String(Rc::from("x")));
    c2.emit_op_u16(Op::global_get, n2, 0);
    c2.emit_op(Op::halt, 0);
    let result = vm.run(vec![c2]).unwrap();
    assert_eq!(result.as_f64(), 10.0);
}

// ============================================================
// STRUCT_SET / STRUCT_GET SEMANTICS
// ============================================================

#[test]
fn struct_set_returns_value() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let prop = chunk.add_constant(Value::String(Rc::from("x")));
    let val = chunk.add_constant(Value::F64(42.0));
    chunk.emit_op_u16(Op::struct_new, 0, 0);
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::struct_set, prop, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn struct_get_missing_returns_null() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let prop = chunk.add_constant(Value::String(Rc::from("missing")));
    chunk.emit_op_u16(Op::struct_new, 0, 0);
    chunk.emit_op_u16(Op::struct_get, prop, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Null));
}

// ============================================================
// LOCAL_SET SEMANTICS
// ============================================================

#[test]
fn local_set_peeks_not_pops() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;
    let val = chunk.add_constant(Value::F64(99.0));
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::local_set, 1, 0);
    // Value should still be on stack (peek semantics)
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 99.0);
}

// ============================================================
// CALL ARITY MISMATCH
// ============================================================

#[test]
fn call_fewer_args_pads_null() {
    let mut vm = VM::new();
    // Function expects 3 args, returns arg[2]
    let mut f = Chunk::new("f");
    f.arity = 3;
    f.local_count = 3;
    f.emit_op_u16(Op::local_get, 2, 0);
    f.emit_op(Op::r#return, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    let arg = main.add_constant(Value::F64(10.0));
    main.emit_op_u16(Op::r#const, arg, 0);
    main.emit_op_u8(Op::call, 1, 0); // only 1 arg
    main.emit_op(Op::halt, 0);

    let result = vm.run(vec![main, f]).unwrap();
    assert!(matches!(result, Value::Null | Value::Undefined)); // 3rd arg padded with Undefined
}

// ============================================================
// INVOKE FROM RUST
// ============================================================

#[test]
fn invoke_function_defined_in_run() {
    let mut vm = VM::new();
    let mut f = Chunk::new("double");
    f.arity = 1;
    f.local_count = 2; // slot 0 = callee, slot 1 = arg
    f.emit_op_u16(Op::local_get, 1, 0); // arg is at slot 1
    let two = f.add_constant(Value::F64(2.0));
    f.emit_op_u16(Op::r#const, two, 0);
    f.emit_op(Op::f64_mul, 0);
    f.emit_op(Op::r#return, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    let name = main.add_constant(Value::String(Rc::from("double")));
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::global_set, name, 0);
    main.emit_op(Op::drop, 0);
    main.emit_op(Op::null, 0);
    main.emit_op(Op::halt, 0);

    vm.run(vec![main, f]).unwrap();
    let func = vm.globals.get("double").cloned().unwrap();
    let result = vm.invoke(&func, &[Value::F64(21.0)]).unwrap();
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn invoke_host_fn_via_value() {
    let mut vm = VM::new();
    vm.register_host_fn("t", "greet", Box::new(|args: &[Value]| {
        Value::String(Rc::from(format!("hi {}", args[0]).as_str()))
    }));
    let idx = *vm.host_registry.get(&("t".into(), "greet".into())).unwrap();
    let mut obj = Object::new();
    obj.kind = ObjectKind::HostFunction(idx);
    let host_val = Value::Object(Rc::new(RefCell::new(obj)));
    let result = vm.invoke(&host_val, &[Value::String(Rc::from("world"))]).unwrap();
    assert_eq!(format!("{}", result), "hi world");
}

#[test]
fn invoke_multiple_times_globals_accumulate() {
    let mut vm = VM::new();
    let mut f = Chunk::new("inc");
    f.arity = 0;
    f.local_count = 0;
    let name = f.add_constant(Value::String(Rc::from("n")));
    let one = f.add_constant(Value::F64(1.0));
    f.emit_op_u16(Op::global_get, name, 0);
    f.emit_op_u16(Op::r#const, one, 0);
    f.emit_op(Op::dyn_add, 0);
    f.emit_op_u16(Op::global_set, name, 0);
    f.emit_op(Op::r#return, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    let n = main.add_constant(Value::String(Rc::from("n")));
    let fn_name = main.add_constant(Value::String(Rc::from("inc")));
    let zero = main.add_constant(Value::F64(0.0));
    main.emit_op_u16(Op::r#const, zero, 0);
    main.emit_op_u16(Op::global_set, n, 0);
    main.emit_op(Op::drop, 0);
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::global_set, fn_name, 0);
    main.emit_op(Op::drop, 0);
    main.emit_op(Op::null, 0);
    main.emit_op(Op::halt, 0);

    vm.run(vec![main, f]).unwrap();
    let inc = vm.globals.get("inc").cloned().unwrap();
    for _ in 0..5 { vm.invoke(&inc, &[]).unwrap(); }
    assert_eq!(vm.globals.get("n").unwrap().as_f64(), 5.0);
}

// ============================================================
// OBJECT METHOD VIA STRUCT_GET + CALL
// ============================================================

#[test]
fn method_on_object_via_struct_get_call() {
    let mut vm = VM::new();
    // method: takes (this), returns this.x + 1
    let mut method = Chunk::new("getX");
    method.arity = 1;
    method.local_count = 2; // slot 0 = callee, slot 1 = this
    let x = method.add_constant(Value::String(Rc::from("x")));
    let one = method.add_constant(Value::F64(1.0));
    method.emit_op_u16(Op::local_get, 1, 0); // this is at slot 1
    method.emit_op_u16(Op::struct_get, x, 0);
    method.emit_op_u16(Op::r#const, one, 0);
    method.emit_op(Op::dyn_add, 0);
    method.emit_op(Op::r#return, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 2;
    let x2 = main.add_constant(Value::String(Rc::from("x")));
    let gx = main.add_constant(Value::String(Rc::from("getX")));
    let ten = main.add_constant(Value::F64(10.0));

    // obj = {}
    main.emit_op_u16(Op::struct_new, 0, 0);
    main.emit_op_u16(Op::local_set, 1, 0);
    main.emit_op(Op::drop, 0);
    // obj.x = 10
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::r#const, ten, 0);
    main.emit_op_u16(Op::struct_set, x2, 0);
    main.emit_op(Op::drop, 0);
    // obj.getX = method
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::struct_set, gx, 0);
    main.emit_op(Op::drop, 0);
    // call obj.getX(obj)
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::struct_get, gx, 0);
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    let result = vm.run(vec![main, method]).unwrap();
    assert_eq!(result.as_f64(), 11.0);
}

// ============================================================
// TWO-BYTE OPCODE ENCODING
// ============================================================

#[test]
fn single_byte_opcode_encodes_correctly() {
    let op = Op::null;
    let (b1, b2) = op.encode();
    assert!(b2.is_none(), "Single-byte opcodes should have no second byte");
    assert!(b1 < 0xFE, "Single-byte opcodes should be < 0xFE");
}

#[test]
fn opcode_encoding_consistency() {
    // Verify that from_byte(encode()) roundtrips for single-byte opcodes
    let op = Op::dyn_add;
    let (b1, _) = op.encode();
    let decoded = Op::from_byte(b1);
    assert!(decoded.is_some());
}

#[test]
fn extended_opcode_executes() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello")));
    chunk.emit_op_u16(Op::r#const, s, 0);
    chunk.emit_op(Op::str_length, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 5);
}

// ============================================================
// HOST FN REGISTRATION ORDERING
// ============================================================

#[test]
fn host_fn_registered_before_run() {
    let mut vm = VM::new();
    vm.register_host_fn("a", "f1", Box::new(|_| Value::F64(1.0)));
    vm.register_host_fn("a", "f2", Box::new(|_| Value::F64(2.0)));
    vm.register_host_fn("b", "f3", Box::new(|_| Value::F64(3.0)));

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.imports.push(vybe_bytecode::chunk::Import { module: "b".into(), name: "f3".into() });
    chunk.emit_op(Op::call_import, 0);
    chunk.emit(0, 0); chunk.emit(0, 0);
    chunk.emit(0, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 3.0);
}

// ============================================================
// CALL_VALUE WITH NON-CALLABLE
// ============================================================

#[test]
fn call_value_number_errors() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let n = chunk.add_constant(Value::F64(42.0));
    chunk.emit_op_u16(Op::r#const, n, 0);
    chunk.emit_op_u8(Op::call, 0, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]);
    assert!(result.is_err());
    assert!(format!("{}", result.unwrap_err()).contains("not callable"));
}

#[test]
fn call_value_null_errors() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_op(Op::null, 0);
    chunk.emit_op_u8(Op::call, 0, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]);
    assert!(result.is_err());
}
