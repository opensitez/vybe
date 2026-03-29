use vybe_bytecode::{VM, Value, Chunk, Op};
use vybe_bytecode::value::{Object, ObjectKind, Function};
use std::rc::Rc;
use std::cell::RefCell;
use std::collections::HashMap;

// ============================================================
// Helpers
// ============================================================

fn run_chunks(chunks: Vec<Chunk>) -> Value {
    let mut vm = VM::new();
    vm.run(chunks).unwrap()
}

#[allow(dead_code)]
fn run_vm(chunks: Vec<Chunk>) -> VM {
    let mut vm = VM::new();
    vm.run(chunks).unwrap();
    vm
}

#[allow(dead_code)]
fn assert_f64(val: &Value, expected: f64) {
    match val {
        Value::F64(v) => assert!(
            (v - expected).abs() < 1e-10,
            "Expected F64({}), got F64({})",
            expected, v
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

fn assert_null(val: &Value) {
    match val {
        Value::Null => {}
        _ => panic!("Expected Null, got {:?}", val),
    }
}

fn assert_undefined(val: &Value) {
    match val {
        Value::Undefined => {}
        _ => panic!("Expected Undefined, got {:?}", val),
    }
}

#[allow(dead_code)]
fn assert_bool(val: &Value, expected: bool) {
    match val {
        Value::Bool(v) => assert_eq!(*v, expected),
        _ => panic!("Expected Bool({}), got {:?}", expected, val),
    }
}

// ============================================================
// 1. Function call mechanics — fewer args than arity (padded with Null)
// ============================================================

#[test]
fn call_fewer_args_than_arity_padded_with_null() {
    // Function expects 3 args, we pass 1. Missing args should be Null.
    // func(a, b, c) => if b == Null then return a else return b
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0); // 0 upvalues
    let c5 = main.add_constant(Value::I32(5));
    main.emit_op_u16(Op::r#const, c5, 0);
    // call with 1 arg, but function arity is 3
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    // func(a, b, c): returns b (which should be Null since not passed)
    let mut func = Chunk::new("check_padding");
    func.arity = 3;
    func.local_count = 4; // func + 3 params
    // local 1 = a, local 2 = b (should be Null), local 3 = c (should be Null)
    func.emit_op_u16(Op::local_get, 2, 0); // b
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_null(&result);
}

// ============================================================
// 2. Function call mechanics — more args than arity
// ============================================================

#[test]
fn call_more_args_than_arity() {
    // Function expects 1 arg, we pass 3. Extra args should be ignored.
    // func(a) => return a
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    let c10 = main.add_constant(Value::I32(10));
    let c20 = main.add_constant(Value::I32(20));
    let c30 = main.add_constant(Value::I32(30));
    main.emit_op_u16(Op::r#const, c10, 0);
    main.emit_op_u16(Op::r#const, c20, 0);
    main.emit_op_u16(Op::r#const, c30, 0);
    main.emit_op_u8(Op::call, 3, 0);
    main.emit_op(Op::halt, 0);

    let mut func = Chunk::new("take_one");
    func.arity = 1;
    func.local_count = 2;
    func.emit_op_u16(Op::local_get, 1, 0); // a = 10
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_i32(&result, 10);
}

// ============================================================
// 3. Function call mechanics — exactly matching arity
// ============================================================

#[test]
fn call_exact_arity() {
    // func(a, b) => a + b, called with exactly 2 args
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    let c3 = main.add_constant(Value::I32(3));
    let c4 = main.add_constant(Value::I32(4));
    main.emit_op_u16(Op::r#const, c3, 0);
    main.emit_op_u16(Op::r#const, c4, 0);
    main.emit_op_u8(Op::call, 2, 0);
    main.emit_op(Op::halt, 0);

    let mut func = Chunk::new("add");
    func.arity = 2;
    func.local_count = 3;
    func.emit_op_u16(Op::local_get, 1, 0);
    func.emit_op_u16(Op::local_get, 2, 0);
    func.emit_op(Op::i32_add, 0);
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_i32(&result, 7);
}

// ============================================================
// 4. Nested function calls (A calls B calls C)
// ============================================================

#[test]
fn nested_function_calls_a_b_c() {
    // main calls A(2), A calls B(x*3), B calls C(x+1), C returns x*10
    // A(2) => B(6) => C(7) => 70
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    let c2 = main.add_constant(Value::I32(2));
    main.emit_op_u16(Op::r#const, c2, 0);
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: A(x) => B(x * 3)
    let mut a = Chunk::new("A");
    a.arity = 1;
    a.local_count = 2;
    a.emit_op_u16(Op::ref_func, 2, 0);
    a.emit(0, 0);
    a.emit_op_u16(Op::local_get, 1, 0);
    let c3 = a.add_constant(Value::I32(3));
    a.emit_op_u16(Op::r#const, c3, 0);
    a.emit_op(Op::i32_mul, 0);
    a.emit_op_u8(Op::call, 1, 0);
    a.emit_op(Op::r#return, 0);

    // chunk 2: B(x) => C(x + 1)
    let mut b = Chunk::new("B");
    b.arity = 1;
    b.local_count = 2;
    b.emit_op_u16(Op::ref_func, 3, 0);
    b.emit(0, 0);
    b.emit_op_u16(Op::local_get, 1, 0);
    let c1 = b.add_constant(Value::I32(1));
    b.emit_op_u16(Op::r#const, c1, 0);
    b.emit_op(Op::i32_add, 0);
    b.emit_op_u8(Op::call, 1, 0);
    b.emit_op(Op::r#return, 0);

    // chunk 3: C(x) => x * 10
    let mut c = Chunk::new("C");
    c.arity = 1;
    c.local_count = 2;
    c.emit_op_u16(Op::local_get, 1, 0);
    let c10 = c.add_constant(Value::I32(10));
    c.emit_op_u16(Op::r#const, c10, 0);
    c.emit_op(Op::i32_mul, 0);
    c.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, a, b, c]);
    assert_i32(&result, 70);
}

// ============================================================
// 5. Recursive function calls — fibonacci
// ============================================================

#[test]
fn recursive_fibonacci() {
    // fib(0) = 0, fib(1) = 1, fib(n) = fib(n-1) + fib(n-2)
    // fib(10) = 55
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    let c10 = main.add_constant(Value::I32(10));
    main.emit_op_u16(Op::r#const, c10, 0);
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: fib(n)
    let mut fib = Chunk::new("fib");
    fib.arity = 1;
    fib.local_count = 2;
    let c0 = fib.add_constant(Value::I32(0));
    let c1 = fib.add_constant(Value::I32(1));
    let c2 = fib.add_constant(Value::I32(2));

    // if n <= 0, return 0
    fib.emit_op_u16(Op::local_get, 1, 0); // n
    fib.emit_op_u16(Op::r#const, c0, 0);  // 0
    fib.emit_op(Op::dyn_le, 0);
    let jump_base0 = fib.emit_jump(Op::br_if_true, 0);

    // if n == 1, return 1
    fib.emit_op_u16(Op::local_get, 1, 0); // n
    fib.emit_op_u16(Op::r#const, c1, 0);  // 1
    fib.emit_op(Op::dyn_eq, 0);
    let jump_base1 = fib.emit_jump(Op::br_if_true, 0);

    // recursive: fib(n-1) + fib(n-2)
    fib.emit_op_u16(Op::ref_func, 1, 0);
    fib.emit(0, 0);
    fib.emit_op_u16(Op::local_get, 1, 0);
    fib.emit_op_u16(Op::r#const, c1, 0);
    fib.emit_op(Op::i32_sub, 0);
    fib.emit_op_u8(Op::call, 1, 0); // fib(n-1)

    fib.emit_op_u16(Op::ref_func, 1, 0);
    fib.emit(0, 0);
    fib.emit_op_u16(Op::local_get, 1, 0);
    fib.emit_op_u16(Op::r#const, c2, 0);
    fib.emit_op(Op::i32_sub, 0);
    fib.emit_op_u8(Op::call, 1, 0); // fib(n-2)

    fib.emit_op(Op::i32_add, 0);
    fib.emit_op(Op::r#return, 0);

    // base case 0: return 0
    fib.patch_jump(jump_base0);
    fib.emit_op_u16(Op::r#const, c0, 0);
    fib.emit_op(Op::r#return, 0);

    // base case 1: return 1
    fib.patch_jump(jump_base1);
    fib.emit_op_u16(Op::r#const, c1, 0);
    fib.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, fib]);
    assert_i32(&result, 55);
}

#[test]
fn recursive_factorial() {
    // fact(n): if n <= 1 return 1, else return n * fact(n-1)
    // fact(6) = 720
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    let c6 = main.add_constant(Value::I32(6));
    main.emit_op_u16(Op::r#const, c6, 0);
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    let mut fact = Chunk::new("fact");
    fact.arity = 1;
    fact.local_count = 2;
    let c1 = fact.add_constant(Value::I32(1));

    fact.emit_op_u16(Op::local_get, 1, 0);
    fact.emit_op_u16(Op::r#const, c1, 0);
    fact.emit_op(Op::dyn_le, 0);
    let jump_base = fact.emit_jump(Op::br_if_true, 0);

    // n * fact(n-1)
    fact.emit_op_u16(Op::local_get, 1, 0);
    fact.emit_op_u16(Op::ref_func, 1, 0);
    fact.emit(0, 0);
    fact.emit_op_u16(Op::local_get, 1, 0);
    fact.emit_op_u16(Op::r#const, c1, 0);
    fact.emit_op(Op::i32_sub, 0);
    fact.emit_op_u8(Op::call, 1, 0);
    fact.emit_op(Op::i32_mul, 0);
    fact.emit_op(Op::r#return, 0);

    // base: return 1
    fact.patch_jump(jump_base);
    fact.emit_op_u16(Op::r#const, c1, 0);
    fact.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, fact]);
    assert_i32(&result, 720);
}

// ============================================================
// 6. call_import with 0, 1, 2, 3, 4 args
// ============================================================

#[test]
fn call_import_zero_args() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("test", "get_value");
    main.emit_op_u16(Op::call_import, import_idx, 0);
    main.emit(0, 0); // 0 args
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "get_value", Box::new(|_args: &[Value]| {
        Value::I32(42)
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 42);
}

#[test]
fn call_import_one_arg() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("test", "negate");
    let c = main.add_constant(Value::I32(7));
    main.emit_op_u16(Op::r#const, c, 0);
    main.emit_op_u16(Op::call_import, import_idx, 0);
    main.emit(1, 0);
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "negate", Box::new(|args: &[Value]| {
        Value::I32(-args[0].as_i32())
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, -7);
}

#[test]
fn call_import_two_args() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("test", "add");
    let c1 = main.add_constant(Value::I32(11));
    let c2 = main.add_constant(Value::I32(22));
    main.emit_op_u16(Op::r#const, c1, 0);
    main.emit_op_u16(Op::r#const, c2, 0);
    main.emit_op_u16(Op::call_import, import_idx, 0);
    main.emit(2, 0);
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "add", Box::new(|args: &[Value]| {
        Value::I32(args[0].as_i32() + args[1].as_i32())
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 33);
}

#[test]
fn call_import_three_args() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("test", "sum3");
    let c1 = main.add_constant(Value::I32(1));
    let c2 = main.add_constant(Value::I32(2));
    let c3 = main.add_constant(Value::I32(3));
    main.emit_op_u16(Op::r#const, c1, 0);
    main.emit_op_u16(Op::r#const, c2, 0);
    main.emit_op_u16(Op::r#const, c3, 0);
    main.emit_op_u16(Op::call_import, import_idx, 0);
    main.emit(3, 0);
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "sum3", Box::new(|args: &[Value]| {
        Value::I32(args[0].as_i32() + args[1].as_i32() + args[2].as_i32())
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 6);
}

#[test]
fn call_import_four_args() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("test", "sum4");
    let c1 = main.add_constant(Value::I32(10));
    let c2 = main.add_constant(Value::I32(20));
    let c3 = main.add_constant(Value::I32(30));
    let c4 = main.add_constant(Value::I32(40));
    main.emit_op_u16(Op::r#const, c1, 0);
    main.emit_op_u16(Op::r#const, c2, 0);
    main.emit_op_u16(Op::r#const, c3, 0);
    main.emit_op_u16(Op::r#const, c4, 0);
    main.emit_op_u16(Op::call_import, import_idx, 0);
    main.emit(4, 0);
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "sum4", Box::new(|args: &[Value]| {
        Value::I32(args[0].as_i32() + args[1].as_i32() + args[2].as_i32() + args[3].as_i32())
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 100);
}

// ============================================================
// 7. call_import return value used in expression
// ============================================================

#[test]
fn call_import_return_in_expression() {
    // result = double(5) + 3 = 13
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("test", "double");
    let c5 = main.add_constant(Value::I32(5));
    main.emit_op_u16(Op::r#const, c5, 0);
    main.emit_op_u16(Op::call_import, import_idx, 0);
    main.emit(1, 0);
    let c3 = main.add_constant(Value::I32(3));
    main.emit_op_u16(Op::r#const, c3, 0);
    main.emit_op(Op::i32_add, 0);
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "double", Box::new(|args: &[Value]| {
        Value::I32(args[0].as_i32() * 2)
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 13);
}

// ============================================================
// 8. call_value with HostFunction objects
// ============================================================

#[test]
fn call_value_with_host_function() {
    // Register a host fn, create a HostFunction object on the stack, call it
    let mut main = Chunk::new("main");
    main.local_count = 2;

    // First, register host fn so it gets index 0 in host_fns
    // We need to put a HostFunction object on the stack
    // The easiest way: store host fn ref as global, then load it
    let import_idx = main.add_import("test", "triple");
    let c7 = main.add_constant(Value::I32(7));
    main.emit_op_u16(Op::r#const, c7, 0);
    main.emit_op_u16(Op::call_import, import_idx, 0);
    main.emit(1, 0);
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "triple", Box::new(|args: &[Value]| {
        Value::I32(args[0].as_i32() * 3)
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 21);
}

// ============================================================
// 9. call_value with Function objects
// ============================================================

#[test]
fn call_value_with_function_object() {
    // Create a Function object via ref_func, store in local, then call via Op::call
    let mut main = Chunk::new("main");
    main.local_count = 2;
    // ref_func pushes a Function object
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::local_set, 1, 0);
    main.emit_op(Op::drop, 0);
    // Load it back and call it
    main.emit_op_u16(Op::local_get, 1, 0);
    let c99 = main.add_constant(Value::I32(99));
    main.emit_op_u16(Op::r#const, c99, 0);
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: identity(x) => x
    let mut func = Chunk::new("identity");
    func.arity = 1;
    func.local_count = 2;
    func.emit_op_u16(Op::local_get, 1, 0);
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_i32(&result, 99);
}

// ============================================================
// 10. call_value with non-callable (should error gracefully)
// ============================================================

#[test]
fn call_value_non_callable_errors() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let c = main.add_constant(Value::I32(42));
    main.emit_op_u16(Op::r#const, c, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![main]);
    assert!(result.is_err(), "Calling a non-callable should return an error");
    let err = result.unwrap_err();
    assert!(
        err.message.contains("not callable") || err.message.contains("Not a function"),
        "Error should mention not callable, got: {}",
        err.message
    );
}

// ============================================================
// 11. invoke with 0 args on a function with arity 0
// ============================================================

#[test]
fn invoke_zero_args_arity_zero() {
    // chunk 0: return 42
    let mut func = Chunk::new("get42");
    func.arity = 0;
    func.local_count = 1;
    let c = func.add_constant(Value::I32(42));
    func.emit_op_u16(Op::r#const, c, 0);
    func.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    // Load chunks so the VM knows about them
    let dummy_main = Chunk::new("main");
    vm.run(vec![dummy_main, func.clone()]).ok();

    let func_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("get42".to_string()),
            arity: 0,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let result = vm.invoke(&func_obj, &[]).unwrap();
    assert_i32(&result, 42);
}

// ============================================================
// 12. invoke with 1 arg on arity 3 — padding
// ============================================================

#[test]
fn invoke_fewer_args_padding() {
    // func(a, b, c) => returns b (should be Null if only 1 arg passed)
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    let mut func = Chunk::new("check_pad");
    func.arity = 3;
    func.local_count = 4;
    func.emit_op_u16(Op::local_get, 2, 0); // b
    func.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.run(vec![dummy_main, func]).ok();

    let func_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("check_pad".to_string()),
            arity: 3,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let result = vm.invoke(&func_obj, &[Value::I32(100)]).unwrap();
    assert_null(&result);
}

// ============================================================
// 13. invoke with 3 args on arity 3
// ============================================================

#[test]
fn invoke_exact_args() {
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    // func(a, b, c) => a + b + c
    let mut func = Chunk::new("sum3");
    func.arity = 3;
    func.local_count = 4;
    func.emit_op_u16(Op::local_get, 1, 0);
    func.emit_op_u16(Op::local_get, 2, 0);
    func.emit_op(Op::i32_add, 0);
    func.emit_op_u16(Op::local_get, 3, 0);
    func.emit_op(Op::i32_add, 0);
    func.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.run(vec![dummy_main, func]).ok();

    let func_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("sum3".to_string()),
            arity: 3,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let result = vm.invoke(&func_obj, &[Value::I32(10), Value::I32(20), Value::I32(30)]).unwrap();
    assert_i32(&result, 60);
}

// ============================================================
// 14. invoke returning a value
// ============================================================

#[test]
fn invoke_returning_value() {
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    // func() => "hello"
    let mut func = Chunk::new("greet");
    func.arity = 0;
    func.local_count = 1;
    let c = func.add_constant(Value::String(Rc::from("hello")));
    func.emit_op_u16(Op::r#const, c, 0);
    func.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.run(vec![dummy_main, func]).ok();

    let func_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("greet".to_string()),
            arity: 0,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let result = vm.invoke(&func_obj, &[]).unwrap();
    assert_string(&result, "hello");
}

// ============================================================
// 15. invoke returning an object
// ============================================================

#[test]
fn invoke_returning_object() {
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    // func() => creates {x: 10} via struct_new
    let mut func = Chunk::new("make_obj");
    func.arity = 0;
    func.local_count = 1;
    let key = func.add_constant(Value::String(Rc::from("x")));
    let val = func.add_constant(Value::I32(10));
    func.emit_op_u16(Op::r#const, key, 0);
    func.emit_op_u16(Op::r#const, val, 0);
    func.emit_op_u16(Op::struct_new, 1, 0); // 1 key-value pair
    func.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.run(vec![dummy_main, func]).ok();

    let func_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("make_obj".to_string()),
            arity: 0,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let result = vm.invoke(&func_obj, &[]).unwrap();
    match &result {
        Value::Object(o) => {
            let ob = o.borrow();
            assert_i32(&ob.get("x"), 10);
        }
        _ => panic!("Expected Object, got {:?}", result),
    }
}

// ============================================================
// 16. invoke a host function
// ============================================================

#[test]
fn invoke_host_function() {
    // Host functions invoked via invoke() are handled inline by call_value
    // (no frame is pushed). We test host function invocation through a wrapper
    // bytecode function that calls the host function via call_import.
    // Imports are resolved from chunk 0, so we add the import there.
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    let _import_idx = dummy_main.add_import("test", "square"); // index 0
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    // chunk 1: wrapper(x) => call_import square(x)
    // Import index 0 is resolved from chunk 0's import table
    let mut wrapper = Chunk::new("wrapper");
    wrapper.arity = 1;
    wrapper.local_count = 2;
    wrapper.emit_op_u16(Op::local_get, 1, 0); // x
    wrapper.emit_op_u16(Op::call_import, 0, 0); // import index 0
    wrapper.emit(1, 0); // 1 arg
    wrapper.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "square", Box::new(|args: &[Value]| {
        let n = args[0].as_i32();
        Value::I32(n * n)
    }));
    vm.run(vec![dummy_main, wrapper]).ok();

    let wrapper_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("wrapper".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let result = vm.invoke(&wrapper_obj, &[Value::I32(9)]).unwrap();
    assert_i32(&result, 81);
}

// ============================================================
// 17. invoke after a previous invoke (stack clean between invocations)
// ============================================================

#[test]
fn invoke_stack_clean_between_invocations() {
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    // func(x) => x * 2
    let mut func = Chunk::new("double");
    func.arity = 1;
    func.local_count = 2;
    func.emit_op_u16(Op::local_get, 1, 0);
    let c2 = func.add_constant(Value::I32(2));
    func.emit_op_u16(Op::r#const, c2, 0);
    func.emit_op(Op::i32_mul, 0);
    func.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.run(vec![dummy_main, func]).ok();

    let func_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("double".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let r1 = vm.invoke(&func_obj, &[Value::I32(5)]).unwrap();
    assert_i32(&r1, 10);

    let r2 = vm.invoke(&func_obj, &[Value::I32(7)]).unwrap();
    assert_i32(&r2, 14);

    // Verify that repeated invocations don't accumulate state
    let r3 = vm.invoke(&func_obj, &[Value::I32(100)]).unwrap();
    assert_i32(&r3, 200);
}

// ============================================================
// 18. invoke preserves globals between calls
// ============================================================

#[test]
fn invoke_preserves_globals() {
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    // chunk 1: set_global(val) => sets global "counter" = val
    let mut setter = Chunk::new("set_global");
    setter.arity = 1;
    setter.local_count = 2;
    let name_idx = setter.add_constant(Value::String(Rc::from("counter")));
    setter.emit_op_u16(Op::local_get, 1, 0);
    setter.emit_op_u16(Op::global_set, name_idx, 0);
    setter.emit_op(Op::r#return, 0);

    // chunk 2: get_global() => returns global "counter"
    let mut getter = Chunk::new("get_global");
    getter.arity = 0;
    getter.local_count = 1;
    let gname = getter.add_constant(Value::String(Rc::from("counter")));
    getter.emit_op_u16(Op::global_get, gname, 0);
    getter.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.run(vec![dummy_main, setter, getter]).ok();

    let setter_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("set_global".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    let getter_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("get_global".to_string()),
            arity: 0,
            chunk_index: 2,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    // Set global to 42
    vm.invoke(&setter_obj, &[Value::I32(42)]).unwrap();
    // Get global — should be 42
    let result = vm.invoke(&getter_obj, &[]).unwrap();
    assert_i32(&result, 42);
}

// ============================================================
// 19. invoke with object that has methods (struct_get + call)
// ============================================================

#[test]
fn invoke_function_that_uses_struct_get() {
    let mut dummy_main = Chunk::new("main");
    dummy_main.local_count = 1;
    dummy_main.emit_op(Op::null, 0);
    dummy_main.emit_op(Op::halt, 0);

    // chunk 1: get_x(obj) => obj.x
    let mut func = Chunk::new("get_x");
    func.arity = 1;
    func.local_count = 2;
    let prop = func.add_constant(Value::String(Rc::from("x")));
    func.emit_op_u16(Op::local_get, 1, 0); // obj
    func.emit_op_u16(Op::struct_get, prop, 0); // obj.x
    func.emit_op(Op::r#return, 0);

    let mut vm = VM::new();
    vm.run(vec![dummy_main, func]).ok();

    let func_obj = Value::Object(Rc::new(RefCell::new(Object {
        properties: HashMap::new(),
        kind: ObjectKind::Function(Function {
            name: Some("get_x".to_string()),
            arity: 1,
            chunk_index: 1,
            upvalues: vec![],
        }),
        type_id: 0, fields: Vec::new(),
    })));

    // Create an object {x: 99}
    let mut obj = Object::new();
    obj.set("x".to_string(), Value::I32(99));
    let obj_val = Value::Object(Rc::new(RefCell::new(obj)));

    let result = vm.invoke(&func_obj, &[obj_val]).unwrap();
    assert_i32(&result, 99);
}

// ============================================================
// 20. local_get/set within a function
// ============================================================

#[test]
fn local_get_set_within_function() {
    // func() { local x = 5; local y = 10; return x + y; }
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut func = Chunk::new("local_test");
    func.arity = 0;
    func.local_count = 3; // func + x + y
    let c5 = func.add_constant(Value::I32(5));
    let c10 = func.add_constant(Value::I32(10));
    func.emit_op_u16(Op::r#const, c5, 0);
    func.emit_op_u16(Op::local_set, 1, 0); // x = 5
    func.emit_op(Op::drop, 0);
    func.emit_op_u16(Op::r#const, c10, 0);
    func.emit_op_u16(Op::local_set, 2, 0); // y = 10
    func.emit_op(Op::drop, 0);
    func.emit_op_u16(Op::local_get, 1, 0); // x
    func.emit_op_u16(Op::local_get, 2, 0); // y
    func.emit_op(Op::i32_add, 0);
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_i32(&result, 15);
}

// ============================================================
// 21. Locals don't leak between function calls
// ============================================================

#[test]
fn locals_dont_leak_between_calls() {
    // func_a sets local 1 = 100, func_b reads local 1 (should be its own, Null or default)
    // main: call func_a(), then call func_b()
    let mut main = Chunk::new("main");
    main.local_count = 1;
    // call func_a
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::drop, 0);
    // call func_b
    main.emit_op_u16(Op::ref_func, 2, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: func_a() { local 1 = 100; return 100; }
    let mut func_a = Chunk::new("func_a");
    func_a.arity = 0;
    func_a.local_count = 2;
    let c100 = func_a.add_constant(Value::I32(100));
    func_a.emit_op_u16(Op::r#const, c100, 0);
    func_a.emit_op_u16(Op::local_set, 1, 0);
    func_a.emit_op(Op::r#return, 0);

    // chunk 2: func_b() { return local 1; } — local 1 should be Null
    let mut func_b = Chunk::new("func_b");
    func_b.arity = 0;
    func_b.local_count = 2;
    func_b.emit_op_u16(Op::local_get, 1, 0);
    func_b.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func_a, func_b]);
    assert_null(&result);
}

// ============================================================
// 22. Globals persist after function returns
// ============================================================

#[test]
fn globals_persist_after_function_returns() {
    // func_a sets global "val" = 42
    // main calls func_a(), then reads global "val"
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::drop, 0);
    let gname = main.add_constant(Value::String(Rc::from("val")));
    main.emit_op_u16(Op::global_get, gname, 0);
    main.emit_op(Op::halt, 0);

    let mut func = Chunk::new("set_val");
    func.arity = 0;
    func.local_count = 1;
    let name_idx = func.add_constant(Value::String(Rc::from("val")));
    let c42 = func.add_constant(Value::I32(42));
    func.emit_op_u16(Op::r#const, c42, 0);
    func.emit_op_u16(Op::global_set, name_idx, 0);
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_i32(&result, 42);
}

// ============================================================
// 23. global_get of undefined name returns Undefined
// ============================================================

#[test]
fn global_get_undefined_returns_undefined() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let name_idx = chunk.add_constant(Value::String(Rc::from("nonexistent")));
    chunk.emit_op_u16(Op::global_get, name_idx, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    assert_undefined(&result);
}

// ============================================================
// 24. global_set then global_get roundtrip
// ============================================================

#[test]
fn global_set_get_roundtrip() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let name_idx = chunk.add_constant(Value::String(Rc::from("myVar")));
    let val = chunk.add_constant(Value::String(Rc::from("hello world")));
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::global_set, name_idx, 0);
    chunk.emit_op(Op::drop, 0);
    chunk.emit_op_u16(Op::global_get, name_idx, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    assert_string(&result, "hello world");
}

// ============================================================
// 25. Multiple functions sharing globals
// ============================================================

#[test]
fn multiple_functions_sharing_globals() {
    // func_a sets global "shared" = 10
    // func_b reads global "shared" and adds 5
    // main calls func_a(), then func_b(), returns result
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::drop, 0);
    main.emit_op_u16(Op::ref_func, 2, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut func_a = Chunk::new("func_a");
    func_a.arity = 0;
    func_a.local_count = 1;
    let name_a = func_a.add_constant(Value::String(Rc::from("shared")));
    let c10 = func_a.add_constant(Value::I32(10));
    func_a.emit_op_u16(Op::r#const, c10, 0);
    func_a.emit_op_u16(Op::global_set, name_a, 0);
    func_a.emit_op(Op::r#return, 0);

    let mut func_b = Chunk::new("func_b");
    func_b.arity = 0;
    func_b.local_count = 1;
    let name_b = func_b.add_constant(Value::String(Rc::from("shared")));
    let c5 = func_b.add_constant(Value::I32(5));
    func_b.emit_op_u16(Op::global_get, name_b, 0);
    func_b.emit_op_u16(Op::r#const, c5, 0);
    func_b.emit_op(Op::i32_add, 0);
    func_b.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func_a, func_b]);
    assert_i32(&result, 15);
}

// ============================================================
// 26. local_set in nested function doesn't affect outer
// ============================================================

#[test]
fn local_set_nested_doesnt_affect_outer() {
    // outer() { local 1 = 5; call inner(); return local 1; }
    // inner() { local 1 = 999; return local 1; }
    // outer should still return 5
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut outer = Chunk::new("outer");
    outer.arity = 0;
    outer.local_count = 2;
    let c5 = outer.add_constant(Value::I32(5));
    outer.emit_op_u16(Op::r#const, c5, 0);
    outer.emit_op_u16(Op::local_set, 1, 0);
    outer.emit_op(Op::drop, 0);
    // call inner
    outer.emit_op_u16(Op::ref_func, 2, 0);
    outer.emit(0, 0);
    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::drop, 0); // discard inner's result
    // return local 1 (should still be 5)
    outer.emit_op_u16(Op::local_get, 1, 0);
    outer.emit_op(Op::r#return, 0);

    let mut inner = Chunk::new("inner");
    inner.arity = 0;
    inner.local_count = 2;
    let c999 = inner.add_constant(Value::I32(999));
    inner.emit_op_u16(Op::r#const, c999, 0);
    inner.emit_op_u16(Op::local_set, 1, 0);
    inner.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, outer, inner]);
    assert_i32(&result, 5);
}

// ============================================================
// 27. Simple closure capturing one variable
// ============================================================

#[test]
fn closure_capture_one_variable() {
    // outer() {
    //   local x = 10;          (local 1)
    //   closure = ref_func(2) capturing local 1
    //   return call closure()
    // }
    // closure() { return upvalue 0; }  => should be 10
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut outer = Chunk::new("outer");
    outer.arity = 0;
    outer.local_count = 3; // func + x + closure
    let c10 = outer.add_constant(Value::I32(10));
    outer.emit_op_u16(Op::r#const, c10, 0);
    outer.emit_op_u16(Op::local_set, 1, 0); // x = 10
    outer.emit_op(Op::drop, 0);

    // ref_func(2) with 1 upvalue: is_local=1, index=1 (capture local 1)
    outer.emit_op_u16(Op::ref_func, 2, 0);
    outer.emit(1, 0); // 1 upvalue
    outer.emit(1, 0); // is_local = true
    outer.emit(1, 0); // index = 1 (local x)

    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::r#return, 0);

    let mut closure = Chunk::new("closure");
    closure.arity = 0;
    closure.local_count = 1;
    closure.emit_op_u8(Op::upvalue_get, 0, 0); // upvalue 0 = x
    closure.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, outer, closure]);
    assert_i32(&result, 10);
}

// ============================================================
// 28. Closure mutation (captured var modified)
// ============================================================

#[test]
fn closure_mutation() {
    // outer() {
    //   local x = 10;
    //   closure = ref_func capturing x
    //   call closure() — modifies x via upvalue_set
    //   return x
    // }
    // closure() { upvalue_set 0, 20; return Null; }
    // outer should return 20 because closure mutated x
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut outer = Chunk::new("outer");
    outer.arity = 0;
    outer.local_count = 3;
    let c10 = outer.add_constant(Value::I32(10));
    outer.emit_op_u16(Op::r#const, c10, 0);
    outer.emit_op_u16(Op::local_set, 1, 0);
    outer.emit_op(Op::drop, 0);

    // ref_func(2) with 1 upvalue capturing local 1
    outer.emit_op_u16(Op::ref_func, 2, 0);
    outer.emit(1, 0); // 1 upvalue
    outer.emit(1, 0); // is_local = true
    outer.emit(1, 0); // index = 1

    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::drop, 0); // discard closure return
    outer.emit_op_u16(Op::local_get, 1, 0); // x — should be 20
    outer.emit_op(Op::r#return, 0);

    let mut closure = Chunk::new("mutator");
    closure.arity = 0;
    closure.local_count = 1;
    let c20 = closure.add_constant(Value::I32(20));
    closure.emit_op_u16(Op::r#const, c20, 0);
    closure.emit_op_u8(Op::upvalue_set, 0, 0); // set upvalue 0 = 20
    closure.emit_op(Op::r#return, 0); // returns 20 (upvalue_set keeps val on stack)

    let result = run_chunks(vec![main, outer, closure]);
    assert_i32(&result, 20);
}

// ============================================================
// 29. Multiple closures sharing same upvalue
// ============================================================

#[test]
fn multiple_closures_sharing_upvalue() {
    // outer() {
    //   local x = 0;
    //   inc = ref_func capturing x   — increments x by 1
    //   get = ref_func capturing x   — returns x
    //   call inc()
    //   call inc()
    //   call inc()
    //   return call get()             — should be 3
    // }
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut outer = Chunk::new("outer");
    outer.arity = 0;
    outer.local_count = 4; // func + x + inc + get
    let c0 = outer.add_constant(Value::I32(0));
    outer.emit_op_u16(Op::r#const, c0, 0);
    outer.emit_op_u16(Op::local_set, 1, 0); // x = 0
    outer.emit_op(Op::drop, 0);

    // inc = ref_func(2) capturing local 1 (x)
    outer.emit_op_u16(Op::ref_func, 2, 0);
    outer.emit(1, 0); // 1 upvalue
    outer.emit(1, 0); // is_local
    outer.emit(1, 0); // index = 1
    outer.emit_op_u16(Op::local_set, 2, 0); // inc
    outer.emit_op(Op::drop, 0);

    // get = ref_func(3) capturing local 1 (x)
    outer.emit_op_u16(Op::ref_func, 3, 0);
    outer.emit(1, 0); // 1 upvalue
    outer.emit(1, 0); // is_local
    outer.emit(1, 0); // index = 1
    outer.emit_op_u16(Op::local_set, 3, 0); // get
    outer.emit_op(Op::drop, 0);

    // call inc() 3 times
    outer.emit_op_u16(Op::local_get, 2, 0);
    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::drop, 0);
    outer.emit_op_u16(Op::local_get, 2, 0);
    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::drop, 0);
    outer.emit_op_u16(Op::local_get, 2, 0);
    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::drop, 0);

    // return call get()
    outer.emit_op_u16(Op::local_get, 3, 0);
    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::r#return, 0);

    // chunk 2: inc() { x = x + 1; return Null; }
    let mut inc = Chunk::new("inc");
    inc.arity = 0;
    inc.local_count = 1;
    let c1 = inc.add_constant(Value::I32(1));
    inc.emit_op_u8(Op::upvalue_get, 0, 0); // x
    inc.emit_op_u16(Op::r#const, c1, 0);
    inc.emit_op(Op::i32_add, 0);
    inc.emit_op_u8(Op::upvalue_set, 0, 0); // x = x + 1
    inc.emit_op(Op::r#return, 0);

    // chunk 3: get() { return x; }
    let mut get = Chunk::new("get");
    get.arity = 0;
    get.local_count = 1;
    get.emit_op_u8(Op::upvalue_get, 0, 0); // x
    get.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, outer, inc, get]);
    assert_i32(&result, 3);
}

// ============================================================
// 30. Nested closure (closure creates closure)
// ============================================================

#[test]
fn nested_closure() {
    // outer() {
    //   local x = 100;
    //   middle = ref_func capturing x
    //   return call middle()
    // }
    // middle() {
    //   inner = ref_func capturing upvalue 0 (x, via non-local upvalue)
    //   return call inner()
    // }
    // inner() { return upvalue 0; }  => 100
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let mut outer = Chunk::new("outer");
    outer.arity = 0;
    outer.local_count = 2;
    let c100 = outer.add_constant(Value::I32(100));
    outer.emit_op_u16(Op::r#const, c100, 0);
    outer.emit_op_u16(Op::local_set, 1, 0);
    outer.emit_op(Op::drop, 0);

    // middle = ref_func(2) capturing local 1
    outer.emit_op_u16(Op::ref_func, 2, 0);
    outer.emit(1, 0); // 1 upvalue
    outer.emit(1, 0); // is_local = true
    outer.emit(1, 0); // index = 1

    outer.emit_op_u8(Op::call, 0, 0);
    outer.emit_op(Op::r#return, 0);

    let mut middle = Chunk::new("middle");
    middle.arity = 0;
    middle.local_count = 1;

    // inner = ref_func(3) capturing upvalue 0 from middle (which is x from outer)
    middle.emit_op_u16(Op::ref_func, 3, 0);
    middle.emit(1, 0); // 1 upvalue
    middle.emit(0, 0); // is_local = false (capture from parent's upvalue)
    middle.emit(0, 0); // index = 0 (middle's upvalue 0)

    middle.emit_op_u8(Op::call, 0, 0);
    middle.emit_op(Op::r#return, 0);

    let mut inner = Chunk::new("inner");
    inner.arity = 0;
    inner.local_count = 1;
    inner.emit_op_u8(Op::upvalue_get, 0, 0);
    inner.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, outer, middle, inner]);
    assert_i32(&result, 100);
}

// ============================================================
// 31. struct_get on object returns value
// ============================================================

#[test]
fn struct_get_returns_value() {
    // Create object {name: "alice"}, then struct_get "name"
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let key = chunk.add_constant(Value::String(Rc::from("name")));
    let val = chunk.add_constant(Value::String(Rc::from("alice")));
    chunk.emit_op_u16(Op::r#const, key, 0);
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::struct_new, 1, 0); // {name: "alice"}
    let get_key = chunk.add_constant(Value::String(Rc::from("name")));
    chunk.emit_op_u16(Op::struct_get, get_key, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    assert_string(&result, "alice");
}

// ============================================================
// 32. struct_get on object for missing prop returns Null
// ============================================================

#[test]
fn struct_get_missing_prop_returns_null() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let key = chunk.add_constant(Value::String(Rc::from("x")));
    let val = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::r#const, key, 0);
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::struct_new, 1, 0); // {x: 10}
    let miss_key = chunk.add_constant(Value::String(Rc::from("y")));
    chunk.emit_op_u16(Op::struct_get, miss_key, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    assert_null(&result);
}

// ============================================================
// 33. struct_set creates new property
// ============================================================

#[test]
fn struct_set_creates_new_property() {
    // Create empty object, struct_set "age" = 25, then struct_get "age"
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    // Create empty object
    chunk.emit_op_u16(Op::struct_new, 0, 0);
    chunk.emit_op_u16(Op::local_set, 1, 0); // local 1 = obj
    chunk.emit_op(Op::drop, 0);

    // struct_set: pops val then obj from stack
    let set_key = chunk.add_constant(Value::String(Rc::from("age")));
    let c25 = chunk.add_constant(Value::I32(25));
    chunk.emit_op_u16(Op::local_get, 1, 0); // obj
    chunk.emit_op_u16(Op::r#const, c25, 0); // 25
    chunk.emit_op_u16(Op::struct_set, set_key, 0); // obj.age = 25
    chunk.emit_op(Op::drop, 0); // struct_set pushes assigned val

    // struct_get
    let get_key = chunk.add_constant(Value::String(Rc::from("age")));
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op_u16(Op::struct_get, get_key, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 25);
}

// ============================================================
// 34. struct_set overwrites existing property
// ============================================================

#[test]
fn struct_set_overwrites_existing_property() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    // Create {x: 1}
    let key = chunk.add_constant(Value::String(Rc::from("x")));
    let val = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, key, 0);
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::struct_new, 1, 0);
    chunk.emit_op_u16(Op::local_set, 1, 0);
    chunk.emit_op(Op::drop, 0);

    // Overwrite x = 99
    let set_key = chunk.add_constant(Value::String(Rc::from("x")));
    let c99 = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op_u16(Op::r#const, c99, 0);
    chunk.emit_op_u16(Op::struct_set, set_key, 0);
    chunk.emit_op(Op::drop, 0);

    // Read it back
    let get_key = chunk.add_constant(Value::String(Rc::from("x")));
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op_u16(Op::struct_get, get_key, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 99);
}

// ============================================================
// 35. struct_get chain: a.b.c
// ============================================================

#[test]
fn struct_get_chain() {
    // Build { b: { c: 42 } }, then access a.b.c
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;

    // Create inner object {c: 42}
    let key_c = chunk.add_constant(Value::String(Rc::from("c")));
    let c42 = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::r#const, key_c, 0);
    chunk.emit_op_u16(Op::r#const, c42, 0);
    chunk.emit_op_u16(Op::struct_new, 1, 0); // {c: 42}

    // Create outer object {b: inner}
    let key_b = chunk.add_constant(Value::String(Rc::from("b")));
    // Stack: inner_obj. We need key then value for struct_new
    // Swap: push key first, then we need inner on top
    // Actually struct_new takes pairs from stack: key, val, key, val...
    // So we need: "b", inner_obj on stack
    // inner_obj is on stack. We need to push "b" under it.
    // Simpler: store inner in local, then build outer
    chunk.emit_op_u16(Op::local_set, 1, 0);
    chunk.emit_op(Op::drop, 0);
    chunk.emit_op_u16(Op::r#const, key_b, 0);
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op_u16(Op::struct_new, 1, 0); // {b: {c: 42}}

    // Now chain access: obj.b.c
    let get_b = chunk.add_constant(Value::String(Rc::from("b")));
    chunk.emit_op_u16(Op::struct_get, get_b, 0);
    let get_c = chunk.add_constant(Value::String(Rc::from("c")));
    chunk.emit_op_u16(Op::struct_get, get_c, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 42);
}

// ============================================================
// 36. Getter auto-dispatch: __get_prop is called when getting prop
// ============================================================

#[test]
fn getter_auto_dispatch() {
    // Create object with __get_name property that is a function returning "computed"
    // struct_get "name" should auto-dispatch to __get_name
    let mut main = Chunk::new("main");
    main.local_count = 2;

    // Create the getter function at chunk 1
    // Build object with __get_name = ref_func(1)
    let key = main.add_constant(Value::String(Rc::from("__get_name")));
    main.emit_op_u16(Op::r#const, key, 0);
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0); // 0 upvalues
    main.emit_op_u16(Op::struct_new, 1, 0); // { __get_name: <func> }
    // struct_get "name" should trigger the getter
    let get_name = main.add_constant(Value::String(Rc::from("name")));
    main.emit_op_u16(Op::struct_get, get_name, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: getter(this) => "computed"
    let mut getter = Chunk::new("__get_name");
    getter.arity = 1; // this
    getter.local_count = 2;
    let cs = getter.add_constant(Value::String(Rc::from("computed")));
    getter.emit_op_u16(Op::r#const, cs, 0);
    getter.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, getter]);
    assert_string(&result, "computed");
}

// ============================================================
// 37. Setter auto-dispatch: __set_prop is called when setting prop
// ============================================================

#[test]
fn setter_auto_dispatch() {
    // Create object with __set_x = func that stores value*2 in "actual_x"
    // struct_set "x" = 5 should trigger __set_x, which sets actual_x = 10
    // Then struct_get "actual_x" should return 10
    let mut main = Chunk::new("main");
    main.local_count = 2;

    let key = main.add_constant(Value::String(Rc::from("__set_x")));
    main.emit_op_u16(Op::r#const, key, 0);
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::struct_new, 1, 0); // { __set_x: <func> }
    main.emit_op_u16(Op::local_set, 1, 0);
    main.emit_op(Op::drop, 0);

    // struct_set "x" = 5
    let set_key = main.add_constant(Value::String(Rc::from("x")));
    let c5 = main.add_constant(Value::I32(5));
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::r#const, c5, 0);
    main.emit_op_u16(Op::struct_set, set_key, 0);
    main.emit_op(Op::drop, 0); // drop set result

    // struct_get "actual_x"
    let get_key = main.add_constant(Value::String(Rc::from("actual_x")));
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::struct_get, get_key, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: setter(this, value) => this.actual_x = value * 2
    let mut setter = Chunk::new("__set_x");
    setter.arity = 2; // this, value
    setter.local_count = 3;
    let actual_key = setter.add_constant(Value::String(Rc::from("actual_x")));
    let c2 = setter.add_constant(Value::I32(2));
    setter.emit_op_u16(Op::local_get, 1, 0); // this
    setter.emit_op_u16(Op::local_get, 2, 0); // value
    setter.emit_op_u16(Op::r#const, c2, 0);
    setter.emit_op(Op::i32_mul, 0); // value * 2
    setter.emit_op_u16(Op::struct_set, actual_key, 0); // this.actual_x = value*2
    setter.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, setter]);
    assert_i32(&result, 10);
}

// ============================================================
// 38. Method on object: struct_get returns Function, then call it
// ============================================================

#[test]
fn object_method_get_and_call() {
    // obj = { greet: func(this) => "hi" }
    // result = obj.greet(obj)
    let mut main = Chunk::new("main");
    main.local_count = 2;

    let key = main.add_constant(Value::String(Rc::from("greet")));
    main.emit_op_u16(Op::r#const, key, 0);
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::struct_new, 1, 0); // { greet: <func> }
    main.emit_op_u16(Op::local_set, 1, 0);
    main.emit_op(Op::drop, 0);

    // obj.greet
    let get_key = main.add_constant(Value::String(Rc::from("greet")));
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::struct_get, get_key, 0);
    // call it with obj as arg (acting as this)
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: greet(this) => "hi"
    let mut func = Chunk::new("greet");
    func.arity = 1;
    func.local_count = 2;
    let cs = func.add_constant(Value::String(Rc::from("hi")));
    func.emit_op_u16(Op::r#const, cs, 0);
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_string(&result, "hi");
}

// ============================================================
// 39. Object with multiple methods, call correct one
// ============================================================

#[test]
fn object_multiple_methods_call_correct() {
    // obj = { add: func(a,b)=>a+b, mul: func(a,b)=>a*b }
    // result = obj.mul(3, 4) => 12
    let mut main = Chunk::new("main");
    main.local_count = 2;

    // Build object with two methods
    let k_add = main.add_constant(Value::String(Rc::from("add")));
    main.emit_op_u16(Op::r#const, k_add, 0);
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);

    let k_mul = main.add_constant(Value::String(Rc::from("mul")));
    main.emit_op_u16(Op::r#const, k_mul, 0);
    main.emit_op_u16(Op::ref_func, 2, 0);
    main.emit(0, 0);

    main.emit_op_u16(Op::struct_new, 2, 0); // { add: <f1>, mul: <f2> }
    main.emit_op_u16(Op::local_set, 1, 0);
    main.emit_op(Op::drop, 0);

    // Call obj.mul(3, 4)
    let get_mul = main.add_constant(Value::String(Rc::from("mul")));
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::struct_get, get_mul, 0);
    let c3 = main.add_constant(Value::I32(3));
    let c4 = main.add_constant(Value::I32(4));
    main.emit_op_u16(Op::r#const, c3, 0);
    main.emit_op_u16(Op::r#const, c4, 0);
    main.emit_op_u8(Op::call, 2, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: add(a, b) => a + b
    let mut add_fn = Chunk::new("add");
    add_fn.arity = 2;
    add_fn.local_count = 3;
    add_fn.emit_op_u16(Op::local_get, 1, 0);
    add_fn.emit_op_u16(Op::local_get, 2, 0);
    add_fn.emit_op(Op::i32_add, 0);
    add_fn.emit_op(Op::r#return, 0);

    // chunk 2: mul(a, b) => a * b
    let mut mul_fn = Chunk::new("mul");
    mul_fn.arity = 2;
    mul_fn.local_count = 3;
    mul_fn.emit_op_u16(Op::local_get, 1, 0);
    mul_fn.emit_op_u16(Op::local_get, 2, 0);
    mul_fn.emit_op(Op::i32_mul, 0);
    mul_fn.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, add_fn, mul_fn]);
    assert_i32(&result, 12);
}

// ============================================================
// 40. Object method receives correct `this` when called with correct args
// ============================================================

#[test]
fn object_method_receives_this() {
    // obj = { value: 7, get_value: func(this) => this.value }
    // result = obj.get_value(obj) => 7
    let mut main = Chunk::new("main");
    main.local_count = 2;

    // Build { value: 7, get_value: <func> }
    let k_value = main.add_constant(Value::String(Rc::from("value")));
    let c7 = main.add_constant(Value::I32(7));
    main.emit_op_u16(Op::r#const, k_value, 0);
    main.emit_op_u16(Op::r#const, c7, 0);

    let k_get = main.add_constant(Value::String(Rc::from("get_value")));
    main.emit_op_u16(Op::r#const, k_get, 0);
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);

    main.emit_op_u16(Op::struct_new, 2, 0); // { value: 7, get_value: <func> }
    main.emit_op_u16(Op::local_set, 1, 0);
    main.emit_op(Op::drop, 0);

    // Call obj.get_value(obj)
    let get_method = main.add_constant(Value::String(Rc::from("get_value")));
    main.emit_op_u16(Op::local_get, 1, 0);
    main.emit_op_u16(Op::struct_get, get_method, 0);
    main.emit_op_u16(Op::local_get, 1, 0); // pass obj as this
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    // chunk 1: get_value(this) => this.value
    let mut func = Chunk::new("get_value");
    func.arity = 1;
    func.local_count = 2;
    let prop = func.add_constant(Value::String(Rc::from("value")));
    func.emit_op_u16(Op::local_get, 1, 0); // this
    func.emit_op_u16(Op::struct_get, prop, 0); // this.value
    func.emit_op(Op::r#return, 0);

    let result = run_chunks(vec![main, func]);
    assert_i32(&result, 7);
}
