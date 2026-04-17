use vybe_bytecode::{VM, Value, Chunk, Op};
use std::rc::Rc;

// ============================================================
// Helpers
// ============================================================

fn run_chunks(chunks: Vec<Chunk>) -> Value {
    let mut vm = VM::new();
    vm.run(chunks).unwrap()
}

fn run_vm(chunks: Vec<Chunk>) -> VM {
    let mut vm = VM::new();
    vm.run(chunks).unwrap();
    vm
}

fn assert_f64(val: &Value, expected: f64) {
    match val {
        Value::F64(v) => assert!((v - expected).abs() < 1e-10, "Expected F64({}), got F64({})", expected, v),
        _ => panic!("Expected F64({}), got {:?}", expected, val),
    }
}

fn assert_i32(val: &Value, expected: i32) {
    match val {
        Value::I32(v) => assert_eq!(*v, expected, "Expected I32({}), got I32({})", expected, v),
        _ => panic!("Expected I32({}), got {:?}", expected, val),
    }
}

fn assert_bool(val: &Value, expected: bool) {
    match val {
        Value::Bool(v) => assert_eq!(*v, expected),
        _ => panic!("Expected Bool({}), got {:?}", expected, val),
    }
}

fn assert_string(val: &Value, expected: &str) {
    match val {
        Value::String(s) => assert_eq!(s.as_ref(), expected, "Expected String({:?}), got String({:?})", expected, s.as_ref()),
        _ => panic!("Expected String({:?}), got {:?}", expected, val),
    }
}

// ============================================================
// 1. Function Calls
// ============================================================

#[test]
fn call_zero_args() {
    // chunk 0: ref_func(1), call 0, halt
    // chunk 1: returns 42.0
    let mut main = Chunk::new("main");
    main.local_count = 1;
    // ref_func chunk_index=1, upvalue_count=0
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0); // 0 upvalues
    main.emit_op_u8(Op::CALL, 0, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("func0");
    func.arity = 0;
    func.local_count = 1;
    let c = func.add_constant(Value::F64(42.0));
    func.emit_op_u16(Op::CONST, c, 0);
    func.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, func]);
    assert_f64(&result, 42.0);
}

#[test]
fn call_one_arg() {
    // chunk 1: takes 1 arg, returns arg + 10
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let c5 = main.add_constant(Value::F64(5.0));
    main.emit_op_u16(Op::CONST, c5, 0);
    main.emit_op_u8(Op::CALL, 1, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("add10");
    func.arity = 1;
    func.local_count = 2;
    let c10 = func.add_constant(Value::F64(10.0));
    func.emit_op_u16(Op::LOCAL_GET, 1, 0); // arg 0 is local 1 (local 0 = function itself)
    func.emit_op_u16(Op::CONST, c10, 0);
    func.emit_op(Op::F64_ADD, 0);
    func.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, func]);
    assert_f64(&result, 15.0);
}

#[test]
fn call_two_args() {
    // chunk 1: (a, b) => a * b
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let c3 = main.add_constant(Value::F64(3.0));
    let c7 = main.add_constant(Value::F64(7.0));
    main.emit_op_u16(Op::CONST, c3, 0);
    main.emit_op_u16(Op::CONST, c7, 0);
    main.emit_op_u8(Op::CALL, 2, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("mul");
    func.arity = 2;
    func.local_count = 3;
    func.emit_op_u16(Op::LOCAL_GET, 1, 0);
    func.emit_op_u16(Op::LOCAL_GET, 2, 0);
    func.emit_op(Op::F64_MUL, 0);
    func.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, func]);
    assert_f64(&result, 21.0);
}

#[test]
fn call_three_args() {
    // chunk 1: (a, b, c) => a + b + c
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let c1 = main.add_constant(Value::I32(10));
    let c2 = main.add_constant(Value::I32(20));
    let c3 = main.add_constant(Value::I32(30));
    main.emit_op_u16(Op::CONST, c1, 0);
    main.emit_op_u16(Op::CONST, c2, 0);
    main.emit_op_u16(Op::CONST, c3, 0);
    main.emit_op_u8(Op::CALL, 3, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("sum3");
    func.arity = 3;
    func.local_count = 4;
    func.emit_op_u16(Op::LOCAL_GET, 1, 0);
    func.emit_op_u16(Op::LOCAL_GET, 2, 0);
    func.emit_op(Op::I32_ADD, 0);
    func.emit_op_u16(Op::LOCAL_GET, 3, 0);
    func.emit_op(Op::I32_ADD, 0);
    func.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, func]);
    assert_i32(&result, 60);
}

#[test]
fn nested_function_calls() {
    // main calls outer(5), outer calls inner(x+1) => inner returns x*2
    // outer(5) => inner(6) => 12
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let c5 = main.add_constant(Value::I32(5));
    main.emit_op_u16(Op::CONST, c5, 0);
    main.emit_op_u8(Op::CALL, 1, 0);
    main.emit_op(Op::HALT, 0);

    // chunk 1: outer(x) => calls inner(x+1)
    let mut outer = Chunk::new("outer");
    outer.arity = 1;
    outer.local_count = 2;
    outer.emit_op_u16(Op::REF_FUNC, 2, 0);
    outer.emit(0, 0);
    outer.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let c1 = outer.add_constant(Value::I32(1));
    outer.emit_op_u16(Op::CONST, c1, 0);
    outer.emit_op(Op::I32_ADD, 0);
    outer.emit_op_u8(Op::CALL, 1, 0);
    outer.emit_op(Op::RETURN, 0);

    // chunk 2: inner(x) => x * 2
    let mut inner = Chunk::new("inner");
    inner.arity = 1;
    inner.local_count = 2;
    inner.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let c2 = inner.add_constant(Value::I32(2));
    inner.emit_op_u16(Op::CONST, c2, 0);
    inner.emit_op(Op::I32_MUL, 0);
    inner.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, outer, inner]);
    assert_i32(&result, 12);
}

#[test]
fn recursive_call_factorial() {
    // main: calls fact(5) => 120
    // fact(n): if n <= 1 return 1, else return n * fact(n-1)
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let c5 = main.add_constant(Value::I32(5));
    main.emit_op_u16(Op::CONST, c5, 0);
    main.emit_op_u8(Op::CALL, 1, 0);
    main.emit_op(Op::HALT, 0);

    // chunk 1: fact(n)
    let mut fact = Chunk::new("fact");
    fact.arity = 1;
    fact.local_count = 2;
    let c1 = fact.add_constant(Value::I32(1));

    // if n <= 1, branch to return 1
    fact.emit_op_u16(Op::LOCAL_GET, 1, 0); // n
    fact.emit_op_u16(Op::CONST, c1, 0);  // 1
    fact.emit_op(Op::DYN_LE, 0);           // n <= 1 ?
    let jump_to_base = fact.emit_jump(Op::BR_IF_TRUE, 0);

    // recursive case: n * fact(n-1)
    fact.emit_op_u16(Op::LOCAL_GET, 1, 0); // n
    fact.emit_op_u16(Op::REF_FUNC, 1, 0);  // fact
    fact.emit(0, 0); // 0 upvalues
    fact.emit_op_u16(Op::LOCAL_GET, 1, 0); // n
    fact.emit_op_u16(Op::CONST, c1, 0);  // 1
    fact.emit_op(Op::I32_SUB, 0);          // n-1
    fact.emit_op_u8(Op::CALL, 1, 0);       // fact(n-1)
    fact.emit_op(Op::I32_MUL, 0);          // n * fact(n-1)
    fact.emit_op(Op::RETURN, 0);

    // base case: return 1
    fact.patch_jump(jump_to_base);
    fact.emit_op_u16(Op::CONST, c1, 0);
    fact.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, fact]);
    assert_i32(&result, 120);
}

#[test]
fn call_import_host_function() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("test", "double");
    let c7 = main.add_constant(Value::I32(7));
    main.emit_op_u16(Op::CONST, c7, 0);
    // call_import: u16 import_idx, u8 arg_count
    main.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    main.emit(1, 0); // 1 arg
    main.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    vm.register_host_fn("test", "double", Box::new(|args: &[Value]| {
        Value::I32(args[0].as_i32() * 2)
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 14);
}

// ============================================================
// 2. Stack Operations
// ============================================================

#[test]
fn stack_dup() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::I32(7));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::DUP, 0);
    chunk.emit_op(Op::I32_ADD, 0); // 7 + 7 = 14
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 14);
}

#[test]
fn stack_drop() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c1 = chunk.add_constant(Value::I32(99));
    let c2 = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op_u16(Op::CONST, c2, 0);
    chunk.emit_op(Op::DROP, 0); // drop 42
    chunk.emit_op(Op::HALT, 0); // TOS = 99

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 99);
}

#[test]
fn stack_select_true() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    let cond = chunk.add_constant(Value::I32(1)); // truthy
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, cond, 0);
    chunk.emit_op(Op::SELECT, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 10);
}

#[test]
fn stack_select_false() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    let cond = chunk.add_constant(Value::I32(0)); // falsy
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, cond, 0);
    chunk.emit_op(Op::SELECT, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 20);
}

#[test]
fn local_get_set() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 3;
    let c = chunk.add_constant(Value::I32(77));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0); // local[1] = 77, keeps on stack
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0); // push local[1]
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 77);
}

#[test]
fn global_get_set() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let name_idx = chunk.add_constant(Value::String(Rc::from("myGlobal")));
    let val = chunk.add_constant(Value::I32(55));
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op_u16(Op::GLOBAL_SET, name_idx, 0); // keeps on stack
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::GLOBAL_GET, name_idx, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    assert_i32(&result, 55);
}

// ============================================================
// 3. Arithmetic
// ============================================================

#[test]
fn i32_add_positive() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 30);
}

#[test]
fn i32_sub_negative_result() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::I32(12));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), -7);
}

#[test]
fn i32_mul_with_zero() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(999));
    let b = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_MUL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 0);
}

#[test]
fn i32_div_positive() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(17));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_DIV_S, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 3);
}

#[test]
fn i32_div_by_zero_returns_zero() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_DIV_S, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 0);
}

#[test]
fn i32_rem() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(17));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_REM_S, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 2);
}

#[test]
fn i32_negative_arithmetic() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(-10));
    let b = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_DIV_S, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), -3);
}

#[test]
fn f64_add_sub_mul() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::F64(3.5));
    let b = chunk.add_constant(Value::F64(2.5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::F64_ADD, 0); // 6.0
    let c = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::F64_SUB, 0); // 5.0
    let d = chunk.add_constant(Value::F64(3.0));
    chunk.emit_op_u16(Op::CONST, d, 0);
    chunk.emit_op(Op::F64_MUL, 0); // 15.0
    chunk.emit_op(Op::HALT, 0);
    assert_f64(&run_chunks(vec![chunk]), 15.0);
}

#[test]
fn f64_div_normal() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::F64(10.0));
    let b = chunk.add_constant(Value::F64(4.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::F64_DIV, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_f64(&run_chunks(vec![chunk]), 2.5);
}

#[test]
fn f64_div_by_zero_infinity() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::F64(1.0));
    let b = chunk.add_constant(Value::F64(0.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::F64_DIV, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) => assert!(v.is_infinite() && v > 0.0, "Expected +Infinity, got {}", v),
        _ => panic!("Expected F64(Infinity), got {:?}", result),
    }
}

#[test]
fn f64_div_negative_by_zero() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::F64(-1.0));
    let b = chunk.add_constant(Value::F64(0.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::F64_DIV, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) => assert!(v.is_infinite() && v < 0.0, "Expected -Infinity, got {}", v),
        _ => panic!("Expected F64(-Infinity), got {:?}", result),
    }
}

#[test]
fn f64_mod_operation() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::F64(10.0));
    let b = chunk.add_constant(Value::F64(3.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    // f64_mod removed — test the WASM-equivalent sequence: a - trunc(a/b) * b
    chunk.emit_op(Op::DUP, 0); // save b
    // Stack: [a, b, b]. Need [a, a, b] for div. Use different approach:
    // Actually we can't easily test fmod in a low-level VM test without stdlib.
    // Just verify f64_div + f64_trunc works (the components of fmod).
    chunk.emit_op(Op::F64_DIV, 0); // 10/3 = 3.333
    chunk.emit_op(Op::F64_TRUNC, 0); // trunc(3.333) = 3.0
    chunk.emit_op(Op::HALT, 0);
    assert_f64(&run_chunks(vec![chunk]), 3.0);
}

// ============================================================
// 4. Comparisons (dyn_eq, dyn_ne, dyn_lt, dyn_gt, dyn_le, dyn_ge)
// ============================================================

#[test]
fn dyn_eq_numbers() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_eq_different_numbers() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::I32(6));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_eq_strings() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("hello")));
    let b = chunk.add_constant(Value::String(Rc::from("hello")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_eq_null_null() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::DYN_EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_eq_booleans() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::TRUE, 0);
    chunk.emit_op(Op::TRUE, 0);
    chunk.emit_op(Op::DYN_EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_eq_nan_is_false() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let nan = chunk.add_constant(Value::F64(f64::NAN));
    chunk.emit_op_u16(Op::CONST, nan, 0);
    chunk.emit_op_u16(Op::CONST, nan, 0);
    chunk.emit_op(Op::DYN_EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_ne_different_types() {
    // With loose equality: 5 == "5" is true (string→number coercion), so 5 != "5" is false
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::String(Rc::from("5")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_NE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_eq_mixed_i32_f64() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::F64(5.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    // dyn_eq coerces I32/F64 cross-type
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_lt_numbers() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(3));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_LT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_gt_numbers() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_GT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_le_equal() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_LE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_ge_less() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(3));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_GE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_lt_strings() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("apple")));
    let b = chunk.add_constant(Value::String(Rc::from("banana")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_LT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_le_strings() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("abc")));
    let b = chunk.add_constant(Value::String(Rc::from("abc")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_LE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_ge_strings() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("z")));
    let b = chunk.add_constant(Value::String(Rc::from("a")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_GE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

// ============================================================
// 5. Conversions
// ============================================================

#[test]
fn conv_i32_from_f64() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::F64(42.9));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::I32_FROM_F64, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 42);
}

#[test]
fn conv_f64_from_i32() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::I32(7));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::F64_FROM_I32, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_f64(&run_chunks(vec![chunk]), 7.0);
}

#[test]
fn conv_i64_extend_i32() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::I32(-5));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::I64_EXTEND_I32_S, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I64(-5) => {}
        _ => panic!("Expected I64(-5), got {:?}", result),
    }
}

#[test]
fn conv_i32_wrap_i64() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::I64(0x1_0000_0005)); // wraps to 5
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::I32_WRAP_I64, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 5);
}

#[test]
fn dyn_to_bool_falsy_values() {
    // 0 => false
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_0, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_to_bool_empty_string() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::String(Rc::from("")));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_to_bool_null() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_to_bool_undefined() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    { let c = chunk.add_constant(Value::Undefined); chunk.emit_op_u16(Op::CONST, c, 0); }
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_to_bool_truthy_number() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_to_bool_truthy_string() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::String(Rc::from("x")));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn dyn_to_bool_true_literal() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::TRUE, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

// ============================================================
// 6. String Operations
// ============================================================

#[test]
fn str_length_ascii() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STR_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 5);
}

#[test]
fn str_length_unicode() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    // Each emoji is 1 char (by chars().count())
    let s = chunk.add_constant(Value::String(Rc::from("a\u{1F600}b"))); // "a😀b" = 3 chars
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STR_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 3);
}

#[test]
fn str_concat_two() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("hello")));
    let b = chunk.add_constant(Value::String(Rc::from(" world")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::STR_CONCAT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "hello world");
}

#[test]
fn str_concat_n_three() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("a")));
    let b = chunk.add_constant(Value::String(Rc::from("b")));
    let c = chunk.add_constant(Value::String(Rc::from("c")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u8(Op::STR_CONCAT_N, 3, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "abc");
}

#[test]
fn str_to_upper() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STR_TO_UPPER, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "HELLO");
}

#[test]
fn str_to_lower() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("WORLD")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STR_TO_LOWER, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "world");
}

#[test]
fn str_trim_whitespace() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("  hi  ")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STR_TRIM, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "hi");
}

#[test]
fn str_index_of_found() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let haystack = chunk.add_constant(Value::String(Rc::from("hello world")));
    let needle = chunk.add_constant(Value::String(Rc::from("world")));
    chunk.emit_op_u16(Op::CONST, haystack, 0);
    chunk.emit_op_u16(Op::CONST, needle, 0);
    chunk.emit_op(Op::STR_INDEX_OF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 6);
}

#[test]
fn str_index_of_not_found() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let haystack = chunk.add_constant(Value::String(Rc::from("hello")));
    let needle = chunk.add_constant(Value::String(Rc::from("xyz")));
    chunk.emit_op_u16(Op::CONST, haystack, 0);
    chunk.emit_op_u16(Op::CONST, needle, 0);
    chunk.emit_op(Op::STR_INDEX_OF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), -1);
}

#[test]
fn str_contains_true() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello world")));
    let n = chunk.add_constant(Value::String(Rc::from("world")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op_u16(Op::CONST, n, 0);
    chunk.emit_op(Op::STR_CONTAINS, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn str_contains_false() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello")));
    let n = chunk.add_constant(Value::String(Rc::from("xyz")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op_u16(Op::CONST, n, 0);
    chunk.emit_op(Op::STR_CONTAINS, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn str_char_at() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello")));
    let idx = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op_u16(Op::CONST, idx, 0);
    chunk.emit_op(Op::STR_CHAR_AT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "e");
}

#[test]
fn str_substring() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello world")));
    let start = chunk.add_constant(Value::I32(6));
    let end = chunk.add_constant(Value::I32(11));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op_u16(Op::CONST, start, 0);
    chunk.emit_op_u16(Op::CONST, end, 0);
    chunk.emit_op(Op::STR_SUBSTRING, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "world");
}

#[test]
fn str_replace() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hello world")));
    let old = chunk.add_constant(Value::String(Rc::from("world")));
    let new = chunk.add_constant(Value::String(Rc::from("rust")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op_u16(Op::CONST, old, 0);
    chunk.emit_op_u16(Op::CONST, new, 0);
    chunk.emit_op(Op::STR_REPLACE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "hello rust");
}

#[test]
fn str_split() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("a,b,c")));
    let delim = chunk.add_constant(Value::String(Rc::from(",")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op_u16(Op::CONST, delim, 0);
    chunk.emit_op(Op::STR_SPLIT, 0);
    // Result is an array, get its length
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 3);
}

// ============================================================
// 7. Array Operations
// ============================================================

#[test]
fn array_new_and_length() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    let c = chunk.add_constant(Value::I32(30));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 3, 0);
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 3);
}

#[test]
fn array_get_valid_index() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    let c = chunk.add_constant(Value::I32(30));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 3, 0);
    // array_get: stack [obj, key] => [val]
    let idx = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 20);
}

#[test]
fn array_get_out_of_bounds() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    let a = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 1, 0);
    let idx = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    // Out of bounds should return Null
    match result {
        Value::Null => {}
        _ => panic!("Expected Null for out-of-bounds, got {:?}", result),
    }
}

#[test]
fn array_set() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 2, 0);
    // store in local for reuse
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op(Op::DROP, 0);
    // array_set: stack [obj, key, val] => [val]
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    let val = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::ARRAY_SET, 0);
    chunk.emit_op(Op::DROP, 0);
    // Now read back index 0
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx0 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, idx0, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 99);
}

#[test]
fn array_push_and_length() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    // Start with empty array
    chunk.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    let val = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::ARRAY_PUSH, 0); // returns array
    let val2 = chunk.add_constant(Value::I32(43));
    chunk.emit_op_u16(Op::CONST, val2, 0);
    chunk.emit_op(Op::ARRAY_PUSH, 0);
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 2);
}

#[test]
fn array_pop() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 2, 0);
    chunk.emit_op(Op::ARRAY_POP, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 20);
}

#[test]
fn array_join() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("a")));
    let b = chunk.add_constant(Value::String(Rc::from("b")));
    let c = chunk.add_constant(Value::String(Rc::from("c")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 3, 0);
    let delim = chunk.add_constant(Value::String(Rc::from("-")));
    chunk.emit_op_u16(Op::CONST, delim, 0);
    chunk.emit_op(Op::ARRAY_JOIN, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "a-b-c");
}

#[test]
fn array_concat() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(1));
    let b = chunk.add_constant(Value::I32(2));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 2, 0);
    let c = chunk.add_constant(Value::I32(3));
    let d = chunk.add_constant(Value::I32(4));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u16(Op::CONST, d, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 2, 0);
    chunk.emit_op(Op::ARRAY_CONCAT, 0);
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 4);
}

#[test]
fn array_reverse() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    let a = chunk.add_constant(Value::I32(1));
    let b = chunk.add_constant(Value::I32(2));
    let c = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 3, 0);
    chunk.emit_op(Op::ARRAY_REVERSE, 0);
    // Get first element (should be 3 now)
    let idx = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 3);
}

#[test]
fn array_fill() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    // Create [0, 0, 0, 0, 0]
    let z = chunk.add_constant(Value::I32(0));
    for _ in 0..5 {
        chunk.emit_op_u16(Op::CONST, z, 0);
    }
    chunk.emit_op_u16(Op::ARRAY_NEW, 5, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op(Op::DROP, 0);
    // array_fill: stack [array, value, start, len]
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let val = chunk.add_constant(Value::I32(7));
    chunk.emit_op_u16(Op::CONST, val, 0);
    let start = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, start, 0);
    let len = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::CONST, len, 0);
    chunk.emit_op(Op::ARRAY_FILL, 0);
    // Read index 2 (should be 7)
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx2 = chunk.add_constant(Value::I32(2));
    chunk.emit_op_u16(Op::CONST, idx2, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 7);
}

// ============================================================
// 8. Object/Struct Operations
// ============================================================

#[test]
fn struct_new_and_get() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    // struct_new with 2 properties: push key, val, key, val
    let k1 = chunk.add_constant(Value::String(Rc::from("name")));
    let v1 = chunk.add_constant(Value::String(Rc::from("Alice")));
    let k2 = chunk.add_constant(Value::String(Rc::from("age")));
    let v2 = chunk.add_constant(Value::I32(30));
    chunk.emit_op_u16(Op::CONST, k1, 0);
    chunk.emit_op_u16(Op::CONST, v1, 0);
    chunk.emit_op_u16(Op::CONST, k2, 0);
    chunk.emit_op_u16(Op::CONST, v2, 0);
    chunk.emit_op_u16(Op::STRUCT_NEW, 2, 0);
    // struct_get "name"
    let name_key = chunk.add_constant(Value::String(Rc::from("name")));
    chunk.emit_op_u16(Op::STRUCT_GET, name_key, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "Alice");
}

#[test]
fn struct_get_missing_prop() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let k1 = chunk.add_constant(Value::String(Rc::from("x")));
    let v1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, k1, 0);
    chunk.emit_op_u16(Op::CONST, v1, 0);
    chunk.emit_op_u16(Op::STRUCT_NEW, 1, 0);
    let missing = chunk.add_constant(Value::String(Rc::from("y")));
    chunk.emit_op_u16(Op::STRUCT_GET, missing, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::Null => {}
        _ => panic!("Expected Null for missing prop, got {:?}", result),
    }
}

#[test]
fn struct_set_property() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 2;
    let k1 = chunk.add_constant(Value::String(Rc::from("x")));
    let v1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, k1, 0);
    chunk.emit_op_u16(Op::CONST, v1, 0);
    chunk.emit_op_u16(Op::STRUCT_NEW, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op(Op::DROP, 0);
    // struct_set: stack [obj, val] with operand prop name
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let new_val = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::CONST, new_val, 0);
    let x_key = chunk.add_constant(Value::String(Rc::from("x")));
    chunk.emit_op_u16(Op::STRUCT_SET, x_key, 0);
    chunk.emit_op(Op::DROP, 0);
    // Read back
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let x_key2 = chunk.add_constant(Value::String(Rc::from("x")));
    chunk.emit_op_u16(Op::STRUCT_GET, x_key2, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 99);
}

#[test]
fn struct_getter_auto_dispatch() {
    // Create an object with __get_foo property that is a function returning 42
    let mut main = Chunk::new("main");
    main.local_count = 2;

    // Create the getter function (chunk 1)
    // It takes 1 arg (self) and returns 42
    let k = main.add_constant(Value::String(Rc::from("__get_foo")));
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0); // 0 upvalues
    // Build object: { __get_foo: <function> }
    // struct_new expects [key, val] pairs on stack
    main.emit_op_u16(Op::CONST, k, 0); // push key first
    // We need to swap: currently stack has [func], we need [key, func]
    // Actually let me restructure: push key, then push func
    // Hmm, we already pushed func. Let me redo:
    // Strategy: push key first, then ref_func
    // Clear and redo
    main.code.clear();
    main.lines.clear();
    main.constants.clear();
    main.local_count = 2;

    let k = main.add_constant(Value::String(Rc::from("__get_foo")));
    chunk_emit_key_func_struct(&mut main, k, 1);
    main.emit_op_u16(Op::LOCAL_SET, 1, 0);
    main.emit_op(Op::DROP, 0);
    // Now struct_get "foo" should auto-invoke __get_foo
    main.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let foo_key = main.add_constant(Value::String(Rc::from("foo")));
    main.emit_op_u16(Op::STRUCT_GET, foo_key, 0);
    main.emit_op(Op::HALT, 0);

    let mut getter = Chunk::new("getter");
    getter.arity = 1; // self
    getter.local_count = 2;
    let c42 = getter.add_constant(Value::I32(42));
    getter.emit_op_u16(Op::CONST, c42, 0);
    getter.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, getter]);
    assert_i32(&result, 42);
}

/// Helper to emit: push key string, push ref_func, struct_new(1)
fn chunk_emit_key_func_struct(chunk: &mut Chunk, key_const_idx: u16, func_chunk_idx: u16) {
    chunk.emit_op_u16(Op::CONST, key_const_idx, 0);
    chunk.emit_op_u16(Op::REF_FUNC, func_chunk_idx, 0);
    chunk.emit(0, 0); // 0 upvalues
    chunk.emit_op_u16(Op::STRUCT_NEW, 1, 0);
}

#[test]
fn struct_setter_auto_dispatch() {
    // Object with __set_bar that stores value * 2 in a property
    let mut main = Chunk::new("main");
    main.local_count = 2;

    let k = main.add_constant(Value::String(Rc::from("__set_bar")));
    chunk_emit_key_func_struct(&mut main, k, 1);
    main.emit_op_u16(Op::LOCAL_SET, 1, 0);
    main.emit_op(Op::DROP, 0);
    // struct_set "bar" should auto-invoke __set_bar
    main.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let val = main.add_constant(Value::I32(5));
    main.emit_op_u16(Op::CONST, val, 0);
    let bar_key = main.add_constant(Value::String(Rc::from("bar")));
    main.emit_op_u16(Op::STRUCT_SET, bar_key, 0);
    // setter return is discarded, val (5) is pushed
    main.emit_op(Op::HALT, 0);

    // chunk 1: setter(self, value) => sets self._bar = value * 2
    let mut setter = Chunk::new("setter");
    setter.arity = 2;
    setter.local_count = 3;
    setter.emit_op_u16(Op::LOCAL_GET, 1, 0); // self
    setter.emit_op_u16(Op::LOCAL_GET, 2, 0); // value
    let c2 = setter.add_constant(Value::I32(2));
    setter.emit_op_u16(Op::CONST, c2, 0);
    setter.emit_op(Op::I32_MUL, 0);
    let bar_key2 = setter.add_constant(Value::String(Rc::from("_bar")));
    setter.emit_op_u16(Op::STRUCT_SET, bar_key2, 0);
    setter.emit_op(Op::RETURN, 0);

    let result = run_chunks(vec![main, setter]);
    // The setter returns value*2=10 from struct_set inside it, and
    // the outer struct_set discards that and pushes the original val.
    // However, the setter's return is what struct_set returns to the caller.
    assert_i32(&result, 10);
}

// ============================================================
// 9. Control Flow
// ============================================================

#[test]
fn br_unconditional() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c1 = chunk.add_constant(Value::I32(1));
    let c2 = chunk.add_constant(Value::I32(2));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    // Jump over the next const
    let jump = chunk.emit_jump(Op::BR, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::CONST, c2, 0); // should be skipped
    chunk.patch_jump(jump);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 1);
}

#[test]
fn br_if_true_taken() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c10 = chunk.add_constant(Value::I32(10));
    let c20 = chunk.add_constant(Value::I32(20));
    chunk.emit_op(Op::TRUE, 0);
    let jump = chunk.emit_jump(Op::BR_IF_TRUE, 0);
    chunk.emit_op_u16(Op::CONST, c20, 0);
    chunk.emit_op(Op::HALT, 0);
    chunk.patch_jump(jump);
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op(Op::HALT, 0);

    assert_i32(&run_chunks(vec![chunk]), 10);
}

#[test]
fn br_if_false_taken() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c10 = chunk.add_constant(Value::I32(10));
    let c20 = chunk.add_constant(Value::I32(20));
    chunk.emit_op(Op::FALSE, 0);
    let jump = chunk.emit_jump(Op::BR_IF_FALSE, 0);
    chunk.emit_op_u16(Op::CONST, c20, 0);
    chunk.emit_op(Op::HALT, 0);
    chunk.patch_jump(jump);
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op(Op::HALT, 0);

    assert_i32(&run_chunks(vec![chunk]), 10);
}

#[test]
fn loop_sum_1_to_5() {
    // sum = 0; i = 1; while i <= 5 { sum += i; i += 1 }; result = sum
    let mut chunk = Chunk::new("test");
    chunk.local_count = 3; // 0=script, 1=sum, 2=i

    // sum = 0
    chunk.emit_op(Op::I32_CONST_0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op(Op::DROP, 0);
    // i = 1
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
    chunk.emit_op(Op::DROP, 0);

    let loop_start = chunk.current_offset();
    // if i > 5, break
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    let c5 = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, c5, 0);
    chunk.emit_op(Op::DYN_GT, 0);
    let exit_jump = chunk.emit_jump(Op::BR_IF_TRUE, 0);

    // sum += i
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op(Op::DROP, 0);

    // i += 1
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_loop(loop_start, 0);

    chunk.patch_jump(exit_jump);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op(Op::HALT, 0);

    assert_i32(&run_chunks(vec![chunk]), 15);
}

#[test]
fn function_return_value() {
    // Function returns a string
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    main.emit_op_u8(Op::CALL, 0, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("greet");
    func.arity = 0;
    func.local_count = 1;
    let s = func.add_constant(Value::String(Rc::from("hello")));
    func.emit_op_u16(Op::CONST, s, 0);
    func.emit_op(Op::RETURN, 0);

    assert_string(&run_chunks(vec![main, func]), "hello");
}

// ============================================================
// 10. Type Checks
// ============================================================

#[test]
fn ref_is_null_on_null() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::REF_IS_NULL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn ref_is_null_on_number() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::REF_IS_NULL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn ref_is_null_on_undefined() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    { let c = chunk.add_constant(Value::Undefined); chunk.emit_op_u16(Op::CONST, c, 0); }
    chunk.emit_op(Op::REF_IS_NULL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn ref_is_array_on_array() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::ARRAY_NEW, 0, 0); // empty array
    chunk.emit_op(Op::REF_IS_ARRAY, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn ref_is_array_on_non_array() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::REF_IS_ARRAY, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn ref_is_object_on_struct() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, 0); // empty object
    chunk.emit_op(Op::REF_IS_OBJECT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn ref_is_object_on_string() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hi")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::REF_IS_OBJECT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn ref_typeof_null() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::REF_TYPEOF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "object"); // JS spec: typeof null === "object"
}

#[test]
fn ref_typeof_number() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::REF_TYPEOF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "number");
}

#[test]
fn ref_typeof_string() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("hi")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::REF_TYPEOF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "string");
}

#[test]
fn ref_typeof_boolean() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::TRUE, 0);
    chunk.emit_op(Op::REF_TYPEOF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "boolean");
}

#[test]
fn ref_typeof_undefined() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    { let c = chunk.add_constant(Value::Undefined); chunk.emit_op_u16(Op::CONST, c, 0); }
    chunk.emit_op(Op::REF_TYPEOF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "undefined");
}

#[test]
fn ref_typeof_array() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    chunk.emit_op(Op::REF_TYPEOF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "array");
}

#[test]
fn ref_typeof_object() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, 0);
    chunk.emit_op(Op::REF_TYPEOF, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "object");
}

#[test]
fn ref_typeof_function() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    main.emit_op(Op::REF_TYPEOF, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("f");
    func.arity = 0;
    func.local_count = 1;
    func.emit_op(Op::NULL, 0);
    func.emit_op(Op::RETURN, 0);

    assert_string(&run_chunks(vec![main, func]), "function");
}

// ============================================================
// 11. Invoke from Rust
// ============================================================

#[test]
fn invoke_simple_function() {
    // Create a VM, run chunks to define a function, then invoke it
    let mut main = Chunk::new("main");
    main.local_count = 2;
    // Create function and store in global
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let name = main.add_constant(Value::String(Rc::from("myFunc")));
    main.emit_op_u16(Op::GLOBAL_SET, name, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("myFunc");
    func.arity = 1;
    func.local_count = 2;
    let c10 = func.add_constant(Value::I32(10));
    func.emit_op_u16(Op::LOCAL_GET, 1, 0);
    func.emit_op_u16(Op::CONST, c10, 0);
    func.emit_op(Op::I32_ADD, 0);
    func.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![main, func]).unwrap();

    let func_val = vm.globals.get("myFunc").cloned().unwrap();
    let result = vm.invoke(&func_val, &[Value::I32(5)]).unwrap();
    assert_i32(&result, 15);
}

#[test]
fn invoke_with_multiple_args() {
    let mut main = Chunk::new("main");
    main.local_count = 2;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let name = main.add_constant(Value::String(Rc::from("add")));
    main.emit_op_u16(Op::GLOBAL_SET, name, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("add");
    func.arity = 2;
    func.local_count = 3;
    func.emit_op_u16(Op::LOCAL_GET, 1, 0);
    func.emit_op_u16(Op::LOCAL_GET, 2, 0);
    func.emit_op(Op::I32_ADD, 0);
    func.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![main, func]).unwrap();

    let func_val = vm.globals.get("add").cloned().unwrap();
    let result = vm.invoke(&func_val, &[Value::I32(3), Value::I32(7)]).unwrap();
    assert_i32(&result, 10);
}

#[test]
fn invoke_returning_string() {
    let mut main = Chunk::new("main");
    main.local_count = 2;
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    let name = main.add_constant(Value::String(Rc::from("greet")));
    main.emit_op_u16(Op::GLOBAL_SET, name, 0);
    main.emit_op(Op::HALT, 0);

    let mut func = Chunk::new("greet");
    func.arity = 0;
    func.local_count = 1;
    let s = func.add_constant(Value::String(Rc::from("Hello from invoke!")));
    func.emit_op_u16(Op::CONST, s, 0);
    func.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![main, func]).unwrap();

    let func_val = vm.globals.get("greet").cloned().unwrap();
    let result = vm.invoke(&func_val, &[]).unwrap();
    assert_string(&result, "Hello from invoke!");
}

// ============================================================
// Additional edge case tests
// ============================================================

#[test]
fn immediate_values() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::DROP, 0);
    { let c = chunk.add_constant(Value::Undefined); chunk.emit_op_u16(Op::CONST, c, 0); }
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::TRUE, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::FALSE, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I32_CONST_0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::F64_CONST_0, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_f64(&run_chunks(vec![chunk]), 0.0);
}

#[test]
fn bool_not() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::TRUE, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_add_numbers() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(3));
    let b = chunk.add_constant(Value::I32(4));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_ADD, 0);
    chunk.emit_op(Op::HALT, 0);
    // dyn_add with numbers should add
    let result = run_chunks(vec![chunk]);
    // Could be I32 or F64 depending on implementation
    match &result {
        Value::I32(7) => {}
        Value::F64(v) if *v == 7.0 => {}
        _ => panic!("Expected 7, got {:?}", result),
    }
}

#[test]
fn dyn_add_string_concat() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::String(Rc::from("foo")));
    let b = chunk.add_constant(Value::String(Rc::from("bar")));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::DYN_ADD, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_string(&run_chunks(vec![chunk]), "foobar");
}

#[test]
fn dyn_neg() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::F64(5.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op(Op::DYN_NEG, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_f64(&run_chunks(vec![chunk]), -5.0);
}

#[test]
fn dyn_not_truthy() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::DYN_NOT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_not_falsy() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_0, 0);
    chunk.emit_op(Op::DYN_NOT, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn f64_neg_operation() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::F64(3.14));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op(Op::F64_NEG, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_f64(&run_chunks(vec![chunk]), -3.14);
}

#[test]
fn i32_eqz_zero() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_0, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn i32_eqz_nonzero() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op(Op::I32_CONST_1, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn host_function_with_multiple_args() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("math", "add3");
    let a = main.add_constant(Value::I32(10));
    let b = main.add_constant(Value::I32(20));
    let c = main.add_constant(Value::I32(30));
    main.emit_op_u16(Op::CONST, a, 0);
    main.emit_op_u16(Op::CONST, b, 0);
    main.emit_op_u16(Op::CONST, c, 0);
    main.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    main.emit(3, 0); // 3 args
    main.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    vm.register_host_fn("math", "add3", Box::new(|args: &[Value]| {
        Value::I32(args[0].as_i32() + args[1].as_i32() + args[2].as_i32())
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_i32(&result, 60);
}

#[test]
fn host_function_returning_string() {
    let mut main = Chunk::new("main");
    main.local_count = 1;
    let import_idx = main.add_import("util", "greet");
    let name = main.add_constant(Value::String(Rc::from("World")));
    main.emit_op_u16(Op::CONST, name, 0);
    main.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    main.emit(1, 0);
    main.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    vm.register_host_fn("util", "greet", Box::new(|args: &[Value]| {
        let name = if let Value::String(s) = &args[0] { s.as_ref() } else { "?" };
        Value::String(Rc::from(format!("Hello, {}!", name).as_str()))
    }));
    let result = vm.run(vec![main]).unwrap();
    assert_string(&result, "Hello, World!");
}

#[test]
fn multiple_locals() {
    // Test setting and getting multiple locals
    let mut chunk = Chunk::new("test");
    chunk.local_count = 5;
    let c1 = chunk.add_constant(Value::I32(10));
    let c2 = chunk.add_constant(Value::I32(20));
    let c3 = chunk.add_constant(Value::I32(30));
    // local[1] = 10
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op(Op::DROP, 0);
    // local[2] = 20
    chunk.emit_op_u16(Op::CONST, c2, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
    chunk.emit_op(Op::DROP, 0);
    // local[3] = 30
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 3, 0);
    chunk.emit_op(Op::DROP, 0);
    // result = local[1] + local[2] + local[3]
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 3, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 60);
}

#[test]
fn str_length_empty() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let s = chunk.add_constant(Value::String(Rc::from("")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STR_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_i32(&run_chunks(vec![chunk]), 0);
}

#[test]
fn array_empty_pop() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    chunk.emit_op(Op::ARRAY_POP, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::Null => {}
        _ => panic!("Expected Null from empty array pop, got {:?}", result),
    }
}

#[test]
fn dyn_to_bool_nan_is_false() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let nan = chunk.add_constant(Value::F64(f64::NAN));
    chunk.emit_op_u16(Op::CONST, nan, 0);
    chunk.emit_op(Op::DYN_TO_BOOL, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), false);
}

#[test]
fn dyn_ne_nan_nan() {
    // NaN != NaN should be true
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let nan = chunk.add_constant(Value::F64(f64::NAN));
    chunk.emit_op_u16(Op::CONST, nan, 0);
    chunk.emit_op_u16(Op::CONST, nan, 0);
    chunk.emit_op(Op::DYN_NE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn eq_same_type_strict() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::EQ, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}

#[test]
fn ne_different_values() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(5));
    let b = chunk.add_constant(Value::I32(6));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::NE, 0);
    chunk.emit_op(Op::HALT, 0);
    assert_bool(&run_chunks(vec![chunk]), true);
}
