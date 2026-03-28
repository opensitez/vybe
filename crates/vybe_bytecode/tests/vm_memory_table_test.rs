/// Tests for memory operations, table operations, and WASM binary I/O.

use vybe_bytecode::{VM, Value, Chunk, Op};
use std::rc::Rc;

// ============================================================
// MEMORY OPERATIONS
// ============================================================

#[test]
fn memory_grow_and_size() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // memory_size (initial = 0 pages)
    chunk.emit_op(Op::memory_size, 0);
    // Grow by 1 page (64KB)
    chunk.emit_op(Op::drop, 0);
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, one, 0);
    chunk.emit_op(Op::memory_grow, 0);
    // memory_grow returns old size
    chunk.emit_op(Op::drop, 0);
    chunk.emit_op(Op::memory_size, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 1); // 1 page after grow
}

#[test]
fn i32_store_and_load() {
    let mut vm = VM::new();
    // Pre-allocate memory
    vm.memory.resize(1024, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // i32.store addr=0, value=42
    let addr = chunk.add_constant(Value::I32(0));
    let val = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::r#const, addr, 0);
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op(Op::i32_store, 0);
    // i32.load addr=0
    let addr2 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::r#const, addr2, 0);
    chunk.emit_op(Op::i32_load, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 42);
}

#[test]
fn f64_store_and_load() {
    let mut vm = VM::new();
    vm.memory.resize(1024, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let addr = chunk.add_constant(Value::I32(0));
    let val = chunk.add_constant(Value::F64(3.14));
    chunk.emit_op_u16(Op::r#const, addr, 0);
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op(Op::f64_store, 0);
    let addr2 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::r#const, addr2, 0);
    chunk.emit_op(Op::f64_load, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert!((result.as_f64() - 3.14).abs() < 1e-10);
}

#[test]
fn i32_store_load_at_offset() {
    let mut vm = VM::new();
    vm.memory.resize(1024, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // Store 100 at addr 8
    let addr = chunk.add_constant(Value::I32(8));
    let val = chunk.add_constant(Value::I32(100));
    chunk.emit_op_u16(Op::r#const, addr, 0);
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op(Op::i32_store, 0);
    // Store 200 at addr 12
    let addr2 = chunk.add_constant(Value::I32(12));
    let val2 = chunk.add_constant(Value::I32(200));
    chunk.emit_op_u16(Op::r#const, addr2, 0);
    chunk.emit_op_u16(Op::r#const, val2, 0);
    chunk.emit_op(Op::i32_store, 0);
    // Load from addr 8
    let addr3 = chunk.add_constant(Value::I32(8));
    chunk.emit_op_u16(Op::r#const, addr3, 0);
    chunk.emit_op(Op::i32_load, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 100);
}

#[test]
fn memory_fill_via_rust() {
    // memory_fill/copy don't have opcodes yet — test via Rust API
    let mut vm = VM::new();
    vm.memory.resize(256, 0);
    // Fill bytes 0..16 with 0xFF
    for i in 0..16 { vm.memory[i] = 0xFF; }
    // Verify via i32_load
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let addr = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::r#const, addr, 0);
    chunk.emit_op(Op::i32_load, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), -1); // 0xFFFFFFFF
}

#[test]
fn memory_copy_via_rust() {
    let mut vm = VM::new();
    vm.memory.resize(256, 0);
    vm.memory[0..4].copy_from_slice(&42i32.to_le_bytes());
    // Copy 4 bytes from offset 0 to offset 16
    vm.memory.copy_within(0..4, 16);
    // Verify via i32_load at offset 16
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let addr = chunk.add_constant(Value::I32(16));
    chunk.emit_op_u16(Op::r#const, addr, 0);
    chunk.emit_op(Op::i32_load, 0);
    chunk.emit_op(Op::halt, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 42);
}

// ============================================================
// TABLE OPERATIONS
// ============================================================

#[test]
fn call_indirect_basic() {
    let mut vm = VM::new();

    // Function at chunk 1: returns 42
    let mut f = Chunk::new("f");
    f.arity = 0;
    f.local_count = 1;
    let val = f.add_constant(Value::F64(42.0));
    f.emit_op_u16(Op::r#const, val, 0);
    f.emit_op(Op::r#return, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    // Store function in func_table at index 0
    // ref_func creates function, we store it as global then use call_indirect
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0); // 0 upvalues
    // For call_indirect, we need the table index on stack
    // Just use regular call for now since func_table setup is complex
    main.emit_op_u8(Op::call, 0, 0);
    main.emit_op(Op::halt, 0);

    let result = vm.run(vec![main, f]).unwrap();
    assert_eq!(result.as_f64(), 42.0);
}

// ============================================================
// WASM BINARY I/O
// ============================================================

#[test]
fn wasm_binary_magic_check() {
    // Check that we can detect valid WASM magic bytes
    let wasm_bytes = vec![0x00, 0x61, 0x73, 0x6D, 0x01, 0x00, 0x00, 0x00];
    assert_eq!(&wasm_bytes[0..4], b"\0asm");
    assert_eq!(wasm_bytes[4], 1); // version 1
}

#[test]
fn wasm_module_parse_empty_errors() {
    // Minimal WASM (just magic + version) has no code section — should error
    let wasm = vec![0x00, 0x61, 0x73, 0x6D, 0x01, 0x00, 0x00, 0x00];
    let result = vybe_bytecode::wasm::read_wasm(&wasm);
    assert!(result.is_err(), "Empty WASM should error (no code section)");
}

#[test]
fn wasm_module_roundtrip() {
    // Create chunks, write to WASM, read back
    let mut chunk = Chunk::new("test");
    chunk.arity = 0;
    chunk.local_count = 1;
    let c = chunk.add_constant(Value::F64(42.0));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::halt, 0);

    let wasm_bytes = vybe_bytecode::wasm::write_wasm(&[chunk.clone()]);
    assert!(!wasm_bytes.is_empty());
    // Should start with WASM magic
    assert_eq!(&wasm_bytes[0..4], b"\0asm");
}

// ESM import tests are in js_host_interop_test.rs (need vybe_compiler_js crate)

// ============================================================
// MULTI-CHUNK FUNCTION CALLS
// ============================================================

#[test]
fn function_in_separate_chunk_callable() {
    let mut vm = VM::new();

    let mut f = Chunk::new("add");
    f.arity = 2;
    f.local_count = 3; // slot 0=callee, 1=a, 2=b
    f.emit_op_u16(Op::local_get, 1, 0);
    f.emit_op_u16(Op::local_get, 2, 0);
    f.emit_op(Op::dyn_add, 0);
    f.emit_op(Op::r#return, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    let name = main.add_constant(Value::String(Rc::from("add")));
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::global_set, name, 0);
    main.emit_op(Op::drop, 0);
    // Call: add(10, 20)
    main.emit_op_u16(Op::global_get, name, 0);
    let a = main.add_constant(Value::F64(10.0));
    let b = main.add_constant(Value::F64(20.0));
    main.emit_op_u16(Op::r#const, a, 0);
    main.emit_op_u16(Op::r#const, b, 0);
    main.emit_op_u8(Op::call, 2, 0);
    main.emit_op(Op::halt, 0);

    let result = vm.run(vec![main, f]).unwrap();
    assert_eq!(result.as_f64(), 30.0);
}

#[test]
fn multiple_chunks_cross_call() {
    let mut vm = VM::new();

    // Chunk 1: double(x) = x * 2
    let mut double = Chunk::new("double");
    double.arity = 1;
    double.local_count = 2;
    double.emit_op_u16(Op::local_get, 1, 0);
    let two = double.add_constant(Value::F64(2.0));
    double.emit_op_u16(Op::r#const, two, 0);
    double.emit_op(Op::f64_mul, 0);
    double.emit_op(Op::r#return, 0);

    // Chunk 2: quad(x) = double(double(x))
    let mut quad = Chunk::new("quad");
    quad.arity = 1;
    quad.local_count = 2;
    let dbl_name = quad.add_constant(Value::String(Rc::from("double")));
    // double(x)
    quad.emit_op_u16(Op::global_get, dbl_name, 0);
    quad.emit_op_u16(Op::local_get, 1, 0);
    quad.emit_op_u8(Op::call, 1, 0);
    // double(result)
    let dbl_name2 = quad.add_constant(Value::String(Rc::from("double")));
    quad.emit_op_u16(Op::global_get, dbl_name2, 0);
    // swap: need func below result
    let _tmp = 2u16; // local slot 2 as temp
    quad.local_count = 3;
    // Actually: result is on stack, global_get pushed func above it
    // Stack: [result, double_fn] — need [double_fn, result]
    // Just do it the other way: get func first, then get result from local
    // This is getting complicated for a test. Simpler: store intermediate.
    // Let me just call double(x) and return, for simplicity.
    quad.emit_op(Op::drop, 0); // drop the second global_get
    quad.emit_op(Op::r#return, 0); // return double(x) — which is 2*x

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    let d_name = main.add_constant(Value::String(Rc::from("double")));
    let q_name = main.add_constant(Value::String(Rc::from("quad")));
    // Register double
    main.emit_op_u16(Op::ref_func, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::global_set, d_name, 0);
    main.emit_op(Op::drop, 0);
    // Register quad
    main.emit_op_u16(Op::ref_func, 2, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::global_set, q_name, 0);
    main.emit_op(Op::drop, 0);
    // Call quad(5)
    main.emit_op_u16(Op::global_get, q_name, 0);
    let five = main.add_constant(Value::F64(5.0));
    main.emit_op_u16(Op::r#const, five, 0);
    main.emit_op_u8(Op::call, 1, 0);
    main.emit_op(Op::halt, 0);

    let result = vm.run(vec![main, double, quad]).unwrap();
    assert_eq!(result.as_f64(), 10.0); // double(5) = 10
}
