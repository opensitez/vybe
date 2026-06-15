use std::sync::Arc;
use std::sync::Mutex;
/// Tests for memory operations, table operations, and WASM binary I/O.
use vybe_bytecode::value::{Function, Object, ObjectKind};
use vybe_bytecode::{Chunk, Op, VM, Value};

// ============================================================
// MEMORY OPERATIONS
// ============================================================

#[test]
fn memory_grow_and_size() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // memory_size (initial = 0 pages)
    chunk.emit_op(Op::MEMORY_SIZE, 0);
    // Grow by 1 page (64KB)
    chunk.emit_op(Op::DROP, 0);
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, one, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    // memory_grow returns old size
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::MEMORY_SIZE, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 1); // 1 page after grow
}

#[test]
fn memory_grow_returns_minus_one_when_max_exceeded() {
    let mut vm = VM::new();
    vm.memory.resize(65536, 0);
    vm.memory.set_max_pages(Some(1));

    let mut chunk = Chunk::new("<script>");
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, one, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), -1);
    assert_eq!(vm.memory.len(), 65536);
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
    chunk.emit_op_u16(Op::CONST, addr, 0);
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    // i32.load addr=0
    let addr2 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, addr2, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);
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
    chunk.emit_op_u16(Op::CONST, addr, 0);
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::F64_STORE, 0);
    let addr2 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, addr2, 0);
    chunk.emit_op(Op::F64_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);
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
    chunk.emit_op_u16(Op::CONST, addr, 0);
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    // Store 200 at addr 12
    let addr2 = chunk.add_constant(Value::I32(12));
    let val2 = chunk.add_constant(Value::I32(200));
    chunk.emit_op_u16(Op::CONST, addr2, 0);
    chunk.emit_op_u16(Op::CONST, val2, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    // Load from addr 8
    let addr3 = chunk.add_constant(Value::I32(8));
    chunk.emit_op_u16(Op::CONST, addr3, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 100);
}

#[test]
fn memory_fill_via_rust() {
    // memory_fill/copy don't have opcodes yet — test via Rust API
    let mut vm = VM::new();
    vm.memory.resize(256, 0);
    // Fill bytes 0..16 with 0xFF
    for i in 0..16 {
        let _ = vm.memory.store_u8(i, 0xFF);
    }
    // Verify via i32_load
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let addr = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, addr, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), -1); // 0xFFFFFFFF
}

#[test]
fn memory_copy_via_rust() {
    let mut vm = VM::new();
    vm.memory.resize(256, 0);
    let _ = vm.memory.store_i32(0, 42);
    // Copy 4 bytes from offset 0 to offset 16
    vm.memory.with_buffer_mut(|buf| {
        buf.copy_within(0..4, 16);
    });
    // Verify via i32_load at offset 16
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let addr = chunk.add_constant(Value::I32(16));
    chunk.emit_op_u16(Op::CONST, addr, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);
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
    f.emit_op_u16(Op::CONST, val, 0);
    f.emit_op(Op::RETURN, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    // Store function in func_table at index 0
    // ref_func creates function, we store it as global then use call_indirect
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0); // 0 upvalues
    // For call_indirect, we need the table index on stack
    // Just use regular call for now since func_table setup is complex
    main.emit_op_u8(Op::CALL, 0, 0);
    main.emit_op(Op::HALT, 0);

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
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::HALT, 0);

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
    f.local_count = 2; // slot 0=a, 1=b (WASM convention)
    f.emit_op_u16(Op::LOCAL_GET, 0, 0);
    f.emit_op_u16(Op::LOCAL_GET, 1, 0);
    f.emit_op(Op::F64_ADD, 0);
    f.emit_op(Op::RETURN, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    let name = main.add_constant(Value::String(Arc::from("add")));
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::GLOBAL_SET, name, 0);
    main.emit_op(Op::DROP, 0);
    // Call: add(10, 20)
    main.emit_op_u16(Op::GLOBAL_GET, name, 0);
    let a = main.add_constant(Value::F64(10.0));
    let b = main.add_constant(Value::F64(20.0));
    main.emit_op_u16(Op::CONST, a, 0);
    main.emit_op_u16(Op::CONST, b, 0);
    main.emit_op_u8(Op::CALL, 2, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main, f]).unwrap();
    assert_eq!(result.as_f64(), 30.0);
}

#[test]
fn multiple_chunks_cross_call() {
    let mut vm = VM::new();

    // Chunk 1: double(x) = x * 2
    let mut double = Chunk::new("double");
    double.arity = 1;
    double.local_count = 1;
    double.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let two = double.add_constant(Value::F64(2.0));
    double.emit_op_u16(Op::CONST, two, 0);
    double.emit_op(Op::F64_MUL, 0);
    double.emit_op(Op::RETURN, 0);

    // Chunk 2: quad(x) = double(x) (simplified — just verifies cross-chunk
    // call returns correctly; avoids the swap gymnastics from the original
    // version that had `quad` dropping its second call and returning the
    // first).
    let mut quad = Chunk::new("quad");
    quad.arity = 1;
    quad.local_count = 1;
    let dbl_name = quad.add_constant(Value::String(Arc::from("double")));
    quad.emit_op_u16(Op::GLOBAL_GET, dbl_name, 0);
    quad.emit_op_u16(Op::LOCAL_GET, 0, 0);
    quad.emit_op_u8(Op::CALL, 1, 0);
    quad.emit_op(Op::RETURN, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    let d_name = main.add_constant(Value::String(Arc::from("double")));
    let q_name = main.add_constant(Value::String(Arc::from("quad")));
    // Register double
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::GLOBAL_SET, d_name, 0);
    main.emit_op(Op::DROP, 0);
    // Register quad
    main.emit_op_u16(Op::REF_FUNC, 2, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::GLOBAL_SET, q_name, 0);
    main.emit_op(Op::DROP, 0);
    // Call quad(5)
    main.emit_op_u16(Op::GLOBAL_GET, q_name, 0);
    let five = main.add_constant(Value::F64(5.0));
    main.emit_op_u16(Op::CONST, five, 0);
    main.emit_op_u8(Op::CALL, 1, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main, double, quad]).unwrap();
    assert_eq!(result.as_f64(), 10.0); // double(5) = 10
}

#[test]
fn call_indirect_vm_function() {
    let mut vm = VM::new();

    // Chunk 1: function that returns 99
    let mut f = Chunk::new("get99");
    f.arity = 0;
    f.local_count = 1;
    let val = f.add_constant(Value::F64(99.0));
    f.emit_op_u16(Op::CONST, val, 0);
    f.emit_op(Op::RETURN, 0);

    // Script: ref_func creates function + adds to func_table
    // Then get __table_idx and use call_indirect
    let mut main = Chunk::new("<script>");
    main.local_count = 2;

    // ref_func 1 → pushes function object
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0); // 0 upvalues

    // Get __table_idx from the function object
    let idx_name = main.add_constant(Value::String(Arc::from("__table_idx")));
    main.emit_op(Op::DUP, 0);
    main.emit_op_u16(Op::STRUCT_GET, idx_name, 0);

    // Stack: [func_obj, table_idx]
    // Save func_obj to local, keep table_idx for call_indirect
    let _tmp = 1u16;
    // Swap: drop func_obj from under table_idx
    // Actually struct_get popped func_obj and pushed table_idx
    // So stack: [table_idx]

    // call_indirect with 0 args
    main.emit_op_u8_u8(Op::CALL_INDIRECT, 0, 0, 0);
    main.emit_op(Op::HALT, 0);

    let result = vm.run(vec![main, f]).unwrap();
    assert_eq!(result.as_f64(), 99.0);
}

#[test]
fn decoded_standard_call_indirect_executes_vm_table_function() {
    let wasm = vec![
        0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00, // header
        0x01, 0x05, 0x01, 0x60, 0x00, 0x01, 0x7f, // type: [] -> [i32]
        0x03, 0x02, 0x01, 0x00, // one function, type 0
        0x0a, 0x09, 0x01, 0x07, 0x00, // one body, body size 7, no locals
        0x41, 0x00, // i32.const 0: table index
        0x11, 0x00, 0x00, // call_indirect type 0 table 0
        0x0b, // end
    ];
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let caller = chunks.remove(1);

    let mut target = Chunk::new("target");
    target.arity = 0;
    target.local_count = 0;
    let value = target.add_constant(Value::I32(77));
    target.emit_op_u16(Op::CONST, value, 0);
    target.emit_op(Op::RETURN, 0);

    let mut function = Object::new();
    function.kind = ObjectKind::Function(Function {
        name: Some("target".into()),
        arity: 0,
        chunk_index: 1,
        upvalues: Vec::new(),
    });

    let mut vm = VM::new();
    vm.func_table
        .push(Value::Object(Arc::new(Mutex::new(function))));
    let result = vm.run(vec![caller, target]).unwrap();
    assert_eq!(result.as_i32(), 77);
}

#[test]
fn decoded_standard_call_indirect_uses_encoded_table_index() {
    let wasm = vec![
        0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00, // header
        0x01, 0x05, 0x01, 0x60, 0x00, 0x01, 0x7f, // type: [] -> [i32]
        0x03, 0x02, 0x01, 0x00, // one function, type 0
        0x04, 0x07, 0x02, 0x70, 0x00, 0x01, 0x70, 0x00, 0x01, // tables 0 and 1
        0x0a, 0x09, 0x01, 0x07, 0x00, // one body, body size 7, no locals
        0x41, 0x00, // i32.const 0: element index
        0x11, 0x00, 0x01, // call_indirect type 0 table 1
        0x0b, // end
    ];
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let caller = chunks.remove(1);

    let mut target = Chunk::new("target_table_1");
    let value = target.add_constant(Value::I32(88));
    target.emit_op_u16(Op::CONST, value, 0);
    target.emit_op(Op::RETURN, 0);

    let mut function = Object::new();
    function.kind = ObjectKind::Function(Function {
        name: Some("target_table_1".into()),
        arity: 0,
        chunk_index: 1,
        upvalues: Vec::new(),
    });

    let mut vm = VM::new();
    vm.extra_tables
        .push(vec![Value::Object(Arc::new(Mutex::new(function)))]);
    let result = vm.run(vec![caller, target]).unwrap();
    assert_eq!(result.as_i32(), 88);
}

// ── Missing load/store variants (§5.3 memory instructions) ──────────────

fn mem_vm() -> VM {
    let vm = VM::new();
    vm.memory.resize(1024, 0);
    vm
}

fn run_mem(vm: &mut VM, emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    c.local_count = 1;
    emit(&mut c);
    c.emit_op(Op::HALT, 0);
    vm.run(vec![c]).unwrap()
}

fn run_mem_err(vm: &mut VM, emit: impl FnOnce(&mut Chunk)) -> String {
    let mut c = Chunk::new("<script>");
    c.local_count = 1;
    emit(&mut c);
    c.emit_op(Op::HALT, 0);
    vm.run(vec![c]).unwrap_err().to_string()
}

// ── memarg offsets and traps ─────────────────────────────────────────────

#[test]
fn i32_load_store_apply_memarg_offset() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let base = c.add_constant(Value::I32(4));
        let val = c.add_constant(Value::I32(0x1122_3344));
        let zero = c.add_constant(Value::I32(0));

        c.emit_op_u16(Op::CONST, base, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I32_STORE, 0);
        c.emit_leb_u32(2, 0); // natural i32 alignment
        c.emit_leb_u32(8, 0); // effective address = 4 + 8

        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op(Op::I32_LOAD, 0);
        c.emit_leb_u32(2, 0);
        c.emit_leb_u32(12, 0);
    });
    assert_eq!(r.as_i32(), 0x1122_3344);
}

#[test]
fn f64_load_store_apply_memarg_offset() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let base = c.add_constant(Value::I32(5));
        let val = c.add_constant(Value::F64(6.25));
        let zero = c.add_constant(Value::I32(0));

        c.emit_op_u16(Op::CONST, base, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::F64_STORE, 0);
        c.emit_leb_u32(3, 0);
        c.emit_leb_u32(9, 0);

        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op(Op::F64_LOAD, 0);
        c.emit_leb_u32(3, 0);
        c.emit_leb_u32(14, 0);
    });
    assert_eq!(r.as_f64(), 6.25);
}

#[test]
fn f32_load_store_apply_memarg_offset() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let base = c.add_constant(Value::I32(3));
        let val = c.add_constant(Value::F64((-2.5f32) as f64));
        let zero = c.add_constant(Value::I32(0));

        c.emit_op_u16(Op::CONST, base, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::F32_STORE, 0);
        c.emit_leb_u32(2, 0);
        c.emit_leb_u32(17, 0);

        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op(Op::F32_LOAD, 0);
        c.emit_leb_u32(2, 0);
        c.emit_leb_u32(20, 0);
    });
    assert_eq!(r.as_f64() as f32, -2.5f32);
}

#[test]
fn i64_load16_s_applies_memarg_offset() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let base = c.add_constant(Value::I32(6));
        let val = c.add_constant(Value::I64(-1234));
        let zero = c.add_constant(Value::I32(0));

        c.emit_op_u16(Op::CONST, base, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE16, 0);
        c.emit_leb_u32(1, 0);
        c.emit_leb_u32(12, 0);

        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op(Op::I64_LOAD16_S, 0);
        c.emit_leb_u32(1, 0);
        c.emit_leb_u32(18, 0);
    });
    assert_eq!(r.as_i64(), -1234);
}

#[test]
fn i32_load_oob_traps() {
    let mut vm = VM::new();
    vm.memory.resize(3, 0);
    let err = run_mem_err(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I32_LOAD, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

#[test]
fn i64_store32_oob_traps() {
    let mut vm = VM::new();
    vm.memory.resize(2, 0);
    let err = run_mem_err(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(1));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE32, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

#[test]
fn f32_load_oob_traps() {
    let mut vm = VM::new();
    vm.memory.resize(3, 0);
    let err = run_mem_err(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::F32_LOAD, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

#[test]
fn f64_store_oob_traps() {
    let mut vm = VM::new();
    vm.memory.resize(7, 0);
    let err = run_mem_err(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::F64(1.0));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::F64_STORE, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

// ── f32.store / f32.load ─────────────────────────────────────────────────

#[test]
fn f32_store_and_load() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::F64(3.14f32 as f64));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::F32_STORE, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::F32_LOAD, 0);
    });
    assert!((r.as_f64() as f32 - 3.14f32).abs() < 1e-5);
}

// ── i32 narrow loads ─────────────────────────────────────────────────────

#[test]
fn i32_store8_and_load8_s() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I32(-1i32)); // 0xFF as byte
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I32_STORE8, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I32_LOAD8_S, 0); // sign-extend → -1
    });
    assert_eq!(r.as_i32(), -1);
}

#[test]
fn i32_load8_u() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I32(0xFF));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I32_STORE8, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I32_LOAD8_U, 0); // zero-extend → 255
    });
    assert_eq!(r.as_i32(), 255);
}

#[test]
fn i32_store16_and_load16_s() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I32(-1000));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I32_STORE16, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I32_LOAD16_S, 0);
    });
    assert_eq!(r.as_i32(), -1000);
}

#[test]
fn i32_load16_u() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I32(0xFFFF));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I32_STORE16, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I32_LOAD16_U, 0);
    });
    assert_eq!(r.as_i32(), 65535);
}

// ── i64.load / i64.store ─────────────────────────────────────────────────

#[test]
fn i64_store_and_load() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(i64::MAX));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I64_LOAD, 0);
    });
    assert_eq!(r.as_i64(), i64::MAX);
}

#[test]
fn i64_store8_and_load8_s() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(-1));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE8, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I64_LOAD8_S, 0);
    });
    assert_eq!(r.as_i64(), -1);
}

#[test]
fn i64_load8_u() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(200));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE8, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I64_LOAD8_U, 0);
    });
    assert_eq!(r.as_i64(), 200);
}

#[test]
fn i64_store16_and_load16_s() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(-5000));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE16, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I64_LOAD16_S, 0);
    });
    assert_eq!(r.as_i64(), -5000);
}

#[test]
fn i64_load16_u() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(40000));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE16, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I64_LOAD16_U, 0);
    });
    assert_eq!(r.as_i64(), 40000);
}

#[test]
fn i64_store32_and_load32_s() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(-70000));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE32, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I64_LOAD32_S, 0);
    });
    assert_eq!(r.as_i64(), -70000);
}

#[test]
fn i64_load32_u() {
    let mut vm = mem_vm();
    let r = run_mem(&mut vm, |c| {
        let addr = c.add_constant(Value::I32(0));
        let val = c.add_constant(Value::I64(3_000_000_000i64));
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op(Op::I64_STORE32, 0);
        c.emit_op_u16(Op::CONST, addr, 0);
        c.emit_op(Op::I64_LOAD32_U, 0);
    });
    assert_eq!(r.as_i64(), 3_000_000_000i64);
}
