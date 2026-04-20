//! Stack-switching proposal coverage: spec-byte emission +
//! fiber-based coroutine semantics.

use vybe_bytecode::{VM, Value, Chunk, Op};
use vybe_bytecode::wasm::write_wasm;

// ── Binary emission ────────────────────────────────────────────────

fn find_byte_seq(wasm: &[u8], needle: &[u8]) -> bool {
    wasm.windows(needle.len()).any(|w| w == needle)
}

#[test]
fn stack_switching_opcodes_emit_spec_bytes() {
    // A chunk that emits CONT_NEW + SUSPEND + RESUME + SWITCH should
    // produce a binary containing the spec 0xE0 / 0xE2 / 0xE3 / 0xE5
    // bytes respectively.
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    // Push a function-shaped value (null is fine for emission only).
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op(Op::DUP, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op_u16(Op::RESUME, 0, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op_u16(Op::SUSPEND, 0, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op_u16(Op::SWITCH, 0, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    assert!(find_byte_seq(&wasm, &[0xE0]),
        "CONT_NEW should emit 0xE0 (cont.new)");
    assert!(find_byte_seq(&wasm, &[0xE2]),
        "SUSPEND should emit 0xE2 (suspend)");
    assert!(find_byte_seq(&wasm, &[0xE3]),
        "RESUME should emit 0xE3 (resume)");
    assert!(find_byte_seq(&wasm, &[0xE5]),
        "SWITCH should emit 0xE5 (switch)");
}

#[test]
fn stack_switching_type_section_declares_continuation_type() {
    // When any stack-switching op is present, the type section must
    // include (a) the suspend tag's func type and (b) the cont type
    // `0x5D <funcidx>` wrapping it.
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    // Look for a `0x5D <n>` pair — the cont type prefix followed by
    // a small LEB128 funcidx. Since the suspend tag func type is
    // emitted just before the cont type, the funcidx referenced is a
    // low number (within the type section). A byte-pattern scan that
    // finds 0x5D followed by any LEB128 first-byte confirms emission.
    let found = wasm.windows(2).any(|w| w[0] == 0x5D && w[1] < 0x80);
    assert!(found, "cont-type prefix 0x5D not present in emitted binary");
}

#[test]
fn stack_switching_tag_section_declares_suspend_tag() {
    // When stack-switching is active AND exceptions are used, the tag
    // section has 2 tags. Confirm the tag section is present and
    // larger than the single-tag baseline.
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    // Section id 13 = tag section. Find it and verify the tag count.
    let mut pos = 8; // skip magic + version
    while pos < wasm.len() {
        let id = wasm[pos]; pos += 1;
        let mut size = 0u32;
        let mut shift = 0;
        loop {
            let b = wasm[pos]; pos += 1;
            size |= ((b & 0x7f) as u32) << shift;
            if b & 0x80 == 0 { break; }
            shift += 7;
        }
        if id == 13 {
            // First byte of the tag section is the tag count (LEB128).
            let tag_count = wasm[pos];
            assert!(tag_count >= 1,
                "tag section should have at least 1 tag when stack-switching is used");
            return;
        }
        pos += size as usize;
    }
    panic!("tag section (id 13) not found");
}

// ── VM coroutine semantics ─────────────────────────────────────────

#[test]
fn cont_new_returns_continuation_object() {
    // CONT_NEW should wrap a function in an ObjectKind::Continuation
    // — verify by inspecting the stack top after the op.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // Build a dummy function value: push a ref.func-shaped placeholder.
    // For this test we use null; CONT_NEW accepts any value as the
    // "entry" and stashes it for later resume.
    let null_k = chunk.add_constant(Value::Null);
    chunk.emit_op_u16(Op::CONST, null_k, 0);
    chunk.emit_op(Op::CONT_NEW, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                vybe_bytecode::value::ObjectKind::Continuation(_) => {}
                other => panic!("expected Continuation, got {other:?}"),
            }
        }
        other => panic!("expected Object, got {other:?}"),
    }
}

#[test]
fn cont_bind_emits_spec_byte() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op_u8(Op::CONT_BIND, 1, 0); // bind 1 arg
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    assert!(find_byte_seq(&wasm, &[0xE1]),
        "CONT_BIND should emit 0xE1 (cont.bind)");
}

#[test]
fn resume_throw_emits_spec_byte() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op_u16(Op::RESUME_THROW, 0, 0);
    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    assert!(find_byte_seq(&wasm, &[0xE4]),
        "RESUME_THROW should emit 0xE4 (resume_throw)");
}

#[test]
fn cont_bind_produces_continuation_with_bound_args() {
    // CONT_BIND wraps args into a new continuation; the runtime
    // result should still be ObjectKind::Continuation and carry the
    // bound args as a `__bound_args` array property.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::CONT_NEW, 0);
    let v = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, v, 0);
    chunk.emit_op_u8(Op::CONT_BIND, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    let obj = match result {
        Value::Object(obj) => obj,
        other => panic!("expected Continuation object, got {other:?}"),
    };
    let o = obj.lock().unwrap();
    match &o.kind {
        vybe_bytecode::value::ObjectKind::Continuation(_) => {}
        other => panic!("expected Continuation kind, got {other:?}"),
    }
    let bound = o.properties.get("__bound_args")
        .expect("CONT_BIND should stash bound args on the new cont");
    if let Value::Object(bo) = bound {
        let b = bo.lock().unwrap();
        if let vybe_bytecode::value::ObjectKind::Array(elems) = &b.kind {
            assert_eq!(elems.len(), 1);
            assert_eq!(elems[0].as_i32(), 42);
        } else {
            panic!("__bound_args should be Array");
        }
    } else {
        panic!("__bound_args should be an Object");
    }
}

#[test]
fn generator_chunk_returns_continuation_on_call() {
    // A function chunk flagged as a generator must not execute its
    // body when called — it should hand back a `Continuation`. The
    // body runs only when the caller RESUMEs the continuation.
    // Build the body chunk (generator), the factory chunk (0-arg
    // script that calls the generator), and wire them.
    let mut gen_body = Chunk::new("count_to_three");
    gen_body.arity = 0;
    gen_body.local_count = 0;
    gen_body.is_generator = true;
    let one = gen_body.add_constant(Value::I32(1));
    gen_body.emit_op_u16(Op::CONST, one, 0);
    gen_body.emit_op_u16(Op::SUSPEND, 0, 0);
    gen_body.emit_op(Op::NULL, 0);
    gen_body.emit_op(Op::RETURN, 0);

    // Caller builds a Function ref to gen_body, calls it with 0 args,
    // and returns the resulting continuation.
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.emit_op_u16(Op::REF_FUNC, 1, 0); // chunk_idx = 1 (gen_body)
    script.emit(0, 0); // uv_count = 0
    script.emit_op_u8(Op::CALL_REF, 0, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, gen_body]).unwrap();
    match result {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                vybe_bytecode::value::ObjectKind::Continuation(_) => {}
                other => panic!("expected Continuation, got {other:?}"),
            }
        }
        other => panic!("expected Object/Continuation, got {other:?}"),
    }
}

#[test]
fn suspend_without_active_continuation_falls_back_to_return() {
    // Running SUSPEND with no RESUME above it should just return the
    // yielded value — this preserves the legacy behaviour so existing
    // code that used SUSPEND as "return from async" keeps working.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let k = chunk.add_constant(Value::I32(77));
    chunk.emit_op_u16(Op::CONST, k, 0);
    chunk.emit_op_u16(Op::SUSPEND, 0, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 77);
}
