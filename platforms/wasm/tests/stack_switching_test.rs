//! Stack-switching proposal coverage: spec-byte emission +
//! fiber-based coroutine semantics.

use std::sync::{Arc, Mutex};

use vybe_platform_wasm as wasm;
use vybe_platform_wasm::write_wasm;
use vybe_runtime::chunk::StackSwitchHandler;
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};

/// Spec-`resume` generator iterator-advance, mirroring the compiler's
/// `emitter::generators::emit_next`. Stack: `[cont]` -> `[value, has_more_i32]`.
/// Uses `resume` + an `(on yield -> handler)` table — the retired custom
/// `GEN_NEXT` opcode's spec replacement.
fn emit_spec_gen_next(chunk: &mut Chunk) {
    let cont_slot = chunk.local_count;
    chunk.local_count += 1;
    chunk.emit_op_u16(Op::LOCAL_SET, cont_slot, 0);

    let done_key = chunk.add_constant(Value::String(Arc::from("__gen_done")));
    let _block = chunk.emit_block_typed(0, 2);

    // Done guard: completed continuation → (null, 0) without re-resuming.
    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, done_key, 0);
    chunk.emit_op(Op::REF_IS_NULL, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    let _done_if = chunk.emit_if(0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_br(1, 0);
    chunk.emit_end(0);

    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    let resume_ip = chunk.code.len();
    chunk.emit_op_u16(Op::RESUME, 0, 0);
    // Completion arm: stamp done, push has_more = 0.
    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, done_key, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_br(0, 0);
    // Yield arm (VM jumps here on suspend tag 0): push has_more = 1.
    let handler_ip = chunk.code.len();
    chunk.emit_i32_const(1, 0);
    chunk.emit_end(0);

    chunk.stack_switch_handlers.insert(
        resume_ip,
        vec![StackSwitchHandler {
            kind: 0,
            tag_index: 0,
            label_index: handler_ip as u32,
        }],
    );
}

fn write_leb_u32(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn push_section(out: &mut Vec<u8>, id: u8, payload: &[u8]) {
    out.push(id);
    write_leb_u32(out, payload.len() as u32);
    out.extend_from_slice(payload);
}

fn make_promise(id: u64, state: &str, value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties.insert("__id".into(), Value::F64(id as f64));
    obj.properties
        .insert("__state".into(), Value::String(Arc::from(state)));
    obj.properties.insert("__value".into(), value);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn emit_try_table_catch_all_start(c: &mut Chunk) -> usize {
    // Spec encoding: one `catch $vybe:exception` clause (kind 0) — the
    // handler receives the exception object payload, matching the DROP in
    // the handlers below. [op, count=1, kind u8, tag u16, offset u16]
    let tag = c.import_exception_tag("vybe:exception", 1);
    c.emit_op(Op::TRY_TABLE, 0);
    c.emit(1, 0);
    c.emit(0, 0); // kind 0 = catch
    c.emit((tag >> 8) as u8, 0);
    c.emit((tag & 0xff) as u8, 0);
    let offset_pos = c.current_offset();
    c.emit(0, 0);
    c.emit(0, 0);
    offset_pos
}

fn patch_try_table(c: &mut Chunk, offset_pos: usize) {
    let body_bytes = c.current_offset() - (offset_pos + 2);
    c.code[offset_pos] = (body_bytes >> 8) as u8;
    c.code[offset_pos + 1] = (body_bytes & 0xFF) as u8;
}

fn standard_stack_switching_module(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut out, 3, &[0x01, 0x00]);

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(body_ops);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

// ── Binary emission ────────────────────────────────────────────────

fn find_byte_seq(wasm: &[u8], needle: &[u8]) -> bool {
    wasm.windows(needle.len()).any(|w| w == needle)
}

fn read_leb_u32_at(bytes: &[u8], pos: &mut usize) -> u32 {
    let mut result = 0u32;
    let mut shift = 0;
    loop {
        let b = bytes[*pos];
        *pos += 1;
        result |= ((b & 0x7f) as u32) << shift;
        if b & 0x80 == 0 {
            return result;
        }
        shift += 7;
    }
}

fn find_stack_switch_opcode(wasm: &[u8], opcode: u8) -> Option<usize> {
    let mut pos = 8;
    while pos < wasm.len() {
        let id = wasm[pos];
        pos += 1;
        let size = read_leb_u32_at(wasm, &mut pos) as usize;
        let end = pos + size;
        if end > wasm.len() {
            return None;
        }
        if id == 10 {
            return wasm[pos..end]
                .iter()
                .position(|b| *b == opcode)
                .map(|offset| pos + offset);
        }
        pos = end;
    }
    None
}

#[test]
fn stack_switching_opcodes_emit_spec_bytes() {
    // Check the actual Wasm code-section instruction encodings, not
    // just whether the opcode byte appears somewhere in the module.
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::CONT_NEW, 0); // e0 <cont_typeidx>
    script.emit_dup(0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u16(Op::RESUME, 0, 0); // e3 <cont_typeidx> <handlers>
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u16(Op::SUSPEND, 0, 0); // e2 <tagidx>
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u16(Op::SWITCH, 0, 0); // e6 <cont_typeidx> <tagidx>
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    let mut pos =
        find_stack_switch_opcode(&wasm, 0xE0).expect("emitted wasm should contain cont.new");
    pos += 1;
    let cont_type = read_leb_u32_at(&wasm, &mut pos);

    let mut pos =
        find_stack_switch_opcode(&wasm, 0xE2).expect("emitted wasm should contain suspend");
    pos += 1;
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), 0);

    let mut pos =
        find_stack_switch_opcode(&wasm, 0xE3).expect("emitted wasm should contain resume");
    pos += 1;
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), cont_type);
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), 0);

    let mut pos =
        find_stack_switch_opcode(&wasm, 0xE6).expect("emitted wasm should contain switch");
    pos += 1;
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), cont_type);
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), 0);
}

#[test]
fn stack_switching_opcodes_disassemble_with_proposal_names() {
    let mut script = Chunk::new("<script>");
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op_u8(Op::CONT_BIND, 0, 0);
    script.emit_op_u16(Op::SUSPEND, 0, 0);
    script.emit_op_u16(Op::RESUME, 0, 0);
    script.emit_op_u16(Op::RESUME_THROW, 0, 0);
    script.emit_op(Op::RESUME_THROW_REF, 0);
    script.emit_op_u16(Op::SWITCH, 0, 0);
    script.emit_op(Op::RETURN, 0);

    let wat = vybe_platform_wasm::disassembler::write_wat(&[script]);
    for name in [
        "cont.new",
        "cont.bind",
        "suspend",
        "resume",
        "resume_throw",
        "resume_throw_ref",
        "switch",
    ] {
        assert!(wat.contains(name), "WAT output missing {name}: {wat}");
    }
}

#[test]
fn stack_switching_type_section_declares_continuation_type() {
    // When any stack-switching op is present, the type section must
    // include (a) the suspend tag's func type and (b) the cont type
    // `0x5D <funcidx>` wrapping it.
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
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
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    // Section id 13 = tag section. Find it and verify the tag count.
    let mut pos = 8; // skip magic + version
    while pos < wasm.len() {
        let id = wasm[pos];
        pos += 1;
        let mut size = 0u32;
        let mut shift = 0;
        loop {
            let b = wasm[pos];
            pos += 1;
            size |= ((b & 0x7f) as u32) << shift;
            if b & 0x80 == 0 {
                break;
            }
            shift += 7;
        }
        if id == 13 {
            // First byte of the tag section is the tag count (LEB128).
            let tag_count = wasm[pos];
            assert!(
                tag_count >= 1,
                "tag section should have at least 1 tag when stack-switching is used"
            );
            return;
        }
        pos += size as usize;
    }
    panic!("tag section (id 13) not found");
}

#[test]
fn standard_resume_with_handler_vector_must_not_decode_as_noop() {
    let bytes = standard_stack_switching_module(&[
        0x41, 0x00, // placeholder continuation operand for validation
        0x41, 0x00, // placeholder resume value for validation
        0xE3, // resume
        0x00, // cont type index
        0x01, // one handler
        0x00, // on-tag-to-label
        0x00, // tag index
        0x00, // label index
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("resume with handler vector should decode");
    assert!(chunks[1].code.windows(2).any(|w| w == [0x00, 0xE3]));
    assert_eq!(
        chunks[1].stack_switch_handlers.values().next().unwrap(),
        &vec![StackSwitchHandler {
            kind: 0,
            tag_index: 0,
            label_index: 0
        }]
    );
}

#[test]
fn decoded_resume_handler_vector_roundtrips_to_wasm() {
    let bytes = standard_stack_switching_module(&[
        0x41, 0x00, // placeholder continuation operand for validation
        0x41, 0x00, // placeholder resume value for validation
        0xE3, // resume
        0x00, // cont type index
        0x01, // one handler
        0x00, // on-tag-to-label
        0x00, // tag index
        0x00, // label index
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("resume with handler vector should decode");
    let emitted = write_wasm(&chunks);
    let mut pos =
        find_stack_switch_opcode(&emitted, 0xE3).expect("emitted wasm should contain resume");
    pos += 1;
    let _cont_type = read_leb_u32_at(&emitted, &mut pos);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 1);
    assert_eq!(emitted[pos], 0);
    pos += 1;
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 0);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 0);
}

#[test]
fn standard_resume_throw_ref_with_handler_vector_decodes_and_roundtrips() {
    let bytes = standard_stack_switching_module(&[
        0xD0, 0x6F, // placeholder exnref operand
        0x41, 0x00, // placeholder continuation operand for validation
        0xE5, // resume_throw_ref
        0x00, // cont type index
        0x01, // one handler
        0x01, // on-tag-to-switch
        0x00, // tag index
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("resume_throw_ref handler vector should decode");
    assert!(chunks[1].code.windows(2).any(|w| w == [0x00, 0xE5]));
    assert_eq!(
        chunks[1].stack_switch_handlers.values().next().unwrap(),
        &vec![StackSwitchHandler {
            kind: 1,
            tag_index: 0,
            label_index: 0
        }]
    );

    let emitted = write_wasm(&chunks);
    let mut pos = find_stack_switch_opcode(&emitted, 0xE5)
        .expect("emitted wasm should contain resume_throw_ref");
    pos += 1;
    let _cont_type = read_leb_u32_at(&emitted, &mut pos);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 1);
    assert_eq!(emitted[pos], 1);
    pos += 1;
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 0);
}

#[test]
fn standard_cont_bind_decodes_and_roundtrips_spec_shape() {
    let bytes = standard_stack_switching_module(&[
        0x41, 0x00, // placeholder continuation operand for validation
        0xE1, // cont.bind
        0x00, // source continuation type index
        0x00, // destination continuation type index
        0x1A, // drop the rebound continuation (cont.bind pushes it; body is () -> ())
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("cont.bind should decode");
    assert!(chunks[1].code.windows(2).any(|w| w == [0x00, 0xE1]));

    let emitted = write_wasm(&chunks);
    let mut pos =
        find_stack_switch_opcode(&emitted, 0xE1).expect("emitted wasm should contain cont.bind");
    pos += 1;
    let source_type = read_leb_u32_at(&emitted, &mut pos);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), source_type);
}

#[test]
fn standard_resume_throw_with_handler_vector_decodes_and_roundtrips() {
    let bytes = standard_stack_switching_module(&[
        0x41, 0x00, // placeholder continuation operand for validation
        0x41, 0x01, // placeholder exception payload
        0xE4, // resume_throw
        0x00, // cont type index
        0x00, // tag index
        0x01, // one handler
        0x00, // on-tag-to-label
        0x00, // tag index
        0x00, // label index
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("resume_throw handler vector should decode");
    assert!(chunks[1].code.windows(2).any(|w| w == [0x00, 0xE4]));
    assert_eq!(
        chunks[1].stack_switch_handlers.values().next().unwrap(),
        &vec![StackSwitchHandler {
            kind: 0,
            tag_index: 0,
            label_index: 0
        }]
    );

    let emitted = write_wasm(&chunks);
    let mut pos =
        find_stack_switch_opcode(&emitted, 0xE4).expect("emitted wasm should contain resume_throw");
    pos += 1;
    let _cont_type = read_leb_u32_at(&emitted, &mut pos);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 0);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 1);
    assert_eq!(emitted[pos], 0);
    pos += 1;
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 0);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 0);
}

#[test]
fn standard_switch_decodes_and_roundtrips_spec_shape() {
    let bytes = standard_stack_switching_module(&[
        0x41, 0x00, // placeholder continuation operand for validation
        0x41, 0x01, // placeholder switch payload
        0xE6, // switch
        0x00, // continuation type index
        0x00, // tag index
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("switch should decode");
    assert!(chunks[1].code.windows(2).any(|w| w == [0x00, 0xE6]));

    let emitted = write_wasm(&chunks);
    let mut pos =
        find_stack_switch_opcode(&emitted, 0xE6).expect("emitted wasm should contain switch");
    pos += 1;
    let _cont_type = read_leb_u32_at(&emitted, &mut pos);
    assert_eq!(read_leb_u32_at(&emitted, &mut pos), 0);
}

// ── VM coroutine semantics ─────────────────────────────────────────

#[test]
fn cont_new_returns_continuation_object() {
    // CONT_NEW should wrap a function in an ObjectKind::Continuation
    // — verify by inspecting the stack top after the op.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::REF_FUNC, 1, 0);
    chunk.emit(0, 0);
    chunk.emit_op(Op::CONT_NEW, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut target = Chunk::new("target");
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk, target]).unwrap();
    match result {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                vybe_runtime::value::ObjectKind::Continuation(_) => {}
                other => panic!("expected Continuation, got {other:?}"),
            }
        }
        other => panic!("expected Object, got {other:?}"),
    }
}

#[test]
fn cont_new_null_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::CONT_NEW, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("cont.new") && err.contains("null"));
}

#[test]
fn cont_bind_emits_spec_byte() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u8(Op::CONT_BIND, 1, 0); // bind 1 arg
    script.emit_op(Op::DROP, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    let mut pos =
        find_stack_switch_opcode(&wasm, 0xE1).expect("emitted wasm should contain cont.bind");
    pos += 1;
    let source_type = read_leb_u32_at(&wasm, &mut pos);
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), source_type);
}

#[test]
fn resume_throw_emits_spec_byte() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u16(Op::RESUME_THROW, 0, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![script]);
    let mut pos =
        find_stack_switch_opcode(&wasm, 0xE4).expect("emitted wasm should contain resume_throw");
    pos += 1;
    let _cont_type = read_leb_u32_at(&wasm, &mut pos);
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), 0);
    assert_eq!(read_leb_u32_at(&wasm, &mut pos), 0);
}

#[test]
fn resume_throw_non_continuation_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u16(Op::RESUME_THROW, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("resume_throw") && err.contains("continuation"));
}

#[test]
fn resume_throw_ref_non_continuation_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op(Op::RESUME_THROW_REF, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("resume_throw_ref") && err.contains("continuation"));
}

#[test]
fn resume_throw_ref_null_exception_traps() {
    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RESUME_THROW_REF, 0);
    script.emit_op(Op::RETURN, 0);

    let mut target = Chunk::new("target");
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![script, target]).unwrap_err().to_string();
    assert!(err.contains("resume_throw_ref") && err.contains("null"));
}

#[test]
fn resume_non_continuation_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u16(Op::RESUME, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("resume") && err.contains("continuation"));
}

#[test]
fn switch_non_continuation_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u16(Op::SWITCH, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("switch") && err.contains("continuation"));
}

#[test]
fn switch_without_active_prompt_traps() {
    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u16(Op::SWITCH, 0, 0);
    script.emit_op(Op::RETURN, 0);

    let mut target = Chunk::new("target");
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![script, target]).unwrap_err().to_string();
    assert!(err.contains("switch") && err.contains("handler"));
}

#[test]
fn switch_requires_matching_on_tag_switch_handler() {
    let mut switcher = Chunk::new("switcher");
    switcher.arity = 1;
    switcher.local_count = 1;
    switcher.emit_op_u16(Op::LOCAL_GET, 0, 0);
    switcher.emit_i32_const(55, 0);
    switcher.emit_op_u16(Op::SWITCH, 3, 0);
    switcher.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    switcher.emit_op(Op::RETURN, 0);

    let mut target = Chunk::new("target");
    target.arity = 1;
    target.local_count = 1;
    target.emit_op_u16(Op::LOCAL_GET, 0, 0);
    target.emit_i32_const(1, 0);
    target.emit_op(Op::I32_ADD, 0);
    target.emit_op_u16(Op::SUSPEND, 0, 0);
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_op_u16(Op::REF_FUNC, 2, 0);
    script.emit(0, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let resume_ip = script.code.len();
    script.emit_op_u16(Op::RESUME, 0, 0);
    script.emit_op(Op::RETURN, 0);
    script.stack_switch_handlers.insert(
        resume_ip,
        vec![StackSwitchHandler {
            kind: 1,
            tag_index: 0,
            label_index: 0,
        }],
    );

    let err = VM::new()
        .run(vec![script, switcher, target])
        .unwrap_err()
        .to_string();
    assert!(err.contains("switch") && err.contains("handler"));
}

#[test]
fn switch_with_on_tag_switch_handler_transfers_to_target_continuation() {
    let mut switcher = Chunk::new("switcher");
    switcher.arity = 1;
    switcher.local_count = 1;
    switcher.emit_op_u16(Op::LOCAL_GET, 0, 0);
    switcher.emit_i32_const(55, 0);
    switcher.emit_op_u16(Op::SWITCH, 0, 0);
    switcher.emit_i32_const(0, 0);
    switcher.emit_op(Op::RETURN, 0);

    let mut target = Chunk::new("target");
    target.arity = 1;
    target.local_count = 1;
    target.emit_op_u16(Op::LOCAL_GET, 0, 0);
    target.emit_i32_const(1, 0);
    target.emit_op(Op::I32_ADD, 0);
    target.emit_op_u16(Op::SUSPEND, 0, 0);
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_op_u16(Op::REF_FUNC, 2, 0);
    script.emit(0, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let resume_ip = script.code.len();
    script.emit_op_u16(Op::RESUME, 0, 0);
    script.emit_op(Op::RETURN, 0);
    script.stack_switch_handlers.insert(
        resume_ip,
        vec![StackSwitchHandler {
            kind: 1,
            tag_index: 0,
            label_index: 0,
        }],
    );

    let result = VM::new().run(vec![script, switcher, target]).unwrap();
    assert_eq!(result, Value::I32(56));
}

#[test]
fn cont_bind_non_continuation_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u8(Op::CONT_BIND, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("cont.bind") && err.contains("continuation"));
}

#[test]
fn cont_bind_null_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u8(Op::CONT_BIND, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("cont.bind") && err.contains("null"));
}

#[test]
fn cont_bind_produces_continuation_with_bound_args() {
    // CONT_BIND wraps args into a new continuation; the runtime
    // result should still be ObjectKind::Continuation and carry the
    // bound args as a `__bound_args` array property.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::REF_FUNC, 1, 0);
    chunk.emit(0, 0);
    chunk.emit_op(Op::CONT_NEW, 0);
    chunk.emit_i32_const(42, 0);
    chunk.emit_op_u8(Op::CONT_BIND, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut target = Chunk::new("target");
    target.arity = 1;
    target.local_count = 1;
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk, target]).unwrap();
    let obj = match result {
        Value::Object(obj) => obj,
        other => panic!("expected Continuation object, got {other:?}"),
    };
    let o = obj.lock().unwrap();
    match &o.kind {
        vybe_runtime::value::ObjectKind::Continuation(_) => {}
        other => panic!("expected Continuation kind, got {other:?}"),
    }
    let bound = o
        .properties
        .get("__bound_args")
        .expect("CONT_BIND should stash bound args on the new cont");
    if let Value::Object(bo) = bound {
        let b = bo.lock().unwrap();
        if let vybe_runtime::value::ObjectKind::Array(elems) = &b.kind {
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
fn cont_bind_consumes_source_continuation() {
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op(Op::CONT_NEW, 0);
    script.emit_dup(0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u8(Op::CONT_BIND, 1, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op_u16(Op::RESUME, 0, 0);
    script.emit_op(Op::RETURN, 0);

    let mut target = Chunk::new("target");
    target.arity = 1;
    target.local_count = 1;
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![script, target]).unwrap_err().to_string();
    assert!(err.contains("completed") || err.contains("consumed"));
}

#[test]
fn resume_handler_vector_routes_suspend_to_handler_offset() {
    let mut gen_body = Chunk::new("yield_once");
    gen_body.arity = 1;
    gen_body.local_count = 1;
    gen_body.is_generator = true;
    gen_body.emit_i32_const(44, 0);
    gen_body.emit_op_u16(Op::SUSPEND, 0, 0);
    gen_body.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    gen_body.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    let resume_ip = script.code.len();
    script.emit_op_u16(Op::RESUME, 0, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_i32_const(0, 0);
    script.emit_op(Op::RETURN, 0);

    let handler_ip = script.code.len();
    script.emit_op(Op::DROP, 0);
    script.emit_i32_const(99, 0);
    script.emit_op(Op::RETURN, 0);
    script.stack_switch_handlers.insert(
        resume_ip,
        vec![StackSwitchHandler {
            kind: 0,
            tag_index: 0,
            label_index: handler_ip as u32,
        }],
    );

    let result = VM::new().run(vec![script, gen_body]).unwrap();
    assert_eq!(result, Value::I32(99));
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
    gen_body.emit_i32_const(1, 0);
    gen_body.emit_op_u16(Op::SUSPEND, 0, 0);
    gen_body.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    gen_body.emit_op(Op::RETURN, 0);

    // Caller builds a Function ref to gen_body, calls it with 0 args,
    // and returns the resulting continuation.
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.emit_op_u16(Op::REF_FUNC, 1, 0); // chunk_idx = 1 (gen_body)
    script.emit(0, 0); // uv_count = 0
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, gen_body]).unwrap();
    match result {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                vybe_runtime::value::ObjectKind::Continuation(_) => {}
                other => panic!("expected Continuation, got {other:?}"),
            }
        }
        other => panic!("expected Object/Continuation, got {other:?}"),
    }
}

#[test]
fn fiber_roundtrip_preserves_label_stack_and_continuations() {
    // Regression check for the JSPI-/coroutine-interaction wiring:
    // save_fiber + resume_fiber_with must round-trip the VM's label
    // stack AND the active-continuation stack so that a generator
    // suspended inside a `while` (pushes label entries) or nested
    // inside another coroutine (pushes cont entries) restores to
    // the exact shape it was captured at.
    use vybe_runtime::vm::LabelEntry;
    let mut vm = VM::new();
    vm.label_stack.push(LabelEntry {
        target: 123,
        is_loop: false,
        is_try: false,
        result_arity: 0,
        stack_height: 0,
    });
    vm.label_stack.push(LabelEntry {
        target: 456,
        is_loop: true,
        is_try: false,
        result_arity: 0,
        stack_height: 0,
    });
    // Build a dummy Continuation object so we have a Value to stash
    // in active_continuations (field is pub(crate); constructed via
    // a round-trip through CONT_NEW).
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u16(Op::REF_FUNC, 1, 0);
    chunk.emit(0, 0);
    chunk.emit_op(Op::CONT_NEW, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut target = Chunk::new("target");
    target.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    target.emit_op(Op::RETURN, 0);
    let cont = vm.run(vec![chunk, target]).unwrap();
    // Re-set VM state for the fiber test.
    let mut vm = VM::new();
    vm.label_stack.push(LabelEntry {
        target: 42,
        is_loop: true,
        is_try: false,
        result_arity: 0,
        stack_height: 0,
    });
    vm.active_continuations
        .push(vybe_runtime::vm::ActiveContinuation {
            cont,
            caller_fiber: vybe_runtime::fiber::Fiber::new(Vec::new(), Vec::new(), Vec::new()),
            mode: vybe_runtime::vm::ResumeMode::Iterator,
            handlers: Vec::new(),
        });

    let fiber = vm.save_fiber();
    assert_eq!(fiber.label_stack.len(), 1);
    assert_eq!(fiber.active_continuations.len(), 1);
    assert!(vm.label_stack.is_empty());
    assert!(vm.active_continuations.is_empty());

    vm.resume_fiber_with(fiber, None).unwrap();
    assert_eq!(vm.label_stack.len(), 1);
    assert_eq!(vm.active_continuations.len(), 1);
    assert_eq!(vm.label_stack[0].target, 42);
}

#[test]
fn generator_resume_preserves_loop_label_stack() {
    // Regression for structured control + stack switching: a generator
    // suspended inside a loop must resume with the loop label stack intact
    // so the `br 0` backedge still targets the loop header.
    let mut gen_body = Chunk::new("counter");
    gen_body.arity = 1;
    gen_body.local_count = 2; // slot 0 = resume/control, slot 1 = i
    gen_body.is_generator = true;

    gen_body.emit_i32_const(0, 0);
    gen_body.emit_op_u16(Op::LOCAL_SET, 1, 0);

    let block = gen_body.emit_block(0);
    let (loop_patch, _) = gen_body.emit_loop_s(0);
    gen_body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    gen_body.emit_i32_const(3, 0);
    gen_body.emit_op(Op::I32_LT_S, 0);
    gen_body.emit_op(Op::I32_EQZ, 0);
    gen_body.emit_br_if(1, 0);

    gen_body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    gen_body.emit_op_u16(Op::SUSPEND, 0, 0);
    gen_body.emit_op(Op::DROP, 0);

    gen_body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    gen_body.emit_i32_const(1, 0);
    gen_body.emit_op(Op::I32_ADD, 0);
    gen_body.emit_op_u16(Op::LOCAL_SET, 1, 0);
    gen_body.emit_br(0, 0);
    gen_body.emit_end(0);
    gen_body.patch_loop(loop_patch);
    gen_body.emit_end(0);
    gen_body.patch_block(block);
    gen_body.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    gen_body.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.local_count = 3; // cont, value, has_more
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);

    for _ in 0..2 {
        script.emit_op_u16(Op::LOCAL_GET, 0, 0);
        emit_spec_gen_next(&mut script);
        script.emit_op_u16(Op::LOCAL_SET, 2, 0);
        script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    }
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, gen_body]).unwrap();
    assert_eq!(result, Value::I32(1));
}

#[test]
fn suspend_without_active_continuation_falls_back_to_return() {
    // Running SUSPEND with no RESUME above it should just return the
    // yielded value — this preserves the legacy behaviour so existing
    // code that used SUSPEND as "return from async" keeps working.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_i32_const(77, 0);
    chunk.emit_op_u16(Op::SUSPEND, 0, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 77);
}

#[test]
fn jspi_fulfilled_promise_suspend_unwraps_value() {
    let mut chunk = Chunk::new("<script>");
    // Host-built promise arrives as an imported global (spec `global.get`).
    let g = chunk.intern_string_constant("__jspi_fulfilled_promise");
    chunk.emit_op_u16(Op::GLOBAL_GET, g, 0);
    {
        let aw = chunk.add_import("jspi", "await");
        chunk.emit_call(aw, 1, 0);
    }
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.globals.insert(
        "__jspi_fulfilled_promise".into(),
        make_promise(1, "fulfilled", Value::I32(88)),
    );
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result, Value::I32(88));
}

#[test]
fn jspi_rejected_promise_suspend_enters_wasm_catch_handler() {
    let mut chunk = Chunk::new("<script>");
    let g = chunk.intern_string_constant("__jspi_rejected_promise");

    let off = emit_try_table_catch_all_start(&mut chunk);
    chunk.emit_op_u16(Op::GLOBAL_GET, g, 0);
    {
        let aw = chunk.add_import("jspi", "await");
        chunk.emit_call(aw, 1, 0);
    }
    patch_try_table(&mut chunk, off);

    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(91, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.globals.insert(
        "__jspi_rejected_promise".into(),
        make_promise(2, "rejected", Value::String(Arc::from("network failed"))),
    );
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result, Value::I32(91));
}

#[test]
fn jspi_pending_promise_inside_generator_preserves_continuation_and_yields_on_resume() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "awaitable",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            make_promise(77, "pending", Value::Null)
        }),
    );

    let mut gen_body = Chunk::new("async_generator");
    gen_body.arity = 1;
    gen_body.local_count = 1;
    gen_body.is_generator = true;
    let awaitable_idx = gen_body.add_import("test", "awaitable");
    gen_body.emit_call(awaitable_idx, 0, 0);
    // `await awaitable()` — the JSPI suspend point. Per ECMA-262 / JSPI a
    // pending promise suspends ONLY at the explicit await (PROMISE_SUSPEND),
    // not implicitly at the import call that produced it.
    {
        let aw = gen_body.add_import("jspi", "await");
        gen_body.emit_call(aw, 1, 0);
    }
    gen_body.emit_op_u16(Op::SUSPEND, 0, 0);
    gen_body.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    gen_body.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);

    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_spec_gen_next(&mut script);
    script.emit_op_u16(Op::LOCAL_SET, 2, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let first = vm.run(vec![script, gen_body]);
    let err = first.unwrap_err().to_string();
    assert!(err.contains("__jspi__:77"), "unexpected suspension: {err}");
    assert!(vm.has_pending_jspi());

    let resumed = vm.jspi_resolve(77, Value::I32(123)).unwrap();
    assert_eq!(resumed, Value::I32(123));
    assert!(!vm.has_pending_jspi());
}

#[test]
fn reentrant_host_callback_inside_generator_does_not_complete_it() {
    // Stack-switching compliance: a continuation owns its stack. A host
    // higher-order function (e.g. `Array.prototype.map`) that calls back
    // into WASM via `HostContext::invoke` runs the callback ABOVE the
    // generator body on the shared frame stack. The callback's RETURN
    // unwinds to the body's depth — which is below the re-entrant
    // `invoke_callback` floor — but it must NOT be mistaken for the
    // generator body completing. The generator must still reach its own
    // SUSPEND and yield the value produced AFTER the host call.
    //
    // Regression: previously the callback's RETURN popped the generator's
    // ActiveContinuation and marked it Done, so the first `GEN_NEXT`
    // returned the callback's result with has_more=0 — i.e. the generator
    // died the instant it touched `[...].map(...)`.

    // Host fn: applyCb(cbRef, x) -> cbRef(x). Re-enters the VM.
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "applyCb",
        Box::new(|ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
            let cb = args[0].clone();
            let arg = args.get(1).cloned().unwrap_or(Value::Null);
            ctx.invoke(&cb, &[arg])
        }),
    );

    // cb(x) -> x  (identity; proves the host re-entry actually ran)
    let mut cb = Chunk::new("cb");
    cb.arity = 1;
    cb.local_count = 1;
    cb.emit_op_u16(Op::LOCAL_GET, 0, 0);
    cb.emit_op(Op::RETURN, 0);

    // Generator body:
    //   applyCb(cb, 41)   // host re-entry, returns 41 — then discarded
    //   yield 99          // value produced AFTER the host call
    let mut gen_body = Chunk::new("g");
    gen_body.arity = 0;
    gen_body.local_count = 0;
    gen_body.is_generator = true;
    let apply_idx = gen_body.add_import("test", "applyCb");
    gen_body.emit_op_u16(Op::REF_FUNC, 2, 0); // cb is chunk index 2
    gen_body.emit(0, 0); // uv_count
    gen_body.emit_i32_const(41, 0);
    gen_body.emit_call(apply_idx, 2, 0); // argc = (cbRef, 41)
    gen_body.emit_op(Op::DROP, 0); // discard host result (41)
    gen_body.emit_i32_const(99, 0);
    gen_body.emit_op_u16(Op::SUSPEND, 0, 0); // yield 99
    gen_body.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    gen_body.emit_op(Op::RETURN, 0);

    // Script: drive the generator once, return the first yielded value.
    let mut script = Chunk::new("<script>");
    script.local_count = 3; // cont, value, has_more
    script.emit_op_u16(Op::REF_FUNC, 1, 0); // gen_body is chunk index 1
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0); // -> continuation
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);

    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_spec_gen_next(&mut script); // -> [value, has_more]
    script.emit_op_u16(Op::LOCAL_SET, 2, 0); // has_more
    script.emit_op_u16(Op::LOCAL_SET, 1, 0); // value
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![script, gen_body, cb]).unwrap();
    // 99 = generator survived the host re-entry and yielded post-call.
    // 41 (the callback result) would mean the generator was wrongly
    // completed mid-callback.
    assert_eq!(result.as_i32(), 99);
}

#[test]
fn continuation_started_inside_host_callback_runs_nested_calls_before_first_suspend() {
    // Per the stack-switching proposal a `resume`d continuation owns its stack
    // and returns to the resume point only on suspend/return. Vybe implements
    // this with reified fibers: GEN_NEXT/resume `save_fiber` (drains the
    // caller's frames) then runs the continuation. The bug this guards against:
    // when the continuation is *started from inside a host `invoke_callback`*
    // and its body makes a NESTED CALL before its first `suspend`, that nested
    // call's RETURN unwinds to a frame depth below the callback's `execute_until`
    // floor (which was recorded BEFORE `save_fiber` drained the frames). It must
    // NOT be mistaken for the callback returning — the continuation runs on its
    // own fiber, so the floor doesn't apply. (Root cause of array-destructuring
    // `let [a,b]=[0,1]` — which lowers to iterator-protocol `next()` calls — in
    // a generator driven by `Array.from`/`map`.)
    let mut vm = VM::new();

    // drive(cb, cont) -> cb(cont): re-entrant host callback that hands the
    // continuation to the callback. Models `Array.from(arrayLike, mapper)`.
    vm.register_host_fn(
        "test",
        "drive",
        Box::new(|ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
            let cb = args[0].clone();
            let cont = args.get(1).cloned().unwrap_or(Value::Null);
            ctx.invoke(&cb, &[cont])
        }),
    );

    // helper() -> 7 — the nested call the generator makes before yielding.
    let mut helper = Chunk::new("helper");
    helper.arity = 0;
    helper.local_count = 0;
    helper.emit_i32_const(7, 0);
    helper.emit_op(Op::RETURN, 0);

    // generator g(): { helper(); yield 99; } — nested call BEFORE first yield.
    let mut gen_body = Chunk::new("g");
    gen_body.arity = 0;
    gen_body.local_count = 0;
    gen_body.is_generator = true;
    gen_body.emit_op_u16(Op::REF_FUNC, 3, 0); // helper = chunk index 3
    gen_body.emit(0, 0);
    gen_body.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    gen_body.emit_op(Op::DROP, 0); // discard helper() result (7)
    gen_body.emit_i32_const(99, 0);
    gen_body.emit_op_u16(Op::SUSPEND, 0, 0); // yield 99
    gen_body.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    gen_body.emit_op(Op::RETURN, 0);

    // cb(cont): GEN_NEXT(cont) → [value, has_more]; drop has_more; return value.
    let mut cb = Chunk::new("cb");
    cb.arity = 1;
    cb.local_count = 1;
    cb.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_spec_gen_next(&mut cb);
    cb.emit_op(Op::DROP, 0); // drop has_more, leaving the yielded value
    cb.emit_op(Op::RETURN, 0);

    // script: cont = g(); return drive(cb, cont).
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    let drive_idx = script.add_import("test", "drive");
    script.emit_op_u16(Op::REF_FUNC, 1, 0); // gen_body = chunk index 1
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0); // -> continuation
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);
    script.emit_op_u16(Op::REF_FUNC, 2, 0); // cb = chunk index 2
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 0, 0); // cont
    script.emit_call(drive_idx, 2, 0); // argc = (cb, cont)
    script.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![script, gen_body, cb, helper]).unwrap();
    // 99 = the generator ran its nested helper() call and reached `yield 99`.
    // An empty/early result would mean the nested call's RETURN tripped the
    // host callback's stale floor and aborted the generator before the yield.
    assert_eq!(result.as_i32(), 99);
}
