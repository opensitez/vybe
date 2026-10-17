//! Runtime-only descriptor probes. No PHP compiler or source parser is involved.
//! Binary validation is a separate gate: internal Chunk tests alone do not prove it.
use vybe_runtime::chunk::TypeEntry;
use vybe_runtime::{Chunk, Op, Value, VM};

fn fixture() -> Chunk {
    let mut chunk = Chunk::new("class-descriptor-identity");
    chunk.local_count = 2;
    chunk.types = vec![
        TypeEntry {
            name: "Identity.Described".into(),
            fields: vec!["value".into()],
            ..Default::default()
        },
        TypeEntry {
            name: "#desc Identity.Described".into(),
            fields: vec!["metadata".into()],
            describes_index: 1,
            ..Default::default()
        },
    ];
    chunk
}

fn descriptor(chunk: &mut Chunk, metadata: i32) {
    chunk.emit_i32_const(metadata, 1);
    chunk.emit_struct_new(2, 1, 1);
}

#[test]
fn descriptor_round_trip_preserves_reference_identity_for_both_allocators() {
    for defaulted in [false, true] {
        let mut chunk = fixture();
        descriptor(&mut chunk, 42);
        chunk.emit_op_u16(Op::LOCAL_SET, 0, 1);
        if !defaulted {
            chunk.emit_i32_const(7, 1);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
        if defaulted {
            chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT_DESC, 1, 1);
        } else {
            chunk.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 1, 1, 1);
        }
        chunk.emit_op_u16(Op::REF_GET_DESC, 1, 1);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
        chunk.emit_op(Op::REF_EQ, 1);
        chunk.emit_op(Op::RETURN, 1);
        assert!(matches!(VM::new().run(vec![chunk]).unwrap(), Value::I32(1)));
    }
}

#[test]
fn descriptors_with_the_same_struct_type_are_distinct_references() {
    let mut chunk = fixture();
    descriptor(&mut chunk, 42);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 1);
    descriptor(&mut chunk, 42);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 1);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
    chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT_DESC, 1, 1);
    chunk.emit_op_u16(Op::REF_GET_DESC, 1, 1);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 1);
    chunk.emit_op(Op::REF_EQ, 1);
    chunk.emit_op(Op::RETURN, 1);
    assert!(matches!(VM::new().run(vec![chunk]).unwrap(), Value::I32(0)));
}

fn dispatch_through_base_slot(derived: bool, overrides: bool) -> Value {
    let mut entry = Chunk::new("descriptor-base-slot-dispatch");
    entry.local_count = 1;
    entry.types = vec![
        TypeEntry {
            name: "Dispatch.Base".into(),
            fields: vec!["value".into()],
            ..Default::default()
        },
        TypeEntry {
            name: "#desc Dispatch.Base".into(),
            fields: vec!["method".into()],
            describes_index: 1,
            ..Default::default()
        },
        TypeEntry {
            name: "Dispatch.Derived".into(),
            parent_index: 1,
            fields: vec!["value".into(), "extra".into()],
            ..Default::default()
        },
        TypeEntry {
            name: "#desc Dispatch.Derived".into(),
            parent_index: 2,
            fields: vec!["method".into()],
            describes_index: 3,
            ..Default::default()
        },
    ];

    entry.emit_i32_const(7, 1);
    if derived {
        entry.emit_i32_const(9, 1);
    }
    entry.emit_op_u16(Op::REF_FUNC, if overrides { 2 } else { 1 }, 1);
    entry.emit(0, 1); // No captured upvalues.
    entry.emit_struct_new(if derived { 4 } else { 2 }, 1, 1);
    entry.emit_struct_new_desc(if derived { 3 } else { 1 }, if derived { 2 } else { 1 }, 1);
    entry.emit_op_u16(Op::LOCAL_SET, 0, 1);

    // Internal IR uses callable-first order; the native reader/writer handles
    // Wasm's argument-first order. Load the callable through the base
    // descriptor's stable field index, then pass the explicit receiver.
    // No named lookup or language-specific VM instruction participates.
    entry.emit_op_u16(Op::LOCAL_GET, 0, 1);
    entry.emit_op_u16(Op::REF_GET_DESC, 1, 1);
    entry.emit_struct_field_op(Op::STRUCT_GET, 2, 0, 1);
    entry.emit_op_u16(Op::LOCAL_GET, 0, 1);
    entry.emit_op_u8_u8(Op::CALL_REF, 1, 1, 1);
    entry.emit_op(Op::RETURN, 1);

    let methods: Vec<Chunk> = [false, true].into_iter().map(|overrides| {
        let mut method = Chunk::new(if overrides { "override" } else { "inherited" });
        method.arity = 1;
        method.local_count = 1;
        method.emit_op_u16(Op::LOCAL_GET, 0, 1);
        method.emit_struct_field_op(Op::STRUCT_GET, 1, 0, 1);
        if overrides {
            method.emit_i32_const(100, 1);
            method.emit_op(Op::I32_ADD, 1);
        }
        method.emit_op(Op::RETURN, 1);
        method
    }).collect();
    let mut chunks = vec![entry];
    chunks.extend(methods);
    VM::new().run(chunks).unwrap()
}

#[test]
fn descriptor_method_slot_passes_the_base_receiver_explicitly() {
    assert!(matches!(dispatch_through_base_slot(false, false), Value::I32(7)));
}

#[test]
fn derived_descriptor_preserves_inherited_method_slot() {
    assert!(matches!(dispatch_through_base_slot(true, false), Value::I32(7)));
}

#[test]
fn base_typed_descriptor_dispatch_selects_the_derived_override() {
    assert!(matches!(dispatch_through_base_slot(true, true), Value::I32(107)));
}
