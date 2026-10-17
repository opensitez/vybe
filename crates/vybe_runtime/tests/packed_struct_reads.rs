//! Standard packed struct reads and unchanged scalar/descriptor controls.
//! All fixtures use public Runtime APIs; no frontend or backend dependency.
use vybe_runtime::chunk::TypeEntry;
use vybe_runtime::native_gc::{
    CompositeType, FieldType, RecGroup, StorageType, SubType, TypeRef, ValueType,
};
use vybe_runtime::{Chunk, Op, VM, Value};

fn fixture(
    width: u32,
    signed: bool,
    value: i32,
    set: bool,
    mutable: bool,
    field: u32,
) -> Vec<Chunk> {
    let storage = if width == 8 {
        StorageType::I8
    } else {
        StorageType::I16
    };
    let fields = vec![FieldType { storage, mutable }; field as usize + 1];
    let group = RecGroup::new(
        vec![SubType {
            final_: true,
            shared: false,
            supertypes: vec![],
            descriptor: None,
            describes: None,
            composite: CompositeType::Struct(fields),
        }],
        vec![],
    );
    let mut script = Chunk::new("packed-edge");
    script.local_count = 1;
    script.types.push(TypeEntry {
        name: format!("packed{width}"),
        native_identity: group.identity(0),
        fields: (0..=field).map(|n| format!("field{n}")).collect(),
        ..Default::default()
    });
    if set {
        script.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 1, 0);
    } else {
        for _ in 0..=field {
            script.emit_i32_const(value, 0);
        }
        script.emit_struct_new(1, (field + 1) as u16, 0);
    }
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);
    if set {
        script.emit_op_u16(Op::LOCAL_GET, 0, 0);
        script.emit_i32_const(value, 0);
        script.emit_struct_field_op(Op::STRUCT_SET, 1, field, 0);
    }
    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    script.emit_struct_field_op(
        if signed {
            Op::STRUCT_GET_S
        } else {
            Op::STRUCT_GET_U
        },
        1,
        field,
        0,
    );
    script.emit_op(Op::RETURN, 0);
    vec![script]
}

#[test]
fn signed_and_unsigned_i8_i16_edges_after_new_and_set() {
    let values = [
        i32::MIN,
        -65537,
        -65536,
        -32769,
        -32768,
        -257,
        -256,
        -129,
        -128,
        -1,
        0,
        1,
        127,
        128,
        255,
        256,
        257,
        32767,
        32768,
        65535,
        65536,
        i32::MAX,
    ];
    for width in [8, 16] {
        for signed in [false, true] {
            for set in [false, true] {
                for value in values {
                    let modulus = 1_i64 << width;
                    let low = i64::from(value).rem_euclid(modulus);
                    let expected = if signed && low >= modulus / 2 {
                        low - modulus
                    } else {
                        low
                    } as i32;
                    let actual = VM::new()
                        .run(fixture(width, signed, value, set, true, 0))
                        .unwrap();
                    assert!(
                        matches!(actual, Value::I32(n) if n == expected),
                        "width={width} signed={signed} set={set} value={value} expected={expected} actual={actual:?}"
                    );
                }
            }
        }
    }
}

#[test]
fn immutable_packed_fields_are_readable() {
    for width in [8, 16] {
        assert!(matches!(
            VM::new()
                .run(fixture(width, true, -1, false, false, 0))
                .unwrap(),
            Value::I32(-1)
        ));
    }
}

#[test]
fn default_packed_fields_are_zero() {
    for width in [8, 16] {
        let mut chunks = fixture(width, true, 0, false, true, 0);
        chunks[0].code.clear();
        chunks[0].emit_op_u16(Op::STRUCT_NEW_DEFAULT, 1, 0);
        chunks[0].emit_struct_field_op(Op::STRUCT_GET_S, 1, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        assert!(matches!(VM::new().run(chunks).unwrap(), Value::I32(0)));
    }
}

#[test]
fn wide_packed_field_immediate_preserves_signedness() {
    assert!(matches!(
        VM::new()
            .run(fixture(16, true, 65535, false, true, 128))
            .unwrap(),
        Value::I32(-1)
    ));
}

#[test]
fn packed_null_read_traps() {
    for width in [8, 16] {
        for op in [Op::STRUCT_GET_S, Op::STRUCT_GET_U] {
            let mut chunks = fixture(width, true, 0, false, true, 0);
            chunks[0].code.clear();
            chunks[0].emit_ref_null(vybe_runtime::opcode::heaptype::HT_ANY, 0);
            chunks[0].emit_struct_field_op(op, 1, 0, 0);
            chunks[0].emit_op(Op::RETURN, 0);
            let error = VM::new().run(chunks).unwrap_err().to_string();
            assert!(error.contains("null"), "{error}");
        }
    }
}

#[test]
fn repeated_module_loads_resolve_the_declared_storage_width() {
    let mut vm = VM::new();
    for _ in 0..4 {
        assert!(matches!(
            vm.run(fixture(8, true, 128, false, true, 0)).unwrap(),
            Value::I32(-128)
        ));
        assert!(matches!(
            vm.run(fixture(16, true, 32768, false, true, 0)).unwrap(),
            Value::I32(-32768)
        ));
    }
}

#[test]
fn packed_field_read_through_declared_supertype() {
    let field = FieldType {
        storage: StorageType::I8,
        mutable: true,
    };
    let group = RecGroup::new(
        vec![
            SubType {
                final_: false,
                shared: false,
                supertypes: vec![],
                descriptor: None,
                describes: None,
                composite: CompositeType::Struct(vec![field.clone()]),
            },
            SubType {
                final_: true,
                shared: false,
                supertypes: vec![TypeRef::Local(0)],
                descriptor: None,
                describes: None,
                composite: CompositeType::Struct(vec![
                    field,
                    FieldType {
                        storage: StorageType::Value(ValueType::Scalar(0x7f)),
                        mutable: true,
                    },
                ]),
            },
        ],
        vec![],
    );
    for op in [Op::STRUCT_GET_S, Op::STRUCT_GET_U] {
        let mut script = Chunk::new("packed-subtype");
        script.types.push(TypeEntry {
            name: "base".into(),
            fields: vec!["value".into()],
            native_identity: group.identity(0),
            ..Default::default()
        });
        script.types.push(TypeEntry {
            name: "child".into(),
            fields: vec!["value".into(), "tail".into()],
            parent_index: 1,
            native_identity: group.identity(1),
            ..Default::default()
        });
        script.emit_i32_const(65535, 0);
        script.emit_i32_const(99, 0);
        script.emit_struct_new(2, 2, 0);
        script.emit_struct_field_op(op, 1, 0, 0);
        script.emit_op(Op::RETURN, 0);
        let expected = if op == Op::STRUCT_GET_S { -1 } else { 255 };
        assert!(matches!(VM::new().run(vec![script]).unwrap(), Value::I32(n) if n == expected));
    }
}

fn subtype(storage: StorageType, mutable: bool) -> SubType {
    SubType {
        final_: true,
        shared: false,
        supertypes: vec![],
        descriptor: None,
        describes: None,
        composite: CompositeType::Struct(vec![FieldType { storage, mutable }]),
    }
}
fn fixtures(mode: &str, n: usize) -> Vec<Chunk> {
    let mut script = Chunk::new("script");
    let mut shape = subtype(StorageType::Value(ValueType::Scalar(0x7f)), true);
    let described = mode == "descriptor" || mode == "new_desc";
    if described {
        shape.descriptor = Some(TypeRef::Local(1));
    }
    if mode == "packed_s" || mode == "packed_u" {
        shape.composite = CompositeType::Struct(vec![FieldType {
            storage: StorageType::I8,
            mutable: true,
        }]);
    }
    let mut types = vec![shape];
    if described {
        let mut desc = subtype(StorageType::Value(ValueType::Scalar(0x7f)), false);
        desc.describes = Some(TypeRef::Local(0));
        types.push(desc);
    }
    let group = RecGroup::new(types, vec![]);
    script.types.push(TypeEntry {
        name: "fixture-struct".into(),
        fields: vec!["value".into()],
        native_identity: group.identity(0),
        ..Default::default()
    });
    if described {
        script.types.push(TypeEntry {
            name: "fixture-descriptor".into(),
            fields: vec!["value".into()],
            native_identity: group.identity(1),
            describes_index: 1,
            ..Default::default()
        });
    }
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op(Op::RETURN, 0);
    let mut body = Chunk::new("work");
    body.local_count = 3;
    if mode == "local" {
        body.emit_i32_const(0, 0);
        body.emit_op_u16(Op::LOCAL_SET, 0, 0);
        for _ in 0..n {
            body.emit_op_u16(Op::LOCAL_GET, 0, 0);
            body.emit_i32_const(1, 0);
            body.emit_op(Op::I32_ADD, 0);
            body.emit_op_u16(Op::LOCAL_SET, 0, 0);
        }
        body.emit_op_u16(Op::LOCAL_GET, 0, 0);
        body.emit_op(Op::RETURN, 0);
        return vec![script, body];
    }
    if mode == "new" || mode == "new_desc" || mode == "new_get_set" {
        if described {
            body.emit_i32_const(97, 0);
            body.emit_struct_new(2, 1, 0);
            body.emit_op_u16(Op::LOCAL_SET, 2, 0);
        }
        for index in 0..n {
            body.emit_i32_const(index as i32, 0);
            if described {
                body.emit_op_u16(Op::LOCAL_GET, 2, 0);
                body.emit_struct_new_desc(1, 1, 0);
            } else {
                body.emit_struct_new(1, 1, 0);
            }
            if mode == "new_get_set" {
                body.emit_op_u16(Op::LOCAL_SET, 0, 0);
                body.emit_op_u16(Op::LOCAL_GET, 0, 0);
                body.emit_i32_const(97, 0);
                body.emit_struct_field_op(Op::STRUCT_SET, 1, 0, 0);
                body.emit_op_u16(Op::LOCAL_GET, 0, 0);
                body.emit_struct_field_op(Op::STRUCT_GET, 1, 0, 0);
                body.emit_op_u16(Op::LOCAL_SET, 1, 0);
            } else {
                body.emit_op(Op::DROP, 0);
            }
        }
        if mode == "new_get_set" {
            body.emit_op_u16(Op::LOCAL_GET, 1, 0);
        } else {
            body.emit_i32_const(42, 0);
        }
        body.emit_op(Op::RETURN, 0);
        return vec![script, body];
    }
    if described {
        body.emit_i32_const(97, 0);
        body.emit_struct_new(2, 1, 0);
        body.emit_op_u16(Op::LOCAL_SET, 2, 0);
        body.emit_i32_const(42, 0);
        body.emit_op_u16(Op::LOCAL_GET, 2, 0);
        body.emit_struct_new_desc(1, 1, 0);
    } else {
        body.emit_i32_const(if mode.starts_with("packed") { 255 } else { 42 }, 0);
        body.emit_struct_new(1, 1, 0);
    }
    body.emit_op_u16(Op::LOCAL_SET, 0, 0);
    for _ in 0..n {
        body.emit_op_u16(Op::LOCAL_GET, 0, 0);
        match mode {
            "set" => {
                body.emit_i32_const(97, 0);
                body.emit_struct_field_op(Op::STRUCT_SET, 1, 0, 0);
            }
            "descriptor" => {
                body.emit_op_u16(Op::REF_GET_DESC, 1, 0);
                body.emit_struct_field_op(Op::STRUCT_GET, 2, 0, 0);
                body.emit_op_u16(Op::LOCAL_SET, 1, 0);
            }
            mode => {
                let op = match mode {
                    "packed_s" => Op::STRUCT_GET_S,
                    "packed_u" => Op::STRUCT_GET_U,
                    _ => Op::STRUCT_GET,
                };
                body.emit_struct_field_op(op, 1, 0, 0);
                body.emit_op_u16(Op::LOCAL_SET, 1, 0);
            }
        }
    }
    if mode == "set" {
        body.emit_op_u16(Op::LOCAL_GET, 0, 0);
        body.emit_struct_field_op(Op::STRUCT_GET, 1, 0, 0);
    } else {
        body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    }
    body.emit_op(Op::RETURN, 0);
    vec![script, body]
}

#[test]
fn scalar_and_descriptor_reads_preserve_instrumented_behavior() {
    for history in [false, true] {
        for counters in [false, true] {
            for (mode, expected, dispatches) in [("get", 42, 14), ("descriptor", 97, 21)] {
                let mut vm = VM::new();
                vm.record_error_context(history);
                let callee = vm.run(fixtures(mode, 3)).unwrap();
                vm.record_runtime_perf(counters);
                assert!(matches!(vm.invoke(&callee, &[]).unwrap(), Value::I32(n) if n == expected));
                if counters {
                    assert_eq!(vm.take_runtime_perf().unwrap().dispatches, dispatches);
                }
            }
        }
    }
}

fn trap_fixture(descriptor: bool) -> Vec<Chunk> {
    let mut chunks = fixtures(if descriptor { "descriptor" } else { "get" }, 1);
    let mut body = Chunk::new("null-read");
    body.local_count = 1;
    body.emit_ref_null(vybe_runtime::opcode::heaptype::HT_ANY, 9);
    body.emit_op_u16(Op::LOCAL_SET, 0, 9);
    body.emit_op_u16(Op::LOCAL_GET, 0, 10);
    if descriptor {
        body.emit_op_u16(Op::REF_GET_DESC, 1, 10);
    } else {
        body.emit_struct_field_op(Op::STRUCT_GET, 1, 0, 10);
    }
    body.emit_op(Op::RETURN, 11);
    chunks[1] = body;
    chunks
}

#[test]
fn null_read_traps_with_and_without_history() {
    for descriptor in [false, true] {
        for history in [false, true] {
            let mut vm = VM::new();
            vm.record_error_context(history);
            let callee = vm.run(trap_fixture(descriptor)).unwrap();
            let error = vm.invoke(&callee, &[]).unwrap_err().to_string();
            assert!(error.contains("null"), "{error}");
        }
    }
}

#[test]
fn wide_field_immediate_retains_original_path() {
    let mut chunks = fixtures("get", 1);
    chunks[0].types[0].fields = (0..129).map(|n| format!("field{n}")).collect();
    chunks[0].types[0].native_identity = RecGroup::new(
        vec![SubType {
            final_: true,
            shared: false,
            supertypes: vec![],
            descriptor: None,
            describes: None,
            composite: CompositeType::Struct(vec![
                FieldType {
                    storage: StorageType::Value(ValueType::Scalar(0x7f)),
                    mutable: true,
                };
                129
            ]),
        }],
        vec![],
    )
    .identity(0);
    let mut body = Chunk::new("wide-read");
    body.local_count = 1;
    for n in 0..129 {
        body.emit_i32_const(n, 0);
    }
    body.emit_struct_new(1, 129, 0);
    body.emit_op_u16(Op::LOCAL_SET, 0, 0);
    body.emit_op_u16(Op::LOCAL_GET, 0, 0);
    body.emit_struct_field_op(Op::STRUCT_GET, 1, 128, 0);
    body.emit_op(Op::RETURN, 0);
    chunks[1] = body;
    let mut vm = VM::new();
    let callee = vm.run(chunks).unwrap();
    assert!(matches!(vm.invoke(&callee, &[]).unwrap(), Value::I32(128)));
}

#[test]
fn unrelated_local_arithmetic_and_repeated_loads_preserve_values() {
    let mut vm = VM::new();
    for _ in 0..4 {
        let callee = vm.run(fixtures("local", 37)).unwrap();
        assert!(matches!(vm.invoke(&callee, &[]).unwrap(), Value::I32(37)));
    }
}
