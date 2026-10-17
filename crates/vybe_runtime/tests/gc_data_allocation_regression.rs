//! Numeric array.new_data dispatch coverage, not a WASM loader conformance suite.
use vybe_runtime::{Chunk, Op, VM, Value};
use vybe_runtime::chunk::{CompositeKind, TypeEntry};
use vybe_runtime::value::ObjectKind;

fn program(kind: &str, data: Vec<u8>, offset: i32, count: i32, dropped: bool) -> Chunk {
    let mut chunk = Chunk::new("numeric-data-allocation");
    chunk.local_count = 1;
    chunk.scratch_high_water = 1;
    chunk.result_arity = 1;
    chunk.types.push(TypeEntry {
        name: format!("data-array-{kind}"),
        kind: CompositeKind::Array,
        fields: vec![kind.into()],
        ..Default::default()
    });
    chunk.data_segments.push(data);
    if dropped {
        chunk.emit_op(Op::DATA_DROP, 1);
        chunk.emit_leb_u32(0, 1);
    }
    chunk.emit_i32_const(offset, 1);
    chunk.emit_i32_const(count, 1);
    chunk.emit_op(Op::ARRAY_NEW_DATA, 1);
    chunk.code.extend_from_slice(&1u16.to_be_bytes());
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 1);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn width(kind: &str) -> usize {
    match kind { "i8" => 1, "i16" => 2, "i32" | "f32" => 4, _ => 8 }
}

#[test]
fn numeric_data_allocation_preserves_values_float_bits_and_rtt() {
    for kind in ["i8", "i16", "i32", "i64", "f32", "f64"] {
        for count in [0, 1, 17, 1024] {
            let data: Vec<u8> = (0usize..3 + count * width(kind))
                .map(|i| (i.wrapping_mul(37) + 129) as u8).collect();
            let mut vm = VM::new();
            let result = vm.run(vec![program(kind, data.clone(), 3, count as i32, false)]).unwrap();
            let Value::Object(object) = result else { panic!("expected GC array"); };
            let object = object.lock().unwrap();
            let definition = vm.type_registry.get(object.type_id).unwrap();
            assert_eq!(definition.kind, CompositeKind::Array);
            assert_eq!(definition.field_defs[0].name, kind);
            let ObjectKind::Array(values) = &object.kind else { panic!("expected array backing"); };
            assert_eq!(values.len(), count);
            for (value, bytes) in values.iter().zip(data[3..].chunks_exact(width(kind))) {
                match (kind, value) {
                    ("i8", Value::I32(v)) => assert_eq!(*v, bytes[0] as i32),
                    ("i16", Value::I32(v)) => assert_eq!(*v, u16::from_le_bytes(bytes.try_into().unwrap()) as i32),
                    ("i32", Value::I32(v)) => assert_eq!(*v, i32::from_le_bytes(bytes.try_into().unwrap())),
                    ("i64", Value::I64(v)) => assert_eq!(*v, i64::from_le_bytes(bytes.try_into().unwrap())),
                    ("f32", Value::F32(v)) => assert_eq!(v.to_bits(), u32::from_le_bytes(bytes.try_into().unwrap())),
                    ("f64", Value::F64(v)) => assert_eq!(v.to_bits(), u64::from_le_bytes(bytes.try_into().unwrap())),
                    _ => panic!("wrong numeric lane variant"),
                }
            }
        }
    }
}

#[test]
fn numeric_data_allocation_checks_element_width_and_dropped_segment_bounds() {
    for kind in ["i8", "i16", "i32", "i64", "f32", "f64"] {
        let len = width(kind) * 3;
        for (offset, count, dropped, trap) in [
            (0, 3, false, false), (len as i32, 0, false, false),
            (1, 3, false, true), (len as i32 + 1, 0, false, true),
            (-1, 0, false, true), (0, -1, false, true),
            (0, 0, true, false), (1, 0, true, true), (0, 1, true, true),
        ] {
            let mut vm = VM::new();
            let result = vm.run(vec![program(kind, vec![0xff; len], offset, count, dropped)]);
            assert_eq!(result.is_err(), trap, "{kind} offset={offset} count={count} dropped={dropped}");
            if let Err(error) = result {
                assert!(error.to_string().contains("out of bounds"), "{error}");
            }
        }
    }
}
