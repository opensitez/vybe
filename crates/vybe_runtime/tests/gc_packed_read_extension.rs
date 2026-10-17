//! Packed GC reads use the common RTT storage metadata, not custom value tags.
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::chunk::{CompositeKind, TypeEntry};

fn program(kind: &str, data: Vec<u8>, signed: bool, overwrite: Option<i32>, index: i32) -> Chunk {
    let mut chunk = Chunk::new("packed-array-read");
    chunk.local_count = 1;
    chunk.scratch_high_water = 1;
    chunk.result_arity = 1;
    chunk.types.push(TypeEntry {
        name: format!("packed-array-{kind}"),
        kind: CompositeKind::Array,
        fields: vec![kind.into()],
        ..Default::default()
    });
    chunk.data_segments.push(data);
    chunk.emit_i32_const(0, 1);
    chunk.emit_i32_const(1, 1);
    chunk.emit_op(Op::ARRAY_NEW_DATA, 1);
    chunk.code.extend_from_slice(&1u16.to_be_bytes());
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 1);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 1);
    if let Some(value) = overwrite {
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
        chunk.emit_i32_const(0, 1);
        chunk.emit_i32_const(value, 1);
        chunk.emit_op(Op::ARRAY_SET, 1);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
    chunk.emit_i32_const(index, 1);
    chunk.emit_op_u16(if signed { Op::ARRAY_GET_S } else { Op::ARRAY_GET_U }, 1, 1);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

#[test]
fn packed_data_reads_extend_to_i32() {
    for raw in [0u16, 127, 128, 255, 256, 32767, 32768, 65535] {
        for kind in ["i8", "i16"] {
            for signed in [false, true] {
                let data = if kind == "i8" { vec![raw as u8] } else { raw.to_le_bytes().to_vec() };
                let expected = match (kind, signed) {
                    ("i8", true) => raw as i8 as i32,
                    ("i8", false) => raw as u8 as i32,
                    (_, true) => raw as i16 as i32,
                    (_, false) => raw as i32,
                };
                let actual = VM::new().run(vec![program(kind, data, signed, None, 0)]).unwrap();
                assert_eq!(actual.as_i32(), expected, "{kind} signed={signed} raw={raw}");
            }
        }
    }
}

#[test]
fn packed_reads_after_array_set_keep_only_storage_bits() {
    for raw in [-1i32, i32::MIN, i32::MAX, 0x1ff, 0x1ffff] {
        for kind in ["i8", "i16"] {
            for signed in [false, true] {
                let data = vec![0; if kind == "i8" { 1 } else { 2 }];
                let expected = match (kind, signed) {
                    ("i8", true) => raw as i8 as i32,
                    ("i8", false) => raw as u8 as i32,
                    (_, true) => raw as i16 as i32,
                    (_, false) => raw as u16 as i32,
                };
                let actual = VM::new().run(vec![program(kind, data, signed, Some(raw), 0)]).unwrap();
                assert_eq!(actual.as_i32(), expected, "{kind} signed={signed} raw={raw}");
            }
        }
    }
}

#[test]
fn packed_read_bounds_traps_are_preserved() {
    for kind in ["i8", "i16"] {
        for signed in [false, true] {
            for index in [-1, 1, i32::MAX] {
                let data = vec![0; if kind == "i8" { 1 } else { 2 }];
                let error = VM::new().run(vec![program(kind, data, signed, None, index)])
                    .unwrap_err().to_string();
                assert!(error.contains("out of bounds array access"), "{error}");
            }
        }
    }
}
