use vybe_runtime::Value;

#[test]
fn numeric_clone_from_preserves_float_bits_and_variant_changes() {
    let mut slot = Value::I32(0);
    let mut state = 0x9e37_79b9_7f4a_7c15u64;
    for bits in [0, 1, 0x8000_0000, 0x7f80_0000, 0xff80_0000, 0x7f80_0001, 0x7fc0_1234, u32::MAX] {
        slot.clone_from(&Value::F32(f32::from_bits(bits)));
        match &slot {
            Value::F32(value) => assert_eq!(value.to_bits(), bits),
            _ => panic!("assignment changed the f32 variant"),
        }
    }
    for bits in [0, 1, 0x8000_0000_0000_0000, 0x7ff0_0000_0000_0000,
        0xfff0_0000_0000_0000, 0x7ff0_0000_0000_0001, 0x7ff8_0000_0000_1234, u64::MAX] {
        slot.clone_from(&Value::F64(f64::from_bits(bits)));
        match &slot {
            Value::F64(value) => assert_eq!(value.to_bits(), bits),
            _ => panic!("assignment changed the f64 variant"),
        }
    }
    for _ in 0..4096 {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        slot.clone_from(&Value::F32(f32::from_bits(state as u32)));
        match &slot {
            Value::F32(value) => assert_eq!(value.to_bits(), state as u32),
            _ => panic!("assignment changed the f32 variant"),
        }
        slot.clone_from(&Value::F64(f64::from_bits(state)));
        match &slot {
            Value::F64(value) => assert_eq!(value.to_bits(), state),
            _ => panic!("assignment changed the f64 variant"),
        }
    }
}

#[test]
fn numeric_clone_from_preserves_integer_width_and_signed_boundaries() {
    let mut slot = Value::F64(-0.0);
    for value in [i32::MIN, -1, 0, 1, i32::MAX] {
        slot.clone_from(&Value::I32(value));
        match &slot {
            Value::I32(actual) => assert_eq!(*actual, value),
            _ => panic!("assignment changed the i32 variant"),
        }
    }
    for value in [i64::MIN, i32::MIN as i64 - 1, -1, 0, 1, i32::MAX as i64 + 1, i64::MAX] {
        slot.clone_from(&Value::I64(value));
        match &slot {
            Value::I64(actual) => assert_eq!(*actual, value),
            _ => panic!("assignment changed the i64 variant"),
        }
    }
}

#[test]
fn bulk_scalar_clone_from_preserves_each_payload() {
    let source: Vec<Value> = (0..1024u64).flat_map(|index| [
        Value::F32(f32::from_bits(0x7f80_0001 | index as u32)),
        Value::F64(f64::from_bits(0x7ff0_0000_0000_0001 | index)),
        Value::I32((index as i32).wrapping_mul(i32::MIN + 1)),
        Value::I64((index as i64).wrapping_mul(i64::MIN + 1)),
    ]).collect();
    let mut destination = vec![Value::I32(0); source.len()];
    for _ in 0..2 {
        destination.clone_from_slice(&source);
        for (actual, expected) in destination.iter().zip(&source) {
            match (actual, expected) {
                (Value::F32(a), Value::F32(b)) => assert_eq!(a.to_bits(), b.to_bits()),
                (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
                (Value::I32(a), Value::I32(b)) => assert_eq!(a, b),
                (Value::I64(a), Value::I64(b)) => assert_eq!(a, b),
                _ => panic!("bulk assignment changed a numeric variant"),
            }
        }
    }
}
