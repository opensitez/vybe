use crate::{dispatch::{allocate_array_values, default_gc_array_element}, value::{Object, Value}};
use std::sync::Arc;

#[test]
fn numeric_defaults_have_the_correct_wasm_value_variant() {
    for name in ["i8", "i16", "i32"] {
        assert_eq!(default_gc_array_element(name), Value::I32(0));
    }
    assert_eq!(default_gc_array_element("i64"), Value::I64(0));
    assert_eq!(default_gc_array_element("f32"), Value::F32(0.0));
    assert_eq!(default_gc_array_element("f64"), Value::F64(0.0));
    assert_eq!(default_gc_array_element("v128"), Value::V128([0; 16]));
    assert_eq!(default_gc_array_element("externref"), Value::Null);
    assert_eq!(default_gc_array_element("funcref"), Value::Null);
    if let Value::F32(value) = default_gc_array_element("f32") {
        assert_eq!(value.to_bits(), 0);
    } else { panic!("f32 default is not a typed zero"); }
    if let Value::F64(value) = default_gc_array_element("f64") {
        assert_eq!(value.to_bits(), 0);
    } else { panic!("f64 default is not a typed zero"); }
}

#[test]
fn excessive_unsigned_lengths_fail_before_initialization() {
    for len in [crate::vm::TABLE_ALLOC_LIMIT as usize + 1, u32::MAX as usize, usize::MAX] {
        let error = allocate_array_values(len, Value::Null).unwrap_err();
        assert!(error.to_string().contains("allocation failure"));
    }
}

#[test]
fn repeated_reference_allocation_preserves_identity_and_releases_clones() {
    let object = crate::heap::alloc(Object::new());
    assert_eq!(Arc::strong_count(&object), 1);
    let values = allocate_array_values(1024, Value::Object(object.clone())).unwrap();
    assert_eq!(values.len(), 1024);
    for value in &values {
        match value {
            Value::Object(actual) => assert!(Arc::ptr_eq(actual, &object)),
            _ => panic!("reference identity was replaced"),
        }
    }
    assert_eq!(Arc::strong_count(&object), 1025);
    drop(values);
    assert_eq!(Arc::strong_count(&object), 1);
    assert!(allocate_array_values(0, Value::Object(object.clone())).unwrap().is_empty());
    assert_eq!(Arc::strong_count(&object), 1);
}

#[test]
fn initialized_numeric_vectors_preserve_every_element() {
    for value in [Value::I32(-37), Value::I64(i64::MIN), Value::F32(-0.0), Value::V128([0x80; 16])] {
        let values = allocate_array_values(257, value.clone()).unwrap();
        assert_eq!(values, vec![value; 257]);
    }
}

#[test]
fn repeated_float_initialization_preserves_exact_bits() {
    for len in [0, 1, 7, 8, 17, 257, 4096] {
        for bits in [0, 0x8000_0000, 0x7f80_0000, 0xff80_0000, 0x7fc0_1234, 0x7f80_0001] {
            let values = allocate_array_values(len, Value::F32(f32::from_bits(bits))).unwrap();
            assert_eq!(values.len(), len);
            for value in values {
                match value {
                    Value::F32(actual) => assert_eq!(actual.to_bits(), bits),
                    _ => panic!("f32 array element changed its Wasm value type"),
                }
            }
        }
        for bits in [
            0, 0x8000_0000_0000_0000, 0x7ff0_0000_0000_0000,
            0xfff0_0000_0000_0000, 0x7ff8_0000_0000_1234, 0x7ff0_0000_0000_0001,
        ] {
            let values = allocate_array_values(len, Value::F64(f64::from_bits(bits))).unwrap();
            assert_eq!(values.len(), len);
            for value in values {
                match value {
                    Value::F64(actual) => assert_eq!(actual.to_bits(), bits),
                    _ => panic!("f64 array element changed its Wasm value type"),
                }
            }
        }
    }
}

#[test]
fn array_elements_keep_references_alive_until_the_last_element_is_released() {
    let object = crate::heap::alloc(Object::new());
    let weak = Arc::downgrade(&object);
    let mut values = allocate_array_values(257, Value::Object(object)).unwrap();
    assert_eq!(weak.strong_count(), 257);
    values.truncate(1);
    assert_eq!(weak.strong_count(), 1);
    values.clear();
    assert!(weak.upgrade().is_none());
}
