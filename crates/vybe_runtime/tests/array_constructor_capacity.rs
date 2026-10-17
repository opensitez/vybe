use vybe_runtime::{Value, value::{Object, ObjectKind}};

#[test]
fn array_constructor_preserves_length_elements_and_named_property_order() {
    let mut array = Object::new_array(vec![Value::I32(7), Value::I32(9)]);
    assert_eq!(array.properties.len(), 1);
    assert!(matches!(array.get("length"), Value::F64(2.0)));
    assert!(matches!(array.get("0"), Value::I32(7)));
    assert!(matches!(array.get("1"), Value::I32(9)));
    array.set("first".into(), Value::I32(11));
    array.set("second".into(), Value::I32(13));
    let keys: Vec<&str> = array.properties.keys().map(String::as_str).collect();
    assert_eq!(keys, ["length", "first", "second"]);
    assert!(matches!(array.get("first"), Value::I32(11)));
    assert!(matches!(array.get("second"), Value::I32(13)));
    assert!(matches!(array.kind, ObjectKind::Array(_)));
}

#[test]
fn array_growth_keeps_length_and_distinct_instances() {
    let mut first = Object::new_array(Vec::new());
    let second = Object::new_array(Vec::new());
    first.set("3".into(), Value::I32(23));
    assert!(matches!(first.get("length"), Value::F64(4.0)));
    assert!(matches!(first.get("3"), Value::I32(23)));
    assert!(matches!(first.get("0"), Value::Null));
    assert!(matches!(second.get("length"), Value::F64(0.0)));
    match second.kind {
        ObjectKind::Array(elements) => assert!(elements.is_empty()),
        _ => panic!("array constructor changed its kind"),
    }
}
