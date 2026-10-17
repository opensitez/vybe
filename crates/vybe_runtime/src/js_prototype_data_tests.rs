use super::*;

fn previous(value: Option<&Value>) -> Vec<u8> {
    array_elems(value).unwrap_or_default().iter().map(|value| match value {
        Value::I32(n) => *n as u8,
        Value::I64(n) => *n as u8,
        Value::F64(n) => *n as u8,
        _ => 0,
    }).collect()
}

fn array(values: Vec<Value>) -> Value {
    Value::Object(crate::heap::alloc(crate::value::Object::new_array(values)))
}

#[test]
fn direct_byte_snapshot_matches_existing_numeric_and_fallback_rules() {
    let mut values: Vec<Value> = (-512..=512).map(Value::I32).collect();
    values.extend([Value::I64(i64::MIN),Value::I64(i64::MAX),Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),Value::F64(f64::NEG_INFINITY),Value::F64(12.75),
        Value::String(Arc::from("ignored")),Value::Null,Value::Undefined]);
    let input=array(values);
    assert_eq!(data_bytes(Some(&input)),previous(Some(&input)));
    for value in [Value::Null,Value::Undefined,Value::I32(1)] {
        assert_eq!(data_bytes(Some(&value)),previous(Some(&value)));
    }
    assert_eq!(data_bytes(None),previous(None));
    let empty=array(Vec::new());
    assert!(data_bytes(Some(&empty)).is_empty());
}

#[test]
fn byte_snapshot_does_not_borrow_or_retain_source_references() {
    let root: Arc<str> = Arc::from("unused reference");
    let input=array(vec![Value::I32(7),Value::String(root.clone()),Value::I32(255)]);
    let snapshot=data_bytes(Some(&input));
    assert_eq!(snapshot,[7,0,255]);
    assert_eq!(Arc::strong_count(&root),2);
    let Value::Object(object)=&input else {unreachable!()};
    let mut object=object.lock().unwrap();
    let ObjectKind::Array(values)=&mut object.kind else {unreachable!()};
    values.clear();
    drop(object);
    assert_eq!(Arc::strong_count(&root),1);
    assert_eq!(snapshot,[7,0,255]);
}

#[test]
#[ignore = "paired prototype data snapshot cost, not whole VM timing"]
fn prototype_byte_snapshot_benchmark() {
    use std::{hint::black_box,time::Instant};
    for len in [0usize,16,4096,65536] {
        let input=array((0..len).map(|i|Value::I32(i as i32)).collect());
        let mut samples:[Vec<f64>;2]=std::array::from_fn(|_|Vec::new());
        for trial in 0..21 {
            for new in if trial%2==0 {[false,true]} else {[true,false]} {
                let begin=Instant::now();
                for _ in 0..32 {
                    black_box(if new {data_bytes(black_box(Some(&input)))}
                        else {previous(black_box(Some(&input)))});
                }
                samples[usize::from(new)].push(begin.elapsed().as_secs_f64()/32.);
            }
        }
        for sample in &mut samples {sample.sort_by(f64::total_cmp)}
        assert_eq!(data_bytes(Some(&input)),previous(Some(&input)));
        println!("elements={len} old={:.3}us direct={:.3}us ratio={:.3}x",samples[0][10]*1e6,samples[1][10]*1e6,samples[0][10]/samples[1][10]);
    }
}
