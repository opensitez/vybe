//! Isolated descriptor access cost; not a WordPress speedup measurement.
use std::hint::black_box;
use std::time::{Duration, Instant};
use vybe_runtime::value::{Object, ObjectKind, Properties};
use vybe_runtime::Value;

#[allow(dead_code)]
struct PreviousObjectLayout {
    properties: Properties,
    kind: ObjectKind,
    type_id: usize,
    fields: Vec<Value>,
}

fn measure(object: &Object, header: bool, iterations: usize) -> Duration {
    let start = Instant::now();
    if header {
        for _ in 0..iterations {
            black_box(black_box(object).descriptor.as_ref().cloned().map(Value::Object).unwrap_or(Value::Null));
        }
    } else {
        for _ in 0..iterations {
            black_box(black_box(object).properties.get(black_box("__descriptor")).cloned().unwrap_or(Value::Null));
        }
    }
    start.elapsed()
}

#[test]
#[ignore = "descriptor access timing; excludes VM dispatch, locks, allocation and compiler"]
fn descriptor_header_versus_previous_property_lookup() {
    let descriptor = vybe_runtime::heap::alloc(Object::new());
    let mut object = Object::new_typed(42);
    object.descriptor = Some(descriptor.clone());
    object.properties.insert("__descriptor".into(), Value::Object(descriptor));
    for index in 0..32 {
        object.properties.insert(format!("member{index}"), Value::I32(index));
    }
    let iterations = 1_000_000;
    black_box(measure(&object, false, 10_000));
    black_box(measure(&object, true, 10_000));
    let mut map = Vec::new();
    let mut header = Vec::new();
    for sample in 0..7 {
        if sample % 2 == 0 {
            map.push(measure(&object, false, iterations));
            header.push(measure(&object, true, iterations));
        } else {
            header.push(measure(&object, true, iterations));
            map.push(measure(&object, false, iterations));
        }
    }
    map.sort();
    header.sort();
    println!("descriptor reads={iterations}, map_median={:?}, header_median={:?}", map[3], header[3]);
    println!("object_bytes_before={}, object_bytes_after={}, descriptor_extra_allocation_bytes=0",
        std::mem::size_of::<PreviousObjectLayout>(),
        std::mem::size_of::<Object>());
}
