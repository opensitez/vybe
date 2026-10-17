use std::hint::black_box;
use std::sync::Arc;
use std::time::Instant;
use vybe_runtime::{canon_value, component::ValType, heap, shared_memory::SharedMemory, value::{Object, Value}};

fn previous_string(memory: &SharedMemory, ptr: u32, len: u32) -> Value {
    assert!((ptr as usize).saturating_add(len as usize) <= memory.len());
    let mut bytes = vec![0u8; len as usize];
    memory.read_bytes(ptr as usize, &mut bytes);
    let text = String::from_utf8(bytes).unwrap();
    Value::String(Arc::from(text.as_str()))
}

fn previous_byte_list(memory: &SharedMemory, ptr: u32, len: u32) -> Value {
    assert!(len < (1 << 28));
    assert!((ptr as usize).saturating_add(len as usize) <= memory.len());
    let mut bytes = vec![0u8; len as usize];
    if len != 0 { memory.read_bytes(ptr as usize, &mut bytes); }
    let items = bytes.into_iter().map(|byte| Value::I32(byte as i32)).collect();
    Value::Object(heap::alloc(Object::new_array(items)))
}

fn paired(label: &str, mut run: impl FnMut(bool) -> u128) {
    run(false);
    run(true);
    let mut before = Vec::with_capacity(21);
    let mut after = Vec::with_capacity(21);
    let mut ratios = Vec::with_capacity(21);
    for round in 0..21 {
        let (a, b) = if round % 2 == 0 { (run(false), run(true)) }
            else { let b = run(true); (run(false), b) };
        before.push(a); after.push(b); ratios.push(a as f64 / b as f64);
    }
    before.sort_unstable(); after.sort_unstable(); ratios.sort_by(f64::total_cmp);
    println!("{label}: before_p50_ns={} after_p50_ns={} before_p95_ns={} after_p95_ns={} ratio_p10={:.3} ratio_p50={:.3} ratio_p90={:.3}", before[10], after[10], before[19], after[19], ratios[2], ratios[10], ratios[18]);
}

#[test]
#[ignore = "alternating loaded-machine canonical string lifting measurement"]
fn canonical_string_read_loaded_benchmark() {
    for (size, count) in [(64usize, 100_000usize), (4096, 20_000), (65536, 2000)] {
        let memory = SharedMemory::new(size + 64);
        let text = "x".repeat(size);
        memory.write_bytes(32, text.as_bytes());
        let ty = ValType::String;
        assert_eq!(previous_string(&memory, 32, size as u32).as_str(), text);
        assert_eq!(canon_value::load_pair_public(&memory, &ty, 32, size as u32).unwrap().as_str(), text);
        paired(&format!("canonical string size={size} count={count}"), |current| {
            let start = Instant::now();
            for _ in 0..count {
                let value = if current {
                    canon_value::load_pair_public(black_box(&memory), &ty, 32, size as u32).unwrap()
                } else {
                    previous_string(black_box(&memory), 32, size as u32)
                };
                drop(black_box(value));
            }
            start.elapsed().as_nanos()
        });
    }
}

#[test]
#[ignore = "alternating loaded-machine canonical byte-list lifting measurement"]
fn canonical_byte_list_read_loaded_benchmark() {
    for (size, count) in [(64usize, 20_000usize), (4096, 2000), (65536, 200)] {
        let memory = SharedMemory::new(size + 64);
        let bytes: Vec<u8> = (0..size).map(|index| index as u8).collect();
        memory.write_bytes(32, &bytes);
        let ty = ValType::List(Box::new(ValType::U8));
        let before = previous_byte_list(&memory, 32, size as u32);
        let current = canon_value::load_pair_public(&memory, &ty, 32, size as u32).unwrap();
        assert_eq!(canon_value::array_items_public(&before), canon_value::array_items_public(&current));
        drop(before);
        drop(current);
        paired(&format!("canonical bytes size={size} count={count}"), |current| {
            let start = Instant::now();
            for _ in 0..count {
                let value = if current {
                    canon_value::load_pair_public(black_box(&memory), &ty, 32, size as u32).unwrap()
                } else {
                    previous_byte_list(black_box(&memory), 32, size as u32)
                };
                drop(black_box(value));
            }
            start.elapsed().as_nanos()
        });
    }
}
