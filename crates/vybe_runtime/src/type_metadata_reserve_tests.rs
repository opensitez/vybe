use super::{Method, TypeDef, TypeRegistry};
use crate::chunk::{PropertyDescriptor, TypeEntry};

fn entry(count: usize, interface: bool) -> TypeEntry {
    let mut entry = TypeEntry::default();
    entry.name = "Metadata_Reserve".into();
    entry.is_interface = interface;
    entry.methods = (0..count).map(|i| (format!("method_{}", i % 4), i)).collect();
    for i in 0..count {
        let name = format!("Field_{i}");
        let mut descriptor = PropertyDescriptor::standard();
        descriptor.writable = i % 2 == 0;
        descriptor.enumerable = i % 3 == 0;
        entry.fields.push(name.clone());
        entry.field_descriptors.insert(name, descriptor);
    }
    entry
}

#[test]
fn reserved_metadata_preserves_fields_descriptors_methods_and_reload_ids() {
    for count in [0usize, 1, 3, 4, 5, 8, 32, 128, 1024] {
        for interface in [false, true] {
            let entry = entry(count, interface);
            let mut registry = TypeRegistry::new();
            registry.load_type_table(&[entry.clone()]);
            let id = registry.get_id(&entry.name).unwrap();
            let ty = registry.get(id).unwrap();
            assert_eq!(ty.is_interface, interface);
            assert_eq!(ty.field_count(), count);
            assert_eq!(ty.methods.len(), count.min(4));
            assert_eq!(ty.required_methods.len(), if interface { count } else { 0 });
            for (index, name) in entry.fields.iter().enumerate() {
                assert_eq!(ty.field_index(name), Some(index));
                let actual = ty.get_field_descriptor(name);
                let expected = &entry.field_descriptors[name];
                assert_eq!(actual.writable, expected.writable);
                assert_eq!(actual.enumerable, expected.enumerable);
                assert_eq!(actual.configurable, expected.configurable);
            }
            for (index, (name, _)) in entry.methods.iter().enumerate() {
                if !entry.methods[index + 1..].iter().any(|(later, _)| later == name) {
                    assert!(matches!(ty.methods.get(name), Some(Method::ChunkFn(actual)) if *actual == index));
                }
            }
            registry.load_type_table(&[entry.clone()]);
            assert_eq!(registry.get_id(&entry.name), Some(id));
            assert_eq!(registry.types.len(), 2);
            let ty = registry.get(id).unwrap();
            assert_eq!(ty.field_count(), count);
            assert_eq!(ty.methods.len(), count.min(4));
            // Preserve existing reload behavior for interface signatures.
            assert_eq!(ty.required_methods.len(), if interface { count * 2 } else { 0 });
        }
    }
}

fn build(entry: &TypeEntry, reserve: bool) -> TypeDef {
    let mut ty = TypeDef::new(&entry.name);
    ty.kind = entry.kind;
    ty.is_interface = entry.is_interface;
    if reserve { ty.reserve_metadata(entry); }
    for (name, descriptor) in &entry.field_descriptors {
        ty.field_descriptors.insert(name.clone(), descriptor.clone());
    }
    if entry.is_interface {
        for (name, _) in &entry.methods { ty.required_methods.push((name.clone(), 0)); }
    }
    for (name, index) in &entry.methods { ty.methods.insert(name.clone(), Method::ChunkFn(*index)); }
    ty
}

#[test]
#[ignore = "paired native type metadata allocation benchmark"]
fn metadata_batch_reservation_benchmark() {
    use std::{hint::black_box, time::Instant};
    for count in [0usize, 1, 4, 8, 32, 128, 1024] {
        for interface in [false, true] {
            let entry = entry(count, interface);
            let rounds = (8192 / count.max(1)).clamp(8, 256);
            let measure = |reserve| {
                let start = Instant::now();
                for _ in 0..rounds { black_box(build(black_box(&entry), black_box(reserve))); }
                start.elapsed().as_secs_f64() * 1e9 / rounds as f64
            };
            let mut old = Vec::new();
            let mut new = Vec::new();
            for trial in 0..31 {
                if trial % 2 == 0 { old.push(measure(false)); new.push(measure(true)); }
                else { new.push(measure(true)); old.push(measure(false)); }
            }
            old.sort_by(f64::total_cmp);
            new.sort_by(f64::total_cmp);
            println!("metadata count={count} interface={interface} old_ns={:.2} new_ns={:.2} ratio={:.2}",
                old[15], new[15], old[15] / new[15]);
        }
    }
}
