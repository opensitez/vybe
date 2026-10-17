use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};

#[test]
fn descriptor_read_ignores_source_property_with_the_old_storage_name() {
    let descriptor = vybe_runtime::heap::alloc(Object::new());
    let mut instance = Object::new_typed(42);
    instance.descriptor = Some(descriptor.clone());
    assert_eq!(std::mem::size_of_val(&instance.descriptor), std::mem::size_of::<usize>());
    instance.properties.insert("__descriptor".into(), Value::I32(99));
    let instance = vybe_runtime::heap::alloc(instance);
    let mut vm = VM::new();
    vm.register_host_fn("header-fixture", "instance", Box::new(move |_, _| Value::Object(instance.clone())));
    let mut chunk = Chunk::new("descriptor-header-storage");
    let factory = chunk.add_import("header-fixture", "instance");
    chunk.emit_call(factory, 0, 1);
    chunk.emit_op_u16(Op::REF_GET_DESC, 0, 1);
    chunk.emit_op(Op::RETURN, 1);
    let Value::Object(actual) = vm.run(vec![chunk]).unwrap() else {
        panic!("descriptor reference was not returned");
    };
    assert!(Arc::ptr_eq(&actual, &descriptor));
}

#[test]
fn snapshot_restores_descriptor_identity_and_collects_temporary_descriptor() {
    std::thread::spawn(|| {
        vybe_runtime::heap::enable_tracking();
        let descriptor = vybe_runtime::heap::alloc(Object::new());
        let mut instance = Object::new_typed(42);
        instance.descriptor = Some(descriptor.clone());
        let instance = vybe_runtime::heap::alloc(instance);
        let snapshot = vybe_runtime::heap::snapshot();
        let temporary = vybe_runtime::heap::alloc(Object::new());
        let weak = Arc::downgrade(&temporary);
        instance.lock().unwrap().descriptor = Some(temporary);
        vybe_runtime::heap::restore(&snapshot);
        assert!(weak.upgrade().is_none());
        let instance = instance.lock().unwrap();
        let Some(actual) = &instance.descriptor else {
            panic!("saved descriptor was not restored");
        };
        assert!(Arc::ptr_eq(actual, &descriptor));
        assert_eq!(instance.type_id, 42);
    }).join().unwrap();
}

#[test]
fn collection_severs_cycles_through_descriptor_headers() {
    std::thread::spawn(|| {
        vybe_runtime::heap::enable_tracking();
        let snapshot = vybe_runtime::heap::snapshot();
        let first = vybe_runtime::heap::alloc(Object::new());
        let second = vybe_runtime::heap::alloc(Object::new());
        first.lock().unwrap().descriptor = Some(second.clone());
        second.lock().unwrap().descriptor = Some(first.clone());
        let first_weak = Arc::downgrade(&first);
        let second_weak = Arc::downgrade(&second);
        drop(first);
        drop(second);
        vybe_runtime::heap::restore(&snapshot);
        assert!(first_weak.upgrade().is_none());
        assert!(second_weak.upgrade().is_none());
    }).join().unwrap();
}
