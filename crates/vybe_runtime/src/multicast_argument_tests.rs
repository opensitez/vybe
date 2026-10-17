use super::*;
use std::sync::atomic::{AtomicUsize, Ordering};

fn handler(vm: &mut VM, name: &str, callback: impl Fn(&mut crate::vm::HostContext<'_>, &[Value]) -> Value + Send + Sync + 'static) -> Value {
    vm.register_host_fn("test:multicast", name, Box::new(callback));
    let index = vm.host_registry[&("test:multicast".to_string(), name.to_string())];
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(index);
    Value::Object(crate::heap::alloc(object))
}

fn list(values: Vec<Value>) -> Value {
    Value::Object(crate::heap::alloc(Object::new_array(values)))
}

#[test]
fn nested_handlers_preserve_order_argument_identity_and_lifetime() {
    let mut vm = VM::new();
    let payload = crate::heap::alloc(Object::new_array(vec![Value::I32(17)]));
    let weak = Arc::downgrade(&payload);
    let identity = Arc::as_ptr(&payload) as usize;
    let order = Arc::new(Mutex::new(Vec::new()));
    let mut handlers = Vec::new();
    for index in 0..3 {
        let order = order.clone();
        handlers.push(handler(&mut vm, &format!("nested_{index}"), move |_, args| {
            assert_eq!(args.len(), 1);
            let Value::Object(actual) = &args[0] else { panic!("object argument lost"); };
            assert_eq!(Arc::as_ptr(actual) as usize, identity);
            order.lock().unwrap().push(index);
            Value::I32(index)
        }));
    }
    let last = handlers.pop().unwrap();
    let callee = list(vec![list(handlers), last]);
    vm.stack = vec![callee, Value::Undefined, Value::Object(payload.clone())];
    vm.call_value(2).unwrap();
    assert_eq!(vm.stack, vec![Value::I32(2)]);
    assert_eq!(*order.lock().unwrap(), vec![0, 1, 2]);
    drop(payload);
    assert!(weak.upgrade().is_none(), "consumed arguments must not remain retained");
}

#[test]
fn mutations_do_not_change_the_current_handler_snapshot() {
    let mut vm = VM::new();
    let callee = crate::heap::alloc(Object::new_array(Vec::new()));
    let weak = Arc::downgrade(&callee);
    let calls = Arc::new(AtomicUsize::new(0));
    let first = handler(&mut vm, "clear_handlers", move |_, args| {
        assert!(args.is_empty());
        let callee = weak.upgrade().unwrap();
        let mut callee = callee.lock().unwrap();
        let ObjectKind::Array(handlers) = &mut callee.kind else { panic!("not a handler list"); };
        handlers.clear();
        Value::I32(1)
    });
    let observed = calls.clone();
    let second = handler(&mut vm, "after_clear", move |_, args| {
        assert!(args.is_empty());
        observed.fetch_add(1, Ordering::Relaxed);
        Value::I32(2)
    });
    callee.lock().unwrap().kind = ObjectKind::Array(vec![first, second]);
    vm.stack = vec![Value::Object(callee.clone()), Value::Undefined];
    vm.call_value(1).unwrap();
    assert_eq!(vm.stack, vec![Value::I32(2)]);
    assert_eq!(calls.load(Ordering::Relaxed), 1);
    vm.stack = vec![Value::Object(callee), Value::Undefined, Value::I32(9)];
    vm.call_value(2).unwrap();
    assert_eq!(vm.stack, vec![Value::Null]);
    assert_eq!(calls.load(Ordering::Relaxed), 1);
}

#[test]
fn non_callable_handler_stops_before_later_handlers() {
    let mut vm = VM::new();
    let calls = Arc::new(AtomicUsize::new(0));
    let observed = calls.clone();
    let later = handler(&mut vm, "must_not_run", move |_, _| {
        observed.fetch_add(1, Ordering::Relaxed);
        Value::I32(42)
    });
    vm.stack = vec![list(vec![Value::I32(-1), later]), Value::Undefined];
    assert!(vm.call_value(1).is_err());
    assert_eq!(calls.load(Ordering::Relaxed), 0);
}

#[test]
fn empty_list_with_no_arguments_returns_null() {
    let mut vm = VM::new();
    vm.stack = vec![list(Vec::new())];
    vm.call_value(0).unwrap();
    assert_eq!(vm.stack, vec![Value::Null]);
}

#[test]
fn inline_boundary_preserves_arguments_and_nested_reentry() {
    for user_count in 0..=6 {
        for nested in [false, true] {
            let mut vm = VM::new();
            let payload = crate::heap::alloc(Object::new_array(vec![Value::I32(17)]));
            let identity = Arc::as_ptr(&payload) as usize;
            let weak = Arc::downgrade(&payload);
            let leaf = handler(&mut vm, "boundary", move |_, args| {
                assert_eq!(args.len(), user_count);
                for arg in args {
                    let Value::Object(actual) = arg else { panic!("object argument lost"); };
                    assert_eq!(Arc::as_ptr(actual) as usize, identity);
                }
                Value::I32(user_count as i32)
            });
            let callee = if nested { list(vec![list(vec![leaf])]) } else { list(vec![leaf]) };
            vm.stack = vec![callee, Value::Undefined];
            vm.stack.extend((0..user_count).map(|_| Value::Object(payload.clone())));
            vm.call_value(user_count + 1).unwrap();
            assert_eq!(vm.stack, vec![Value::I32(user_count as i32)]);
            drop(payload);
            assert!(weak.upgrade().is_none());
        }
    }
}

#[test]
fn snapshot_mutation_preserves_inline_and_heap_handler_lists() {
    for count in [1, 2, 3, 4, 5, 8] {
        let mut vm = VM::new();
        let callee = crate::heap::alloc(Object::new_array(Vec::new()));
        let weak = Arc::downgrade(&callee);
        let calls = Arc::new(AtomicUsize::new(0));
        let first = handler(&mut vm, "snapshot_clear", move |_, args| {
            assert!(args.is_empty());
            let callee = weak.upgrade().unwrap();
            let mut callee = callee.lock().unwrap();
            let ObjectKind::Array(handlers) = &mut callee.kind else { panic!("not a handler list"); };
            handlers.clear();
            Value::I32(1)
        });
        let observed = calls.clone();
        let later = handler(&mut vm, "snapshot_later", move |_, args| {
            assert!(args.is_empty());
            observed.fetch_add(1, Ordering::Relaxed);
            Value::I32(42)
        });
        let mut handlers = vec![first];
        handlers.extend((1..count).map(|_| later.clone()));
        callee.lock().unwrap().kind = ObjectKind::Array(handlers);
        vm.stack = vec![Value::Object(callee.clone()), Value::Undefined];
        vm.call_value(1).unwrap();
        assert_eq!(vm.stack, vec![Value::I32(if count == 1 { 1 } else { 42 })]);
        assert_eq!(calls.load(Ordering::Relaxed), count - 1);
        let callee = callee.lock().unwrap();
        assert!(matches!(&callee.kind, ObjectKind::Array(handlers) if handlers.is_empty()));
    }
}

#[test]
fn empty_handler_lists_consume_all_arguments_without_retaining_references() {
    for argc in [0, 1, 3, 4, 5, 8, 32, 255] {
        let mut vm = VM::new();
        let payload = crate::heap::alloc(Object::new_array(vec![Value::I32(17)]));
        let weak = Arc::downgrade(&payload);
        vm.stack = vec![Value::I32(1234), list(Vec::new())];
        vm.stack.extend((0..argc).map(|_| Value::Object(payload.clone())));
        vm.call_value(argc).unwrap();
        assert_eq!(vm.stack, vec![Value::I32(1234), Value::Null]);
        drop(payload);
        assert!(weak.upgrade().is_none(), "empty multicast retained an argument at argc {argc}");
    }
}
