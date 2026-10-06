use std::sync::{Arc, Mutex};
use vybe_platform_ecma as ecma;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::VM;

fn host_ref(index: usize) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(index);
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn named(vm: &VM, module: &str, name: &str) -> Value {
    host_ref(vm.resolve_host_function_index(module, name).unwrap())
}

fn object(value: &Value) -> Arc<Mutex<Object>> {
    let Value::Object(object) = value else { panic!("expected object, got {value:?}") };
    object.clone()
}

fn values(value: &Value) -> Vec<Value> {
    let object = object(value);
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { panic!("expected array") };
    values.clone()
}

fn same_reference(a: &Value, b: &Value) -> bool {
    matches!((a, b), (Value::Object(a), Value::Object(b)) if Arc::ptr_eq(a, b))
}

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}

fn setup() -> VM {
    let mut vm = VM::new();
    ecma::function::register(&mut vm);
    vm.register_free_fn("test:apply-contract", "record", Box::new(|_, args| array(args.to_vec())));
    vm
}

fn apply(vm: &mut VM, target: Value, receiver: Value, arguments: Value) -> Value {
    let builtin = named(vm, "ecma:function", "apply");
    vm.invoke(&builtin, &[target, receiver, arguments]).unwrap()
}

#[test]
fn dense_argument_lists_keep_order_identity_and_ambient_host_abi() {
    let mut vm = setup();
    let target = named(&vm, "test:apply-contract", "record");
    let receiver = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let shared = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    for count in [0, 1, 4, 8, 9, 32] {
        let mut inputs: Vec<Value> = (0..count).map(Value::I32).collect();
        inputs.push(shared.clone());
        let result = apply(&mut vm, target.clone(), receiver.clone(), array(inputs));
        let actual = values(&result);
        assert_eq!(actual.len(), count as usize + 1, "receiver must not enter free host arguments");
        for (index, value) in actual[..count as usize].iter().enumerate() {
            assert_eq!(value.as_i32(), index as i32);
        }
        assert!(same_reference(actual.last().unwrap(), &shared));
    }
}

#[test]
fn null_undefined_lists_are_empty_and_this_arguments_are_not_coerced() {
    let mut vm = setup();
    vm.register_host_fn("test:apply-contract", "observeReceiver",
        Box::new(|_, args| array(args.to_vec())));
    let index = vm.resolve_host_function_index("test:apply-contract", "observeReceiver").unwrap();
    let target = ecma::receiver_host_fn_ref("test:apply-contract", "observeReceiver", index);
    for empty in [Value::Null, Value::Undefined] {
        for receiver in [Value::Null, Value::Undefined, Value::I32(17)] {
            let result = apply(&mut vm, target.clone(), receiver.clone(), empty.clone());
            let observed = values(&result);
            assert_eq!(observed.len(), 1, "empty list supplies only the declared native receiver");
            let actual = &observed[0];
            match &receiver {
                Value::Null => assert!(matches!(actual, Value::Null)),
                Value::Undefined => assert!(matches!(actual, Value::Undefined)),
                Value::I32(value) => assert!(matches!(actual, Value::I32(actual) if actual == value)),
                _ => unreachable!(),
            }
        }
    }
}

#[test]
fn native_array_snapshot_is_released_before_callback_reentry() {
    let mut vm = setup();
    let arguments = object(&array(vec![Value::I32(11), Value::I32(22)]));
    let source = arguments.clone();
    vm.register_free_fn("test:apply-contract", "mutate", Box::new(move |_, args| {
        let mut input = source.try_lock().expect("apply retained argument-array lock across callback");
        let ObjectKind::Array(values) = &mut input.kind else { panic!("expected array") };
        values[0] = Value::I32(999);
        assert_eq!(args.len(), 2);
        assert_eq!(args[0].as_i32(), 11);
        assert_eq!(args[1].as_i32(), 22);
        Value::I32(31)
    }));
    let target = named(&vm, "test:apply-contract", "mutate");
    assert_eq!(apply(&mut vm, target, Value::Null, Value::Object(arguments.clone())).as_i32(), 31);
    let input = arguments.lock().unwrap();
    assert!(matches!(&input.kind, ObjectKind::Array(values) if values[0].as_i32() == 999));
}

#[test]
fn array_like_gets_length_then_inherited_indices_with_current_receiver_and_mutations() {
    let mut vm = setup();
    let list = vybe_runtime::heap::alloc(Object::new());
    let trace = Arc::new(Mutex::new(Vec::new()));
    let length_trace = trace.clone();
    vm.register_free_fn("test:apply-contract", "length", Box::new(move |_, _| {
        length_trace.lock().unwrap().push(10);
        Value::I32(2)
    }));
    let index_trace = trace.clone();
    let source = list.clone();
    vm.register_host_fn("test:apply-contract", "index0", Box::new(move |_, args| {
        index_trace.lock().unwrap().push(0);
        assert!(same_reference(&args[0], &Value::Object(source.clone())));
        source.try_lock().expect("apply held source lock while invoking getter")
            .properties.insert("1".into(), Value::I32(23));
        Value::I32(11)
    }));
    let mut proto = Object::new();
    let index = vm.resolve_host_function_index("test:apply-contract", "index0").unwrap();
    proto.properties.insert("__get_0".into(),
        ecma::receiver_host_fn_ref("test:apply-contract", "index0", index));
    {
        let mut source = list.lock().unwrap();
        source.properties.insert("__get_length".into(), named(&vm, "test:apply-contract", "length"));
        source.properties.insert("1".into(), Value::I32(22));
        source.properties.insert("__proto__".into(), Value::Object(vybe_runtime::heap::alloc(proto)));
    }
    let target = named(&vm, "test:apply-contract", "record");
    let result = apply(&mut vm, target, Value::Null, Value::Object(list));
    let actual = values(&result);
    assert_eq!(actual.iter().map(Value::as_i32).collect::<Vec<_>>(), vec![11, 23]);
    assert_eq!(*trace.lock().unwrap(), vec![10, 0]);
}

#[test]
fn receiver_style_host_targets_get_one_receiver_before_the_user_arguments() {
    let mut vm = setup();
    vm.register_host_fn("test:apply-contract", "method", Box::new(|_, args| {
        array(args.to_vec())
    }));
    let index = vm.resolve_host_function_index("test:apply-contract", "method").unwrap();
    let target = ecma::receiver_host_fn_ref("test:apply-contract", "method", index);
    let receiver = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let result = apply(&mut vm, target, receiver.clone(), array(vec![Value::I32(7)]));
    let actual = values(&result);
    assert_eq!(actual.len(), 2);
    assert!(same_reference(&actual[0], &receiver));
    assert_eq!(actual[1].as_i32(), 7);
}

#[test]
fn argument_getter_mutations_are_seen_by_final_receiver_dispatch() {
    let mut vm = setup();
    vm.register_host_fn("test:apply-contract", "lateReceiver", Box::new(|_, args| array(args.to_vec())));
    let target = named(&vm, "test:apply-contract", "lateReceiver");
    let target_object = object(&target);
    vm.register_free_fn("test:apply-contract", "markReceiver", Box::new(move |_, _| {
        target_object.lock().unwrap().properties.insert("__vybe_method_receiver".into(), Value::Bool(true));
        Value::I32(7)
    }));
    let mut arguments = Object::new();
    arguments.properties.insert("length".into(), Value::I32(1));
    arguments.properties.insert("__get_0".into(), named(&vm, "test:apply-contract", "markReceiver"));
    let arguments = Value::Object(vybe_runtime::heap::alloc(arguments));
    let receiver = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let result = apply(&mut vm, target, receiver.clone(), arguments);
    let result = values(&result);
    assert_eq!(result.len(), 2);
    assert!(same_reference(&result[0], &receiver));
    assert!(matches!(result[1], Value::I32(7)));
}
