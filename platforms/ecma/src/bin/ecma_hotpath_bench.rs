//! Native registered ECMA call-path baselines, without guest compilation.
//! Timings include VM invocation; they are not exclusive adapter CPU time.
use std::time::Instant;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::VM;

fn function(vm: &VM, module: &str, name: &str) -> Value {
    let index = *vm.host_registry.get(&(module.to_owned(), name.to_owned())).expect("registered adapter");
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(index);
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn measure(vm: &mut VM, label: &str, callable: &Value, args: &[Value], iterations: usize) -> Value {
    for _ in 0..100 {
        std::hint::black_box(vm.invoke(callable, args).expect("warmup succeeds"));
    }
    let start = Instant::now();
    let mut last = Value::Undefined;
    for _ in 0..iterations {
        last = vm.invoke(callable, args).expect("invocation succeeds");
        std::hint::black_box(&last);
    }
    eprintln!("{label} operations={iterations} elapsed={:?}", start.elapsed());
    last
}

fn main() {
    let assign_only = std::env::args().any(|argument| argument == "--assign-only");
    let mut vm = VM::new();
    vybe_platform_ecma::function::register(&mut vm);
    vybe_platform_ecma::object::register(&mut vm);
    vybe_platform_ecma::map::register(&mut vm);
    vm.register_free_fn("ecma:benchmark", "first", Box::new(|_ctx, args| {
        args.first().cloned().unwrap_or(Value::Undefined)
    }));
    let target = function(&vm, "ecma:benchmark", "first");
    let apply = function(&vm, "ecma:function", "apply");
    for length in [0, 1, 4, 16, 64, 256].into_iter().filter(|_| !assign_only) {
        let values: Vec<Value> = (0..length).map(|index| Value::I32(index as i32 + 1)).collect();
        let expected = values.first().cloned().unwrap_or(Value::Undefined);
        let array = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values.clone())));
        let args = [target.clone(), Value::Undefined, array];
        for sample in 0..4 {
            for indirect in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                let label = format!("{} length={length} sample={sample}", if indirect { "apply" } else { "direct-native" });
                let result = if indirect {
                    measure(&mut vm, &label, &apply, &args, 20_000)
                } else {
                    measure(&mut vm, &label, &target, &values, 20_000)
                };
                assert_eq!(result, expected);
            }
        }
    }

    let assign = function(&vm, "ecma:object", "assign");
    for count in [16, 128, 1024] {
        let mut source = Object::new();
        let mut destination = Object::new();
        for index in 0..count {
            source.properties.insert(format!("k{index}"), Value::I32(index));
            destination.properties.insert(format!("k{index}"), Value::I32(-1));
        }
        let source = Value::Object(vybe_runtime::heap::alloc(source));
        let destination = Value::Object(vybe_runtime::heap::alloc(destination));
        let args = [destination.clone(), source];
        for sample in 0..4 {
            let result = measure(&mut vm, &format!("assign-existing keys={count} sample={sample}"), &assign, &args, 100);
            assert_eq!(result, destination);
            let Value::Object(object) = &destination else { unreachable!(); };
            assert_eq!(object.lock().unwrap().properties.get(&format!("k{}", count - 1)), Some(&Value::I32(count - 1)));
        }
    }

    let mut entries = indexmap::IndexMap::new();
    for index in 0..2048 {
        entries.insert(Value::I32(index), Value::I32(index + 1));
    }
    let mut object = Object::new();
    object.kind = ObjectKind::Map(entries);
    let map = Value::Object(vybe_runtime::heap::alloc(object));
    let get = function(&vm, "ecma:map", "get");
    for key in [0, 1024, 2047, 4096].into_iter().filter(|_| !assign_only) {
        let args = [map.clone(), Value::I32(key)];
        let expected = if key < 2048 { Value::I32(key + 1) } else { Value::Undefined };
        for sample in 0..4 {
            assert_eq!(measure(&mut vm, &format!("map-get key={key} sample={sample}"), &get, &args, 20_000), expected);
        }
    }
    println!("Native ECMA hot-path result checks passed; timings include VM dispatch");
}
