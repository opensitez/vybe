//! Focused receiver/exception checks entered through actual VM call frames.
//! No source compiler or language-specific emitter is involved.
use vybe_platform_ecma as ecma;
use vybe_runtime::chunk::ReceiverAbi;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};

fn function_ref(vm: &VM, name: &str, receiver: bool) -> Value {
    let index = vm
        .resolve_host_function_index("receiver-check", name)
        .unwrap();
    if receiver {
        ecma::receiver_host_fn_ref("receiver-check", name, index)
    } else {
        let mut object = Object::new();
        object.kind = ObjectKind::HostFunction(index);
        Value::Object(vybe_runtime::heap::alloc(object))
    }
}

fn builtin_ref(vm: &VM, module: &str, name: &str) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn check(abi: ReceiverAbi) {
    let mut vm = VM::new();
    ecma::register(&mut vm);
    vm.register_host_fn(
        "receiver-check",
        "method",
        Box::new(|_, args| {
            if args.len() != 3 {
                return Value::I32(-1);
            }
            Value::I32(args[0].as_i32() * 100 + args[1].as_i32() * 10 + args[2].as_i32())
        }),
    );
    vm.register_free_fn(
        "receiver-check",
        "free",
        Box::new(|_, args| {
            if args.len() != 2 {
                return Value::I32(-1);
            }
            Value::I32(args[0].as_i32() * 10 + args[1].as_i32())
        }),
    );
    vm.register_host_fn(
        "receiver-check",
        "throw",
        Box::new(|ctx, args| {
            if args.len() != 2 || args[0].as_i32() != 41 || args[1].as_i32() != 2 {
                ctx.throw_value(Value::I32(-1));
            } else {
                ctx.throw_value(Value::I32(73));
            }
            Value::Undefined
        }),
    );
    let method = function_ref(&vm, "method", true);
    let free = function_ref(&vm, "free", false);
    let throwing = function_ref(&vm, "throw", true);
    let call = builtin_ref(&vm, "ecma:function", "call");
    let apply = builtin_ref(&vm, "ecma:function", "apply");
    let reflect_apply = builtin_ref(&vm, "ecma:reflect", "apply");
    vm.register_free_fn(
        "receiver-check",
        "check",
        Box::new(move |ctx, _| {
            assert_eq!(ctx.receiver_is_parameter(), abi == ReceiverAbi::Parameter);
            ctx.set_js_this(Value::I32(777));
            let args = [Value::I32(2), Value::I32(3)];
            let result =
                ecma::function::invoke_with_explicit_this(ctx, &method, Value::I32(41), &args);
            assert_eq!(
                result.as_i32(),
                4123,
                "{abi:?} normal method receiver/order"
            );
            assert_eq!(
                ctx.current_js_this().as_i32(),
                777,
                "{abi:?} normal receiver restoration"
            );
            let result =
                ecma::function::try_invoke_with_explicit_this(ctx, &method, Value::I32(41), &args);
            assert_eq!(
                result.expect("method must not throw").as_i32(),
                4123,
                "{abi:?} fallible method receiver/order"
            );
            assert_eq!(
                ctx.current_js_this().as_i32(),
                777,
                "{abi:?} fallible receiver restoration"
            );
            let result =
                ecma::function::invoke_with_explicit_this(ctx, &free, Value::I32(41), &args);
            assert_eq!(result.as_i32(), 23, "{abi:?} free-function arguments");
            let result =
                ecma::function::try_invoke_with_explicit_this(ctx, &free, Value::I32(41), &args);
            assert_eq!(
                result.expect("free function must not throw").as_i32(),
                23,
                "{abi:?} fallible free arguments"
            );
            let thrown = ecma::function::try_invoke_with_explicit_this(
                ctx,
                &throwing,
                Value::I32(41),
                &[Value::I32(2)],
            )
            .expect_err("throwing callback must propagate");
            assert_eq!(thrown.as_i32(), 73, "{abi:?} thrown value/receiver/order");
            assert_eq!(
                ctx.current_js_this().as_i32(),
                777,
                "{abi:?} restoration after throw"
            );
            for (target, expected) in [(&method, 4123), (&free, 23)] {
                let result = ctx.try_invoke_with_receiver(
                    &call,
                    target.clone(),
                    &[Value::I32(41), Value::I32(2), Value::I32(3)],
                );
                assert_eq!(
                    result.expect("Function.call must not throw").as_i32(),
                    expected,
                    "{abi:?} public Function.call placement"
                );
                assert_eq!(ctx.current_js_this().as_i32(), 777);
                for (name, api) in [
                    ("Function.apply", &apply),
                    ("Reflect.apply", &reflect_apply),
                ] {
                    let list =
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(args.to_vec())));
                    let result = if name == "Function.apply" {
                        ctx.try_invoke_with_receiver(api, target.clone(), &[Value::I32(41), list])
                    } else {
                        ctx.try_invoke(api, &[target.clone(), Value::I32(41), list])
                    };
                    assert_eq!(
                        result.expect("apply must not throw").as_i32(),
                        expected,
                        "{abi:?} public {name} placement"
                    );
                    assert_eq!(
                        ctx.current_js_this().as_i32(),
                        777,
                        "{abi:?} {name} restoration"
                    );
                }
            }
            let result = ctx.try_invoke_with_receiver(
                &call,
                throwing.clone(),
                &[Value::I32(41), Value::I32(2)],
            );
            assert_eq!(
                result
                    .expect_err("Function.call must propagate throw")
                    .as_i32(),
                73
            );
            assert_eq!(ctx.current_js_this().as_i32(), 777);
            for (name, api) in [
                ("Function.apply", &apply),
                ("Reflect.apply", &reflect_apply),
            ] {
                let list = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                    Value::I32(2),
                ])));
                let result = if name == "Function.apply" {
                    ctx.try_invoke_with_receiver(api, throwing.clone(), &[Value::I32(41), list])
                } else {
                    ctx.try_invoke(api, &[throwing.clone(), Value::I32(41), list])
                };
                assert_eq!(result.expect_err("apply must propagate throw").as_i32(), 73);
                assert_eq!(ctx.current_js_this().as_i32(), 777);
            }
            Value::Bool(true)
        }),
    );
    let mut entry = Chunk::new("receiver-check");
    entry.module_receiver_abi = abi;
    entry.local_count = 1;
    let check = entry.add_import("receiver-check", "check");
    entry.emit_call(check, 0, 1);
    entry.emit_op(Op::RETURN, 1);
    assert!(matches!(
        vm.run(vec![entry]).expect("receiver check must complete"),
        Value::Bool(true)
    ));
    println!(
        "{abi:?}: helper/public call/apply receiver placement, free arguments, throws and restoration passed"
    );
}

fn main() {
    check(ReceiverAbi::Ambient);
    check(ReceiverAbi::Parameter);
}
