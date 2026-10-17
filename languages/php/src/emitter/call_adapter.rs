use std::sync::Arc;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_php_dynamic_method_call(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 3 {
        return;
    }
    let chunk = &mut chunks[current];
    let args_slot = chunk.alloc_scratch(1);
    let method_slot = chunk.alloc_scratch(1);
    let receiver_slot = chunk.alloc_scratch(1);
    let target_slot = chunk.alloc_scratch(1);
    let call_slot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    let lookup = chunk.add_import("ecma:value", "getMethodForCall");
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_call(lookup, 3, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, target_slot, line);
    vybe_compiler::primitives::reflection::emit_typeof_is(chunk, "function", line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    let apply = chunk.add_import("ecma:function", "apply");
    chunk.emit_call(apply, 3, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_string_const(&Arc::<str>::from("__call"), line);
    chunk.emit_bool_const(false, line);
    chunk.emit_call(lookup, 3, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, call_slot, line);
    vybe_compiler::primitives::reflection::emit_typeof_is(chunk, "function", line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, call_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_array_new_fixed(0, 2, line);
    chunk.emit_call(apply, 3, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_call(apply, 3, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_php_call_user_func_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        return;
    }

    let chunk = &mut chunks[current];
    let args_slot = chunk.alloc_scratch(1);
    let callable_slot = chunk.alloc_scratch(1);
    let receiver_slot = chunk.alloc_scratch(1);
    let method_slot = chunk.alloc_scratch(1);
    let target_slot = chunk.alloc_scratch(1);

    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, callable_slot, line);

    // PHP arrays can be backed by an ordered Map after an offset write.
    // Function.apply only expands JS arrays/array-like objects, so normalize
    // the argument values here rather than changing ECMA apply semantics.
    let object_values = chunk.add_import("ecma:object", "values");
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_call(object_values, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);

    let is_array = chunk.add_import("ecma:array", "isArray");
    let array_get = chunk.add_import("ecma:array", "get");
    let str_concat = chunk.add_import("wasm:js-string", "concat");
    let global_function = chunk.add_import("vybe:php", "global_function");
    let method_lookup = chunk.add_import("ecma:value", "getMethodForCall");
    let apply = chunk.add_import("ecma:function", "apply");

    chunk.emit_op_u16(Op::LOCAL_GET, callable_slot, line);
    chunk.emit_call(is_array, 1, line);
    let boolean_cast = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(boolean_cast, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, callable_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_call(array_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, callable_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_call(array_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    // A PHP callback pair can name a class loaded by an include. Resolve that
    // name against the live module globals before looking up its static method.
    // Looking up the member on the string itself silently produced no callable.
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    vybe_compiler::primitives::reflection::emit_typeof_is(chunk, "string", line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let class_lookup = chunk.add_import("php:dynamic", "globalByName");
    chunk.emit_call(class_lookup, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_call(method_lookup, 3, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_call(apply, 3, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, callable_slot, line);
    vybe_compiler::primitives::reflection::emit_typeof_is(chunk, "string", line);
    chunk.emit_if(line);

    chunk.emit_string_const(&Arc::<str>::from("__php_fn_"), line);
    chunk.emit_op_u16(Op::LOCAL_GET, callable_slot, line);
    chunk.emit_call(str_concat, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_call(global_function, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, callable_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);

    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_call(apply, 3, line);
    chunk.emit_end(line);
}

pub fn emit_php_method_exists(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 2 {
        return;
    }

    let chunk = &mut chunks[current];
    let method_slot = chunk.alloc_scratch(1);
    let receiver_slot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, receiver_slot, line);
    let is_string = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    let class_lookup = chunk.add_import("php:dynamic", "methodExists");
    chunk.emit_call(class_lookup, 2, line);
    chunk.emit_else(line);

    let lookup = chunk.add_import("ecma:value", "getMethodForCall");
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_call(lookup, 3, line);

    vybe_compiler::primitives::reflection::emit_typeof_is(chunk, "function", line);
    chunk.emit_end(line);
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_compiler::primitives::{bundle, globals, platforms, reflection};
    use vybe_runtime::value::{Object, ObjectKind};
    use vybe_runtime::{VM, Value};

    #[test]
    fn method_exists_tests_the_receiver_without_a_reload() {
        let mut chunks = vec![Chunk::new("script")];
        emit_php_method_exists(&mut chunks, 0, 2, 0);
        let mut prefix = Chunk::new("prefix");
        prefix.emit_op_u16(Op::LOCAL_SET, 0, 0);
        prefix.emit_op_u16(Op::LOCAL_TEE, 1, 0);
        let is_string = prefix.add_import("wasm:js-string", "test");
        prefix.emit_call(is_string, 1, 0);
        assert!(chunks[0].code.starts_with(&prefix.code));
        assert_eq!(chunks[0].local_count, 2);
    }

    #[test]
    fn dynamic_method_lookup_keeps_its_result_for_the_predicate() {
        let mut chunks = vec![Chunk::new("script")];
        emit_php_dynamic_method_call(&mut chunks, 0, 3, 0);
        let mut prefix = Chunk::new("prefix");
        prefix.emit_op_u16(Op::LOCAL_SET, 0, 0);
        prefix.emit_op_u16(Op::LOCAL_SET, 1, 0);
        prefix.emit_op_u16(Op::LOCAL_SET, 2, 0);
        let lookup = prefix.add_import("ecma:value", "getMethodForCall");
        prefix.emit_op_u16(Op::LOCAL_GET, 2, 0);
        prefix.emit_op_u16(Op::LOCAL_GET, 1, 0);
        prefix.emit_bool_const(false, 0);
        prefix.emit_call(lookup, 3, 0);
        prefix.emit_op_u16(Op::LOCAL_TEE, 3, 0);
        assert!(chunks[0].code.starts_with(&prefix.code));
        assert_eq!(chunks[0].local_count, 5);
    }

    fn array(values: Vec<Value>) -> Value {
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
    }

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut value = Object::new();
        for (name, field) in fields {
            value.properties.insert((*name).into(), field.clone());
        }
        Value::Object(vybe_runtime::heap::alloc(value))
    }

    fn predicate(kind: &str, direct: bool) -> Chunk {
        let mut chunk = Chunk::new("php-callback-type-predicate");
        chunk.arity = 1;
        chunk.local_count = 1;
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        if direct {
            reflection::emit_typeof_is(&mut chunk, kind, 0);
        } else {
            // Frozen old emission, used only to compare bytecode size.
            let typeof_fn = chunk.add_import("ecma:value", "typeof");
            chunk.emit_call(typeof_fn, 1, 0);
            chunk.emit_string_const(kind, 0);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunk, 0);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunk, 0);
        }
        chunk.emit_op(Op::RETURN, 0);
        chunk
    }

    fn load(mut body: Chunk) -> (VM, Value, Value) {
        let mut root = Chunk::new("php-callback-adapter-test");
        root.emit_op_u16(Op::REF_FUNC, 1, 0);
        root.emit(0, 0);
        root.emit_op_u16(Op::REF_FUNC, 2, 0);
        root.emit(0, 0);
        root.emit_array_new_fixed(0, 2, 0);
        root.emit_op(Op::RETURN, 0);
        body.emit_op(Op::RETURN, 0);
        let mut echo = Chunk::new("php-test-echo");
        echo.arity = 1;
        echo.local_count = 1;
        echo.emit_op_u16(Op::LOCAL_GET, 0, 0);
        echo.emit_op(Op::RETURN, 0);
        let mut chunks = vec![root, body, echo];
        bundle::finalize_with_runtime_helpers(&mut chunks);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        let mut vm = VM::new();
        platforms::register_platforms_all(&mut vm);
        vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
        let packed = vm.run(chunks).unwrap();
        let Value::Object(packed) = packed else { panic!("expected packed functions") };
        let ObjectKind::Array(functions) = &packed.lock().unwrap().kind else {
            panic!("expected function array")
        };
        (vm, functions[0].clone(), functions[1].clone())
    }

    fn adapter(emitter: fn(&mut [Chunk], usize, u8, u32), argc: u8) -> Chunk {
        let mut body = Chunk::new("php-callback-adapter");
        body.arity = argc;
        body.local_count = u16::from(argc);
        for argument in 0..argc {
            body.emit_op_u16(Op::LOCAL_GET, u16::from(argument), 0);
        }
        let mut chunks = vec![body];
        emitter(&mut chunks, 0, argc, 0);
        chunks.pop().unwrap()
    }

    #[test]
    fn callback_type_predicates_shrink_emitted_bytecode() {
        for kind in ["function", "string"] {
            let before = predicate(kind, false);
            let after = predicate(kind, true);
            eprintln!("PHP {kind} predicate bytes: {} -> {}", before.code.len(), after.code.len());
            assert!(after.code.len() < before.code.len());
        }
    }

    #[test]
    fn callback_type_predicates_preserve_scalar_and_function_results() {
        for kind in ["function", "string"] {
            let (mut vm, predicate, function) = load(predicate(kind, true));
            for input in [Value::Null, Value::Undefined, Value::Bool(false), Value::I32(7), array(vec![]), object(&[])] {
                assert_eq!(vm.invoke(&predicate, &[input]).unwrap(), Value::I32(0));
            }
            assert_eq!(vm.invoke(&predicate, &[Value::String("name".into())]).unwrap(), Value::I32(i32::from(kind == "string")));
            assert_eq!(vm.invoke(&predicate, &[function]).unwrap(), Value::I32(i32::from(kind == "function")));
        }
    }

    #[test]
    fn callback_adapter_preserves_function_and_callback_pair_arguments() {
        let (mut vm, invoke, echo) = load(adapter(emit_php_call_user_func_array, 2));
        let args = array(vec![Value::I32(7), Value::I32(8)]);
        assert_eq!(vm.invoke(&invoke, &[echo.clone(), args.clone()]).unwrap(), Value::I32(7));
        let receiver = object(&[("echo", echo)]);
        let pair = array(vec![receiver, Value::String("echo".into())]);
        assert_eq!(vm.invoke(&invoke, &[pair, args]).unwrap(), Value::I32(7));
    }

    #[test]
    fn dynamic_method_adapter_preserves_lookup_and_arguments() {
        let (mut vm, invoke, echo) = load(adapter(emit_php_dynamic_method_call, 3));
        let receiver = object(&[("echo", echo)]);
        assert_eq!(vm.invoke(&invoke, &[receiver, Value::String("echo".into()), array(vec![Value::I32(11)])]).unwrap(), Value::I32(11));
    }
}
