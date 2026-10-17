use vybe_language_php::emitter::numeric_adapter::{emit_php_inc, emit_php_string_inc_proven};
use vybe_runtime::{Chunk, Op, Value, VM};

fn program(shared: bool, sites: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("owner"),
        vybe_compiler::primitives::functions::create_function_chunk("increment", 1)];
    chunks[1].local_count = 1;
    chunks[1].param_count = 1;
    for site in 0..sites {
        chunks[1].emit_op_u16(Op::LOCAL_GET, 0, 1);
        if shared { emit_php_string_inc_proven(&mut chunks, 1, 1, 1); }
        else { emit_php_inc(&mut chunks, 1, 1, 1); }
        if site + 1 != sites { chunks[1].emit_op(Op::DROP, 1); }
    }
    chunks[1].emit_op(Op::RETURN, 1);
    vybe_compiler::primitives::functions::emit_ref_func(&mut chunks[0], 1, 0, 1);
    chunks[0].emit_op(Op::RETURN, 1);
    chunks
}

fn invoke(chunks: Vec<Chunk>, input: &str) -> Value {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let function = vm.run(chunks).unwrap();
    assert!(vm.global("primitive.php.string_increment").is_none());
    vm.invoke_callback(&function, &[Value::String(input.into())])
}

#[test]
fn string_protocol_preserves_public_increment_results_and_tags() {
    for input in ["", "0", "-0", " 12 ", "12e2", "12bad", "a", "az", "zz",
        "Zz", "ZZ", "a9", "zz9", "z!9", "!9", "é", "éz", "😀z", "\0z"] {
        let actual = invoke(program(true, 1), input);
        let generic = invoke(program(false, 1), input);
        match (&actual, &generic) {
            (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits(), "{input:?}"),
            _ => assert_eq!(actual, generic, "{input:?}"),
        }
    }
    assert_eq!(invoke(program(true, 1), "az"), Value::String("ba".into()));
    assert_eq!(invoke(program(true, 1), "zz"), Value::String("aaa".into()));
}

#[test]
fn repeated_sites_share_one_private_body_and_shrink_complete_bytecode() {
    let inline = program(false, 64);
    let shared = program(true, 64);
    assert_eq!(shared.iter().filter(|c| c.name == "primitive.php.string_increment").count(), 1);
    assert!(shared.iter().map(|c| c.code.len()).sum::<usize>()
        < inline.iter().map(|c| c.code.len()).sum::<usize>());
    assert_eq!(invoke(shared, "zz"), Value::String("aaa".into()));
}

#[test]
fn source_named_function_cannot_impersonate_private_increment_body() {
    let mut chunks = vec![Chunk::new("caller"),
        vybe_compiler::primitives::functions::create_function_chunk("primitive.php.string_increment", 1)];
    chunks[0].local_count = 1;
    chunks[1].local_names.push(vybe_runtime::chunk::LocalName::new(0,
        "<php-private-string-increment-body>", vybe_runtime::chunk::LocalOrigin::Source));
    emit_php_string_inc_proven(&mut chunks, 0, 1, 1);
    assert_eq!(chunks.len(), 3, "source function must not supply the compiler body");
    emit_php_string_inc_proven(&mut chunks, 0, 1, 1);
    assert_eq!(chunks.len(), 3, "compiler body must be reused");
}
