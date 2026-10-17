use vybe_compiler::{Compiler, primitives::platforms};
use vybe_runtime::{Chunk, Op, VM, Value};

fn compile(source: &str, optimized: bool) -> Vec<Chunk> {
    compile_with(source, optimized.then_some(vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS))
}

fn compile_with(source: &str, definitions: Option<&'static [vybe_compiler::primitives::argument_emission::Definition]>) -> Vec<Chunk> {
    static INIT: std::sync::Once = std::sync::Once::new();
    INIT.call_once(vybe_language_php::register);
    let ast = vybe_language_php::parse(source).unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let mut compiler = Compiler::with_profile(profile);
    if let Some(definitions) = definitions {
        compiler = compiler.with_argument_emitters(vybe_language_php::parse, definitions);
    }
    compiler.compile(&ast).unwrap()
}

fn size(chunks: &[Chunk]) -> (usize, usize) {
    let mut instructions = 0;
    let bytes = chunks.iter().map(|chunk| chunk.code.len()).sum();
    for chunk in chunks {
        let mut at = 0;
        while at < chunk.code.len() {
            let op = Op::from_code_at(&chunk.code, at);
            instructions += 1;
            at += if op == Op::REF_FUNC {
                7 + usize::from(chunk.code[at + 6] & 0x7f) * 3
            } else {
                4 + op.operand_format().size_in(&chunk.code, at + 4)
            };
        }
        assert_eq!(at, chunk.code.len());
    }
    (bytes, instructions)
}

fn run(chunks: Vec<Chunk>) -> Value {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap()
}

fn validate_wasm(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).expect("string adapter emission must remain standard Wasm");
}

#[test]
fn literal_string_operations_remove_conversion_ladders() {
    for (source, expected) in [
        ("<?php return strrev('abc');", "cba"),
        ("<?php return str_repeat('x',3);", "xxx"),
        ("<?php return strtolower(\"AB\\0C\");", "ab\0c"),
        ("<?php return strtoupper(\"ab\\0c\");", "AB\0C"),
    ] {
        let old = compile(source, false);
        let new = compile(source, true);
        let before = size(&old);
        let after = size(&new);
        eprintln!("{source} old_bytes={} new_bytes={} old_ops={} new_ops={}",
            before.0, after.0, before.1, after.1);
        assert!(before.0 >= after.0 + 1_000);
        assert!(before.1 > after.1);
        assert_eq!(run(old), Value::String(expected.into()));
        assert_eq!(run(new), Value::String(expected.into()));
    }
}

#[test]
fn unknown_operands_keep_the_original_coercion_results() {
    for source in [
        "<?php return strrev(false);",
        "<?php return str_repeat(true,3);",
        "<?php return strtolower(null);",
        "<?php $s='abc'; return strrev($s);",
    ] {
        let old = compile(source, false);
        let new = compile(source, true);
        assert_eq!(run(old), run(new));
    }
}

#[test]
fn repeated_unknown_string_operations_share_the_cold_conversion_body() {
    let source = "<?php $s='Abc'; return strrev($s).strtolower($s).strtoupper($s).str_repeat($s,2);";
    let old = compile(source, false);
    let new = compile(source, true);
    let before = size(&old);
    let after = size(&new);
    eprintln!("mixed variable wrappers old_bytes={} new_bytes={} old_ops={} new_ops={}",
        before.0, after.0, before.1, after.1);
    assert!(before.0 >= after.0 + 4_000);
    assert!(before.1 > after.1);
    assert_eq!(run(old), Value::String("cbAabcABCAbcAbc".into()));
    assert_eq!(run(new), Value::String("cbAabcABCAbcAbc".into()));
}

#[test]
fn production_literal_operations_remove_conversion_ladders() {
    for (source, expected) in [
        ("<?php return strrev('abc');", "cba"),
        ("<?php return str_repeat('x',3);", "xxx"),
        ("<?php return strtolower('ABC');", "abc"),
        ("<?php return strtoupper('abc');", "ABC"),
    ] {
        let old = compile(source, false);
        let new = compile_with(source, Some(vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS));
        let before = size(&old);
        let after = size(&new);
        eprintln!("production {source} old_bytes={} new_bytes={} old_ops={} new_ops={}",
            before.0, after.0, before.1, after.1);
        validate_wasm(&old);
        validate_wasm(&new);
        assert!(before.0 >= after.0 + 1_000);
        assert!(before.1 > after.1);
        assert_eq!(run(old), Value::String(expected.into()));
        assert_eq!(run(new), Value::String(expected.into()));
    }
}

#[test]
fn production_unknown_values_preserve_original_coercion_results() {
    for source in [
        "<?php return strrev(false);",
        "<?php return str_repeat(true,3);",
        "<?php return strtolower(null);",
        "<?php $s='abc'; return strrev($s);",
    ] {
        let old = compile(source, false);
        let new = compile_with(source, Some(vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS));
        assert_eq!(run(old), run(new));
    }
}

#[test]
fn outlined_coercion_preserves_object_conversion() {
    for source in [
        "<?php class TextValue { public function __toString(): string { return 'Abc'; } } return strrev(new TextValue());",
        "<?php namespace TextScope; class TextValue { private function render() { return 'Abc'; } public function __toString(): string { return $this->render(); } } return strrev(new TextValue());",
    ] {
        assert_eq!(run(compile(source, false)), Value::String("cbA".into()));
        assert_eq!(run(compile(source, true)), Value::String("cbA".into()));
    }
}

#[test]
fn outlined_coercion_cannot_bind_a_user_function_with_a_helper_name() {
    let source = "<?php function __stdlib_php_concat_stringify_one($value) { return 'wrong'; } return strrev(42);";
    assert_eq!(run(compile(source, false)), Value::String("24".into()));
    assert_eq!(run(compile(source, true)), Value::String("24".into()));
    let production = compile_with(source, Some(vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS));
    assert_eq!(run(production), Value::String("24".into()));
}

#[test]
fn production_unknown_string_operations_share_one_owned_conversion_body() {
    let source = "<?php $s='Abc'; return strrev($s).strtolower($s).strtoupper($s).str_repeat($s,2);";
    let old = compile(source, false);
    let new = compile_with(source, Some(vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS));
    let before = size(&old);
    let after = size(&new);
    eprintln!("production mixed wrappers old_bytes={} new_bytes={} old_ops={} new_ops={}",
        before.0, after.0, before.1, after.1);
    assert!(before.0 >= after.0 + 4_000);
    assert!(before.1 > after.1);
    assert_eq!(new.iter().filter(|chunk| chunk.name == "__stdlib_php_concat_stringify_one").count(), 1);
    validate_wasm(&old);
    validate_wasm(&new);
    assert_eq!(run(old), Value::String("cbAabcABCAbcAbc".into()));
    assert_eq!(run(new), Value::String("cbAabcABCAbcAbc".into()));
}

#[test]
#[ignore = "execution benchmark; run explicitly with --ignored --nocapture"]
fn production_string_fast_path_execution_benchmark() {
    // Compile outside the timing window. Clone input chunks before each run so
    // cloning is not attributed to execution. Both paths retain identical PHP
    // loop, result, and VM initialization semantics.
    let source = "<?php $s='Abc'; for ($i=0; $i<10000; $i++) { $a=strrev($s); $b=strtolower($s); $c=strtoupper($s); $d=str_repeat($s,2); } return $a.$b.$c.$d;";
    let baseline = compile(source, false);
    let optimized = compile(source, true);
    let before = size(&baseline);
    let after = size(&optimized);
    eprintln!("string execution workload bytes {} -> {}, ops {} -> {}",
        before.0, after.0, before.1, after.1);
    for round in 0..7 {
        // Alternate order to reduce systematic warm-up/order bias. Report raw
        // pairs, not a flaky speed assertion or a whole-WordPress claim.
        for candidate in if round % 2 == 0 { [false, true] } else { [true, false] } {
            let input = if candidate { &optimized } else { &baseline }.clone();
            let start = std::time::Instant::now();
            let result = run(input);
            let elapsed = start.elapsed();
            assert_eq!(result, Value::String("cbAabcABCAbcAbc".into()));
            eprintln!("round={round} optimized={candidate} elapsed_ns={}", elapsed.as_nanos());
        }
    }
}
