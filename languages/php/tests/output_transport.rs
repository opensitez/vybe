use vybe_compiler::{primitives::platforms, Compiler};
use vybe_runtime::{Value, VM};

#[test]
fn output_transport_preserves_buffering_and_conversion_order() {
    vybe_language_php::register();
    let source = r#"<?php
class OutputValue {
    public function __construct(private $text) {}
    public function __toString(): string { echo '['.$this->text.']'; return $this->text; }
}
ob_start();
echo new OutputValue('a'), new OutputValue('b'), new OutputValue('c');
$printed = print new OutputValue('p');
echo ':', $printed, ':', false, null, true;
return ob_get_clean();
"#;
    let ast = vybe_language_php::parse(source).unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let chunks = Compiler::with_profile(profile).compile(&ast).unwrap();
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm).expect("output transport must remain standard Wasm");
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    assert_eq!(vm.run(chunks).unwrap(), Value::String("[a]a[b]b[c]c[p]p:1:1".into()));

    for (source, expected) in [
        ("<?php ob_start(); echo 'literal'; echo ' output'; return ob_get_clean();", "literal output"),
        ("<?php ob_start(); $s=print_r('hidden',true); echo 'marker:'.$s; return ob_get_clean();", "marker:hidden"),
        ("<?php ob_start(); $n=print_r('hidden'); return ob_get_clean().':'.($n?'yes':'no');", "hidden:yes"),
        ("<?php ob_start(); $n=print_r('hidden','0'); return ob_get_clean().':'.($n?'yes':'no');", "hidden:yes"),
        ("<?php ob_start(); echo 'outer'; ob_start(); echo 'inner'; $inner=ob_get_clean(); echo $inner; return ob_get_clean();", "outerinner"),
        ("<?php ob_start(); ob_start(); echo 'inner'; ob_end_flush(); return ob_get_clean();", "inner"),
        ("<?php ob_start(); $n=print 'literal'; return ob_get_clean().':'.$n;", "literal:1"),
        ("<?php ob_start(); echo \"a\\0b\"; echo ''; return ob_get_clean();", "a\0b"),
        ("<?php function __stdlib_php_string_write_one($x) { return 'wrong'; } ob_start(); echo 'safe'; return ob_get_clean();", "safe"),
    ] {
        let mut byte_sizes = Vec::new();
        for optimized in [false, true] {
            let ast = vybe_language_php::parse(source).unwrap();
            let profile = vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
            let mut compiler = Compiler::with_profile(profile);
            if optimized {
                compiler = compiler.with_argument_emitters(vybe_language_php::parse,
                    vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS);
            }
            let chunks = compiler.compile(&ast).unwrap();
            byte_sizes.push(chunks.iter().map(|chunk| chunk.code.len()).sum::<usize>());
            if optimized {
                // Mixed output may legitimately retain the generic writer.
                // Prove the literal path has one owned writer without
                // mistaking a guest function with its name for that body.
                assert_eq!(chunks.iter().filter(|chunk| chunk.name == "__stdlib_php_string_write_one"
                    && chunk.local_names.iter().any(|local| local.name == "<php-string-output:v1:line0>"
                        && local.origin == vybe_runtime::chunk::LocalOrigin::Compiler)).count(), 1);
            }
            let wasm = vybe_platform_wasm::write_wasm(&chunks);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&wasm).expect("literal output must remain standard Wasm");
            let mut vm = VM::new();
            platforms::register_platforms_all(&mut vm);
            assert_eq!(vm.run(chunks).unwrap(), Value::String(expected.into()));
        }
        eprintln!("literal output bytes {} -> {}: {source}", byte_sizes[0], byte_sizes[1]);
        assert!(byte_sizes[1] < byte_sizes[0]);
    }
}
