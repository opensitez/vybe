//! Private-entry snapshot iteration must preserve PHP strtr results.
use vybe_runtime::{VM, Value};

#[test]
fn strtr_private_snapshot_preserves_key_and_value_results() {
    vybe_language_php::register();
    for (source, expected) in [
        ("<?php return strtr('abba',['a'=>'x','b'=>'y']);", "xyyx"),
        ("<?php return strtr('abc',[]);", "abc"),
        ("<?php return strtr('cab',['a'=>'x']);", "cxb"),
        ("<?php return strtr('0a0',[0=>'z']);", "zaz"),
    ] {
        let module = vybe_language_php::parse(source).unwrap();
        let profile = vybe_compiler::profile::parse_profile(
            vybe_language_php::profile_source(),
        ).unwrap();
        let chunks = vybe_compiler::Compiler::with_profile(profile)
            .compile(&module).unwrap();
        let wasm = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false, include_name_section: true,
            });
        vybe_platform_wasm::validate_binary(&wasm).unwrap();
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        assert_eq!(vm.run(chunks).unwrap(), Value::String(expected.into()), "{source}");
    }
}
