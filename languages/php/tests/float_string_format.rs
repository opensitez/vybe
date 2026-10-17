//! PHP scalar formatter semantics across VM and complete standard native output.
use vybe_runtime::{Value,VM};
#[test]fn dynamic_float_string_format_preserves_negative_zero_and_special_values(){
 vybe_language_php::register();let ast=vybe_language_php::parse("<?php return function($x){return (string)$x;};").unwrap();let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&chunks)).unwrap();let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);let f=vm.run(chunks).unwrap();
 for (input,want) in [(Value::F64(-0.0),"-0"),(Value::F64(0.0),"0"),(Value::F64(f64::NAN),"NAN"),(Value::F64(f64::INFINITY),"INF"),(Value::F64(f64::NEG_INFINITY),"-INF"),(Value::F64(1.5),"1.5"),(Value::bigint_i64(i64::MIN),"-9223372036854775808")]{assert_eq!(vm.invoke(&f,&[input]).unwrap(),Value::String(want.into()));assert!(vm.last_exception.is_none());}
}
