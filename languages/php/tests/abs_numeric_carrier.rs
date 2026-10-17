//! Public PHP routing, exact BigInt sign, and complete native validation.
use vybe_runtime::{VM,Value};
#[test]
fn parsed_abs_preserves_numeric_values_and_emits_boxed_operands(){
 vybe_language_php::register();
 for (source,expected) in [
  ("<?php return abs(-7.5);",Value::F64(7.5)),
  ("<?php return abs(7.5);",Value::F64(7.5)),
  ("<?php return abs(-0.0);",Value::F64(0.0)),
  ("<?php return abs(-2147483649);",Value::F64(2147483649.0)),
  ("<?php return abs(2147483649);",Value::F64(2147483649.0)),
 ]{
  let module=vybe_language_php::parse(source).unwrap();
  let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
  let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
  wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&chunks)).unwrap();
  let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
  assert_eq!(vm.run(chunks).unwrap(),expected,"{source}");assert!(vm.last_exception.is_none());
 }
}

#[test]
fn parsed_dynamic_abs_preserves_actual_bigint_arguments(){
 vybe_language_php::register();
 let ast=vybe_language_php::parse("<?php return function($x){return abs($x);};").unwrap();
 let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
 let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();
 wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&chunks)).unwrap();
 let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);let f=vm.run(chunks).unwrap();
 for (input,want) in [(Value::bigint_i64(-2147483649),Value::bigint_i64(2147483649)),(Value::bigint_i64(2147483649),Value::bigint_i64(2147483649)),(Value::bigint_i64(i64::MIN),Value::bigint_u64(1u64<<63)),(Value::F64(-7.5),Value::F64(7.5))]{assert_eq!(vm.invoke(&f,&[input]).unwrap(),want);assert!(vm.last_exception.is_none());}
}
