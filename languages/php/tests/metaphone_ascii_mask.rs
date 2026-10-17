//! Preserve the existing shared MVP phonetic contract; this is not an oracle
//! for PHP's full Metaphone algorithm.
use vybe_runtime::{Value,VM};
#[test]
fn actual_php_metaphone_preserves_coercion_limits_and_callback_order() {
 vybe_language_php::register();
 for (source,expected) in [
  ("<?php return metaphone('Smith');","SMT"),
  ("<?php return metaphone('AEIOUHWYBCDFGJKLMNPQRSTVXZ');","ABCDFGJKLMNPQRSTVXZ"),
  ("<?php return metaphone('Smith',2);","SM"),
  ("<?php return metaphone('Smith',0);","SMT"),
  ("<?php return metaphone('Smith',-1);","SMT"),
  ("<?php return metaphone('Xéß🦀Z');","XSSZ"),
  ("<?php return metaphone(false);",""),
  ("<?php return metaphone(12345);","1"),
  ("<?php $trace=''; class Word { public function __toString(){global $trace;$trace.='S';return 'Smith';}} function limit(){global $trace;$trace.='L';return 2;} $r=metaphone(new Word(),limit());return $r.'|'.$trace;","SM|LS"),
 ] {
  let ast=vybe_language_php::parse(source).unwrap();
  let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
  let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();
  wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&chunks)).unwrap();
  let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
  assert_eq!(vm.run(chunks).unwrap(),Value::String(expected.into()),"{source}");
  assert!(vm.last_exception.is_none());
 }
}
