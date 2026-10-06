//! Actual scoped compiler registry/finalizer, no substituted call dispatcher.
include!("fixtures/numeric_optimization_similar_text_fixture.rs");
#[test]
fn actual_shared_registry_behavior_native_and_invalid_caller_errors() {
 let pair=[program(0,1),program(2,1)];
 assert_eq!(pair[1].iter().filter(|c|c.name=="__stdlib_similar_text").count(),0);
 assert!(!pair[1][0].globals.iter().any(|g|g=="__vybe_similar_text"));
 for p in &pair {validate(p);}
 for (a,b) in [("Hello".into(),"World".into()),("aba".repeat(32),"bab".repeat(32)),
 ("a".repeat(64),"a".repeat(128)),("A😀é".repeat(32),"😀AA".repeat(32))] {
  let a=Value::String(a.into());let b=Value::String(b.into());
  assert_eq!(result(a.clone(),b.clone(),pair[0].clone()),result(a,b,pair[1].clone()));
 }
 for a in [Value::Null,Value::Undefined,Value::Bool(true),Value::F64(2.),Value::TypedNull(0)] {
  for b in [Value::String("abc".into()),Value::Null] {
   assert_eq!(result(a.clone(),b.clone(),pair[0].clone()),result(a.clone(),b,pair[1].clone()));
  }
 }
}
#[test]
fn actual_shared_calls_static_counts_capacity_and_idempotence() {
 for calls in [1,64] {
  let before=program(0,calls);let mut after=program(2,calls);
  let snapshot:Vec<_>=after.iter().map(|c|c.code.clone()).collect();
  bundle::finalize_with_runtime_helpers(&mut after);
  assert_eq!(snapshot,after.iter().map(|c|c.code.clone()).collect::<Vec<_>>());
  eprintln!("shared calls{calls} frozen bytes/ops={}/{} actual={}/{} chunks={}/{}",before.iter().map(|c|c.code.len()).sum::<usize>(),before.iter().map(count).sum::<usize>(),after.iter().map(|c|c.code.len()).sum::<usize>(),after.iter().map(count).sum::<usize>(),before.len(),after.len());
  validate(&after);
 }
 let mut c=vec![Chunk::new("maximum-caller")];c[0].local_count=u16::MAX-8;
 emit(2,&mut c,0,2,59);assert_eq!(c[0].local_count,u16::MAX-6);
 bundle::finalize_with_runtime_helpers(&mut c);assert_eq!(c[0].local_count,u16::MAX);
 // A valid caller with fewer than six free locals retains the module body.
 let mut c=vec![Chunk::new("nearly-full-caller")];c[0].local_count=u16::MAX-5;
 emit(2,&mut c,0,2,59);bundle::finalize_with_runtime_helpers(&mut c);
 assert_eq!(c[0].local_count,u16::MAX-3);
 assert_eq!(c.iter().filter(|chunk|chunk.name=="__stdlib_similar_text").count(),1);
}
#[test]
fn actual_php_shared_scalar_percentage_and_coercion_controls() {
 static REGISTER:std::sync::Once=std::sync::Once::new();REGISTER.call_once(vybe_language_php::register);
 for body in ["return similar_text($a,$b);","similar_text($a,$b,$percent);return $percent;"] {
  for (a,b) in [("World".into(),"Word".into()),("aba".repeat(32),"bab".repeat(32)),("😀".repeat(64),"é".repeat(64))] {
   let source=format!("<?php function similar_subject($a,$b){{{body}}}return similar_subject('{a}','{b}');");
   let ast=vybe_language_php::parse(&source).unwrap();let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
   let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();validate(&chunks);
   let mut vm=runtime(Value::Null,Value::Null);let value=vm.run(chunks).unwrap();
   let expected=runtime(Value::String(a.clone().into()),Value::String(b.clone().into())).run(program(0,1)).unwrap();
   let Value::F64(count)=expected else {panic!("frozen result ABI")};
   assert_eq!(value,Value::F64(if body.starts_with("return") {count}else {count*200./(a.encode_utf16().count()+b.encode_utf16().count()) as f64}));
  }
 }
 let source="<?php $trace=''; class Probe {public $label;function __construct($v){$this->label=$v;}function __toString(){global $trace;$trace.=$this->label;return str_repeat($this->label,64);}}function similar_subject($a,$b){return similar_text($a,$b);} $answer=similar_subject(new Probe('a'),new Probe('b'));return $trace;";
 let ast=vybe_language_php::parse(source).unwrap();let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
 let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();validate(&chunks);
 assert_eq!(runtime(Value::Null,Value::Null).run(chunks).unwrap(),Value::String("ab".into()));
}
