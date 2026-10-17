//! Actual PHP grammar commit -> complete common AST -> one prepared compiler.
//! No pre-EOF statement publication or native JIT is asserted.
use vybe_language_php::{parse,parse_committed_source_unit,parse_with_peg,parse_with_source_units};
use std::sync::Arc;

#[test]
fn typed_consumer_gets_original_complete_scoped_manifest_exactly_once(){
 for source in [
  "<?php $r=f(2,3,4); function f($x,...$rest){return $x+$rest[0]+$rest[1];} return $r;",
  "<?php namespace StreamUnit; class Base {public $x=4;public function get(){return $this->x;}} class Child extends Base {public function value(){return $this->get()+3;}} $c=new Child();$f=function($o){return $o->value();};return $f($c);",
  "<?php $a=1;$b=&$a;$b+=2;return $a;",
  "<?php function scoped($x){global $shared;static $calls=0;$calls++;return $x+$calls;}return scoped(2);",
  "<b>prefix</b><?php $x=1;echo $x;?><i>suffix</i>",
  "<?php if(false){include 'never.php';} function later(){require_once 'later.php';} return 7;",
 ] {
  let expected=parse(source).unwrap();let expected=format!("{expected:?}");
  assert_eq!(expected,format!("{:?}",parse_with_peg(source).unwrap()));
  let mut calls=0;
  let actual=parse_committed_source_unit(source,|unit|{
   calls+=1;let receipt=unit.grammar_commit();assert_eq!(receipt.source_bytes(),source.len());assert_eq!(receipt.consumed(),receipt.parser_bytes());
   let original_body=unit.manifest().body.as_ptr();let original_len=unit.manifest().body.len();
   let (module,_requests,moved_receipt)=unit.into_parts();assert_eq!(module.body.as_ptr(),original_body,"handoff must move original statement allocation");assert_eq!(module.body.len(),original_len);assert_eq!(receipt,moved_receipt);
   Ok(module)
  }).unwrap();
  assert_eq!(calls,1);assert_eq!(format!("{actual:?}"),expected,"complete Module {source}");
 }
}
#[test]
fn all_include_forms_keep_original_policy_scope_spans_and_preparation(){
 let source="<?php include 'a.php';include_once 'b.php';require 'c.php';function deferred(){require_once 'd.php';}return 1;";
 let (expected,requests)=parse_with_source_units(source).unwrap();assert_eq!(requests.len(),4);
 let (module,actual)=parse_committed_source_unit(source,|unit|{let(module,requests,_)=unit.into_parts();Ok((module,requests))}).unwrap();
 assert_eq!(actual,requests);assert_eq!(format!("{module:?}"),format!("{expected:?}"));
}
#[test]
fn failed_grammar_never_invokes_compiler_consumer_and_errors_match(){
 for source in ["<?php $a + ;","<?php return $a *;","<?php $a='unterminated"]{
  let expected=parse(source).unwrap_err();let mut called=false;
  let actual=parse_committed_source_unit(source,|_|{called=true;Ok(())}).unwrap_err();
  assert!(!called);assert_eq!(actual,expected);
 }
 let mut called=false;
 assert_eq!(parse_committed_source_unit("<?php return 2;",|_|{called=true;Err::<(),_>("consumer rejected unpublished unit".into())}).unwrap_err(),"consumer rejected unpublished unit");assert!(called);
}
#[cfg(feature="service-verification")]
#[test]
fn actual_frontend_builder_census_proves_no_duplicate_parse_for_handoff(){
 use vybe_language_php::with_pratt_admission_census_for_validation as observed;
 let source="<?php $r=f(2);function f($x){return $x+3;}return $r;";
 parse(source).unwrap(); // warm the unchanged shared bootstrap selection
 let (_,before)=observed(||parse(source).unwrap());
 let (_,after)=observed(||parse_committed_source_unit(source,|unit|Ok(unit.into_parts().0)).unwrap());
 assert!(before.builders>0);assert_eq!(before,after,"typed consumer must not trigger a second parser/Pratt build");
}
fn compiler()->vybe_compiler::Compiler{
 vybe_language_php::register();let language=vybe_compiler::languages::find_by_name("php").unwrap();
 vybe_compiler::Compiler::with_profile(vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap())
  .with_optimizations(vybe_runtime::registry::compiler_optimizations(language))
  .with_helper_templates(vybe_language_php::parse,vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS,Arc::new(vybe_compiler::primitives::helper_templates::Pool::default()))
  .with_argument_emitters(vybe_language_php::parse,vybe_language_php::emitter::argument_emission::ARGUMENT_DEFINITIONS)
  .with_run_module_end_hooks(false)
}
#[path="../examples/emission_import_interning/canonical.rs"]mod metadata;
#[test]
fn actual_php_commit_drives_one_reusable_common_ast_emission_and_vm(){
 let source="<?php $r=f(2,3,4);function f($x,...$rest){return $x+$rest[0]+$rest[1];}return $r;";
 let expected=compiler().compile_with_imports(&parse(source).unwrap()).unwrap();let mut calls=0;
 let actual=parse_committed_source_unit(source,|unit|{
  calls+=1;let(module,_,_)=unit.into_parts();let mut emission=compiler().begin_emission_owned(module)?;let mut steps=0;
  while emission.emit_next()?.is_some(){steps+=1;}assert!(steps>0);emission.finish()
 }).unwrap();assert_eq!(calls,1);assert_eq!(expected.chunks.len(),actual.chunks.len());
 for (a,b) in expected.chunks.iter().zip(&actual.chunks){assert_eq!(a.code,b.code);assert_eq!(a.lines,b.lines);assert_eq!(metadata::canonical(&format!("{a:?}")),metadata::canonical(&format!("{b:?}")));}
 let mut vm=vybe_runtime::VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
 let resolved=actual.chunks[0].imports.iter().map(|i|vm.resolve_import_target(&i.module,&i.name).or_else(|error|if matches!(i.module.as_str(),"env"|"*"){Ok(vybe_runtime::ImportTarget::StdlibRedirect(i.name.clone()))}else{Err(error)})).collect::<Result<Vec<_>,_>>().unwrap();
 let result=vm.run_linked(actual.chunks,resolved).unwrap();let number=match result{vybe_runtime::Value::I32(n)=>n as f64,vybe_runtime::Value::I64(n)=>n as f64,vybe_runtime::Value::F64(n)=>n,other=>panic!("unexpected {other:?}")};assert_eq!(number,9.0);
}
