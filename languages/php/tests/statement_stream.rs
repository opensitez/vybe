use vybe_language_php::{parse,produce_grameo_statements_with_receipt};
fn gate(source:&str)->vybe_language_php::StatementStreamReceipt {
 let manifest=parse(source).unwrap();let expected=format!("{:?}",manifest.body);let mut body=vec![];
 let receipt=produce_grameo_statements_with_receipt(source,manifest,&mut |index,statement|{assert_eq!(index,body.len());body.push(statement);Ok(())}).unwrap();
 assert_eq!(format!("{body:?}"),expected);assert_eq!(receipt.emitted_statements,body.len());assert!(receipt.parser_committed);receipt
}
#[test]fn fresh_typed_statements_arrive_before_parser_eof(){
 let receipt=gate("<?php $x=1; $x+=2; return $x;");assert!(receipt.completed_captures>=3);assert!(receipt.emitted_before_parser_eof>0);assert!(receipt.freshly_lowered_statements>0);assert!(receipt.largest_capture_nodes>0);
}
#[test]fn late_function_and_class_facts_preserve_normalized_sequence(){
 gate("<?php $x=later(2); class Child extends Base{} function later($x){return $x+1;} class Base{public $n=4;function get(){return $this->n;}} return $x;");
}
#[test]fn namespace_scope_references_and_loads_preserve_complete_ast(){
 gate("<?php namespace Gate; use Vendor\\Thing as Alias; $a=1; $b=&$a; function f(&$x){include 'x.php';return $x;} $b=3; return $a;");
 gate("before<?php echo 1;?>after<?php echo 2;");
}
#[test]fn compiler_sink_rejection_aborts_source_unit(){
 let source="<?php $a=1;return $a;";let manifest=parse(source).unwrap();let mut calls=0;
 let e=produce_grameo_statements_with_receipt(source,manifest,&mut |_,_|{calls+=1;Err("reject unpublished statement".into())}).unwrap_err();assert!(e.contains("reject unpublished statement"));assert_eq!(calls,1);
}
#[test]fn syntax_failure_has_no_successful_unit_receipt(){
 let manifest=parse("<?php return 1;").unwrap();let mut emitted=0;
 assert!(produce_grameo_statements_with_receipt("<?php return ; ???",manifest,&mut |_,_|{emitted+=1;Ok(())}).is_err());assert_eq!(emitted,0);
}
#[test]fn source_pratt_recipes_are_generated_and_consumed_across_boundary_generations(){
 let receipt=gate("<?php $a=1+2*3;$b=(4+5)*6;return $a+$b;");
 assert_eq!(receipt.completed_captures,3);
 assert!(receipt.expression_recipes>0,"advancing parser must transfer real AST recipes: {receipt:?}");
 assert!(receipt.consumed_expression_recipes>0,"contextual walker must use the transferred recipes: {receipt:?}");
 assert!(receipt.source_pratt_expression_recipes>0,"actual committed Pratt reductions required: {receipt:?}");
 assert_eq!(receipt.lowering_failures,0);
 println!("advancing_source_pratt={receipt:?}");
}
#[test]fn streaming_pratt_keeps_function_namespace_reference_and_control_context(){
 gate("<?php namespace Stream; function sum(&$a){$a=$a+2*3;return ($a+4)*5;} $n=1;$f=function($x)use(&$n){return $n+$x*2;};if($n<4){$n=sum($n);}return $f($n);");
 gate("<?php class Base{public $n=3;function read(){return $this->n+2*3;}}class Child extends Base{}$a=new Child;return $a->read();");
}
#[test]fn advancing_pratt_preserves_complete_wordpress_statement_ast(){
 for file in ["functions.php","class-wp-query.php","formatting.php","script-loader.php"] {
  let source=std::fs::read_to_string(format!("examples/webroot/wordpress/wp-includes/{file}")).unwrap();
  let manifest=parse(&source).unwrap();let expected=manifest.body.clone();let mut seen=0;
  let receipt=produce_grameo_statements_with_receipt(&source,manifest,&mut |index,statement|{
   assert_eq!(index,seen);assert_eq!(format!("{statement:?}"),format!("{:?}",expected[index]),"complete normalized statement {file} index{index}");seen+=1;Ok(())
  }).unwrap();
  assert_eq!(seen,expected.len());assert!(receipt.parser_committed);assert!(receipt.source_pratt_expression_recipes>0,"real Pratt acceptance required for {file}: {receipt:?}");assert!(receipt.consumed_expression_recipes>0,"actual recipe consumption required for {file}: {receipt:?}");
  println!("wordpress_advancing_pratt file={file} receipt={receipt:?}");
 }
}

fn prepared_gate(source: &str) -> vybe_language_php::StatementStreamReceipt {
 let (prepared,context)=vybe_language_php::prepare_source_units_with_statement_context(source,std::path::Path::new("gate.php"),&mut |_|{}).unwrap();
 let expected=format!("{:?}",prepared.module.body);let mut body=Vec::new();
 let receipt=vybe_language_php::produce_grameo_statements_prepared_with_receipt(source,prepared.module,&context,&mut |index,statement| {assert_eq!(index,body.len());body.push(statement);Ok(())}).unwrap();
 assert_eq!(format!("{body:?}"),expected);assert_eq!(receipt.declaration_header_parses,0);assert!(receipt.reused_declaration_context);assert!(receipt.parser_committed);assert_eq!(receipt.producer_body_parses,0);assert_eq!(receipt.normalized_handoff_statements,body.len());assert_eq!(receipt.emitted_before_parser_eof,0);receipt
}
#[test] fn prepared_declaration_context_preserves_scopes_and_mixed_sources() {
 for source in ["<?php namespace Prepared; $n=later(1); function later(&$x){return $x+2;} class C extends B{} class B{} return $n;",
 "before<?php echo 1;?>after<?php echo 2;", "<?php $s='<?php';return $s;", "$x=1;return $x+2;"] {prepared_gate(source);}
}
#[test] fn prepared_context_rejects_other_source_and_type_before_sink() {
 let source="<?php return 1;";
 let (prepared,context)=vybe_language_php::prepare_source_units_with_statement_context(source,std::path::Path::new("gate.php"),&mut |_|{}).unwrap();
 let mut calls=0;
 let error=vybe_language_php::produce_grameo_statements_prepared_with_receipt("<?php return 2;",prepared.module,&context,&mut |_,_|{calls+=1;Ok(())}).unwrap_err();
 assert!(error.contains("source mismatch"));assert_eq!(calls,0);
 let (prepared,_)=vybe_language_php::prepare_source_units_with_statement_context(source,std::path::Path::new("gate.php"),&mut |_|{}).unwrap();
 let wrong: std::sync::Arc<dyn std::any::Any+Send+Sync>=std::sync::Arc::new(true);
 assert!(vybe_language_php::produce_grameo_statements_prepared_with_receipt(source,prepared.module,&wrong,&mut |_,_|{calls+=1;Ok(())}).unwrap_err().contains("wrong type"));assert_eq!(calls,0);
}
#[test] fn wordpress_prepared_context_skips_header_parse_and_preserves_ast() {
 let root=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/webroot/wordpress/wp-includes");
 for file in ["functions.php","class-wp-query.php","formatting.php","script-loader.php"] {
  let source=std::fs::read_to_string(root.join(file)).unwrap();let receipt=prepared_gate(&source);
  // Recipes were already consumed by ORIGINAL preparation; this producer
  // must not reparse/reconsume them to rediscover the same normalized body.
  assert_eq!(receipt.producer_body_parses,0);assert_eq!(receipt.consumed_expression_recipes,0);
  assert_eq!(receipt.normalized_handoff_statements,receipt.emitted_statements);
  assert_eq!(receipt.lowering_failures,0);
  println!("prepared_wp={file} {receipt:?}");
 }
}

#[test] fn prepared_owned_handoff_propagates_sink_abort_and_rejects_bad_source_preparation() {
 let source="<?php $a=1; $b=2; return $a+$b;";
 let (prepared,context)=vybe_language_php::prepare_source_units_with_statement_context(source,std::path::Path::new("abort.php"),&mut |_|{}).unwrap();
 let mut calls=0;
 let error=vybe_language_php::produce_grameo_statements_prepared_with_receipt(source,prepared.module,&context,&mut |_,_|{calls+=1;Err("abort unpublished unit".into())}).unwrap_err();
 assert_eq!(calls,1);assert_eq!(error,"abort unpublished unit");
 assert!(vybe_language_php::prepare_source_units_with_statement_context("<?php function incomplete(",std::path::Path::new("bad.php"),&mut |_|{}).is_err());
}

#[cfg(feature="service-verification")]
#[test] fn owned_preparation_keeps_commit_without_capturing_replay_facts() {
 for source in ["<?php namespace Own; $x=later(); function later(){return 3;} class A{} return $x;", "before<?php echo 1;?>after"] {
  let (module, captures, rejected)=vybe_language_php::verify_owned_statement_preparation(source,std::path::Path::new("probe.php")).unwrap();
  assert_eq!(captures,0,"owned preparation must not clone replay declaration maps");
  assert!(rejected,"wrong source, wrong context type and missing normalization seal must reject");
  assert_eq!(format!("{module:?}"),format!("{:?}",parse(source).unwrap()));
 }
 assert!(vybe_language_php::verify_owned_statement_preparation("<?php function broken(",std::path::Path::new("probe.php")).is_err());
}
