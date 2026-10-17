//! Actual parsed PHP relational/equality routes and complete native validation.
use vybe_runtime::{Chunk,Value,VM};
fn compile(source:&str)->Vec<Chunk>{static INIT:std::sync::Once=std::sync::Once::new();INIT.call_once(vybe_language_php::register);let ast=vybe_language_php::parse(&format!("<?php {source}")).unwrap();let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap()}
fn run(source:&str)->Value{let c=compile(source);wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&c)).unwrap();let mut v=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut v);v.run(c).unwrap()}
#[test]fn actual_php_dynamic_numeric_relational_routes(){for (expression,want) in [("$a<$b",true),("$a>$b",false),("$a<=$b",true),("$a>=$b",false),("$a==$b",false),("$a!=$b",true)]{assert_eq!(run(&format!("function compare($a,$b){{return {expression};}}return compare(17,23);")),if ["$a==$b","$a!=$b"].contains(&expression){Value::Bool(want)}else{Value::I32(i32::from(want))});}}
#[test]fn actual_php_mixed_relational_routes(){for (a,b,want) in [("'17'","23",true),("null","1",true),("false","true",true),("'z'","'a'",false)]{assert_eq!(run(&format!("function compare($a,$b){{return $a<$b;}}return compare({a},{b});")),Value::I32(i32::from(want)));}}
#[test]fn datetime_operands_keep_chronological_order(){
 for (a,b,want) in [("2004-02-29","2004-03-01",1),("2004-03-01","2004-02-29",0),("2004-02-29","2004-02-29",0)]{let source=format!("$a=new DateTimeImmutable('{a}',new DateTimeZone('UTC'));$b=new DateTimeImmutable('{b}',new DateTimeZone('UTC'));return $a<$b;");assert_eq!(run(&source),Value::I32(want));}
}
