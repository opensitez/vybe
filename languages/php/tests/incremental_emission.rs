//! Public incremental compiler-service and shared emitter execution gates.
use vybe_runtime::{Chunk, VM};
#[derive(Clone,Copy)]
enum Expected { Number(i32) }
fn assert_expected(result: &Result<vybe_runtime::Value,vybe_runtime::VMError>, expected:Expected) {
 let Expected::Number(expected)=expected;
 let actual=match result.as_ref().unwrap() {
  vybe_runtime::Value::I32(n)=>*n as f64,
  vybe_runtime::Value::I64(n)=>*n as f64,
  vybe_runtime::Value::F64(n)=>*n,
  other=>panic!("expected number, got {other:?}"),
 };
 assert_eq!(actual,expected as f64);
}
#[test]
fn incremental_submissions_keep_prior_code_and_execute_only_new_units() {
    vybe_language_php::register();let mut vm=VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let php=vybe_compiler::languages::find_by_name("php").unwrap();
    let mut service=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    assert_expected(&service.submit_source(
        "<?php function incremental_add($x){return $x+1;}return incremental_add(40);",
        php,"interactive/first.php").map_err(vybe_runtime::VMError::new),Expected::Number(41));
    let installed=service.vm().chunks.iter().map(|chunk|(chunk.name.clone(),chunk.code.clone())).collect::<Vec<_>>();
    let callable=service.vm().global("incremental_add").unwrap().clone();
    for index in 0..3 {
        let compiled=service.compile_submission_source(
            "<?php return incremental_add(41);",php,format!("interactive/next-{index}.php")).unwrap();
        assert!(!compiled.chunks.iter().any(|chunk|chunk.name=="incremental_add"),
            "previous function body must not be emitted in the new unit");
        let value=service.run_submission(compiled).unwrap();
        assert!(matches!(value,vybe_runtime::Value::F64(n) if n==42.0)
            || matches!(value,vybe_runtime::Value::I32(42))
            || matches!(value,vybe_runtime::Value::I64(42)),"{value:?}");
        for (chunk,(name,code)) in service.vm().chunks.iter().zip(&installed) {
            assert_eq!(&chunk.name,name);assert_eq!(&chunk.code,code);
        }
    }
    let before_error=service.vm().chunks.len();
    assert!(service.compile_submission_source("<?php function (",php,"interactive/bad.php").is_err());
    assert_eq!(service.vm().chunks.len(),before_error,"parse failure must not install partial code");
    let value=service.vm().invoke(&callable,&[vybe_runtime::Value::I32(9)]).unwrap();
    assert!(matches!(value,vybe_runtime::Value::F64(n) if n==10.0)
        || matches!(value,vybe_runtime::Value::I32(10))
        || matches!(value,vybe_runtime::Value::I64(10)),"{value:?}");
}

#[test]
fn direct_concat_stringify_calls_preserve_scalar_coercion_and_private_identity() {
    use vybe_language_php::emitter::string_adapter::{emit_concat_stringify,emit_concat_stringify_call};
    let mut chunks=vec![Chunk::new("caller")];
    emit_concat_stringify(&mut chunks,0,17);
    let mut indirect=Chunk::new("indirect");
    emit_concat_stringify_call(&mut indirect,1,17);
    assert_eq!(chunks.len(),2);assert_eq!(chunks[0].local_count,0);
    assert!(chunks[0].code.len()<indirect.code.len());
    println!("concat-callsite direct_bytes={} indirect_bytes={} direct_locals={} indirect_locals={}",
        chunks[0].code.len(),indirect.code.len(),chunks[0].local_count,indirect.local_count);
    for _ in 0..255 {emit_concat_stringify(&mut chunks,0,17);}
    assert_eq!(chunks.len(),2);assert_eq!(chunks[0].local_count,0);
    let mut collision=Chunk::new(chunks[1].name.clone());collision.arity=1;
    let mut collision_chunks=vec![Chunk::new("caller"),collision];
    emit_concat_stringify(&mut collision_chunks,0,17);
    assert_eq!(collision_chunks.len(),3,"source spelling cannot claim private helper identity");
    vybe_language_php::register();let mut vm=VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let php=vybe_compiler::languages::find_by_name("php").unwrap();
    let source="<?php return 'v=' . null . '|' . false . '|' . true . '|' . (42) . '|' . (1.25);";
    let value=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_and_run_source(source,php,"concat-parity.php").unwrap();
    assert!(matches!(value,vybe_runtime::Value::String(ref text) if text.as_ref()=="v=||1|42|1.25"),"{value:?}");
}

#[test]
fn shared_implode_preserves_values_evaluation_order_and_object_lifetimes() {
    vybe_language_php::register();
    let php=vybe_compiler::languages::find_by_name("php").unwrap();
    let cases=[
        ("<?php return implode(['left'=>'a','right'=>'b']);", "ab"),
        ("<?php return implode('|',['a'=>true,'b'=>false,'c'=>null,'d'=>'007','e'=>3]);", "1|||007|3"),
        ("<?php $x='a';$values=[&$x,'b'];$x='c';return implode('-',$values);", "c-b"),
        ("<?php class ImplodeStringValue{function __toString(){return 'obj';}}$o=new ImplodeStringValue;return implode(':',[$o,'x']);", "obj:x"),
        ("<?php $events='';class ImplodeLifetime{function __toString(){global $events;$events.='S';return 'v';}function __destruct(){global $events;$events.='D';}}$result=implode('-',[new ImplodeLifetime,'x']);return $result.'|'.$events;", "v-x|SD"),
        ("<?php $events='';function implodeGlue(){global $events;$events.='G';return ':';}function implodeItems(){global $events;$events.='A';return ['a','b'];}$result=implode(implodeGlue(),implodeItems());return $result.'|'.$events;", "a:b|GA"),
    ];
    for (index,(source,expected)) in cases.iter().enumerate() {
        let mut vm=VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
        vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
        let value=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
            .compile_and_run_source(*source,php,format!("implode-parity-{index}.php")).unwrap();
        assert!(matches!(value,vybe_runtime::Value::String(ref text) if text.as_ref()==*expected),
            "case {index}: expected {expected:?}, got {value:?}");
    }
    for argc in [1,2] {
        let mut shared=vec![Chunk::new("caller")];let mut inline=vec![Chunk::new("caller")];
        for _ in 0..20 {
            assert!(vybe_language_php::emitter::dispatch::dispatch("php.implode",&mut shared,0,argc,7));
            vybe_language_php::emitter::array_adapter::emit_php_implode(&mut inline,0,argc,7);
        }
        let bytes:usize=shared.iter().map(|chunk|chunk.code.len()).sum();
        println!("implode-callsite argc={argc} calls=20 shared_bytes={bytes} inline_bytes={} shared_locals={} inline_locals={}",
            inline[0].code.len(),shared[0].local_count,inline[0].local_count);
        assert!(bytes*5<inline[0].code.len());
    }
}


#[test]
fn incremental_submissions_preserve_variable_storage_and_reference_captures() {
    vybe_language_php::register();let mut vm=VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm,&[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let php=vybe_compiler::languages::find_by_name("php").unwrap();
    let mut service=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    for (index,(source,expected)) in [
        ("<?php $shared=4; return $shared;",4),
        ("<?php $shared+=3; return $shared;",7),
        ("<?php $read_shared=function()use(&$shared){return $shared;};return $read_shared();",7),
        ("<?php $shared=9;return $read_shared();",9),
    ].into_iter().enumerate(){
        assert_expected(&service.submit_source(source,php,format!("interactive/storage-{index}.php"))
            .map_err(vybe_runtime::VMError::new),Expected::Number(expected));
    }
}
