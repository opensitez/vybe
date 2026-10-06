//! Exact public reflection emitter and metadata-only candidate.
use vybe_compiler::primitives::{bundle,globals,platforms,reflection};
use vybe_runtime::{Chunk,Op,VM,Value};
#[path="fixtures/reflection_predicate_before.rs"] mod baseline;
#[path="fixtures/reflection_predicate_expected.rs"] mod expected;
fn emit(chunks:&mut[Chunk],index:usize,mode:usize){if mode==0{baseline::emit_is_instance_of(chunks,index,0,"Widget",3)}else{reflection::emit_is_instance_of(chunks,index,0,"Widget",3)}}
fn kernel(mode:usize,entry:bool)->Vec<Chunk>{
 let mut body=Chunk::new("<reflection-type-test>");body.arity=1;body.local_count=1;
 let mut chunks=if entry{let mut root=Chunk::new("<host-entry>");root.emit_op_u16(Op::REF_FUNC,1,3);root.emit(0,3);root.emit_op(Op::RETURN,3);vec![root,body]}else{vec![body]};
 let index=usize::from(entry);emit(&mut chunks,index,mode);chunks[index].emit_op(Op::RETURN,3);
 bundle::finalize_with_runtime_helpers(&mut chunks);globals::declare_free_globals(&mut chunks);globals::normalize_global_table(&mut chunks);chunks
}
#[test]
fn exact_frozen_public_emission_changes_only_two_native_result_markers(){
 let mut c=Chunk::new("probe");c.arity=1;c.local_count=1;
 let mut actual=vec![c.clone()];reflection::emit_is_instance_of(&mut actual,0,0,"Widget",3);
 let mut old=vec![c.clone()];emit(&mut old,0,0);let mut new=vec![c];emit(&mut new,0,1);
 let mut expected=vec![actual[0].clone()];
 expected[0]=Chunk::new("probe");expected[0].arity=1;expected[0].local_count=1;
 expected::emit_is_instance_of(&mut expected,0,0,"Widget",3);
 assert_eq!(actual[0].code,old[0].code);assert_eq!(actual[0].code,expected[0].code);
 assert_eq!(actual[0].block_i32_results,expected[0].block_i32_results);
 assert_eq!(actual[0].local_count,expected[0].local_count);
 assert_eq!(new[0].code,old[0].code);assert_eq!(new[0].local_count,old[0].local_count);
 assert_eq!(format!("{:?}",new[0].imports),format!("{:?}",old[0].imports));
 assert_eq!(new[0].block_i32_results.difference(&old[0].block_i32_results).count(),2);
 assert!(old[0].block_i32_results.is_subset(&new[0].block_i32_results));
}
#[test]
fn complete_candidate_native_module_validates_and_frozen_gap_is_explicit(){
 let old=vybe_platform_wasm::write_wasm(&kernel(0,false));
 let old_error=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&old).err().expect("known frozen value-join raw-IF defect");
 eprintln!("frozen reflection native gap: {old_error}");assert!(old_error.to_string().contains("expected i32"),"{old_error}");
 let new=vybe_platform_wasm::write_wasm(&kernel(1,false));
 wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&new).unwrap_or_else(|e|panic!("candidate native: {e}"));
}
fn object(fields:Vec<(&str,Value)>)->Value{let mut o=vybe_runtime::value::Object::new();for(k,v)in fields{o.properties.insert(k.into(),v);}Value::Object(vybe_runtime::heap::alloc(o))}
fn array(items:Vec<Value>)->Value{let mut o=vybe_runtime::value::Object::new();o.kind=vybe_runtime::value::ObjectKind::Array(items);Value::Object(vybe_runtime::heap::alloc(o))}
#[test]
fn actual_vm_metadata_does_not_coerce_boolean_or_checked_i32_arms(){
 for mode in [0,1]{let mut vm=VM::new();platforms::register_platforms_all(&mut vm);let f=vm.run(kernel(mode,true)).unwrap();
  for fields in [vec![],vec![("__types",Value::Null)],vec![("__types",Value::Undefined)],vec![("__type",Value::String("Other".into()))],vec![("__types",array(vec![Value::String("Other".into())]))]]{assert_eq!(vm.invoke(&f,&[object(fields)]).unwrap(),Value::Bool(false),"mode={mode} negative");}
  for fields in [vec![("__type",Value::String("Widget".into()))],vec![("__types",array(vec![Value::String("Widget".into())]))]]{assert_eq!(vm.invoke(&f,&[object(fields)]).unwrap(),Value::Bool(true),"mode={mode} positive");}
 }
}
