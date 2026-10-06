//! Positive native-metadata-transparent VM getter and writable helper dispatch.
use std::sync::{Arc,Mutex};
use vybe_compiler::primitives::{bundle,globals,platforms};
use vybe_runtime::{Chunk,Op,VM,Value};
#[path="fixtures/reflection_predicate_before.rs"]mod baseline;

fn loaded(mode:usize)->(VM,Value){
 let mut root=Chunk::new("entry");root.emit_op_u16(Op::REF_FUNC,1,3);root.emit(0,3);root.emit_op(Op::RETURN,3);
 let mut body=Chunk::new("reflection-getter");body.arity=1;body.local_count=1;
 let mut chunks=vec![root,body];if mode==0{baseline::emit_is_instance_of(&mut chunks,1,0,"Widget",3)}else{vybe_compiler::primitives::reflection::emit_is_instance_of(&mut chunks,1,0,"Widget",3)}
 chunks[1].emit_op(Op::RETURN,3);bundle::finalize_with_runtime_helpers(&mut chunks);globals::declare_free_globals(&mut chunks);globals::normalize_global_table(&mut chunks);
 let mut vm=VM::new();platforms::register_platforms_all(&mut vm);let f=vm.run(chunks).unwrap();(vm,f)
}
fn helper(vm:&mut VM)->Value{
 let mut root=Chunk::new("legal-helper-entry");root.emit_op_u16(Op::REF_FUNC,1,3);root.emit(0,3);root.emit_op(Op::RETURN,3);
 let mut body=Chunk::new("legal-writable-dyneq");body.arity=2;body.local_count=2;
 body.emit_op_u16(Op::LOCAL_GET,0,3);body.emit_op_u16(Op::LOCAL_GET,1,3);let idx=body.add_import("test:reflection","equal");body.emit_call(idx,2,3);body.emit_op(Op::RETURN,3);root.imports=body.imports.clone();
 let target=vm.resolve_host_function_index("test:reflection","equal").unwrap();vm.run_linked(vec![root,body],vec![vybe_runtime::vm::ImportTarget::Host(target)]).unwrap()
}
#[test]
fn real_metadata_getters_then_writable_checked_helper_keep_positive_order(){
 for mode in [0,1]{for answer in [0,-7,19]{
  let(mut vm,f)=loaded(mode);let trace=Arc::new(Mutex::new(Vec::new()));
  let t=trace.clone();vm.register_host_fn("test:reflection","types",Box::new(move|_,_|{t.lock().unwrap().push("types");Value::Undefined}));
  let t=trace.clone();vm.register_host_fn("test:reflection","type",Box::new(move|_,_|{t.lock().unwrap().push("type");Value::String("Widget".into())}));
  let t=trace.clone();vm.register_host_fn("test:reflection","equal",Box::new(move|_,args|{assert_eq!(args.len(),2);assert_eq!(args[0],Value::String("Widget".into()));assert_eq!(args[1],Value::String("Widget".into()));t.lock().unwrap().push("equal");Value::I32(answer)}));
  let h=helper(&mut vm);vm.set_global("__vybe_dyneq",h);
  // Genuine typed mismatching RTT suppresses VM name fallback, ensuring
  // metadata branch/getters and writable helper genuinely execute.
  let other=vm.type_registry.register(vybe_runtime::TypeDef::new("Different"));
  let mut o=vybe_runtime::value::Object::new_typed(other);
  o.properties.insert("__get___types".into(),vm.func_table[vm.resolve_host_function_index("test:reflection","types").unwrap()].clone());
  o.properties.insert("__get___type".into(),vm.func_table[vm.resolve_host_function_index("test:reflection","type").unwrap()].clone());
  let object=Value::Object(vybe_runtime::heap::alloc(o));
  assert_eq!(vm.invoke(&f,&[object]).unwrap(),Value::Bool(answer!=0),"mode={mode} answer={answer}");
  assert_eq!(*trace.lock().unwrap(),vec!["types","type","equal"],"positive callbacks mode={mode}");
 }}
}
type ErrorKey=(String,Option<u32>,Vec<(String,usize,Option<u32>)>);
fn failure_outcome(behavior:usize,mode:usize)->(Option<String>,Option<ErrorKey>,bool,Vec<&'static str>){
   let(mut vm,f)=loaded(mode);let trace=Arc::new(Mutex::new(Vec::new()));
   let t=trace.clone();vm.register_host_fn("test:reflection","types",Box::new(move|ctx,_|{t.lock().unwrap().push("types");if behavior==0{ctx.throw_value(Value::String("metadata-error".into()));}Value::Undefined}));
   let t=trace.clone();vm.register_host_fn("test:reflection","type",Box::new(move|_,_|{t.lock().unwrap().push("type");Value::String("Widget".into())}));
   let t=trace.clone();vm.register_host_fn("test:reflection","equal",Box::new(move|ctx,args|{assert_eq!(args.len(),2);t.lock().unwrap().push("equal");if behavior==1{ctx.throw_value(Value::String("helper-error".into()));Value::I32(0)}else{Value::String("not-a-number".into())}}));
   let h=helper(&mut vm);vm.set_global("__vybe_dyneq",h);
   let other=vm.type_registry.register(vybe_runtime::TypeDef::new("Different"));let mut o=vybe_runtime::value::Object::new_typed(other);
   for(name,host)in[("__get___types","types"),("__get___type","type")]{let idx=vm.resolve_host_function_index("test:reflection",host).unwrap();o.properties.insert(name.into(),vm.func_table[idx].clone());}
   let result=vm.invoke(&f,&[Value::Object(vybe_runtime::heap::alloc(o))]);
   let observed=trace.lock().unwrap().clone();
   let value=result.as_ref().ok().map(|v|format!("{v}"));
   let error=result.err().map(|e|(e.message,e.line,e.call_stack.into_iter().map(|f|(f.chunk_name,f.offset,f.line)).collect::<Vec<_>>()));
   (value,error,vm.last_exception.is_some(),observed)
}
fn check_failure_parity(behavior:usize){
 let before=failure_outcome(behavior,0);let candidate=failure_outcome(behavior,1);
 eprintln!("reflection failure behavior={behavior} before={before:?} candidate={candidate:?}");
 assert_eq!(before,candidate,"exact error/value/trace/last_exception behavior={behavior}");
 assert_eq!(before.3,vec!["types","type","equal"],"positive actual callbacks behavior={behavior}");
 assert!(before.1.is_some()||before.2,"strict actual failure behavior={behavior}");
}
#[test]fn recorded_getter_throw_continuation_gap_preserves_exact_outcome(){check_failure_parity(0);}
#[test]fn checked_helper_throw_preserves_exact_outcome(){check_failure_parity(1);}
#[test]fn checked_helper_illegal_string_result_preserves_exact_outcome(){check_failure_parity(2);}
#[test]
#[ignore="known frozen baseline gap: getter ctx.throw continues into type/equal; normative strict gate remains held"]
fn normative_getter_throw_must_stop_before_following_getter_and_helper(){
 for mode in [0,1]{let outcome=failure_outcome(0,mode);assert_eq!(outcome.3,vec!["types"],"getter throw must stop mode={mode}");assert!(outcome.1.is_some()||outcome.2);}
}
#[test]
fn genuine_matching_rtt_skips_metadata_getter_and_helper(){
 for mode in [0,1]{let(mut vm,f)=loaded(mode);let trace=Arc::new(Mutex::new(Vec::new()));
  let t=trace.clone();vm.register_host_fn("test:reflection","forbidden",Box::new(move|ctx,_|{t.lock().unwrap().push("unexpected");ctx.throw_value(Value::String("unexpected-fallback".into()));Value::Null}));
  let id=vm.type_registry.get_id("Widget").expect("actual module type registered");let mut o=vybe_runtime::value::Object::new_typed(id);
  let idx=vm.resolve_host_function_index("test:reflection","forbidden").unwrap();o.properties.insert("__get___types".into(),vm.func_table[idx].clone());
  vm.set_global("__vybe_dyneq",Value::Null);
  assert_eq!(vm.invoke(&f,&[Value::Object(vybe_runtime::heap::alloc(o))]).unwrap(),Value::Bool(true));assert!(trace.lock().unwrap().is_empty(),"matching RTT must skip fallback mode={mode}");assert!(vm.last_exception.is_none());
 }
}
