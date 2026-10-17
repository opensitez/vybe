use super::*;
use std::sync::atomic::{AtomicUsize,Ordering};
fn script(value:i32)->Chunk{let mut c=Chunk::new("resident-script");c.emit_i32_const(value,27);c.emit_op(Op::RETURN,27);c}
fn load(vm:&mut VM,chunks:Vec<Chunk>,imports:Vec<ImportTarget>,prepared:bool)->(Result<Value,VMError>,Option<LoadedRequestUnit>){
 if prepared{vm.run_linked_nested_prepared_request(crate::PreparedLinkedCode::prepare(chunks,64*1024*1024).ok().unwrap(),imports)}
 else{vm.run_linked_nested_request(chunks,imports)}
}
fn resolve(vm:&VM,unit:&LoadedRequestUnit)->Vec<ImportTarget>{vm.request_unit_imports(unit).unwrap().iter().map(|i|vm.resolve_import_target(&i.module,&i.name).unwrap()).collect()}
#[cfg_attr(test,test)]
fn repeat_effects_and_fresh_locals_no_code_or_decode_growth(){
 for prepared in [false,true]{
  let counter=Arc::new(AtomicUsize::new(0));let count=counter.clone();let mut vm=VM::new();
  vm.register_host_fn("resident-test","tick",Box::new(move|_,_|Value::I32(count.fetch_add(1,Ordering::SeqCst)as i32+1)));
  let mut c=Chunk::new("effect");c.local_count=2;let i=c.add_import("resident-test","tick");c.emit_call(i,0,4);c.emit_op(Op::RETURN,4);
  let target=vm.resolve_import_target("resident-test","tick").unwrap();let(value,unit)=load(&mut vm,vec![c],vec![target],prepared);assert_eq!(value.unwrap(),Value::I32(1));let unit=unit.unwrap();
  let len=vm.chunks.len();let decoded=vm.decoded_ops[unit.0.first].as_ptr();let ops=vm.chunks[unit.0.first].code.as_ptr();
  for expected in 2..5{let imports=resolve(&vm,&unit);assert_eq!(vm.run_resident_nested(&unit,imports).unwrap(),Value::I32(expected));}
  assert_eq!(vm.chunks.len(),len);assert_eq!(vm.decoded_ops[unit.0.first].as_ptr(),decoded);assert_eq!(vm.chunks[unit.0.first].code.as_ptr(),ops);
  let mut c=Chunk::new("fresh-locals");c.local_count=1;c.emit_op_u16(Op::LOCAL_GET,0,8);c.emit_op(Op::RETURN,8);
  let(value,local)=load(&mut vm,vec![c],vec![],prepared);assert_eq!(value.unwrap(),Value::Null);let local=local.unwrap();
  vm.stack.push(Value::I32(99));assert_eq!(vm.run_resident_nested(&local,vec![]).unwrap(),Value::Null);assert_eq!(vm.stack,vec![Value::I32(99)]);vm.stack.clear();
 }
}
#[cfg_attr(test,test)]
fn source_function_and_class_type_context_bind_once(){
 for prepared in [false,true]{
  let mut vm=VM::new();vm.run_linked_nested(vec![script(1)],vec![]).unwrap();let base=vm.chunks.len();
  let mut entry=Chunk::new("source-class");entry.types.push(crate::chunk::TypeEntry{name:"ResidentClass".into(),fields:vec!["x".into()],methods:vec![("get".into(),1)],constructor_chunk:Some(1),..Default::default()});
  entry.emit_op_u16(Op::REF_FUNC,1,10);entry.emit(0,10);entry.emit_op(Op::RETURN,10);let body=script(9);
  assert!(LoadedRequestUnit::accepts(&[entry.clone(),body.clone()]));
  let(value,unit)=load(&mut vm,vec![entry,body],vec![],prepared);let unit=unit.unwrap();let function=value.unwrap();assert_eq!(vm.invoke(&function,&[]).unwrap(),Value::I32(9));
  let count=vm.module_type_ids.len();let type_base=vm.chunk_type_base[base];let id=vm.get_type_id("ResidentClass");
  let again=vm.run_resident_nested(&unit,vec![]).unwrap();assert_eq!(vm.invoke(&again,&[]).unwrap(),Value::I32(9));assert_eq!(vm.module_type_ids.len(),count);assert_eq!(vm.chunk_type_base[base],type_base);assert_eq!(vm.get_type_id("ResidentClass"),id);
  let ty=vm.type_registry.get(id).unwrap();assert!(matches!(ty.methods.get("get"),Some(crate::typedef::Method::ChunkFn(i))if *i==base+1));
 }
}
#[cfg_attr(test,test)]
fn source_native_gc_identity_is_not_a_native_storage_denial(){
 use crate::native_gc::{CompositeType,RecGroup,SubType};
 let group=RecGroup::new(vec![SubType{final_:true,shared:false,supertypes:vec![],descriptor:None,describes:None,composite:CompositeType::Struct(vec![])}],vec![]);
 let mut c=script(7);c.types.push(crate::chunk::TypeEntry{name:"SourceGcClass".into(),native_identity:group.identity(0),..Default::default()});
 assert_eq!(LoadedRequestUnit::admission(&[c]),Ok(()));
}
#[cfg_attr(test,test)]
fn wrong_vm_reset_reload_remove_global_and_explicit_generation_invalidate(){
 let mut vm=VM::new();let baseline=vm.snapshot();let(_,unit)=load(&mut vm,vec![script(7)],vec![],false);let unit=unit.unwrap();
 assert!(!VM::new().request_unit_valid(&unit));let other=unit.clone();assert!(std::thread::spawn(move||!VM::new().request_unit_valid(&other)).join().unwrap());
 vm.reset_to(&baseline);let(_,current)=load(&mut vm,vec![script(8)],vec![],false);let current=current.unwrap();assert!(!vm.request_unit_valid(&unit));assert!(vm.request_unit_valid(&current));
 vm.set_global("slot",Value::I32(1));vm.set_global("slot",Value::I32(2));assert!(vm.request_unit_valid(&current));vm.remove_global("slot");vm.set_global("slot",Value::I32(3));assert!(!vm.request_unit_valid(&current));
 let mut reload=VM::new();let(_,unit)=load(&mut reload,vec![script(7)],vec![],false);let unit=unit.unwrap();reload.apply_reload(vec![script(9)]).unwrap();assert!(!reload.request_unit_valid(&unit));
 let(_,unit)=load(&mut reload,vec![script(7)],vec![],false);let unit=unit.unwrap();reload.invalidate_request_units();assert!(!reload.request_unit_valid(&unit));assert!(reload.run_resident_nested(&unit,vec![]).is_err());
}
#[cfg_attr(test,test)]
fn fresh_host_imports_and_native_projection_denial(){
 let mut vm=VM::new();vm.register_host_fn("resident-test","value",Box::new(|_,_|Value::I32(7)));
 let mut c=Chunk::new("changing-host");let i=c.add_import("resident-test","value");c.emit_call(i,0,1);c.emit_op(Op::RETURN,1);
 let imports=vec![vm.resolve_import_target("resident-test","value").unwrap()];let(value,unit)=load(&mut vm,vec![c],imports,false);assert_eq!(value.unwrap(),Value::I32(7));let unit=unit.unwrap();
 vm.register_host_fn("resident-test","value",Box::new(|_,_|Value::I32(9)));let imports=resolve(&vm,&unit);assert!(vm.request_unit_accepts_imports(&unit,&imports));assert_eq!(vm.run_resident_nested(&unit,imports).unwrap(),Value::I32(9));
 let mut c=Chunk::new("native-import");c.emit_string_const("abc",2);let i=c.add_import("wasm:js-string","length");c.emit_call(i,1,2);c.emit_op(Op::RETURN,2);
 let imports=vec![vm.resolve_import_target("wasm:js-string","length").unwrap()];let(value,unit)=load(&mut vm,vec![c],imports,false);assert_eq!(value.unwrap(),Value::I32(3));let unit=unit.unwrap();
 let changed=vec![vm.resolve_import_target("resident-test","value").unwrap()];let count=vm.chunks.len();assert!(!vm.request_unit_accepts_imports(&unit,&changed));assert!(vm.run_resident_nested(&unit,changed).is_err());assert_eq!(vm.chunks.len(),count);
 assert_eq!(vm.run_resident_nested(&unit,resolve(&vm,&unit)).unwrap(),Value::I32(3));
}
#[cfg_attr(test,test)]
fn stable_global_reads_refresh_but_removed_slot_reentry_is_rejected(){
 let mut vm=VM::new();vm.set_global("source_value",Value::I32(7));let mut c=Chunk::new("source-global");c.globals=Arc::new(vec!["source_value".into()].into());c.emit_op_u32(Op::GLOBAL_GET,0,1);c.emit_op(Op::RETURN,1);
 let(value,unit)=load(&mut vm,vec![c.clone()],vec![],false);assert_eq!(value.unwrap(),Value::I32(7));let unit=unit.unwrap();vm.set_global("source_value",Value::I32(9));assert!(vm.request_unit_valid(&unit));assert_eq!(vm.run_resident_nested(&unit,vec![]).unwrap(),Value::I32(9));
 vm.remove_global("source_value");vm.set_global("source_value",Value::I32(11));let len=vm.chunks.len();assert!(!vm.request_unit_valid(&unit));assert!(vm.run_resident_nested(&unit,vec![]).is_err());assert_eq!(len,vm.chunks.len());assert_eq!(vm.run_linked_nested(vec![c],vec![]).unwrap(),Value::I32(11));
}
#[cfg_attr(test,test)]
fn failed_link_import_binding_restores_caller_without_manufacturing_handle(){
 for prepared in [false,true]{let mut vm=VM::new();vm.register_host_fn("resident-test","unused",Box::new(|_,_|Value::I32(7)));vm.import_table=vec![ImportTarget::Host(987)];vm.rebuild_resolved_import_table();
  let mut c=Chunk::new("unresolved-import");let i=c.add_import("missing-unit","missing-function");c.emit_call(i,0,44);c.emit_op(Op::RETURN,44);
  let(value,unit)=load(&mut vm,vec![c],vec![ImportTarget::Host(0)],prepared);assert!(value.unwrap_err().message.contains("Unresolved import"));assert!(unit.is_none());assert!(matches!(vm.import_table.as_slice(),[ImportTarget::Host(987)]));assert!(matches!(vm.resolved_import_table.as_slice(),[ResolvedCallTarget::Host(987)]));assert!(vm.frames.is_empty());assert!(vm.stack.is_empty());
 }
}
#[cfg_attr(test,test)]
fn replay_binds_exact_unit_globals_and_reruns_unassigned_initializers(){
 let mut vm=VM::new();let mut own=Chunk::new("own-string");own.emit_string_const("owned",1);own.emit_op(Op::RETURN,1);own.global_inits.push(crate::chunk::GlobalInit{name:"init_value".into(),init:crate::chunk::ConstExpr::Value(Value::I32(7))});
 let(value,unit)=load(&mut vm,vec![own],vec![],false);assert_eq!(value.unwrap(),Value::String(Arc::from("owned")));let unit=unit.unwrap();
 let mut later=Chunk::new("later-string");later.emit_string_const("later",1);later.emit_op(Op::RETURN,1);vm.run_linked_nested(vec![later],vec![]).unwrap();
 let owned=crate::chunk::imported_global_key(crate::chunk::STRING_CONSTANTS_MODULE,"owned");let later=crate::chunk::imported_global_key(crate::chunk::STRING_CONSTANTS_MODULE,"later");vm.set_global(&owned,Value::I32(8));vm.set_global(&later,Value::I32(9));
 let slot=*vm.global_index.get("init_value").unwrap()as usize;vm.globals[slot]=Value::I32(0);vm.globals_assigned[slot]=false;
 assert_eq!(vm.run_resident_nested(&unit,vec![]).unwrap(),Value::String(Arc::from("owned")));assert_eq!(vm.global(&later),Some(&Value::I32(9)));assert_eq!(vm.global("init_value"),Some(&Value::I32(7)));
 vm.set_global("init_value",Value::I32(42));vm.run_resident_nested(&unit,vec![]).unwrap();assert_eq!(vm.global("init_value"),Some(&Value::I32(42)));
}
#[cfg_attr(test,test)]
fn initializer_error_restores_context_without_entering_script(){
 for prepared in [false,true]{
  let count=Arc::new(AtomicUsize::new(0));let counter=count.clone();let mut vm=VM::new();
  vm.register_host_fn("resident-test","init-tick",Box::new(move|_,_|Value::I32(counter.fetch_add(1,Ordering::SeqCst)as i32+1)));
  vm.set_global("good_import",Value::I32(77));
  let mut c=Chunk::new("initializer-error");c.add_import("source","initializer");let tick=c.add_import("resident-test","init-tick");
  c.global_inits.push(crate::chunk::GlobalInit{name:"captured_fn".into(),init:crate::chunk::ConstExpr::ImportedRefFunc(0)});
  c.emit_call(tick,0,71);c.emit_op(Op::RETURN,71);
  let imports=vec![ImportTarget::StdlibRedirect("good_import".into()),vm.resolve_import_target("resident-test","init-tick").unwrap()];
  let(value,unit)=load(&mut vm,vec![c.clone()],imports,prepared);assert_eq!(value.unwrap(),Value::I32(1));let unit=unit.unwrap();
  let slot=*vm.global_index.get("captured_fn").unwrap()as usize;vm.globals_assigned[slot]=false;
  vm.import_table=vec![ImportTarget::Host(987)];vm.rebuild_resolved_import_table();vm.stack.push(Value::I32(99));
  let imports=vec![ImportTarget::StdlibRedirect("missing_import".into()),vm.resolve_import_target("resident-test","init-tick").unwrap()];
  assert!(vm.request_unit_accepts_imports(&unit,&imports));let before=(vm.frames.len(),vm.label_stack.len());
  let actual=vm.run_resident_nested(&unit,imports).unwrap_err();assert_eq!(actual.message,"unresolved imported constant function");
  assert_eq!(count.load(Ordering::SeqCst),1);assert_eq!(vm.stack,vec![Value::I32(99)]);assert_eq!((vm.frames.len(),vm.label_stack.len()),before);
  assert!(matches!(vm.import_table.as_slice(),[ImportTarget::Host(987)]));assert!(matches!(vm.resolved_import_table.as_slice(),[ResolvedCallTarget::Host(987)]));
  let mut old=VM::new();old.register_host_fn("resident-test","init-tick",Box::new(|_,_|panic!("initializer failure must precede entry")));
  old.import_table=vec![ImportTarget::Host(987)];old.rebuild_resolved_import_table();old.stack.push(Value::I32(99));
  let imports=vec![ImportTarget::StdlibRedirect("missing_import".into()),old.resolve_import_target("resident-test","init-tick").unwrap()];
  let expected=if prepared{old.run_linked_nested_prepared(crate::PreparedLinkedCode::prepare(vec![c],64*1024*1024).ok().unwrap(),imports)}else{old.run_linked_nested(vec![c],imports)}.unwrap_err();
  assert_eq!(format!("{actual:?}"),format!("{expected:?}"));assert_eq!(old.stack,vec![Value::I32(99)]);assert!(old.frames.is_empty());assert!(matches!(old.import_table.as_slice(),[ImportTarget::Host(987)]));
 }
}
#[cfg_attr(test,test)]
fn nested_child_range_and_caller_import_context_are_exact(){
 let mut vm=Box::new(VM::new());let pointer=(&mut*vm as*mut VM)as usize;let first=Arc::new(AtomicUsize::new(0));let counter=first.clone();
 vm.register_host_fn("resident-test","append-child",Box::new(move|_,_|{if counter.fetch_add(1,Ordering::SeqCst)==0{unsafe{(&mut*(pointer as*mut VM)).run_linked_nested(vec![script(99)],vec![]).unwrap();}}Value::I32(7)}));
 let host=vm.resolve_import_target("resident-test","append-child").unwrap();vm.import_table=vec![ImportTarget::Host(987)];vm.rebuild_resolved_import_table();
 let mut c=Chunk::new("parent-unit");let i=c.add_import("resident-test","append-child");c.emit_call(i,0,3);c.emit_op(Op::RETURN,3);
 let(value,unit)=load(&mut vm,vec![c],vec![host],false);assert_eq!(value.unwrap(),Value::I32(7));let unit=unit.unwrap();assert_eq!(unit.chunk_range(),0..1);assert_eq!(vm.chunks.len(),2);assert!(matches!(vm.import_table.as_slice(),[ImportTarget::Host(987)]));
 let imports=resolve(&vm,&unit);assert_eq!(vm.run_resident_nested(&unit,imports).unwrap(),Value::I32(7));assert_eq!(vm.chunks.len(),2);assert!(matches!(vm.import_table.as_slice(),[ImportTarget::Host(987)]));assert!(vm.frames.is_empty());assert!(vm.stack.is_empty());
}
#[cfg_attr(test,test)]
fn native_storage_and_mutable_payloads_decline_but_raw_execution_remains(){
 let mut vm=VM::new();vm.memory.resize(8,9);vm.wasm_tables=vec![vec![Value::I32(71)]];
 let mut c=script(7);c.native_global_import_count=Some(0);c.memory_min_pages=vec![1];assert_eq!(LoadedRequestUnit::admission(&[c.clone()]),Err(RequestUnitDenial::NativeInstance));
 let(value,unit)=load(&mut vm,vec![c],vec![],false);assert_eq!(value.unwrap(),Value::I32(7));assert!(unit.is_none());assert_eq!(vm.memory.len(),8);assert_eq!(vm.wasm_tables,vec![vec![Value::I32(71)]]);
 let mut c=script(7);c.constants.push(Value::Object(crate::heap::alloc(Object::new())));assert_eq!(LoadedRequestUnit::admission(&[c.clone()]),Err(RequestUnitDenial::MutableConstant));let(value,unit)=load(&mut vm,vec![c],vec![],false);assert_eq!(value.unwrap(),Value::I32(7));assert!(unit.is_none());
 let mut c=script(7);c.global_inits.push(crate::chunk::GlobalInit{name:"object".into(),init:crate::chunk::ConstExpr::Value(Value::Object(crate::heap::alloc(Object::new())))});assert_eq!(LoadedRequestUnit::admission(&[c]),Err(RequestUnitDenial::MutableConstant));
 let mut c=script(7);c.table_is_64=vec![true];assert_eq!(LoadedRequestUnit::admission(&[c]),Err(RequestUnitDenial::Storage));
}
#[cfg_attr(test,test)]
fn malformed_install_has_no_handle_and_runtime_errors_match_original_loader(){
 for prepared in [false,true]{
  let mut c=Chunk::new("bad");c.code=vec![0xFF,0xFF,0xFF,0xFF];c.lines.resize(4,42);let mut vm=VM::new();let(value,unit)=load(&mut vm,vec![c],vec![],prepared);assert!(value.is_err());assert!(unit.is_none());
  let mut c=Chunk::new("trap-line");c.emit_op(Op::UNREACHABLE,42);let mut old=VM::new();let mut selected=VM::new();let expected=old.run_linked_nested(vec![c.clone()],vec![]).unwrap_err();let(value,unit)=load(&mut selected,vec![c.clone()],vec![],prepared);assert_eq!(format!("{expected:?}"),format!("{:?}",value.unwrap_err()));let unit=unit.unwrap();
  let expected=old.run_linked_nested(vec![c],vec![]).unwrap_err();let actual=selected.run_resident_nested(&unit,vec![]).unwrap_err();assert_eq!(format!("{expected:?}"),format!("{actual:?}"));
 }
}
#[cfg_attr(test,test)]
fn nested_pending_exit_propagates_to_outer_loop(){
 let mut vm=Box::new(VM::new());let pointer=(&mut*vm as*mut VM)as usize;
 vm.register_host_fn("resident-test","exit",Box::new(|ctx,_|{ctx.request_exit_with_code(23);Value::Null}));
 let mut child=Chunk::new("exit-child");let i=child.add_import("resident-test","exit");child.emit_call(i,0,7);child.emit_op(Op::RETURN,7);
 let imports=vec![vm.resolve_import_target("resident-test","exit").unwrap()];let(value,unit)=load(&mut vm,vec![child],imports,false);assert_eq!(value.unwrap(),Value::Null);assert!(vm.pending_exit);let unit=unit.unwrap();vm.pending_exit=false;
 vm.register_host_fn("resident-test","enter",Box::new(move|_,_|unsafe{let vm=&mut*(pointer as*mut VM);let imports=resolve(vm,&unit);vm.run_resident_nested(&unit,imports).unwrap()}));
 let count=Arc::new(AtomicUsize::new(0));let counter=count.clone();vm.register_host_fn("resident-test","after",Box::new(move|_,_|{counter.fetch_add(1,Ordering::SeqCst);Value::I32(99)}));
 let mut caller=Chunk::new("exit-caller");let enter=caller.add_import("resident-test","enter");let after=caller.add_import("resident-test","after");caller.emit_call(enter,0,11);caller.emit_op(Op::DROP,11);caller.emit_call(after,0,12);caller.emit_op(Op::RETURN,12);
 let imports=caller.imports.iter().map(|i|vm.resolve_import_target(&i.module,&i.name).unwrap()).collect();assert_eq!(vm.run_linked(vec![caller],imports).unwrap(),Value::Null);assert_eq!(count.load(Ordering::SeqCst),0);assert!(!vm.pending_exit);assert_eq!(vm.pending_exit_code,23);assert!(vm.frames.is_empty());
}
pub(super)fn run_all(){
 for(name,test)in [
  ("repeated effects/fresh locals",repeat_effects_and_fresh_locals_no_code_or_decode_growth as fn()),
  ("source function/class type context",source_function_and_class_type_context_bind_once),
  ("source native GC identity admission",source_native_gc_identity_is_not_a_native_storage_denial),
  ("managed generation/VM identity",wrong_vm_reset_reload_remove_global_and_explicit_generation_invalidate),
  ("fresh imports/native projection denial",fresh_host_imports_and_native_projection_denial),
  ("stable global reads/removed slot rejection",stable_global_reads_refresh_but_removed_slot_reentry_is_rejected),
  ("failed import installation restores caller",failed_link_import_binding_restores_caller_without_manufacturing_handle),
  ("exact global binding/init replay",replay_binds_exact_unit_globals_and_reruns_unassigned_initializers),
  ("initializer error restores context/no entry",initializer_error_restores_context_without_entering_script),
  ("nested child range/import restoration",nested_child_range_and_caller_import_context_are_exact),
  ("native/mutable fallback",native_storage_and_mutable_payloads_decline_but_raw_execution_remains),
  ("malformed/error parity",malformed_install_has_no_handle_and_runtime_errors_match_original_loader),
  ("nested pending exit",nested_pending_exit_propagates_to_outer_loop),
 ]{test();println!("PASS {name}")}
}
