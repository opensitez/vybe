// Isolated factory gates. Actual compiler APIs/protocols, not a replacement dispatcher.
mod primitives { pub use vybe_compiler::primitives::{ops,math}; }
#[path="fixtures/numeric_optimization_step_literal_rhs_before.rs"] mod before;
#[path="fixtures/numeric_optimization_step_literal_rhs_candidate.rs"] mod candidate;
use vybe_runtime::{Chunk,Op,Value,VM};
fn factory(old:bool,add:bool)->Chunk {
 let mut c=Chunk::new("literal_rhs_step");c.arity=1;c.param_count=1;c.local_count=1;
 c.emit_op_u16(Op::LOCAL_GET,0,31);if old{before::emit_step(&mut c,add,31)}else{candidate::emit_step(&mut c,add,31)}c.emit_op(Op::RETURN,31);c
}
fn program(old:bool,add:bool)->Vec<Chunk>{let mut root=Chunk::new("<literal-rhs-owner>");vybe_compiler::primitives::functions::emit_ref_func(&mut root,1,0,31);root.emit_op(Op::RETURN,31);vec![root,factory(old,add)]}
fn exact(a:Value,b:Value){match(&a,&b){(Value::F64(a),Value::F64(b))=>assert_eq!(a.to_bits(),b.to_bits()),(Value::F32(a),Value::F32(b))=>assert_eq!(a.to_bits(),b.to_bits()),_=>assert_eq!(a,b)}}
#[test] fn current_frozen_factory_guard_and_emission_budget(){let mut live=Chunk::new("literal_rhs_step");live.arity=1;live.param_count=1;live.local_count=1;live.emit_op_u16(Op::LOCAL_GET,0,31);vybe_compiler::primitives::bigint::emit_step(&mut live,true,31);live.emit_op(Op::RETURN,31);let selected=factory(false,true);assert_eq!(live.code,selected.code);assert_eq!(live.imports,selected.imports);assert_eq!(live.local_count,selected.local_count);let old=factory(true,true);let new=factory(false,true);assert!(new.code.len()<old.code.len());assert!(new.local_count<=old.local_count);eprintln!("step emitter bytes {} -> {}, locals {} -> {}",old.code.len(),new.code.len(),old.local_count,new.local_count);
 let old_dec=factory(true,false);let new_dec=factory(false,false);assert_eq!(old_dec.code,new_dec.code);assert_eq!(old_dec.imports,new_dec.imports);assert_eq!(old_dec.local_count,new_dec.local_count);
}
#[test] fn exact_numeric_tags_specials_and_primitive_passthrough(){
 let inputs=[Value::I32(0),Value::I32(i32::MAX),Value::I64(i64::MAX),Value::I64(9007199254740993),Value::F32(-0.0),Value::F64(-0.0),Value::F64(f64::NAN),Value::F64(f64::INFINITY),Value::F64(f64::NEG_INFINITY),Value::Bool(true),Value::Bool(false),Value::Null,Value::Undefined,Value::String("7".into()),Value::String("a".into())];
 for input in inputs {let mut outputs=Vec::new();for old in [true,false]{let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);let callee=vm.run(program(old,true)).unwrap();outputs.push(vm.invoke_callback(&callee,&[input.clone()]));}exact(outputs.remove(0),outputs.remove(0));}
}
#[test] fn complete_standard_native_modules_remain_strict(){for old in [true,false]{wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&program(old,true))).unwrap();}}

fn root_call(vm:&mut VM, callee:Value,input:Value)->Result<Value,vybe_runtime::VMError> {
 let mut entry=Chunk::new("<step-input-owner>");
 entry.add_global_init("__step_input",vybe_runtime::chunk::ConstExpr::Value(input));
 entry.add_global_init("__step_callee",vybe_runtime::chunk::ConstExpr::Value(callee));
 vybe_compiler::primitives::globals::emit_read(&mut entry,"__step_callee",31);
 vybe_compiler::primitives::globals::emit_read(&mut entry,"__step_input",31);
 vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut entry,1,31);entry.emit_op(Op::RETURN,31);
 let mut chunks=vec![entry];vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);vm.run(chunks)
}
#[test] fn guest_left_methods_errors_aliases_and_retained_input_lifetime(){
 use vybe_runtime::value::Object;
 for result in [Value::F64(7.0),Value::String("z".into()),Value::Null,Value::Bool(true)] {
  let mut outputs=vec![];
  for old in [true,false] {
   let mut chunks=program(old,true);let mut method=Chunk::new("left_valueof_guest");method.arity=1;method.param_count=1;method.local_count=1;
   chunks[0].add_global_init("__method_result",vybe_runtime::chunk::ConstExpr::Value(result.clone()));
   chunks[0].add_global_init("__method_calls",vybe_runtime::chunk::ConstExpr::Value(Value::I32(0)));
   chunks[0].add_global_init("__left_method",vybe_runtime::chunk::ConstExpr::RefFunc(2));
   vybe_compiler::primitives::globals::emit_read(&mut method,"__method_calls",31);method.emit_i32_const(1,31);method.emit_op(Op::I32_ADD,31);vybe_compiler::primitives::globals::emit_write(&mut method,"__method_calls",31);
   vybe_compiler::primitives::globals::emit_read(&mut method,"__method_result",31);method.emit_op(Op::RETURN,31);chunks.push(method);
   vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);
   let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);let callee=vm.run(chunks).unwrap();
   let mut object=Object::new();object.properties.insert(vybe_compiler::ast::protocol_slot_key(vybe_compiler::ast::ProtocolSlot::ValueOf),vm.global("__left_method").unwrap().clone());
   let input=vybe_runtime::heap::alloc(object);let weak=std::sync::Arc::downgrade(&input);
   let answer=root_call(&mut vm,callee,Value::Object(input)).unwrap();
   assert!(weak.upgrade().is_some(),"original input remains retained by caller-owned storage");assert_eq!(vm.global("__method_calls"),Some(&Value::I32(1)));drop(vm);assert!(weak.upgrade().is_none(),"no new retained root after the owning VM/module is released");outputs.push(answer);
  }
  exact(outputs.remove(0),outputs.remove(0));
 }
 // Real malformed user method: compare exact message/source frames, not Null returned by invoke_callback.
 let mut errors=vec![];
 for old in [true,false] {let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);let callee=vm.run(program(old,true)).unwrap();let mut object=Object::new();object.properties.insert(vybe_compiler::ast::protocol_slot_key(vybe_compiler::ast::ProtocolSlot::ValueOf),Value::I32(3));let input=Value::Object(vybe_runtime::heap::alloc(object));let error=root_call(&mut vm,callee,input).err().expect("malformed method must remain an error");errors.push((error.message,error.line,error.call_stack.into_iter().map(|f|(f.chunk_name,f.line)).collect::<Vec<_>>()));}
 assert_eq!(errors[0],errors[1],"no new frame or error waiver; byte offsets alone may shift");
}
