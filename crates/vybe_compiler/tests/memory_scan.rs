#[path="fixtures/memory_scan_scalar_reference.rs"] mod algorithm;
mod inline_vector { pub use vybe_compiler::primitives::memory_scan::emit_zero_terminated_length as emit; }
use vybe_runtime::{Chunk,Op,VM,Value};
fn inline_caller(bounded:bool)->Chunk {
 let mut c=Chunk::new("scalar_scan");c.arity=if bounded {2} else {1};c.param_count=c.arity;c.local_count=c.arity as u16;c.memory_min_pages=vec![1];c.func_sig=Some((vec!["externref";c.arity as usize].join(","),"externref".into()));c
}
fn replace_scan(chunks:&mut Vec<Chunk>,index:usize,bounded:bool) {
 chunks[index]=inline_caller(bounded);let len=inline_vector::emit(chunks,index,0,bounded.then_some(1),1);chunks[index].emit_op_u16(Op::LOCAL_GET,len,1);chunks[index].emit_op(Op::RETURN,1);
 assert_eq!(chunks[index].local_count,u16::from(chunks[index].arity)+1);
 let mut ip=0;let mut vector_loads=0;while ip<chunks[index].code.len(){let op=Op::from_code_at(&chunks[index].code,ip);vector_loads+=usize::from(op==Op::V128_LOAD);ip+=4;ip+=op.operand_format().size_in(&chunks[index].code,ip);}assert_eq!(vector_loads,if bounded {1} else {2},"positive actual vector selection");
}
fn native_module(simd:bool,bounded:bool)->Vec<Chunk> {
 let mut chunks=algorithm::native_module(false,bounded);if simd {replace_scan(&mut chunks,0,bounded);}chunks
}
fn vm_module(simd:bool,bounded:bool)->Vec<Chunk> {
 let mut chunks=algorithm::vm_module(false,bounded);if simd {replace_scan(&mut chunks,1,bounded);}chunks
}
fn wide_vm_module(bounded:bool)->Vec<Chunk> {
 let mut chunks=vm_module(true,bounded);while chunks.len()<=u16::MAX as usize {let mut filler=Chunk::new(format!("inline_wide_filler_{}",chunks.len()));filler.emit_i32_const(0,1);filler.emit_op(Op::RETURN,1);chunks.push(filler);}chunks
}
fn deep_vm_module(simd:bool)->Vec<Chunk> {
 let mut chunks=algorithm::deep_vm_module(false);if simd {replace_scan(&mut chunks,1,false);}chunks
}
fn reload_vm_module(simd:bool,bounded:bool,preceding:usize)->Vec<Chunk> {
 let mut chunks=algorithm::reload_vm_module(false,bounded,preceding);if simd {replace_scan(&mut chunks,1,bounded);}chunks
}
#[test]
fn complete_scans_validate_as_standard_wasm(){
 for simd in [false,true] {for bounded in [false,true] {
  let chunks=native_module(simd,bounded);let bytes=vybe_platform_wasm::write_wasm(&chunks);
  if let Err(error)=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes) {
   let offset=error.offset();eprintln!("native simd={simd} bounded={bounded} error={error}");
   for payload in wasmparser::Parser::new(0).parse_all(&bytes) {if let wasmparser::Payload::CodeSectionEntry(body)=payload.unwrap() {
    let mut reader=body.get_operators_reader().unwrap();while !reader.eof(){let (op,at)=reader.read_with_offset().unwrap();if at+50>=offset&&at<offset+30 {eprintln!("{at}: {op:?}");}}
   }}
  }
  wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
  let path=std::env::var_os("VYBE_SCAN_NATIVE_OUTPUT").map(std::path::PathBuf::from);if let Some(path)=path {std::fs::write(path.join(format!("inline-scan-{}-{}.wasm",simd,bounded)),bytes).unwrap();}
 }}
}
#[test]
fn scalar_and_simd_scans_keep_limits_nuls_and_boundary_traps(){
 for bounded in [false,true] {for simd in [false,true] {
  let mut vm=VM::new();let f=vm.run(vm_module(simd,bounded)).unwrap();
  for address in [0usize,1,7,15,65500,65520,65521,65535] {
   let max=(65536-address).min(80);
   for length in 0..max {
    vm.memory.with_buffer_mut(|m|{m[address..address+max].fill(0x91);m[address+length]=0;});
    let args=if bounded {vec![Value::I32(address as i32),Value::I32(max as i32)]} else {vec![Value::I32(address as i32)]};
    assert_eq!(vm.invoke(&f,&args).unwrap().as_f64(),length as f64,"simd={simd} bounded={bounded} address={address} len={length}");
   }
  }
  vm.memory.with_buffer_mut(|m|m.fill(0x91));
  assert!(vm.invoke(&f,&if bounded {vec![Value::I32(65521),Value::I32(16)]} else {vec![Value::I32(65521)]}).is_err());
  if bounded {for (address,limit) in [(65521,0),(65521,15),(65536,0),(0,31)] {assert_eq!(vm.invoke(&f,&[Value::I32(address),Value::I32(limit)]).unwrap().as_f64(),limit as f64);}}
 }}
}

#[test]
fn direct_helpers_do_not_shift_the_next_public_function_table_index(){
 for simd in [false,true] {
  let mut chunks=vm_module(simd,false);let index=chunks.len();
  let mut after=vybe_runtime::Chunk::new("public_after_scan");after.emit_i32_const(42,1);after.emit_op(vybe_runtime::Op::RETURN,1);chunks.push(after);
  let root=&mut chunks[0];root.code.clear();root.lines.clear();
  root.emit_op_u16(vybe_runtime::Op::REF_FUNC,1,1);root.emit(0,1);root.emit_i32_const(0,1);
  root.emit_op_u8_u8(vybe_runtime::Op::CALL_REF,1,1,1);root.emit_op(vybe_runtime::Op::DROP,1);
  root.emit_op_u16(vybe_runtime::Op::REF_FUNC,index as u16,1);root.emit(0,1);root.emit_op(vybe_runtime::Op::RETURN,1);
  let value=VM::new().run(chunks).unwrap();let Value::Object(object)=value else {panic!("actual public function")};
  assert_eq!(object.lock().unwrap().properties.get("__table_idx").unwrap().as_f64(),1.0,"simd={simd}");
 }
}

#[test]
fn inline_scan_accepts_complete_wide_definition_tables(){
 for bounded in [false,true] {
  let chunks=wide_vm_module(bounded);
  let helpers:Vec<_>=chunks.iter().enumerate().filter(|(_,c)|c.name.starts_with("primitive.cstring_scan")||c.name=="primitive.zero_mask_16").collect();
  assert!(helpers.is_empty(),"this variant emits the vector island in the caller, with no helper edges");
  let bytes=vybe_platform_wasm::write_wasm(&chunks);
  wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
  let mut vm=VM::new();let f=vm.run(chunks).unwrap();
  vm.memory.with_buffer_mut(|m|{m[..81].fill(0x91);m[80]=0;});
  let args=if bounded {vec![Value::I32(0),Value::I32(81)]} else {vec![Value::I32(0)]};
  assert_eq!(vm.invoke(&f,&args).unwrap().as_f64(),80.0);
 }
}

#[test]
fn scan_helpers_preserve_near_limit_recursive_callers(){
 for depth in [16_378,16_381,16_382] {
  let mut outcomes=Vec::new();
  for simd in [false,true] {
   let chunks=deep_vm_module(simd);
   let bytes=vybe_platform_wasm::write_wasm(&chunks);
   wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
   let mut vm=VM::new();let f=vm.run(chunks).unwrap();vm.memory.with_buffer_mut(|m|{m[..65].fill(0x91);m[64]=0;});
   outcomes.push(vm.invoke(&f,&[Value::I32(depth)]).map(|v|v.as_f64()).map_err(|e|e.to_string()));
  }
  let summary:Vec<_>=outcomes.iter().map(|r|r.as_ref().map(|n|*n).map_err(|e|e.lines().next().unwrap_or("").to_owned())).collect();
  eprintln!("actual frame boundary depth={depth} scalar={:?} candidate={:?}",summary[0],summary[1]);
  assert_eq!(summary[0],Ok(64.0),"original must genuinely fit");
  assert_eq!(summary[1],summary[0],"inline vector scan must preserve accepted caller capacity");
 }
}

#[test]
fn retained_scan_callers_survive_supported_same_structure_reload(){
 for simd in [false,true] {for bounded in [false,true] {
  let mut vm=VM::new();let f=vm.run(reload_vm_module(simd,bounded,0)).unwrap();
  vm.memory.with_buffer_mut(|m|{m[..81].fill(0x91);m[80]=0;});
  let args=if bounded {vec![Value::I32(0),Value::I32(81)]} else {vec![Value::I32(0)]};
  assert_eq!(vm.invoke(&f,&args).unwrap().as_f64(),80.0);
  let replacement=reload_vm_module(simd,bounded,0);
  wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&replacement)).unwrap();
  vm.set_reload_hook(Box::new(move |_|Ok(replacement.clone())));vm.debug_reload().unwrap();
  assert_eq!(vm.invoke(&f,&args).unwrap().as_f64(),80.0);
  vm.memory.with_buffer_mut(|m|m[4]=0);assert_eq!(vm.invoke(&f,&args).unwrap().as_f64(),4.0);
  if bounded {assert_eq!(vm.invoke(&f,&[Value::I32(0),Value::I32(3)]).unwrap().as_f64(),3.0);}
 }}
}


#[test]
fn inline_scan_preserves_u16_local_boundary_and_high_native_frames(){
 for frame in [49_900u16,u16::MAX] {for simd in [false,true] {
  let mut chunks=vm_module(simd,false);chunks[1]=inline_caller(false);chunks[1].local_count=frame-1;
  let len=if simd {inline_vector::emit(&mut chunks,1,0,None,1)} else {
   let mut scalar=algorithm::scanner(false,false,0,&algorithm::probe());
   // Generate the original scalar loop at the same maximum caller frame.
   scalar.local_count=frame;let old_slot=1u16;let new_slot=frame-1;
   let mut ip=0;while ip<scalar.code.len(){let op=Op::from_code_at(&scalar.code,ip);if matches!(op,Op::LOCAL_GET|Op::LOCAL_SET)&&u16::from_be_bytes([scalar.code[ip+4],scalar.code[ip+5]])==old_slot {scalar.code[ip+4..ip+6].copy_from_slice(&new_slot.to_be_bytes());}ip+=4;ip+=op.operand_format().size_in(&scalar.code,ip);}
   chunks[1]=scalar;new_slot
  };
  if simd {chunks[1].emit_op_u16(Op::LOCAL_GET,len,1);chunks[1].emit_op(Op::RETURN,1);}
  assert_eq!(chunks[1].local_count,frame);
  let bytes=vybe_platform_wasm::write_wasm(&chunks);
  let native=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes);
  if frame==49_900 {native.unwrap();} else {assert!(native.err().expect("original native resource cap").to_string().contains("locals exceed maximum"),"retain the same wasmparser50000 resource limit on both graphs");}
  let mut vm=VM::new();let f=vm.run(chunks).unwrap();vm.memory.with_buffer_mut(|m|{m[..65].fill(0x91);m[64]=0;});assert_eq!(vm.invoke(&f,&[Value::I32(0)]).unwrap().as_f64(),64.0);
 }}
}

#[test]
fn memory64_route_retains_original_scalar_bytes() {
 for bounded in [false,true] {
  let original=algorithm::scanner(false,bounded,0,&algorithm::probe());
  let mut chunks=vec![inline_caller(bounded)];chunks[0].memory_is_64=vec![true];
  let len=inline_vector::emit(&mut chunks,0,0,bounded.then_some(1),1);
  chunks[0].emit_op_u16(Op::LOCAL_GET,len,1);chunks[0].emit_op(Op::RETURN,1);
  assert_eq!(chunks[0].code,original.code);
  assert_eq!(chunks[0].local_count,original.local_count);
 }
}
