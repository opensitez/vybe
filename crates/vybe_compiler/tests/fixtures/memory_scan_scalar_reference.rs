use vybe_runtime::{Chunk,Op};
fn get(c:&mut Chunk,s:u16){c.emit_op_u16(Op::LOCAL_GET,s,1);}
fn set(c:&mut Chunk,s:u16){c.emit_op_u16(Op::LOCAL_SET,s,1);}
fn iconst(c:&mut Chunk,n:i32){c.emit_i32_const(n,1);}
pub fn probe()->Chunk {
 let mut c=Chunk::new("primitive.zero_mask_16");c.arity=1;c.param_count=1;c.local_count=1;
 c.func_sig=Some(("i32".into(),"i32".into()));c.memory_min_pages=vec![1];
 get(&mut c,0);c.emit_op(Op::V128_LOAD,1);c.emit_leb_u32(0x80,1);c.emit_leb_u32(0,1);
 c.emit_op(Op::V128_CONST,1);for _ in 0..16 {c.emit(0,1);}
 c.emit_op(Op::I8X16_EQ,1);c.emit_op(Op::I8X16_BITMASK,1);c.emit_op(Op::RETURN,1);c
}
pub fn scanner(simd:bool,bounded:bool,helper_index:usize,helper:&Chunk)->Chunk {
 let mut c=Chunk::new(if simd {"simd_scan"} else {"scalar_scan"});
 c.arity=if bounded {2} else {1};c.param_count=c.arity;c.local_count=c.arity as u16;
 c.func_sig=Some((vec!["externref";c.arity as usize].join(","),"externref".into()));c.memory_min_pages=vec![1];
 let len=c.alloc_scratch(1);iconst(&mut c,0);set(&mut c,len);
 let extras=if simd {Some(c.alloc_scratch(3))} else {None};
 let done=c.emit_block(1);let (repeat,_)=c.emit_loop_s(1);
 if bounded {get(&mut c,len);get(&mut c,1);c.emit_op(Op::I32_GE_U,1);c.emit_br_if(1,1);}
 if let Some(addr)=extras {
  let mask=addr+1;let eligible=addr+2;
  get(&mut c,0);get(&mut c,len);c.emit_op(Op::I32_ADD,1);set(&mut c,addr);
  // A vector never crosses a Wasm page. Page numbers are unsigned and
  // memory.size can represent all 65,536 pages of a 4GiB wasm32 memory.
  get(&mut c,addr);iconst(&mut c,16);c.emit_op(Op::I32_SHR_U,1);
  c.emit_op_u8(Op::MEMORY_SIZE,0,1);c.emit_op(Op::I32_LT_U,1);
  get(&mut c,addr);iconst(&mut c,65535);c.emit_op(Op::I32_AND,1);iconst(&mut c,65520);c.emit_op(Op::I32_LE_U,1);c.emit_op(Op::I32_AND,1);
  if bounded {get(&mut c,1);get(&mut c,len);c.emit_op(Op::I32_SUB,1);iconst(&mut c,16);c.emit_op(Op::I32_GE_U,1);c.emit_op(Op::I32_AND,1);}
  set(&mut c,eligible);get(&mut c,eligible);c.emit_if(1);
  get(&mut c,addr);c.emit_direct_defined_call(helper_index as u32,helper,false,1).unwrap();c.emit_op_u16(Op::LOCAL_TEE,mask,1);c.emit_if(1);
  get(&mut c,len);get(&mut c,mask);c.emit_op(Op::I32_CTZ,1);c.emit_op(Op::I32_ADD,1);set(&mut c,len);c.emit_br(3,1);c.emit_end(1);
  get(&mut c,len);iconst(&mut c,16);c.emit_op(Op::I32_ADD,1);set(&mut c,len);c.emit_br(1,1);c.emit_end(1);
 }
 get(&mut c,0);get(&mut c,len);c.emit_op(Op::I32_ADD,1);c.emit_op(Op::I32_LOAD8_U,1);c.emit_op(Op::I32_EQZ,1);c.emit_br_if(1,1);
 get(&mut c,len);iconst(&mut c,1);c.emit_op(Op::I32_ADD,1);set(&mut c,len);c.emit_br(0,1);
 c.emit_end(1);c.patch_loop(repeat);c.emit_end(1);c.patch_block(done);get(&mut c,len);c.emit_op(Op::RETURN,1);c
}
fn caller(bounded:bool)->Chunk {
 let mut c=Chunk::new("scan_caller");c.arity=if bounded {2} else {1};c.param_count=c.arity;c.local_count=c.arity as u16;
 c.func_sig=Some((vec!["externref";c.arity as usize].join(","),"externref".into()));c.memory_min_pages=vec![1];c
}
fn emit_call(chunks:&mut Vec<Chunk>,index:usize,bounded:bool) {
 let len=super::inline_vector::emit(chunks,index,0,bounded.then_some(1),1);
 chunks[index].emit_op_u16(Op::LOCAL_GET,len,1);chunks[index].emit_op(Op::RETURN,1);
}
pub fn native_module(simd:bool,bounded:bool)->Vec<Chunk>{
 if !simd {return vec![scanner(false,bounded,0,&probe())];}
 let mut chunks=vec![caller(bounded)];emit_call(&mut chunks,0,bounded);chunks
}
pub fn vm_module(simd:bool,bounded:bool)->Vec<Chunk>{
 let mut root=Chunk::new("owner");root.memory_min_pages=vec![1];root.emit_op_u16(Op::REF_FUNC,1,1);root.emit(0,1);root.emit_op(Op::RETURN,1);
 if !simd {return vec![root,scanner(false,bounded,0,&probe())];}
 let mut chunks=vec![root,caller(bounded)];emit_call(&mut chunks,1,bounded);chunks
}

/// Preserve the real caller at index one while linking helpers beyond u16.
pub fn wide_vm_module(bounded:bool)->Vec<Chunk> {
 let mut root=Chunk::new("wide_owner");root.memory_min_pages=vec![1];root.emit_op_u16(Op::REF_FUNC,1,1);root.emit(0,1);root.emit_op(Op::RETURN,1);
 let mut chunks=vec![root,caller(bounded)];
 while chunks.len()<=u16::MAX as usize {let mut filler=Chunk::new(format!("wide_filler_{}",chunks.len()));filler.emit_i32_const(0,1);filler.emit_op(Op::RETURN,1);chunks.push(filler);}
 emit_call(&mut chunks,1,bounded);chunks
}

/// Actual non-tail recursion reaches the VM frame boundary before scanning.
pub fn deep_vm_module(simd:bool)->Vec<Chunk> {
 let mut root=Chunk::new("depth_owner");root.memory_min_pages=vec![1];root.emit_op_u16(Op::REF_FUNC,2,1);root.emit(0,1);root.emit_op(Op::RETURN,1);
 let scan=if simd {caller(false)} else {scanner(false,false,0,&probe())};
 let mut recurse=Chunk::new("recursive_scan_caller");recurse.arity=1;recurse.param_count=1;recurse.local_count=1;recurse.func_sig=Some(("externref".into(),"externref".into()));
 let self_target=recurse.clone();
 get(&mut recurse,0);recurse.emit_op(Op::I32_EQZ,1);recurse.emit_if_value(1);
 iconst(&mut recurse,0);recurse.emit_direct_defined_call(1,&scan,false,1).unwrap();
 recurse.emit_else(1);get(&mut recurse,0);iconst(&mut recurse,1);recurse.emit_op(Op::I32_SUB,1);recurse.emit_direct_defined_call(2,&self_target,false,1).unwrap();recurse.emit_end(1);recurse.emit_op(Op::RETURN,1);
 let mut chunks=vec![root,scan,recurse];if simd {emit_call(&mut chunks,1,false);}chunks
}

/// Recompile an unchanged scan while preceding unrelated definitions move helpers.
pub fn reload_vm_module(simd:bool,bounded:bool,preceding:usize)->Vec<Chunk> {
 let mut root=Chunk::new("reload_scan_owner");root.memory_min_pages=vec![1];root.emit_op_u16(Op::REF_FUNC,1,1);root.emit(0,1);root.emit_op(Op::RETURN,1);
 let mut chunks=vec![root,if simd {caller(bounded)} else {scanner(false,bounded,0,&probe())}];
 for index in 0..preceding {let mut unrelated=Chunk::new(format!("unrelated_before_helpers_{index}"));unrelated.emit_i32_const(0,1);unrelated.emit_op(Op::RETURN,1);chunks.push(unrelated);}
 if simd {emit_call(&mut chunks,1,bounded);}chunks
}
