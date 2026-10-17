use vybe_runtime::{VM,Value};
#[test]
fn standard_wide_equality_keeps_all_masks_in_actual_vm(){
 use vybe_runtime::{Chunk,Op};
 for (width,eq) in [(2usize,Op::I16X8_EQ),(4,Op::I32X4_EQ)] {
 let mut probe=Chunk::new("standard_byte_equality_probe");probe.arity=1;probe.param_count=1;probe.local_count=1;probe.func_sig=Some(("i32".into(),"i32".into()));probe.memory_min_pages=vec![1];
 for offset in [0,16] {
  probe.emit_op_u16(Op::LOCAL_GET,0,1);if offset!=0 {probe.emit_i32_const(offset,1);probe.emit_op(Op::I32_ADD,1);}
  probe.emit_op(Op::V128_LOAD,1);probe.emit_leb_u32(0x80,1);probe.emit_leb_u32(0,1);
 }
 probe.emit_op(eq,1);probe.emit_op(Op::I8X16_BITMASK,1);probe.emit_op(Op::RETURN,1);
 let mut root=Chunk::new("byte_equality_owner");root.memory_min_pages=vec![1];root.emit_op_u16(Op::REF_FUNC,1,1);root.emit(0,1);root.emit_op(Op::RETURN,1);
 let mut vm=VM::new();let f=vm.run(vec![root,probe]).unwrap();
 for mask in 0..(1usize<<(16/width)) {
  let address=mask&15;
  vm.memory.with_buffer_mut(|m|{for lane in 0..16 {let a=(mask as u8).wrapping_add(lane as u8*13);m[address+lane]=a;m[address+16+lane]=if mask&(1<<(lane/width))!=0 {a} else {a^0x80};}});
  assert_eq!(vm.invoke(&f,&[Value::I32(address as i32)]).unwrap().as_f64(),(0..16).fold(0u32,|out,lane|out|(((mask>>(lane/width))&1) as u32)<<lane) as f64);
 }
 assert!(vm.invoke(&f,&[Value::I32(65505)]).is_err());
 }
}
