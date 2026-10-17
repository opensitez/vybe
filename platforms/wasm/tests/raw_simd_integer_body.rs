use vybe_runtime::{Chunk,Op,VM,Value};
fn probe()->Chunk {
 let mut c=Chunk::new("zero_mask_16");c.arity=1;c.param_count=1;c.local_count=1;
 c.func_sig=Some(("i32".into(),"i32".into()));c.memory_min_pages=vec![1];
 c.emit_op_u16(Op::LOCAL_GET,0,1);c.emit_op(Op::V128_LOAD,1);c.emit_leb_u32(0x80,1);c.emit_leb_u32(0,1);
 c.emit_op(Op::V128_CONST,1);for _ in 0..16 {c.emit(0,1);}
 c.emit_op(Op::I8X16_EQ,1);c.emit_op(Op::I8X16_BITMASK,1);c.emit_op(Op::RETURN,1);c
}
#[test]
fn standard_simd_probe_complete_native_validation(){
 let bytes=vybe_platform_wasm::write_wasm(&[probe()]);
 wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
 if let Some(path)=std::env::var_os("VYBE_SCAN_NATIVE_OUTPUT"){std::fs::write(path,bytes).unwrap();}
}

#[test]
fn standard_simd_reader_tracks_nested_comparison_stack() {
    let mut chunk = Chunk::new("nested_simd");
    chunk.func_sig = Some(("".into(), "i32".into()));
    let block = chunk.emit_block(1);
    for _ in 0..2 {
        chunk.emit_op(Op::V128_CONST, 1);
        for _ in 0..16 { chunk.emit(0, 1); }
    }
    chunk.emit_op(Op::I8X16_EQ, 1);
    chunk.emit_op(Op::I8X16_BITMASK, 1);
    chunk.emit_op(Op::DROP, 1);
    chunk.emit_end(1);
    chunk.patch_block(block);
    chunk.emit_i32_const(0, 1);
    chunk.emit_op(Op::RETURN, 1);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&[chunk],
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    vybe_platform_wasm::read_wasm(&bytes).unwrap();
}
#[test]
fn standard_simd_probe_vm_matches_every_zero_lane_and_unaligned_load(){
 let mut root=Chunk::new("root");root.memory_min_pages=vec![1];root.emit_op_u16(Op::REF_FUNC,1,1);root.emit(0,1);root.emit_op(Op::RETURN,1);
 let mut vm=VM::new();let function=vm.run(vec![root,probe()]).unwrap();
 for address in [0usize,1,7,15,65520] {
  for bits in [0u16,1,2,4,8,16,32,64,128,256,512,1024,2048,4096,8192,16384,32768,65535,0x5555,0xaaaa] {
   vm.memory.with_buffer_mut(|m|{m[address..address+16].fill(0x91);for lane in 0..16 {if bits&(1<<lane)!=0 {m[address+lane]=0;}}});
   assert_eq!(vm.invoke(&function,&[Value::I32(address as i32)]).unwrap(),Value::I32(bits as i32));
  }
 }
 assert!(vm.invoke(&function,&[Value::I32(65521)]).is_err());
}
