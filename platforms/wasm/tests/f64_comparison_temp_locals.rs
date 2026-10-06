//! Standalone boxed F64 equality consumers: no unrelated temp-reserving op.
use vybe_runtime::{Chunk,Op};
fn module(op:Op,mode:u8)->Vec<Chunk> {
 let mut root=Chunk::new("<comparison-owner>");
 root.emit_op_u16(Op::REF_FUNC,1,59);root.emit(0,59);root.emit_op(Op::RETURN,59);
 let mut c=Chunk::new("compare");c.arity=2;c.param_count=2;c.local_count=2;
 if mode==3 {c.emit_block_i32(59);}
 c.emit_op_u16(Op::LOCAL_GET,0,59);c.emit_op_u16(Op::LOCAL_GET,1,59);c.emit_op(op,59);
 match mode {
  0=>{},
  1=>{c.emit_op(Op::DROP,59);c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN,59);},
  2=>{c.emit_if_value(59);c.emit_bool_const(true,59);c.emit_else(59);c.emit_bool_const(false,59);c.emit_end(59);},
  3=>c.emit_end(59),_=>unreachable!()
 }
 c.emit_op(Op::RETURN,59);vec![root,c]
}
#[test]
fn standalone_f64_eq_ne_native_eight_actual_consumer_contexts() {
 let mut failures=vec![];
 for op in [Op::F64_EQ,Op::F64_NE] {for mode in 0..4 {
  let c=module(op,mode);let bytes=vybe_platform_wasm::write_wasm(&c);
  if let Err(e)=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes) {
   failures.push(format!("{op:?} mode{mode}: {e}"));
  }
 }}
 assert!(failures.is_empty(),"{}",failures.join("\n"));
}
