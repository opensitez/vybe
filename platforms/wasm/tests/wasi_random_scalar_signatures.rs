use vybe_runtime::{Chunk, Op};
#[test]
fn full_scalar_random_modules_have_actual_i64_results() {
 for (module,name) in [("wasi:random/insecure","get-insecure-random-u64"),("wasi:random/random","get-random-u64")] {
  let mut root=Chunk::new("root");root.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN,0);root.emit_op(Op::RETURN,0);
  let mut leaf=Chunk::new("random");let i=leaf.add_import(module,name);leaf.emit_call(i,0,0);leaf.emit_op(Op::F64_CONVERT_I64_U,0);leaf.emit_op(Op::RETURN,0);
  let bytes=vybe_platform_wasm::write_wasm(&[root,leaf]);
  wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
 }
}
