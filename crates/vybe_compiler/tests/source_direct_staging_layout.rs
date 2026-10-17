use vybe_runtime::{Chunk,Op,VM,Value};
use vybe_compiler::primitives::functions;
#[test]
fn repeated_source_direct_calls_keep_native_staging_after_environment_and_cells() {
 for captured in [false,true] {
  let mut root=Chunk::new("source-direct-staging");root.local_count=3;
  root.emit_string_const("capture",1);root.emit_op_u16(Op::LOCAL_SET,1,1);
  let mut helper=functions::create_function_chunk("private-source-identity",3);helper.emit_op_u16(Op::LOCAL_GET,0,1);helper.emit_op(Op::RETURN,1);
  for value in ["first","second"] {
   root.emit_string_const(value,1);root.emit_string_const("delimiter",1);root.emit_string_const("enclosure",1);
   root.emit_direct_defined_call(1,&helper,false,1).unwrap();root.emit_op(Op::DROP,1);
  }
  let mut child=functions::create_function_chunk("actual-captured-child",0);child.local_count=1;child.capture_count=1;child.capture_base=0;child.emit_op_u16(Op::LOCAL_GET,0,1);child.emit_op(Op::RETURN,1);
  if captured {functions::emit_ref_func(&mut root,2,1,1);functions::emit_closure_upvalue(&mut root,true,1,1);} else {root.emit_string_const("done",1);}
  root.emit_op(Op::RETURN,1);let chunks=vec![root,helper,child];
  let bytes=vybe_platform_wasm::write_wasm(&chunks);wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
  let mut vm=VM::new();let result=vm.run(chunks).unwrap();
  if captured {assert_eq!(vm.invoke(&result,&[]).unwrap(),Value::String("capture".into()));}else {assert_eq!(result,Value::String("done".into()));}
 }
}
