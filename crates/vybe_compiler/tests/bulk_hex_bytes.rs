use vybe_compiler::primitives;
use vybe_runtime::{Chunk,Op,VM,Value};
fn chunks(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("hex");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut chunks = vec![root, c];
    if fast {
        primitives::string_encoding::emit_hex2bytes(&mut chunks, 1, 1, 0);
    } else {
        primitives::string_encoding::emit_hex2bin(&mut chunks, 1, 1, 0);
        primitives::base64::emit_binary_string_to_byte_array(&mut chunks, 1, 0);
    }
    chunks[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
    primitives::globals::declare_free_globals(&mut chunks);
    primitives::globals::normalize_global_table(&mut chunks);
    chunks
}
fn worker(fast: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(chunks(fast)).unwrap();
    (vm, f)
}

fn result(worker:&mut (VM,Value),input:&str)->Result<Vec<i32>,String>{
 match worker.0.invoke(&worker.1,&[Value::String(input.into())]){
  Ok(Value::Object(o))=>{let o=o.lock().unwrap();let vybe_runtime::value::ObjectKind::Array(a)=&o.kind else {panic!("plain array")};Ok(a.iter().map(Value::as_i32).collect())},
  Ok(v)=>panic!("unexpected {v:?}"),Err(e)=>Err(e.message),
 }
}
#[test]
fn fused_hex_bytes_preserves_values_errors_and_fresh_arrays(){
 for fast in [false,true] {wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&chunks(fast))).unwrap();}
 let mut workers=[worker(false),worker(true)];
 let all=(0..=255).map(|b|format!("{b:02x}")).collect::<String>();
 for w in &mut workers {assert_eq!(result(w,&all),Ok((0..=255).collect()));}
 for input in ["", "00abFF436174", "0", "abggcd", "0é", "ab🙂cd", "ab cd", "ab\0cd"] {
  assert_eq!(result(&mut workers[0],input),result(&mut workers[1],input),"{input:?}");
 }
 for fast in [false,true] {let(mut vm,f)=worker(fast);let first=vm.invoke(&f,&[Value::String("61".into())]).unwrap();let second=vm.invoke(&f,&[Value::String("61".into())]).unwrap();let(Value::Object(a),Value::Object(b))=(first,second)else{panic!("arrays")};assert!(!std::sync::Arc::ptr_eq(&a,&b));}
}
