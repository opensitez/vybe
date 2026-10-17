use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, Value, VM};
fn module() -> Vec<Chunk> {
 let mut c=vec![Chunk::new("bin2hex-owner"),Chunk::new("bin2hex-call")];
 c[1].arity=1;c[1].param_count=1;c[1].local_count=1;
 c[1].emit_op_u16(Op::LOCAL_GET,0,0);
 primitives::string_encoding::emit_bin2hex(&mut c,1,1,0);
 c[1].emit_op(Op::RETURN,0);
 primitives::functions::emit_ref_func(&mut c[0],1,0,0);c[0].emit_op(Op::RETURN,0);
 primitives::bundle::finalize_with_runtime_helpers(&mut c);
 primitives::globals::declare_free_globals(&mut c);primitives::globals::normalize_global_table(&mut c);c
}
fn vm()->VM{let mut v=VM::new();primitives::platforms::register_platforms_all(&mut v);v}
fn expected(s:&str)->String{
 let mut out=String::new();let table=b"0123456789abcdef";
 for code in s.encode_utf16(){let hi=code as usize/16;if hi<16 {out.push(table[hi] as char);}out.push(table[code as usize%16] as char);}
 out
}
#[test] fn exact_public_inventory_full_module_native(){
 let c=module();let h=c.iter().find(|c|c.name=="__stdlib_encoding_bin2hex").expect("actual registered helper");
 let mut at=0;let mut ops=0;let mut names=std::collections::BTreeMap::new();
 while at<h.code.len(){let op=Op::from_code_at(&h.code,at);*names.entry(format!("{op:?}")).or_insert(0u32)+=1;ops+=1;at+=4+op.operand_format().size_in(&h.code,at+4);}
 let bytes=vybe_platform_wasm::write_wasm(&c);
 if let Err(e)=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes){std::fs::write(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("whole-native-failure.wasm"),bytes).unwrap();panic!("unwaived completebin2hex: {e}");}
 println!("inventory {}",serde_json::json!({"bytes":h.code.len(),"ops":ops,"locals":h.local_count,"wasm":bytes.len(),"opcodeCounts":names}));
}
#[test] fn all_bytes_bmp_and_surrogate_pair_legacy_boundaries(){
 let c=module();let mut v=vm();let f=v.run(c).unwrap();
 let mut cases=vec![String::new(),"vybe\0\x7f".into(),"éĀ€😀".into()];
 cases.push((0..256).map(|n|char::from_u32(n).unwrap()).collect());
 for n in [0,1,15,16,127,128,255,256,4095,4096,0xd7ff,0xe000,0xffff,0x10000,0x10ffff] {cases.push(char::from_u32(n).unwrap().to_string());}
 for s in cases{let want=expected(&s);assert_eq!(v.invoke(&f,&[Value::String(s.clone().into())]).unwrap(),Value::String(want.into()),"{s:?}");}
}
#[test] fn strict_input_errors_remain_errors_with_no_character_loop(){
 for value in [Value::Null,Value::Undefined,Value::I32(4),Value::Bool(true)]{
 let mut v=vm();let f=v.run(module()).unwrap();assert!(v.invoke(&f,&[value]).is_err());
 }
}
#[test] fn generated_helper_body_executes_in_real_vm_caller(){
 let mut c=module();let index=c.iter().position(|c|c.name=="__stdlib_encoding_bin2hex").unwrap();
 let n=c[index].code.len();assert_eq!(Op::from_code_at(&c[index].code,n-4),Op::RETURN);
 assert!(c[index].stack_switch_handlers.is_empty());assert!(c[index].call_indirect_sigs.is_empty());
 assert!(c[index].block_i32_results.iter().all(|p|*p<n-4));
 // Insert a pure test host immediately before the original final return. All
 // earlier code and offset metadata are untouched; no dispatcher/body substitute.
 let probe=c[index].add_import("test:codec","generatedBodyReturn");
 c[index].code.truncate(n-4);c[index].lines.truncate(n-4);
 c[index].emit_call(probe,0,0);c[index].emit_op(Op::DROP,0);c[index].emit_op(Op::RETURN,0);
 let hits=std::sync::Arc::new(std::sync::atomic::AtomicUsize::new(0));let copy=hits.clone();
 let mut v=vm();v.register_host_fn("test:codec","generatedBodyReturn",Box::new(move|_,args|{assert!(args.is_empty());copy.fetch_add(1,std::sync::atomic::Ordering::SeqCst);Value::Null}));
 let f=v.run(c).unwrap();assert_eq!(v.invoke(&f,&[Value::String("Az".into())]).unwrap(),Value::String("417a".into()));
 assert_eq!(hits.load(std::sync::atomic::Ordering::SeqCst),1,"actual generated body must execute, not registered hostreplacement");
}
