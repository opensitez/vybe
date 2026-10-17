//! Actual public primitive and complete module; independent legacy contract.
use vybe_compiler::primitives;
use vybe_runtime::{Chunk,Op,Value,VM};
fn module()->Vec<Chunk>{
 let mut c=vec![Chunk::new("similar-exit-owner"),Chunk::new("similar-exit-call")];
 c[1].arity=2;c[1].param_count=2;c[1].local_count=2;
 for k in 0..2{for s in [0,1]{c[1].emit_op_u16(Op::LOCAL_GET,s,0);}primitives::string_similarity::emit_similar_text(&mut c,1,2,0);if k==0{c[1].emit_op(Op::DROP,0);}}
 c[1].emit_op(Op::RETURN,0);primitives::functions::emit_ref_func(&mut c[0],1,0,0);c[0].emit_op(Op::RETURN,0);
 primitives::bundle::finalize_with_runtime_helpers(&mut c);primitives::globals::declare_free_globals(&mut c);primitives::globals::normalize_global_table(&mut c);c
}
#[test]fn installed_private_all_consumed_exit_retains_frame_and_standard_module(){
 let c=module();let h=c.iter().find(|h|h.name=="__stdlib_similar_text").unwrap();assert_eq!(h.local_count,8);
 let mut at=0;let mut exit_conditions=0;
 while at<h.code.len(){let op=Op::from_code_at(&h.code,at);let next=at+4+op.operand_format().size_in(&h.code,at+4);
 if op==Op::F64_GE&&next<h.code.len()&&Op::from_code_at(&h.code,next)==Op::BR_IF{exit_conditions+=1;}
 at=next;}
 assert_eq!(exit_conditions,1,"one proven private all-consumed exit, not a replacement comparison dispatcher");
 assert!(h.code.len()<1500,"bounded +21-byte control change, no cache arrays/new slots");
 wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&c)).unwrap();
}
#[test]fn legacy_count_dense_disjoint_threshold_unicode_and_empty(){
 let mut vm=VM::new();primitives::platforms::register_platforms_all(&mut vm);let f=vm.run(module()).unwrap();
 for(a,b)in[("a".into(),"aaa".into()),("ab".into(),"ba".into()),("".into(),"abc".into()),("abc".into(),"".into()),("a".repeat(63),"a".repeat(65)),("a".repeat(64),"a".repeat(64)),("é".repeat(512),"é".into()),("é".repeat(128),"é".repeat(128)),("é".repeat(128),"€".repeat(128)),("😀".repeat(32),"🦀".repeat(32))]{
  let units=|s:&str|s.encode_utf16().map(|u|if(0xd800..=0xdfff).contains(&u){0xfffd}else{u}).collect::<Vec<_>>();let au=units(&a);let expected=units(&b).iter().filter(|u|au.contains(u)).count();
  assert_eq!(vm.invoke(&f,&[Value::String(a.into()),Value::String(b.into())]).unwrap(),Value::F64(expected as f64),"actual legacy frequency contract, separate from normative PHP recursive algorithm");
 }
}
