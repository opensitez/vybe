//! True emitted-site counts and actual helper-index fallback, separate proofs.
include!("fixtures/numeric_optimization_similar_text_fixture.rs");
fn sites(which:u8,n:usize)->Vec<Chunk> {
 let mut c=vec![Chunk::new("<sites-owner>"),Chunk::new("similar_callback")];
 c[1].arity=2;c[1].param_count=2;c[1].local_count=2;
 for i in 0..n {c[1].emit_op_u16(Op::LOCAL_GET,0,59);c[1].emit_op_u16(Op::LOCAL_GET,1,59);
  emit(which,&mut c,1,2,59);if i+1<n {c[1].emit_op(Op::DROP,59);}}
 c[1].emit_op(Op::RETURN,59);
 c[0].global_inits.push(GlobalInit{name:"similar_callback".into(),init:ConstExpr::RefFunc(1)});
 globals::emit_read(&mut c[0],"similar_callback",59);globals::emit_read(&mut c[0],"left_value",59);globals::emit_read(&mut c[0],"right_value",59);
 callable::emit_direct_invoke_chunk(&mut c[0],2,59);c[0].emit_op(Op::RETURN,59);
 bundle::finalize_with_runtime_helpers(&mut c);globals::declare_free_globals(&mut c);globals::normalize_global_table(&mut c);c
}
#[test]
fn true_64_emitted_sites_static_native_and_behavior() {
 for n in [1,64] {
  let old=sites(0,n);let new=sites(2,n);
  let before=(old[1].code.len(),count(&old[1]));let after=(new[1].code.len(),count(&new[1]));
  assert_eq!(new.iter().filter(|c|c.name=="__stdlib_similar_text").count(),usize::from(n>1));
  eprintln!("REAL emitted sites{n} caller bytes/ops {before:?}->{after:?}; whole bytes {}->{}; chunks {}->{}; caller locals {}->{}",old.iter().map(|c|c.code.len()).sum::<usize>(),new.iter().map(|c|c.code.len()).sum::<usize>(),old.len(),new.len(),old[1].local_count,new[1].local_count);
  if n==64 {assert!(new.iter().map(|c|c.code.len()).sum::<usize>()<old.iter().map(|c|c.code.len()).sum::<usize>());}
  validate(&old);validate(&new);
  let a=Value::String("aba".repeat(32).into());let b=Value::String("bab".repeat(32).into());
  assert_eq!(result(a.clone(),b.clone(),old),result(a,b,new));
 }
}
#[test]
fn actual_php_32_algorithm_sites_static_native_and_result() {
 static REGISTER:std::sync::Once=std::sync::Once::new();REGISTER.call_once(vybe_language_php::register);
 let mut source=String::from("<?php function similar_subject($a,$b){$sum=0;");
 for _ in 0..32 {source.push_str("$sum+=similar_text($a,$b);");}
 source.push_str("return $sum;}return similar_subject('abaaba','babbab');");
 let ast=vybe_language_php::parse(&source).unwrap();let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
 let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();validate(&chunks);
 let helper=chunks.iter().position(|c|c.name=="__stdlib_similar_text").unwrap();
 let caller=chunks.iter().find(|c|c.name.ends_with("similar_subject")).unwrap();
 let mut pc=0;let mut refs=0;
 while pc<caller.code.len() {let op=Op::decode(caller.read_u16(pc),caller.read_u16(pc+2)).unwrap();
  if op==Op::REF_FUNC && caller.read_u16(pc+4) as usize==helper {refs+=1;}
  pc+=4+op.operand_format().size_in(&caller.code,pc+4);
 }
 assert_eq!(refs,32,"32 actual primitive callsites, not repeated execution of one site");
 eprintln!("actualPHP32 sites caller bytes/ops={}/{} helperbytes={} total={}",caller.code.len(),count(caller),chunks[helper].code.len(),chunks.iter().map(|c|c.code.len()).sum::<usize>());
 assert_eq!(runtime(Value::Null,Value::Null).run(chunks).unwrap(),Value::F64(192.));
}
fn wide_module(base:u16)->Vec<Chunk> {
 let mut c=vec![Chunk::new("<wide-sites-owner>"),Chunk::new("similar_callback")];
 c[1].arity=2;c[1].param_count=2;c[1].local_count=base;
 c[1].emit_op_u16(Op::LOCAL_GET,0,59);c[1].emit_op_u16(Op::LOCAL_GET,1,59);
 emit(2,&mut c,1,2,59);assert_eq!(c[1].local_count,base+2);c[1].emit_op(Op::RETURN,59);
 for _ in c.len()..65536 {let mut p=Chunk::new("<valid-wide-padding>");p.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN,59);p.emit_op(Op::RETURN,59);c.push(p);}
 c[0].global_inits.push(GlobalInit{name:"similar_callback".into(),init:ConstExpr::RefFunc(1)});
 globals::emit_read(&mut c[0],"similar_callback",59);globals::emit_read(&mut c[0],"left_value",59);globals::emit_read(&mut c[0],"right_value",59);
 callable::emit_direct_invoke_chunk(&mut c[0],2,59);c[0].emit_op(Op::RETURN,59);
 bundle::finalize_with_runtime_helpers(&mut c);globals::declare_free_globals(&mut c);globals::normalize_global_table(&mut c);
 assert_eq!(c.len(),65536);assert_eq!(c[1].local_count,base+8);
 assert!(!c.iter().any(|p|p.name=="__stdlib_similar_text"));
 assert!(!c[1].local_names.iter().any(|n|n.name.starts_with("<private-helper-ref:")));
 assert!(!c[0].globals.iter().any(|n|n=="__vybe_similar_text"));
 let mut pc=0;while pc<c[1].code.len(){let op=Op::decode(c[1].read_u16(pc),c[1].read_u16(pc+2)).unwrap();assert_ne!(op,Op::REF_FUNC);pc+=4+op.operand_format().size_in(&c[1].code,pc+4);}
 c
}
#[test]
fn true_65536_helper_index_falls_back_at_original_maximum_frame() {
 // Actual oversized private-helper fallback at the original MAX-8 caller base.
 let c=wide_module(u16::MAX-8);
 let a=Value::String("aba".repeat(32).into());let b=Value::String("bab".repeat(32).into());
 assert_eq!(runtime(a,b).run(c).unwrap(),Value::F64(96.));
}
#[test]
fn true_65536_helper_index_fallback_full_native_at_ordinary_frame() {
 // Independently emitted ordinary frame avoids validator's total-local ceiling;
 // same >65535 helper index, exact real fallback, entire module stays intact.
 let c=wide_module(2);validate(&c);
 let a=Value::String("aba".repeat(32).into());let b=Value::String("bab".repeat(32).into());
 assert_eq!(runtime(a,b).run(c).unwrap(),Value::F64(96.));
}
