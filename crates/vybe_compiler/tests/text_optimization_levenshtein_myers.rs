use vybe_compiler::primitives;
use vybe_runtime::{Chunk,Op,Value,VM};
fn module()->Vec<Chunk>{
 let mut c=vec![Chunk::new("lev-owner"),Chunk::new("lev-call")];
 c[1].arity=2;c[1].param_count=2;c[1].local_count=2;
 for slot in [0,1]{c[1].emit_op_u16(Op::LOCAL_GET,slot,0);}
 primitives::string_similarity::emit_levenshtein(&mut c,1,2,0);c[1].emit_op(Op::RETURN,0);
 primitives::functions::emit_ref_func(&mut c[0],1,0,0);c[0].emit_op(Op::RETURN,0);
 primitives::bundle::finalize_with_runtime_helpers(&mut c);
 primitives::globals::declare_free_globals(&mut c);primitives::globals::normalize_global_table(&mut c);c
}
fn vm()->VM{let mut v=VM::new();primitives::platforms::register_platforms_all(&mut v);v}
fn units(s:&str)->Vec<u16>{s.encode_utf16().map(|u|if (0xd800..=0xdfff).contains(&u){0xfffd}else{u}).collect()}
fn dp(a:&str,b:&str)->usize{let(a,b)=(units(a),units(b));let mut row:Vec<usize>=(0..=b.len()).collect();for(i,&x)in a.iter().enumerate(){let mut diag=row[0];row[0]=i+1;for(j,&y)in b.iter().enumerate(){let old=row[j+1];row[j+1]=(row[j]+1).min(old+1).min(diag+usize::from(x!=y));diag=old;}}row[b.len()]}
fn distance(v:&mut VM,f:&Value,a:&str,b:&str){assert_eq!(v.invoke(f,&[Value::String(a.into()),Value::String(b.into())]).unwrap(),Value::F64(dp(a,b) as f64),"{a:?} vs{b:?}");}
#[test]fn exhaustive_independent_global_dp_binary_alphabet(){let mut v=vm();let f=v.run(module()).unwrap();let mut strings=vec![String::new()];for len in 1..=5{for bits in 0..1usize<<len{strings.push((0..len).map(|i|if bits&(1<<i)==0{'a'}else{'b'}).collect());}}for a in &strings{for b in &strings{distance(&mut v,&f,a,b);}}println!("exhaustive {} real emitted pairs",strings.len()*strings.len());}
#[test]fn high_bit_wrap_random_short_long_and_unicode_fallback(){let mut v=vm();let f=v.run(module()).unwrap();
 for(a,b)in [("",""),("a",""),("","a"),("kitten","sitting"),("😀","🦀"),("😀","��"),("é€Ā","e€a"),("a\u{2028}","a\u{2029}")]{distance(&mut v,&f,a,b);}
 for n in [1,2,7,8,15,16,30,31,32,33,64]{for m in [0,1,7,16,31,32,33,96]{for (a,b)in [("a".repeat(m),"b".repeat(n)),("ab".repeat(m).chars().take(m).collect::<String>(),"ba".repeat(n).chars().take(n).collect::<String>())]{distance(&mut v,&f,&a,&b);}}}
 let mut seed=0x12345678u32;for _ in 0..256{let mut make=|len|->String{(0..len).map(|_|{seed=seed.wrapping_mul(1664525).wrapping_add(1013904223);(b'a'+(seed%6)as u8)as char}).collect()};let a=make(32);let b=make(32);distance(&mut v,&f,&a,&b);}
}
#[test]fn actual_private_helper_inventory_and_complete_standard_module(){let c=module();let h=c.iter().find(|c|c.name=="__stdlib_levenshtein").expect("actual emitted helper");let mut at=0;let mut ops=0;let mut counts=std::collections::BTreeMap::new();while at<h.code.len(){let op=Op::from_code_at(&h.code,at);*counts.entry(format!("{op:?}")).or_insert(0u32)+=1;ops+=1;at+=4+op.operand_format().size_in(&h.code,at+4);}let b=vybe_platform_wasm::write_wasm(&c);if let Err(e)=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&b){std::fs::write(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("full-native-failure.wasm"),b).unwrap();panic!("complete actual helper unwaived:{e}");}println!("inventory {}",serde_json::json!({"bytes":h.code.len(),"ops":ops,"locals":h.local_count,"wasm":b.len(),"counts":counts}));}
#[test]fn initial_strict_length_error_order(){for args in [[Value::Null,Value::String("a".into())],[Value::String("a".into()),Value::I32(2)],[Value::Undefined,Value::Null]]{let mut v=vm();let f=v.run(module()).unwrap();assert!(v.invoke(&f,&args).is_err());}}
#[test]fn generated_private_body_executes_in_real_caller(){let mut c=module();let index=c.iter().position(|c|c.name=="__stdlib_levenshtein").unwrap();let n=c[index].code.len();assert_eq!(Op::from_code_at(&c[index].code,n-4),Op::RETURN);assert!(c[index].stack_switch_handlers.is_empty());assert!(c[index].call_indirect_sigs.is_empty());assert!(c[index].block_i32_results.iter().all(|p|*p<n-4));let probe=c[index].add_import("test:myers","bodyReturn");c[index].code.truncate(n-4);c[index].lines.truncate(n-4);c[index].emit_call(probe,0,0);c[index].emit_op(Op::DROP,0);c[index].emit_op(Op::RETURN,0);let hits=std::sync::Arc::new(std::sync::atomic::AtomicUsize::new(0));let copy=hits.clone();let mut v=vm();v.register_host_fn("test:myers","bodyReturn",Box::new(move|_,args|{assert!(args.is_empty());copy.fetch_add(1,std::sync::atomic::Ordering::SeqCst);Value::Null}));let f=v.run(c).unwrap();distance(&mut v,&f,"kitten","sitting");assert_eq!(hits.load(std::sync::atomic::Ordering::SeqCst),1);}

#[test]
fn exhaustive_binary_suffixes_force_threshold_selected_route() {
 let mut v=vm();let f=v.run(module()).unwrap();
 let mut strings=vec!["________".to_string()];
 for len in 1..=5 {for bits in 0..1usize<<len {let mut s="________".to_string();s.extend((0..len).map(|i|if bits&(1<<i)==0{'a'}else{'b'}));strings.push(s);}}
 for a in &strings {for b in &strings {assert!((8..=32).contains(&a.len())&&(8..=32).contains(&b.len()));distance(&mut v,&f,a,b);}}
 println!("{} independent DP pairs all satisfy actual ASCII fast-route length gates",strings.len()*strings.len());
}

#[test]
fn reversed_short_pattern_against_long_text_matches_independent_dp() {
 let mut v=vm();let f=v.run(module()).unwrap();
 for short in [8,16,31,32] {for long in [33,64,256,2048] {
  let a="ab".repeat(short).chars().take(short).collect::<String>();
  let b="ba".repeat(long).chars().take(long).collect::<String>();
  distance(&mut v,&f,&a,&b);distance(&mut v,&f,&b,&a);
 }}
}
