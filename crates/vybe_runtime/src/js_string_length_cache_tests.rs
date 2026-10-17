use std::sync::Arc;
use crate::{Chunk,Op,Value,VM};
fn program()->Vec<Chunk>{
 let mut root=Chunk::new("length-owner");root.emit_op_u16(Op::REF_FUNC,1,0);root.emit(0,0);root.emit_op(Op::RETURN,0);
 let mut f=Chunk::new("entry");f.arity=1;f.param_count=1;f.local_count=1;f.emit_op_u16(Op::LOCAL_GET,0,0);let idx=f.add_import("wasm:js-string","length");f.emit_call(idx,1,0);f.emit_op(Op::RETURN,0);
 root.imports=f.imports.clone();vec![root,f]
}
fn worker(fast:bool)->(VM,Value){
 let mut v=VM::new();super::register(&mut v);
 v.benchmark_uncached_string_lengths(!fast);
 let f=v.run(program()).unwrap();(v,f)
}
#[test]
fn unicode_lengths_identity_eviction_and_strict_errors(){
 let mut ws=[worker(false),worker(true)];let mut checks=0;
 let roots:Vec<Arc<str>>=["","a","é","😀","a\0Ωé中😀𐐀"].into_iter().map(Arc::from).collect();
 for cp in 0..=0x10ffff{let Some(ch)=char::from_u32(cp)else{continue;};let s:Arc<str>=Arc::from(format!("prefix{ch}suffix"));let expected=s.encode_utf16().count()as i32;
  for(v,f)in &mut ws{assert_eq!(v.invoke(f,&[Value::String(s.clone())]).unwrap(),Value::I32(expected));checks+=1;}
 }
 for n in [1,7,8,9,255,256,65535,65536,65537]{for seed in ["A","Ω","😀","A\0é中😀"]{let s:Arc<str>=Arc::from(seed.repeat(n));for _ in 0..3{for(v,f)in &mut ws{assert_eq!(v.invoke(f,&[Value::String(s.clone())]).unwrap(),Value::I32(s.encode_utf16().count()as i32));checks+=1;}}}}
 for _ in 0..5{for s in &roots{for(v,f)in &mut ws{assert_eq!(v.invoke(f,&[Value::String(s.clone())]).unwrap(),Value::I32(s.encode_utf16().count()as i32));checks+=1;}}}
 for value in [Value::Null,Value::Undefined,Value::I32(7),Value::Bool(true)]{let errors:Vec<_>=ws.iter_mut().map(|(v,f)|v.invoke(f,&[value.clone()]).unwrap_err().message).collect();assert_eq!(errors[0],errors[1]);checks+=2;}
 let mut counts=vec![];
 for(v,f)in &mut ws {
  let source:Arc<str>=Arc::from("Ω😀".repeat(1024));let weak=Arc::downgrade(&source);
  assert_eq!(Arc::strong_count(&source),1);v.invoke(f,&[Value::String(source.clone())]).unwrap();counts.push(Arc::strong_count(&source));drop(source);
  // Both VMs retain the argument in their last frame until its next use.
  v.invoke(f,&[Value::String("other".into())]).unwrap();assert!(weak.upgrade().is_none());
 }
 assert_eq!(counts[0],counts[1],"cache must not add a strong root");
 eprintln!("standard length cache checks={checks}");
}
#[test]
#[ignore]
fn debug_pairs(){
 use std::time::Instant;
 let mut rows=vec![];
 for(label,text,control)in [
  ("control","Ωé中😀".repeat(1024),true),
  ("empty",String::new(),false),
  ("tiny","abc".into(),false),
  ("ascii8","abcdefgh".into(),false),
  ("ascii256","a".repeat(256),false),
  ("ascii16k","a".repeat(16384),false),
  ("unicode","Ωé中😀".repeat(1024),false),
  ("large_uncached","Ωé中😀".repeat(16384),false),
  ("large_ascii","a".repeat(65537),false),
  ("large_astral","😀".repeat(32768),false),
  ("large_ascii_tail",format!("{}😀", "a".repeat(65536)),false),
  ("rotate3_ascii8","abcdefgh".into(),false),
  ("rotate3_ascii256","a".repeat(256),false),
  ("rotate3_unicode","Ωé中😀".repeat(1024),false),
 ]{
  let(va,fa)=worker(false);let(vb,fb)=worker(!control);let mut vs=[va,vb];let fs=[fa,fb];let args=[Value::String(text.clone().into())];let rotating:Vec<[Value;1]>=(0..3).map(|_|[Value::String(Arc::from(text.as_str()))]).collect();
  for side in 0..2{for _ in 0..8{assert_eq!(vs[side].invoke(&fs[side],&args).unwrap(),Value::I32(text.encode_utf16().count()as i32));}}
  let mut wall=[vec![],vec![]];let mut cpu=[vec![],vec![]];let mut call_counts=vec![];
  for pair in 0..7{let mut secs=[0.0,0.0];let mut cpus=[0.0,0.0];let mut counts=[0usize,0];let mut turn=pair;
   // Strings are immutable and this operation has no observable per-call
   // state. Each side independently collects >=100ms to keep fast samples
   // meaningful without forcing the slow scan through the cache-hit count.
   while secs[0]<0.1||secs[1]<0.1||counts.iter().any(|n|*n<8){for side in if turn%2==0{[0,1]}else{[1,0]}{if secs[side]>=0.1&&counts[side]>=8{continue;}let c=thread_cpu_seconds();let t=Instant::now();let input=if label.starts_with("rotate3_"){&rotating[counts[side]%3]}else{&args};std::hint::black_box(vs[side].invoke(&fs[side],input).unwrap());secs[side]+=t.elapsed().as_secs_f64();cpus[side]+=thread_cpu_seconds()-c;counts[side]+=1;}turn+=1;}
   for side in 0..2{wall[side].push(secs[side]/counts[side] as f64);cpu[side].push(cpus[side]/counts[side] as f64);}call_counts.push(counts);
  }
  let mut ratios:Vec<_>=wall[0].iter().zip(&wall[1]).map(|(a,b)|a/b).collect();ratios.sort_by(f64::total_cmp);let mut cr:Vec<_>=cpu[0].iter().zip(&cpu[1]).map(|(a,b)|a/b).collect();cr.sort_by(f64::total_cmp);
  let row=format!("{{\"label\":\"{label}\",\"control\":{control},\"bytes\":{},\"median_speedup\":{},\"median_thread_cpu_speedup\":{},\"wall\":{:?},\"cpu\":{:?},\"call_counts\":{:?}}}",text.len(),ratios[3],cr[3],wall,cpu,call_counts);eprintln!("{row}");rows.push(row);
 }
 let path=std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap();std::fs::create_dir_all(&path).unwrap();std::fs::write(std::path::Path::new(&path).join("STANDARD_LENGTH_MEASUREMENTS.json"),format!("{{\"pairs\":7,\"separate_vms\":true,\"cases\":[{}]}}",rows.join(","))).unwrap();
}
#[cfg(target_os="macos")]
fn thread_cpu_seconds()->f64 {#[repr(C)]struct Timespec{sec:i64,nsec:i64}unsafe extern "C"{fn clock_gettime(clock:i32,result:*mut Timespec)->i32;}let mut t=Timespec{sec:0,nsec:0};assert_eq!(unsafe{clock_gettime(16,&mut t)},0);t.sec as f64+t.nsec as f64*1e-9}
#[cfg(not(target_os="macos"))]
fn thread_cpu_seconds()->f64 {0.0}
#[test]
fn bounded_weak_entries_and_snapshot_reset(){
 let(mut v,f)=worker(true);let snap=v.snapshot();
 let roots:Vec<Arc<str>>=(0..3).map(|i|Arc::from(format!("immutable-Ω😀-{i}"))).collect();
 for _ in 0..3{for root in &roots{assert_eq!(v.invoke(&f,&[Value::String(root.clone())]).unwrap(),Value::I32(root.encode_utf16().count()as i32));}}
 assert_eq!(roots.iter().map(Arc::weak_count).sum::<usize>(),2,"actual builtin must use exactly two weak cache entries");
 let oversized:Arc<str>=Arc::from("x".repeat(crate::string_lengths::MAX_SOURCE_BYTES+1));
 v.invoke(&f,&[Value::String(oversized.clone())]).unwrap();assert_eq!(Arc::weak_count(&oversized),0);
 v.reset_to(&snap);assert_eq!(roots.iter().map(Arc::weak_count).sum::<usize>(),0);
 assert_eq!(v.invoke(&f,&[Value::String(roots[0].clone())]).unwrap(),Value::I32(roots[0].encode_utf16().count()as i32));
 assert_eq!(Arc::weak_count(&roots[0]),1);
}
