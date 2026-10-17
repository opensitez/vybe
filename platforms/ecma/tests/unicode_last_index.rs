#![cfg(feature = "primitive-benchmarks")]
use std::sync::Arc;
use vybe_runtime::{Chunk,Op,Value,VM};
#[path="fixtures/last_index_before.rs"] mod before;
fn program()->Vec<Chunk>{let mut root=Chunk::new("last-index-owner");root.emit_op_u16(Op::REF_FUNC,1,0);root.emit(0,0);root.emit_op(Op::RETURN,0);let mut c=Chunk::new("entry");c.arity=3;c.param_count=3;c.local_count=3;for i in 0..3{c.emit_op_u16(Op::LOCAL_GET,i,0);}let idx=c.add_import("ecma:string","lastIndexOf");c.emit_call(idx,3,0);c.emit_op(Op::RETURN,0);root.imports=c.imports.clone();vec![root,c]}
fn worker(fast:bool)->(VM,Value){let mut v=VM::new();vybe_platform_ecma::string::register(&mut v);vybe_platform_ecma::string::benchmark_last_index_of(&mut v,fast);let f=v.run(program()).unwrap();(v,f)}
fn output(name:&str,v:&serde_json::Value){let dir=std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap();std::fs::create_dir_all(&dir).unwrap();std::fs::write(std::path::Path::new(&dir).join(format!("LAST_INDEX_{name}.json")),serde_json::to_vec_pretty(v).unwrap()).unwrap();}
fn check(ws:&mut[(VM,Value);2],args:&[Value])->usize{let expected=before::invoke(args);for(v,f)in ws{assert_eq!(v.invoke(f,args).unwrap(),expected,"{args:?}");}2}
#[test]
fn utf16_positions_and_adapter_coercions(){
 let mut ws=[worker(false),worker(true)];let mut checks=0;
 let sources=["","ABC","Ωé中","😀A😀","𐐀Ω😀A","A\0ΩA","😀😀😀","éééé","ababΩababΩ"];
 let needles=["","A","Ω","😀","𐐀","\0","Ω😀","😀A","missing","éé","babΩ"];
 for s in sources{for n in needles{for p in[-1.,-0.,0.,0.9,1.,1.9,2.,3.,4.,5.,6.,20.,i32::MAX as f64,f64::INFINITY,f64::NEG_INFINITY,f64::NAN]{checks+=check(&mut ws,&[Value::String(s.into()),Value::String(n.into()),Value::F64(p)]);}}}
 for cp in 0..=0x10ffff{let Some(ch)=char::from_u32(cp)else{continue;};let s=format!("😀{ch}A{ch}😀");let needle=ch.to_string();let args=[Value::String(s.into()),Value::String(needle.into()),Value::I32(3)];checks+=check(&mut ws,&args);}
 for s in [Value::Null,Value::Undefined,Value::I32(123),Value::Bool(true)]{for n in [Value::Undefined,Value::Null,Value::I32(2),Value::Bool(false)]{for p in [Value::Undefined,Value::Null,Value::Bool(true),Value::String("2.9".into()),Value::String("Infinity".into()),Value::String("bad".into())]{checks+=check(&mut ws,&[s.clone(),n.clone(),p]);}}}
 output("GATES",&serde_json::json!({"checks":checks,"unicode_scalars":1112064,"scope":"Actual ecma:string.lastIndexOf import vs frozen UTF16-window algorithm, all scalar needles with astral prefix and middle-surrogate positions; primitive coercions and legacy NaN behavior"}));
}
#[test]
fn overlapping_patterns_and_js_oracle_corpus(){
 let mut ws=[worker(false),worker(true)];let mut rows=vec![];let mut state=0x12345678u64;let chars=['A','a','Ω','é','中','😀','𐐀','\0','👩','\u{200d}','%','\n'];
 for case in 0..12000{
  let mut next=||{state^=state<<13;state^=state>>7;state^=state<<17;state};
  let len=(next()%16)as usize;let source:String=(0..len).map(|_|chars[next()as usize%chars.len()]).collect();
  let needle=if case%3==0&&!source.is_empty(){let cs:Vec<_>=source.chars().collect();let a=next()as usize%cs.len();let b=(a+1+(next()%4)as usize).min(cs.len());cs[a..b].iter().collect::<String>()}else{let n=(next()%5)as usize;(0..n).map(|_|chars[next()as usize%chars.len()]).collect()};
  let position=if case%7==0{f64::INFINITY}else{(next()%24)as f64-2.};let args=[Value::String(source.clone().into()),Value::String(needle.clone().into()),Value::F64(position)];check(&mut ws,&args);
  let f=ws[1].1.clone();let value=ws[1].0.invoke(&f,&args).unwrap().as_f64();let p=if position.is_infinite(){serde_json::json!("Infinity")}else{serde_json::json!(position)};
  rows.push(serde_json::json!({"source":source,"needle":needle,"position":p,"value":value}));
 }
 output("JS_CORPUS",&serde_json::json!({"queries":rows}));
}
#[test]
#[ignore]
fn debug_pairs(){
 use std::time::Instant;let mut rows=vec![];
 for(label,source,needle,position,control)in[
  ("control","Ω".repeat(4096),"missing".into(),f64::INFINITY,true),
  ("ascii-tail","A".repeat(4096)+"!","!".into(),f64::INFINITY,false),
  ("unicode-missing","Ω".repeat(4096),"x".into(),f64::INFINITY,false),
  ("unicode-tail","Ω".repeat(4096)+"!","!".into(),f64::INFINITY,false),
  ("unicode-zero","Ω".repeat(4096),"Ω".into(),0.,false),
  ("astral-middle","😀".repeat(2048),"😀".into(),2049.,false),
  ("overlap-missing","Ω".repeat(4096),"Ω".repeat(128)+"Z",f64::INFINITY,false),
  ("tiny-unicode","Ωa".into(),"Ω".into(),f64::INFINITY,false),
 ]{
  let(va,fa)=worker(false);let(vb,fb)=worker(!control);let mut vs=[va,vb];let fs=[fa,fb];let args=[Value::String(Arc::from(source.as_str())),Value::String(Arc::from(needle.as_str())),Value::F64(position)];let expected=before::invoke(&args);
  for side in 0..2{for _ in 0..8{assert_eq!(vs[side].invoke(&fs[side],&args).unwrap(),expected);}}
  let mut wall=[vec![],vec![]];let mut cpu=[vec![],vec![]];let mut call_counts=vec![];
  for pair in 0..7{let mut secs=[0.0,0.0];let mut cpus=[0.0,0.0];let mut counts=[0usize,0];let mut turn=pair;
   while secs[0]<0.1||secs[1]<0.1||counts.iter().any(|n|*n<8){for side in if turn%2==0{[0,1]}else{[1,0]}{if secs[side]>=0.1&&counts[side]>=8{continue;}let c=thread_cpu_seconds();let t=Instant::now();std::hint::black_box(vs[side].invoke(&fs[side],&args).unwrap());secs[side]+=t.elapsed().as_secs_f64();cpus[side]+=thread_cpu_seconds()-c;counts[side]+=1;}turn+=1;}
   for side in 0..2{wall[side].push(secs[side]/counts[side]as f64);cpu[side].push(cpus[side]/counts[side]as f64);}call_counts.push(counts);
  }
  let mut ratios:Vec<_>=wall[0].iter().zip(&wall[1]).map(|(a,b)|a/b).collect();ratios.sort_by(f64::total_cmp);let mut cr:Vec<_>=cpu[0].iter().zip(&cpu[1]).map(|(a,b)|a/b).collect();cr.sort_by(f64::total_cmp);
  rows.push(serde_json::json!({"label":label,"control":control,"source_bytes":source.len(),"needle_bytes":needle.len(),"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"call_counts":call_counts}));
 }
 output("MEASUREMENTS",&serde_json::json!({"pairs":7,"separate_vms":true,"minimum_side_ms":100,"independent_counts_for_pure_calls":true,"cases":rows}));
}
#[cfg(target_os="macos")]
fn thread_cpu_seconds()->f64 {#[repr(C)]struct Timespec{sec:i64,nsec:i64}unsafe extern "C"{fn clock_gettime(clock:i32,result:*mut Timespec)->i32;}let mut t=Timespec{sec:0,nsec:0};assert_eq!(unsafe{clock_gettime(16,&mut t)},0);t.sec as f64+t.nsec as f64*1e-9}
#[cfg(not(target_os="macos"))]
fn thread_cpu_seconds()->f64{0.0}
