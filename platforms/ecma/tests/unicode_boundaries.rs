#![cfg(feature="primitive-benchmarks")]
use std::sync::{Arc,Mutex};
use vybe_runtime::{Chunk,Op,Value,VM,value::Object};
const OPS:[&str;3]=["includes","startsWith","endsWith"];
fn program(op:usize,arity:u8)->Vec<Chunk>{let mut root=Chunk::new("boundary-owner");root.emit_op_u16(Op::REF_FUNC,1,0);root.emit(0,0);root.emit_op(Op::RETURN,0);let mut c=Chunk::new("entry");c.arity=arity;c.param_count=arity;c.local_count=arity as u16;for i in 0..arity{c.emit_op_u16(Op::LOCAL_GET,i as u16,0);}let idx=c.add_import("ecma:string",OPS[op]);c.emit_call(idx,arity,0);c.emit_op(Op::RETURN,0);root.imports=c.imports.clone();vec![root,c]}
fn worker(op:usize,arity:u8,fast:bool)->(VM,Value){let mut v=VM::new();vybe_platform_ecma::string::register(&mut v);vybe_platform_ecma::string::benchmark_utf8_boundaries(&mut v,fast);let f=v.run(program(op,arity)).unwrap();(v,f)}
fn oracle(op:usize,s:&str,n:&str,p:Option<&Value>)->Value{let h:Vec<_>=s.encode_utf16().collect();let n:Vec<_>=n.encode_utf16().collect();let default=if op==2{h.len()as i32}else{0};let raw=match p{Some(Value::I32(n))=>*n,Some(Value::F64(n))=>*n as i32,_=>default};let pos=if op==2&&p.is_none(){h.len()}else{(raw.max(0)as usize).min(h.len())};let yes=match op{0=>n.is_empty()||h[pos..].windows(n.len()).any(|w|w==n),1=>n.len()<=h.len()-pos&&h[pos..pos+n.len()]==n,2=>n.len()<=pos&&h[pos-n.len()..pos]==n,_=>unreachable!()};Value::Bool(yes)}
fn output(name:&str,v:&serde_json::Value){let d=std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap();std::fs::create_dir_all(&d).unwrap();std::fs::write(std::path::Path::new(&d).join(format!("BOUNDARIES_{name}.json")),serde_json::to_vec_pretty(v).unwrap()).unwrap();}
#[test]
fn utf16_boundaries_empty_needles_and_coercions(){
 let mut checks=0;
 for op in 0..3{for arity in[2,3]{let mut ws=[worker(op,arity,false),worker(op,arity,true)];
  for s in["","ASCII","Ωé中","😀A😀","𐐀Ω😀A","A\0ΩA","😀😀😀","éééé"]{for n in["","A","Ω","😀","𐐀","\0","Ω😀","😀A","missing","éé"]{for p in[Value::I32(-1),Value::I32(0),Value::I32(1),Value::I32(2),Value::I32(3),Value::I32(4),Value::I32(5),Value::I32(6),Value::I32(20),Value::I32(i32::MAX),Value::F64(1.9),Value::F64(f64::INFINITY),Value::F64(f64::NEG_INFINITY),Value::F64(f64::NAN),Value::Undefined,Value::Null,Value::Bool(true),Value::String("2".into())]{
   let mut args=vec![Value::String(s.into()),Value::String(n.into())];if arity==3{args.push(p.clone());}let expected=oracle(op,s,n,args.get(2));for(v,f)in &mut ws{assert_eq!(v.invoke(f,&args).unwrap(),expected,"{} {s:?} {n:?} {p:?}",OPS[op]);checks+=1;}
  }}}
  for s in[Value::Null,Value::Undefined,Value::I32(123),Value::Bool(true)]{for n in[Value::Undefined,Value::Null,Value::I32(2),Value::Bool(false)]{let mut args=vec![s.clone(),n];if arity==3{args.push(Value::String("2".into()));}let outcomes:Vec<_>=ws.iter_mut().map(|(v,f)|v.invoke(f,&args).map_err(|e|e.message)).collect();assert_eq!(outcomes[0],outcomes[1]);checks+=2;}}
 }}
 output("GATES",&serde_json::json!({"checks":checks,"scope":"Actual registered old/new adapters with identical coercion paths; independent UTF16 reference for empty, clamped and surrogate-middle boundaries; absent vs provided positions and legacy nonnumeric conversions"}));
}
#[test]
fn every_scalar_after_an_astral_prefix(){
 let mut workers:Vec<_>=(0..3).map(|op|[worker(op,3,false),worker(op,3,true)]).collect();let mut checks=0;
 for cp in 0..=0x10ffff{let Some(ch)=char::from_u32(cp)else{continue;};let s:Arc<str>=Arc::from(format!("😀{ch}A{ch}😀"));let n:Arc<str>=Arc::from(ch.to_string());for op in 0..3{let pos=match op{0=>1,1=>2,_=>2+ch.len_utf16()as i32};let args=[Value::String(s.clone()),Value::String(n.clone()),Value::I32(pos)];for(v,f)in &mut workers[op]{assert_eq!(v.invoke(f,&args).unwrap(),Value::Bool(true),"{} U+{cp:04X}",OPS[op]);checks+=1;let negative=[Value::String(s.clone()),Value::String(n.clone()),Value::I32(if op==0{i32::MAX}else{1})];assert_eq!(v.invoke(f,&negative).unwrap(),Value::Bool(false),"{} negative U+{cp:04X}",OPS[op]);checks+=1;}}}
 output("SCALARS",&serde_json::json!({"checks":checks,"scalars":1112064}));
}
#[test]
fn includes_preserves_regexp_rejection(){
 let mut obj=Object::new();obj.properties.insert("__type".into(),Value::String("RegExp".into()));let re=Value::Object(Arc::new(Mutex::new(obj)));let args=[Value::String("abc".into()),re,Value::I32(0)];let mut errors=vec![];for fast in[false,true]{let(mut v,f)=worker(0,3,fast);errors.push(v.invoke(&f,&args).unwrap_err().message);}assert_eq!(errors[0],errors[1]);
}

#[test]
fn randomized_js_oracle_corpus(){
 let mut workers:Vec<_>=(0..3).map(|op|[worker(op,3,false),worker(op,3,true)]).collect();
 let mut rows=vec![];let mut state=0x12345678u64;let chars=['A','a','Ω','é','中','😀','𐐀','\0','👩','\u{200d}','%','\n'];
 for case in 0..12000{
  let mut next=||{state^=state<<13;state^=state>>7;state^=state<<17;state};
  let len=(next()%24)as usize;let source:String=(0..len).map(|_|chars[next()as usize%chars.len()]).collect();
  let needle=if case%3==0&&!source.is_empty(){let cs:Vec<_>=source.chars().collect();let a=next()as usize%cs.len();let b=(a+1+(next()%5)as usize).min(cs.len());cs[a..b].iter().collect::<String>()}else{let n=(next()%6)as usize;(0..n).map(|_|chars[next()as usize%chars.len()]).collect()};
  let position=match case%13{0=>f64::INFINITY,1=>f64::NEG_INFINITY,2=>f64::NAN,_=>(next()%44)as f64-2.5};
  let args=[Value::String(source.clone().into()),Value::String(needle.clone().into()),Value::F64(position)];
  for op in 0..3{let expected=oracle(op,&source,&needle,args.get(2));for(v,f)in &mut workers[op]{assert_eq!(v.invoke(f,&args).unwrap(),expected);}
   let value=matches!(expected,Value::Bool(true));let p=if position.is_nan(){serde_json::json!("NaN")}else if position==f64::INFINITY{serde_json::json!("Infinity")}else if position==f64::NEG_INFINITY{serde_json::json!("-Infinity")}else{serde_json::json!(position)};
   rows.push(serde_json::json!({"op":OPS[op],"source":source,"needle":needle,"position":p,"value":value}));
  }
 }
 output("JS_CORPUS",&serde_json::json!({"queries":rows}));
}
#[test]
#[ignore]
fn debug_pairs(){
 use std::time::Instant;let mut rows=vec![];
 for(op,label,source,needle,position,control)in[
  (0,"control","Ω".repeat(4096),String::from("missing"),1.,true),
  (0,"ascii-includes","A".repeat(4096)+"!","!".into(),1.,false),
  (0,"unicode-includes-missing","Ω".repeat(4096),"x".into(),1.,false),
  (0,"unicode-includes-tail","Ω".repeat(4096)+"!","!".into(),1.,false),
  (0,"astral-includes","😀".repeat(2048),"😀".into(),1.,false),
  (1,"ascii-starts-zero","A".repeat(4096),"A".into(),0.,false),
  (1,"unicode-starts-zero","Ω".repeat(4096),"Ω".into(),0.,false),
  (1,"unicode-starts-position","Ω".repeat(4096),"Ω".into(),2.,false),
  (1,"astral-starts-middle","😀".repeat(2048),"😀".into(),1.,false),
  (2,"unicode-ends-default","Ω".repeat(4096),"Ω".into(),f64::INFINITY,false),
  (2,"unicode-ends-position","Ω".repeat(4096),"Ω".into(),2.,false),
  (2,"astral-ends-middle","😀".repeat(2048),"😀".into(),1.,false),
  (0,"tiny-includes","Ωa".into(),"a".into(),1.,false),
  (1,"tiny-starts","Ωa".into(),"a".into(),1.,false),
  (2,"tiny-ends","Ωa".into(),"a".into(),2.,false),
  (2,"unicode-empty","Ω".repeat(4096),"".into(),1.,false),
 ]{
  let(va,fa)=worker(op,3,false);let(vb,fb)=worker(op,3,!control);let mut vs=[va,vb];let fs=[fa,fb];let args=[Value::String(Arc::from(source.as_str())),Value::String(Arc::from(needle.as_str())),Value::F64(position)];let expected=oracle(op,&source,&needle,args.get(2));
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
