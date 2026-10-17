#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk,VM};
#[path="fixtures/primitive_native_export.rs"] mod native_export;
fn compile(source:&str,fast:bool)->Vec<Chunk>{
    static REGISTER:std::sync::Once=std::sync::Once::new();
    REGISTER.call_once(vybe_language_csharp::register);
    let ast=vybe_language_csharp::parse(source).unwrap();
    let profile=vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source()).unwrap();
    let mut cc=vybe_compiler::Compiler::with_profile(profile);
    cc.benchmark_static_class_member_writes(fast);
    cc.compile(&ast).unwrap()
}
fn vm()->VM {let mut v=VM::new();primitives::platforms::register_platforms_all(&mut v);v}
fn output(name:&str,value:&serde_json::Value){std::fs::write(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("../../STATIC_CLASS_WRITE_{name}.json")),serde_json::to_vec_pretty(value).unwrap()).unwrap();}
fn bytes(cs:&[Chunk])->usize {cs.iter().map(|c|c.code.len()).sum()}
fn source(repeats:usize)->String{
 format!("class Box {{public double X;}} double Run(){{var b=new Box(); {} return b.X;}} return Run;",(0..repeats).map(|i|format!("b.X={i};")).collect::<String>())
}
#[test]
fn values_and_native_sizes(){
 let mut rows=vec![];
 for (label,source) in [
  ("field","class Box {public double X;} var b=new Box(); b.X=7; return b.X;"),
  ("this","class Box {public double X; public Box(double x){this.X=x;}} var b=new Box(9); return b.X;"),
  ("parameter","class Box {public double X;} double Write(Box b){b.X=11;return b.X;} return Write(new Box());"),
  ("inherited","class A {public double X;} class B:A {} var b=new B(); b.X=13; return b.X;"),
  ("property","class Box {public double X {get;set;}} var b=new Box(); b.X=17; return b.X;"),
  ("reference-marker","class Box {public double X;public string __ref_kind;public Box __value;} var inner=new Box();var outer=new Box();outer.__ref_kind=\"cell\";outer.__value=inner;outer.X=7;return inner.X;"),
  ("custom-setter","class Box {public double V;public double X {get{return V;}set{V=value+1;}}} var b=new Box(); b.X=19; return b.X;"),
 ]{
  let mut outcomes=vec![];let mut sizes=vec![];let mut native=vec![];
  for fast in [false,true]{let cs=compile(source,fast);sizes.push(bytes(&cs));native_export::export_chunks(&cs,&format!("static-write-{label}-{fast}"));outcomes.push(vm().run(cs.clone()).map(|v|format!("{v:?}")).map_err(|e|e.message));
   let wasm=vybe_platform_wasm::write_wasm_with_options(&cs,vybe_platform_wasm::WasmWriteOptions{include_name_section:false,include_vybe_metadata:false});
   wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&wasm).unwrap();native.push(wasm.len());
   let root=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../artifacts/static-class-writes");std::fs::create_dir_all(&root).unwrap();std::fs::write(root.join(format!("{label}-{fast}.wasm")),wasm).unwrap();
  }
  assert!(outcomes[0].is_ok(),"{label}: {outcomes:?}");assert_eq!(outcomes[0],outcomes[1],"{label}");
  let expected=match label {"field"|"reference-marker"=>7.,"this"=>9.,"parameter"=>11.,"inherited"=>13.,"property"=>17.,"custom-setter"=>20.,_=>unreachable!()};
  assert_eq!(outcomes[1],Ok(format!("F64({expected:?})")),"{label}: numeric oracle");
  eprintln!("{label}: {outcomes:?} {sizes:?} {native:?}");rows.push(serde_json::json!({"label":label,"outcomes":outcomes,"internal_bytes":sizes,"native_bytes":native}));
 }
 output("GATES",&serde_json::json!({"rows":rows}));
}
#[test]
fn repeated_write_kernel(){
 for n in [1,4,32]{
  let mut sizes=vec![];
  for fast in [false,true]{let cs=compile(&source(n),fast);sizes.push(bytes(&cs));let mut v=vm();let f=v.run(cs).unwrap();assert_eq!(v.invoke(&f,&[]).unwrap().as_f64(),(n-1)as f64);}
  assert!(sizes[1]<sizes[0],"{n}: {sizes:?}");
 }
}

#[test]
fn dynamic_and_pointer_receivers_retain_the_probe(){
 static REGISTER:std::sync::Once=std::sync::Once::new();REGISTER.call_once(vybe_language_csharp::register);
 for mode in [0,1,2] {
  let mut ast=vybe_language_csharp::parse("class Box {public double X;} double Write(Box b){b.X=7;return 0;} return Write;").unwrap();
  if mode==0 {ast.directives.type_resolution=Some(vybe_ast::TypeResolution::Hybrid);}
  if mode==1 {for st in &mut ast.body {if let vybe_ast::StmtKind::FunctionDecl{params,..}=&mut st.kind {for p in params {if p.name=="b" {p.type_hint=Some(vybe_ast::TypeHint::descriptive("Box*"));}}}}}
  if mode==2 {ast.directives.type_resolution=Some(vybe_ast::TypeResolution::Dynamic);}
  let mut variants=vec![];
  for fast in [false,true]{let profile=vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source()).unwrap();let mut cc=vybe_compiler::Compiler::with_profile(profile);cc.benchmark_static_class_member_writes(fast);variants.push(cc.compile(&ast).unwrap());}
  assert_eq!(variants[0].len(),variants[1].len());
  for(a,b)in variants[0].iter().zip(&variants[1]){assert_eq!(a.code,b.code,"{mode}: {}",a.name);assert_eq!(a.constants,b.constants);assert_eq!(a.imports,b.imports);assert_eq!(a.global_imports,b.global_imports);assert_eq!(a.local_count,b.local_count);}
 }
}
#[test]
#[ignore]
fn debug_pairs(){
 use std::time::Instant;
 let mut rows=vec![];
 for (n,control) in [(32,true),(1,false),(4,false),(32,false)]{
  let mut vs=[vm(),vm()];let fs=[vs[0].run(compile(&source(n),false)).unwrap(),vs[1].run(compile(&source(n),!control)).unwrap()];
  for side in 0..2 {for _ in 0..8 {assert_eq!(vs[side].invoke(&fs[side],&[]).unwrap().as_f64(),(n-1)as f64);}}
  let mut wall=[vec![],vec![]];let mut cpu=[vec![],vec![]];
  for pair in 0..7 {
   let mut secs=[0.0,0.0];let mut cs=[0.0,0.0];let mut count=0;
   while secs[0]<0.25||secs[1]<0.25||count<8 {
    for side in if (pair+count)%2==0 {[0,1]}else{[1,0]}{
     let c=thread_cpu_seconds();let t=Instant::now();std::hint::black_box(vs[side].invoke(&fs[side],&[]).unwrap());secs[side]+=t.elapsed().as_secs_f64();cs[side]+=thread_cpu_seconds()-c;
    }count+=1;
   }
   for side in 0..2 {wall[side].push(secs[side]/count as f64);cpu[side].push(cs[side]/count as f64);}
  }
  let mut ratios:Vec<_>=wall[0].iter().zip(&wall[1]).map(|(a,b)|a/b).collect();ratios.sort_by(f64::total_cmp);
  let mut cr:Vec<_>=cpu[0].iter().zip(&cpu[1]).map(|(a,b)|a/b).collect();cr.sort_by(f64::total_cmp);
  rows.push(serde_json::json!({"writes":n,"control":control,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"seconds_per_call":wall,"thread_cpu_seconds_per_call":cpu}));
 }
 output("MEASUREMENTS",&serde_json::json!({"pairs":7,"separate_vms":true,"cases":rows}));
}
#[cfg(target_os="macos")]
fn thread_cpu_seconds()->f64 {#[repr(C)]struct Timespec{sec:i64,nsec:i64}unsafe extern "C"{fn clock_gettime(clock:i32,result:*mut Timespec)->i32;}let mut t=Timespec{sec:0,nsec:0};assert_eq!(unsafe{clock_gettime(16,&mut t)},0);t.sec as f64+t.nsec as f64*1e-9}
#[cfg(not(target_os="macos"))]
fn thread_cpu_seconds()->f64 {0.0}

#[test]
fn null_receiver_baseline_behavior_is_preserved(){
 let src="class Box {public double X;} Box b=null; b.X=7; return 0;";
 let outcomes:Vec<_>=[false,true].into_iter().map(|fast|vm().run(compile(src,fast)).map(|v|format!("{v:?}")).map_err(|e|e.message)).collect();
 // Baseline currently ignores this null write; preserve it, without claiming C# correctness.
 assert_eq!(outcomes[0],Ok("F64(0.0)".into()));assert_eq!(outcomes[0],outcomes[1]);
}

#[test]
fn parsed_vb_field_write(){
 use vybe_ast::{Argument,ExprKind,Expression,Statement,StmtKind};
 static REGISTER:std::sync::Once=std::sync::Once::new();REGISTER.call_once(vybe_language_vb::register);
 let source="Public Class Box\n Public X As Double\nEnd Class\nModule Program\n Public Function Run() As Double\n  Dim b As New Box()\n  b.X = 23\n  Return b.X\n End Function\nEnd Module\n";
 let mut ast=vybe_language_vb::parse(source).unwrap();
 ast.body.push(Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Call {callee:Box::new(Expression::new(ExprKind::Member {object:Box::new(Expression::ident("Program")),field:"Run".into(),null_safe:false})),args:Vec::<Argument>::new(),optional:false})))));
 let mut sizes=vec![];let mut outcomes=vec![];
 for fast in [false,true]{let profile=vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source()).unwrap();let mut cc=vybe_compiler::Compiler::with_profile(profile);cc.benchmark_static_class_member_writes(fast);let cs=cc.compile(&ast).unwrap();sizes.push(bytes(&cs));outcomes.push(vm().run(cs).map(|v|v.as_f64()).map_err(|e|e.message));}
 assert_eq!(outcomes,vec![Ok(23.),Ok(23.)]);assert!(sizes[1]<sizes[0],"{sizes:?}");output("VB_GATES",&serde_json::json!({"internal_bytes":sizes,"outcomes":outcomes,"scope":"Parsed VB plus diagnostic call to public Program.Run; no CLI claim"}));
}
