use vybe_compiler::primitives::{bundle, collections::{self,ZipLen}, globals, platforms};
use vybe_runtime::{Chunk,Op,VM,Value};
use vybe_runtime::value::{Object,ObjectKind};
#[path="fixtures/zip_dynamic_bounds.rs"] mod before;
fn array(values:Vec<Value>)->Value{Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))}
fn values(v:&Value)->Vec<Value>{let Value::Object(o)=v else{panic!("array")};let o=o.lock().unwrap();let ObjectKind::Array(v)=&o.kind else{panic!("array")};v.clone()}
fn snapshot(v:&Value)->String{if let Value::Object(o)=v{let o=o.lock().unwrap();if let ObjectKind::Array(items)=&o.kind{return format!("{:?}:[{}]",o.properties.get("__tuple"),items.iter().map(snapshot).collect::<Vec<_>>().join(","));}}format!("{v:?}")}
pub(crate) fn program(fast:bool,n:u8,mode:ZipLen)->Vec<Chunk>{
    let mut root=Chunk::new("zip-owner");root.emit_op_u16(Op::REF_FUNC,1,0);root.emit(0,0);root.emit_op(Op::RETURN,0);
    let mut zip=Chunk::new("zip");zip.arity=n;zip.param_count=n;zip.local_count=n as u16;for i in 0..n{zip.emit_op_u16(Op::LOCAL_GET,i as u16,0);}
    let mut code=vec![root,zip];if fast{collections::emit_zip(&mut code,1,n,mode,0);}else{before::emit_zip(&mut code,1,n,mode,0);}code[1].emit_op(Op::RETURN,0);bundle::finalize_with_runtime_helpers(&mut code);globals::declare_free_globals(&mut code);globals::normalize_global_table(&mut code);code
}
fn worker(fast:bool,n:u8,mode:ZipLen)->(VM,Value){let mut vm=VM::new();platforms::register_platforms_all(&mut vm);let f=vm.run(program(fast,n,mode)).unwrap();(vm,f)}
#[test]
fn modes_lengths_tuple_tags_and_values_are_preserved(){
    let cases=vec![vec![array(vec![])],vec![array(vec![Value::I32(1),Value::Null])],vec![array(vec![Value::I32(1),Value::I32(2),Value::I32(3)]),array(vec![Value::String("x".into())])],vec![array(vec![]),array(vec![Value::Bool(true)]),array(vec![Value::I32(4),Value::I32(5)])],vec![Value::String("abc".into()),array(vec![Value::I32(1),Value::I32(2)])],vec![Value::Null,array(vec![Value::I32(2)])]];
    for mode in[ZipLen::First,ZipLen::Shortest,ZipLen::Longest]{for args in&cases{
        let mut outputs=vec![];for fast in[false,true]{let(mut vm,f)=worker(fast,args.len() as u8,mode);outputs.push(vm.invoke(&f,args).map(|v|snapshot(&v)).map_err(|e|e.message));}
        assert_eq!(outputs[0],outputs[1],"arity {}",args.len());
    }}
    for mode in[ZipLen::First,ZipLen::Shortest,ZipLen::Longest]{let(mut vm,f)=worker(true,2,mode);let out=vm.invoke(&f,&[array(vec![Value::I32(1),Value::I32(2)]),array(vec![Value::I32(9)])]).unwrap();let rows=values(&out);assert_eq!(rows.len(),if mode==ZipLen::Shortest{1}else{2});assert_eq!(values(&rows[0]),vec![Value::I32(1),Value::I32(9)]);if rows.len()==2{assert_eq!(values(&rows[1]),vec![Value::I32(2),Value::Undefined]);}}
}
#[test]
fn native_validation_and_size(){
    let mut rows=vec![];
    for(mode,label)in[(ZipLen::First,"first"),(ZipLen::Shortest,"shortest"),(ZipLen::Longest,"longest")]{for n in[1,2,3,8]{let mut sizes=vec![];for fast in[false,true]{let code=program(fast,n,mode);let wasm=vybe_platform_wasm::write_wasm(&code);wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&wasm).unwrap();sizes.push((code.iter().map(|c|c.code.len()).sum::<usize>(),wasm.len(),code[1].local_count));}assert!(sizes[1].0<sizes[0].0);assert!(sizes[1].1<sizes[0].1);rows.push(serde_json::json!({"mode":label,"arity":n,"before":sizes[0],"after":sizes[1]}));}}
    if let Ok(path)=std::env::var("ZIP_GATES_PATH"){std::fs::write(path,serde_json::to_vec_pretty(&rows).unwrap()).unwrap();}
}
#[test]#[ignore="Seven paired opt0 zip materialization batches"]
fn debug_pairs(){
    let args=[array((0..4096).map(Value::I32).collect()),array((0..4096).map(Value::I32).collect())];let mut workers=[worker(false,2,ZipLen::Shortest),worker(true,2,ZipLen::Shortest)];let mut counts=[1,1];let mut raw=vec![];let mut ratios=vec![];
    for pair in 0..7{let mut ns=[0,0];for side in[pair%2,1-pair%2]{loop{let start=std::time::Instant::now();for _ in 0..counts[side]{let w=&mut workers[side];let out=w.0.invoke(&w.1,&args).unwrap();assert_eq!(values(&out).len(),4096);std::hint::black_box(out);}ns[side]=start.elapsed().as_nanos();if ns[side]>=100_000_000{break;}counts[side]*=2;}}ratios.push((ns[0]as f64/counts[0]as f64)/(ns[1]as f64/counts[1]as f64));raw.push(serde_json::json!({"counts":counts,"ns":ns}));}
    let mut sorted=ratios.clone();sorted.sort_by(f64::total_cmp);let result=serde_json::json!({"median":sorted[3],"ratios":ratios,"raw":raw,"scope":"opt0 zip two arrays of 4096 elements, shortest mode; VM invocation, eager tuple materialization, result length copy included"});eprintln!("{result}");if let Ok(path)=std::env::var("ZIP_MEASUREMENTS_PATH"){std::fs::write(path,serde_json::to_vec_pretty(&result).unwrap()).unwrap();}
}
