use vybe_compiler::primitives::{bundle,globals,platforms,tuples};
use vybe_runtime::{Chunk,Op,Value,VM};
use vybe_runtime::value::{Object,ObjectKind};
#[path="fixtures/tuple_reloaded_aliases.rs"] mod before;
pub(crate) fn program(fast:bool,n:u8)->Vec<Chunk>{let mut root=Chunk::new("tuple-owner");root.emit_op_u16(Op::REF_FUNC,1,0);root.emit(0,0);root.emit_op(Op::RETURN,0);let mut tuple=Chunk::new("tuple");tuple.arity=n;tuple.param_count=n;tuple.local_count=n as u16;for i in 0..n{tuple.emit_op_u16(Op::LOCAL_GET,i as u16,0);}let mut code=vec![root,tuple];if fast{tuples::emit_tuple(&mut code,1,n as u16,0);}else{before::emit_tuple(&mut code,1,n as u16,0);}code[1].emit_op(Op::RETURN,0);bundle::finalize_with_runtime_helpers(&mut code);globals::declare_free_globals(&mut code);globals::normalize_global_table(&mut code);code}
fn worker(fast:bool,n:u8)->(VM,Value){let mut vm=VM::new();platforms::register_platforms_all(&mut vm);let f=vm.run(program(fast,n)).unwrap();(vm,f)}
fn check(v:&Value,args:&[Value]){let Value::Object(o)=v else{panic!("tuple")};let o=o.lock().unwrap();let ObjectKind::Array(items)=&o.kind else{panic!("tuple array")};assert_eq!(items,args);assert_eq!(o.properties.get("__tuple"),Some(&Value::Bool(true)));for(i,arg)in args.iter().enumerate(){for key in[format!("Item{}",i+1),format!("item{}",i+1)]{assert_eq!(o.properties.get(key.as_str()),Some(arg));if let Value::Object(expected)=arg{let Value::Object(actual)=o.properties.get(key.as_str()).unwrap() else{panic!("reference")};assert!(std::sync::Arc::ptr_eq(expected,actual));}}}}
#[test]
fn tuple_values_aliases_and_identity_are_preserved(){let object=Value::Object(vybe_runtime::heap::alloc(Object::new()));let inputs=vec![vec![],vec![Value::Null],vec![Value::I32(7),Value::String("x".into())],vec![Value::Undefined,Value::Bool(true),Value::F64(-0.0),object.clone(),Value::I32(-9),Value::String("λ".into()),Value::Bool(false),Value::Null]];for args in inputs{for fast in[false,true]{let(mut vm,f)=worker(fast,args.len()as u8);let out=vm.invoke(&f,&args).unwrap();check(&out,&args);}}}
#[test]
fn native_validation_and_size(){let mut rows=vec![];for n in[0,1,2,3,8,16]{let mut sizes=vec![];for fast in[false,true]{let code=program(fast,n);let wasm=vybe_platform_wasm::write_wasm(&code);wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&wasm).unwrap();sizes.push((code[1].code.len(),wasm.len(),code[1].local_count));}if n>0{assert!(sizes[1].0<sizes[0].0);assert!(sizes[1].1<sizes[0].1);}assert_eq!(sizes[0].2,sizes[1].2);rows.push(serde_json::json!({"arity":n,"before":sizes[0],"after":sizes[1]}));}if let Ok(path)=std::env::var("TUPLE_GATES_PATH"){std::fs::write(path,serde_json::to_vec_pretty(&rows).unwrap()).unwrap();}}
#[test]#[ignore="Seven paired opt0 tuple construction batches"]
fn debug_pairs(){let args=[Value::I32(7),Value::String("x".into())];let mut workers=[worker(false,2),worker(true,2)];let mut counts=[1,1];let mut raw=vec![];let mut ratios=vec![];for pair in 0..7{let mut ns=[0,0];for side in[pair%2,1-pair%2]{loop{let start=std::time::Instant::now();for _ in 0..counts[side]{let w=&mut workers[side];std::hint::black_box(w.0.invoke(&w.1,&args).unwrap());}ns[side]=start.elapsed().as_nanos();if ns[side]>=100_000_000{break;}counts[side]*=2;}}ratios.push((ns[0]as f64/counts[0]as f64)/(ns[1]as f64/counts[1]as f64));raw.push(serde_json::json!({"counts":counts,"ns":ns}));}let mut sorted=ratios.clone();sorted.sort_by(f64::total_cmp);let result=serde_json::json!({"median":sorted[3],"ratios":ratios,"raw":raw,"scope":"opt0 two-element tuple construction, VM invocation and full array/tag/alias materialization included"});if let Ok(path)=std::env::var("TUPLE_MEASUREMENTS_PATH"){std::fs::write(path,serde_json::to_vec_pretty(&result).unwrap()).unwrap();}}
#[test]
fn pack_preserves_saved_locals_and_large_arity_fallback(){
    for n in[0u16,1,2,255,256,300]{for fast in[false,true]{
        let mut c=Chunk::new("pack-boundary");c.local_count=n.max(1);
        for i in 0..n{c.emit_i32_const(i as i32,0);}
        let mut code=vec![c];if fast{vybe_compiler::primitives::collections::emit_pack_n(&mut code,0,n,0,0);}else{before::emit_pack_n(&mut code,0,n,0,0);}
        if n>0{code[0].emit_op_u16(Op::LOCAL_GET,n-1,0);}else{code[0].emit_i32_const(-1,0);}
        code[0].emit_array_new_fixed(0,2,0);code[0].emit_op(Op::RETURN,0);
        let wasm=vybe_platform_wasm::write_wasm(&code);wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&wasm).unwrap();
        let mut vm=VM::new();platforms::register_platforms_all(&mut vm);let out=vm.run(code).unwrap();let Value::Object(o)=out else{panic!("wrapper")};let o=o.lock().unwrap();let ObjectKind::Array(pair)=&o.kind else{panic!("wrapper")};assert_eq!(pair[1],Value::I32(n as i32-1));let Value::Object(a)=&pair[0] else{panic!("packed")};let a=a.lock().unwrap();let ObjectKind::Array(values)=&a.kind else{panic!("packed")};assert_eq!(values,&(0..n).map(|i|Value::I32(i as i32)).collect::<Vec<_>>());
    }}
}
pub(crate) fn pack_program(fast:bool,n:u16)->Vec<Chunk>{
    let mut root=Chunk::new("tuple-owner");root.emit_op_u16(Op::REF_FUNC,1,0);root.emit(0,0);root.emit_op(Op::RETURN,0);
    let mut c=Chunk::new("tuple");c.local_count=n.max(1);for i in 0..n{c.emit_i32_const(i as i32,0);}
    let mut code=vec![root,c];if fast{vybe_compiler::primitives::collections::emit_pack_n(&mut code,1,n,0,0);}else{before::emit_pack_n(&mut code,1,n,0,0);}code[1].emit_op(Op::RETURN,0);bundle::finalize_with_runtime_helpers(&mut code);globals::declare_free_globals(&mut code);globals::normalize_global_table(&mut code);code
}
