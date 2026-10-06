use vybe_runtime::{Chunk, Op, Value, VM};
use vybe_runtime::chunk::{ConstExpr, GlobalInit};
use vybe_compiler::primitives::{bundle, callable, globals, platforms};
#[path="numeric_optimization_similar_text_frozen.rs"] mod baseline;
fn emit(which:u8,chunks:&mut [Chunk],current:usize,argc:u8,line:u32){
    match which {0=>baseline::emit_similar_text(chunks,current,argc,line),2=>vybe_compiler::primitives::string_similarity::emit_similar_text(chunks,current,argc,line),_=>unreachable!()}
}
fn program(which:u8,calls:usize)->Vec<Chunk>{
    let mut c=vec![Chunk::new("<similar-owner>"),Chunk::new("similar_callback")];
    c[1].arity=2;c[1].param_count=2;c[1].local_count=2;
    c[1].emit_op_u16(Op::LOCAL_GET,0,59);c[1].emit_op_u16(Op::LOCAL_GET,1,59);emit(which,&mut c,1,2,59);c[1].emit_op(Op::RETURN,59);
    c[0].global_inits.push(GlobalInit{name:"similar_callback".into(),init:ConstExpr::RefFunc(1)});
    for i in 0..calls {globals::emit_read(&mut c[0],"similar_callback",59);globals::emit_read(&mut c[0],"left_value",59);globals::emit_read(&mut c[0],"right_value",59);callable::emit_direct_invoke_chunk(&mut c[0],2,59);if i+1!=calls{c[0].emit_op(Op::DROP,59);}}
    c[0].emit_op(Op::RETURN,59);bundle::finalize_with_runtime_helpers(&mut c);globals::declare_free_globals(&mut c);globals::normalize_global_table(&mut c);c
}
fn runtime(a:Value,b:Value)->VM {let mut v=VM::new();platforms::register_platforms_all(&mut v);v.set_global("left_value",a);v.set_global("right_value",b);v}
fn result(a:Value,b:Value,c:Vec<Chunk>)->Result<String,String>{runtime(a,b).run(c).map(|x| match x{Value::F64(f)=>format!("F64:{:x}",f.to_bits()),x=>format!("{x:?}")}).map_err(|mut e|{for f in &mut e.call_stack{f.offset=0;}e.to_string()})}
fn validate(c:&[Chunk]){wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(c)).unwrap();}
fn count(c:&Chunk)->usize{let(mut pc,mut n)=(0,0);while pc<c.code.len(){let op=Op::decode(c.read_u16(pc),c.read_u16(pc+2)).unwrap();pc+=4+op.operand_format().size_in(&c.code,pc+4);n+=1;}assert_eq!(pc,c.code.len());n}
