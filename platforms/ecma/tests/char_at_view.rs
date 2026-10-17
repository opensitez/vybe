//! Actual registered charAt provider through guest bytecode: index and output identity contracts.
use std::sync::Arc;
use vybe_runtime::{Chunk,Op,Value,VM};
fn callable()->(VM,Value) {
 let mut vm=VM::new();vybe_platform_ecma::register(&mut vm);
 let mut c=vec![Chunk::new("charAt-view-owner"),Chunk::new("charAt-view-call")];
 c[1].arity=2;c[1].param_count=2;c[1].local_count=2;
 c[1].emit_op_u16(Op::LOCAL_GET,0,0);c[1].emit_op_u16(Op::LOCAL_GET,1,0);
 let import=c[1].add_import("ecma:string","charAt");c[1].emit_call(import,2,0);c[1].emit_op(Op::RETURN,0);
 vybe_compiler::primitives::functions::emit_ref_func(&mut c[0],1,0,0);c[0].emit_op(Op::RETURN,0);
 let f=vm.run(c).unwrap();(vm,f)
}
fn expected(text:&str,pos:i32)->Value {
 if pos<0{return Value::String("".into());}
 if text.is_ascii() {return Value::String(text.as_bytes().get(pos as usize).map(|x|(*x as char).to_string()).unwrap_or_default().into());}
 let u:Vec<u16>=text.encode_utf16().collect();
 Value::String(u.get(pos as usize).map(|u|String::from_utf16_lossy(&[*u])).unwrap_or_default().into())
}
#[test]fn real_registered_provider_alternating_strings_full_index_contract() {
 let (mut v,f)=callable();
 for (a,b) in [("α😀\0".repeat(128),"β🦀\u{ffff}".repeat(128)),("ascii".repeat(32),"β".repeat(128)),("é".repeat(32768),"€".repeat(21846))] {
  let a:Arc<str>=a.into();let b:Arc<str>=b.into();
  for pos in [-1,0,1,2,3,5,127,128,256,65535,i32::MAX] {
   for s in [&a,&b] {
    assert_eq!(v.invoke(&f,&[Value::String(s.clone()),Value::I32(pos)]).unwrap(),expected(s,pos));
   }
  }
 }
}
#[test]fn old_numeric_index_casts_are_preserved() {
 let (mut v,f)=callable();let s:Arc<str>="α😀z".into();
 for n in [f64::NAN,-0.,0.5,1.9,-1.9,f64::INFINITY,f64::NEG_INFINITY,2147483648.] {
  assert_eq!(v.invoke(&f,&[Value::String(s.clone()),Value::F64(n)]).unwrap(),expected(&s,n as i32));
 }
 for n in [Value::Null,Value::Undefined,Value::Bool(true),Value::String("2".into())] {
  assert_eq!(v.invoke(&f,&[Value::String(s.clone()),n]).unwrap(),expected(&s,0),"existing nonnumeric index defaults0");
 }
}
#[test]fn weak_cache_key_does_not_retain_source_or_change_pool_identity() {
 let (mut v,f)=callable();
 for (text,pos,shared) in [("ascii",0,true),("αβγ",0,false),("aα",0,false),("α",99,true)] {
  let s:Arc<str>=text.into();
  let first=v.invoke(&f,&[Value::String(s.clone()),Value::I32(pos)]).unwrap();
  let second=v.invoke(&f,&[Value::String(s.clone()),Value::I32(pos)]).unwrap();
  if let (Value::String(a),Value::String(b))=(first,second) {
   assert_eq!(Arc::ptr_eq(&a,&b),shared,"exact old key policy for {text:?}");
  } else {panic!("actual String result");}
 }
 // Exact source lifetime is separately tested directly on cache state, because
 // ordinary VM call-frame/source retention must be compared to frozen baseline.
}
