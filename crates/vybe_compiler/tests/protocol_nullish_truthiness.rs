use vybe_runtime::{Value,VM,Chunk};
fn compile(source:&str)->Vec<Chunk>{vybe_language_python::register();let m=vybe_language_python::parse(source).unwrap();vybe_compiler::Compiler::with_profile(vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap()).compile(&m).unwrap()}
fn vm()->VM{let mut v=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut v);v}
fn enc(v:&Value)->String{match v{Value::Bool(b)=>b.to_string(),Value::F64(n)=>n.to_string(),Value::I32(n)=>n.to_string(),Value::String(s)=>format!("{:?}",s.as_ref()),Value::Null|Value::Undefined=>"null".into(),Value::Object(o)=>{let o=o.lock().unwrap();let vybe_runtime::value::ObjectKind::Array(a)=&o.kind else{panic!("Array result")};format!("[{}]",a.iter().map(enc).collect::<Vec<_>>().join(","))},_=>panic!("result {v:?}")}}
fn run(s:&str)->VM{let c=compile(s);let mut v=vm();v.run(c).unwrap();v}
#[test] fn actual_python_builtin_empty_nonempty_and_nested_not(){let v=run("answer=[bool([]),bool([1]),bool({}),bool({'a':1}),bool(''),bool('a'),bool(None),not [],not (not []),not [1],not (not [1])]\n");assert_eq!(enc(v.global("answer").unwrap()),"[false,true,false,true,false,true,false,true,false,false,true]");}
#[test] fn actual_python_bool_precedes_len_once_each_conversion(){let v=run(r#"events=[]
class Probe:
    def __bool__(self):
        events.append('bool')
        return False
    def __len__(self):
        events.append('len')
        return 3
p=Probe()
answer=[bool(p),not p,not (not p),events]
"#);assert_eq!(enc(v.global("answer").unwrap()),r#"[false,true,false,["bool","bool","bool"]]"#);}
#[test] fn actual_python_absent_bool_falls_to_len_once(){let v=run(r#"events=[]
class Probe:
    def __len__(self):
        events.append('len')
        return 0
p=Probe()
answer=[bool(p),not p,not (not p),events]
"#);assert_eq!(enc(v.global("answer").unwrap()),r#"[false,true,false,["len","len","len"]]"#);}
#[test] fn false_and_zero_method_slots_are_not_absent(){let c=compile("def truth_subject(x):\n    return bool(x)\n");for slot in [vybe_ast::ProtocolSlot::Bool,vybe_ast::ProtocolSlot::Len]{for answer in [Value::Bool(false),Value::I32(0)]{let mut v=vm();v.run(c.clone()).unwrap();let f=v.global("truth_subject").or_else(||v.global("__vybe_func$truth_subject")).unwrap().clone();let mut o=vybe_runtime::value::Object::new();o.properties.insert(vybe_ast::protocol_slot_key(slot),answer);let e=v.invoke(&f,&[Value::Object(vybe_runtime::heap::alloc(o))]).err().expect("noncallable present protocol slot must attempt invocation");assert!(!e.message.contains("Undefined"),"{e:?}");assert!(e.call_stack.iter().any(|s|s.chunk_name=="truth_subject"),"real invocation error retains source frames: {e:?}");}}}
#[test] fn method_throw_keeps_callback_count_and_source_frame(){let s=r#"events=[]
class Probe:
    def __bool__(self):
        events.append('bool')
        return 1/0
p=Probe()
answer=bool(p)
"#;let mut v=vm();let e=v.run(compile(s)).err().expect("actual method failure");assert_eq!(enc(v.global("events").unwrap()),r#"["bool"]"#);assert!(e.call_stack.iter().any(|f|f.chunk_name=="__bool__"),"{e:?}");}
#[test] #[ignore="strict whole actual Python Source native diagnostic, no function pruning"] fn complete_source_native(){for s in ["answer=bool([])\n", "answer=bool(None)\n", "class Probe:\n    def __bool__(self):\n        return False\np=Probe()\nanswer=not (not p)\n"]{let c=compile(s);let b=vybe_platform_wasm::write_wasm(&c);wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&b).unwrap();}}
