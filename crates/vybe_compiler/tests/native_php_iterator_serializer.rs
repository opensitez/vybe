#[path="fixtures/native_php_request_context.rs"] mod http_request_control;
use vybe_runtime::{VM,Value,value::ObjectKind};
fn function(vm:&VM,name:&str)->Value{vm.global(&format!("__vybe_func${name}")).or_else(||vm.global(name)).unwrap().clone()}
#[test]
fn parsed_iterator_key_value_and_null_serializer() {
    vybe_language_php::register();
    let source=include_str!("fixtures/native_php_iterator_serializer.php");
    let ast=vybe_language_php::parse(source).unwrap();
    let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();
    let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);vm.run(chunks.clone()).unwrap();
    let collected=function(&vm,"collected_values");let result=vm.invoke(&collected,&[]).unwrap();
    let Value::Object(object)=result else{panic!("actual array")};let object=object.lock().unwrap();let ObjectKind::Array(values)=&object.kind else{panic!("actual array kind")};
    assert_eq!(values,&vec![Value::F64(3.),Value::F64(7.)]);drop(object);
    for (name,expected) in [("nullable_serialized","N;"),("scalar_serialized","i:3;")]{let f=function(&vm,name);assert_eq!(vm.invoke(&f,&[]).unwrap(),Value::String(expected.into()));}
    let f=function(&vm,"native_in_array_controls");let value=vm.invoke(&f,&[]).unwrap();
    let Value::Object(object)=value else{panic!("actual control array")};let object=object.lock().unwrap();let ObjectKind::Array(values)=&object.kind else{panic!("actual array kind")};
    assert_eq!(values,&vec![Value::Bool(true),Value::Bool(false),Value::Bool(true),Value::Bool(false),Value::Bool(true),Value::Bool(false)]);drop(object);
    let f=function(&vm,"native_pathinfo_controls");let value=vm.invoke(&f,&[]).unwrap();
    let Value::Object(object)=value else{panic!("actual path control array")};let object=object.lock().unwrap();let ObjectKind::Array(values)=&object.kind else{panic!("array kind")};
    assert_eq!(values,&vec!["/a","b.txt","txt","b",""].into_iter().map(|s|Value::String(s.into())).collect::<Vec<_>>());drop(object);
    let f=function(&vm,"native_ini_controls");let value=vm.invoke(&f,&[]).unwrap();
    let Value::Object(object)=value else{panic!("actual ini control array")};let object=object.lock().unwrap();let ObjectKind::Array(values)=&object.kind else{panic!("array kind")};
    assert_eq!(values,&vec![Value::String("14".into()),Value::Bool(false),Value::String("14".into()),Value::String("!".into()),Value::Null,Value::String("14".into())]);drop(object);
    let request=http_request_control::request();
    let _request_guard=vybe_platform_node::http::install_context(request.clone());
    let header=function(&vm,"native_header_controls");assert_eq!(vm.invoke(&header,&[]).unwrap(),Value::Null);
    let response=request.response.lock().unwrap();
    assert_eq!(response.status,204);
    assert_eq!(response.headers,vec![("X-Probe".into(),"yes".into()),("X-Probe".into(),"again".into())]);
    drop(response);
    // Retain every Source function/import, including unused serializer branches.
    let echo=function(&vm,"native_echo_controls");let value=vm.invoke(&echo,&[]).unwrap();
    let Value::Object(object)=value else{panic!("echo control array")};let object=object.lock().unwrap();let ObjectKind::Array(values)=&object.kind else{panic!("array kind")};
    assert_eq!(values,&vec![Value::String("a3b".into()),Value::I32(1)]);drop(object);
    let random=function(&vm,"native_random_controls");let value=vm.invoke(&random,&[]).unwrap();
    let Value::Object(object)=value else{panic!("random control array")};let object=object.lock().unwrap();let ObjectKind::Array(values)=&object.kind else{panic!("array kind")};
    assert_eq!(values[0],Value::F64(7.));
    let Value::F64(draw)=values[1] else{panic!("randNumber")};assert!((1.0..=3.0).contains(&draw));assert_eq!(draw.trunc(),draw);
    let Value::F64(draw)=values[2] else{panic!("lcgNumber")};assert!((0.0..1.0).contains(&draw));drop(object);
    let size=function(&vm,"native_file_size_control");
    let source_path=concat!(env!("CARGO_MANIFEST_DIR"),"/tests/fixtures/native_php_iterator_serializer.php");
    assert_eq!(vm.invoke(&size,&[Value::String(source_path.into())]).unwrap(),Value::F64(source.len() as f64));
    let native=vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&native).unwrap();
}

#[test]
#[ignore="existing PHP generator string-key fallback bug: nativePHP uses0, emitted iterator uses1"]
fn normative_string_key_then_implicit_numeric_key() {
    vybe_language_php::register();
    let ast=vybe_language_php::parse(include_str!("fixtures/native_php_iterator_serializer.php")).unwrap();
    let profile=vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();
    let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);vm.run(chunks).unwrap();
    let keyed=function(&vm,"collected_keys");let result=vm.invoke(&keyed,&[]).unwrap();
    let Value::Object(object)=result else{panic!("actualmap")};let object=object.lock().unwrap();let ObjectKind::Map(entries)=&object.kind else{panic!("map")};
    assert_eq!(entries.get(&Value::F64(0.)),Some(&Value::F64(7.)));
}

#[test]
fn actual_dynamic_runtime_string_callback_controls() {
    vybe_language_php::register();
    let source=format!("{}\nreturn native_callback_controls();\n",include_str!("fixtures/native_php_iterator_serializer.php"));
    let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let mut service=vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    let compiled=service.compile_source_by_name(source,"php","native-callback-controls.php").unwrap();
    let native=vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&native).unwrap();
    let value=service.run_compiled(compiled).unwrap();
    let Value::Object(object)=value else{panic!("callback control array")};let object=object.lock().unwrap();let ObjectKind::Array(values)=&object.kind else{panic!("array kind")};
    assert_eq!(values,&vec![Value::F64(7.),Value::F64(11.)]);
}
