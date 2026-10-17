use vybe_runtime::{VM,Value};
fn function(vm:&VM,name:&str)->Value {vm.global(name).or_else(||vm.global(&format!("__vybe_func${name}"))).expect("actual Source binding").clone()}
#[test]
fn actual_source_fresh_close_does_not_start_and_throw_closes() {
    vybe_language_python::register();
    let ast=vybe_language_python::parse(include_str!("fixtures/native_generator_lifecycle.py")).unwrap();
    let profile=vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
    let chunks=vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();
    let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);vm.run(chunks.clone()).unwrap();
    let fresh=function(&vm,"fresh_close_count");assert_eq!(vm.invoke(&fresh,&[]).unwrap(),Value::F64(0.));
    let values=Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new_array(vec![Value::F64(7.),Value::F64(3.)])));
    let make=function(&vm,"generator_make");let next=function(&vm,"generator_next");let throwing=function(&vm,"generator_throw");
    let generator=vm.invoke(&make,&[values]).unwrap();assert_eq!(vm.invoke(&next,&[generator.clone()]).unwrap(),Value::F64(7.));
    let error=vm.invoke(&throwing,&[generator.clone()]).err().expect("actual throw");assert!(format!("{error:?}").contains("boom"));
    // Runtime consumption is independently observable without re-entering
    // VM::invoke with its retained unhandled-exception slot (known host seam).
    let Value::Object(object)=generator else {panic!("actual continuation")};
    let object=object.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Continuation(state)=&object.kind else {panic!("actual continuation kind")};
    assert_eq!(*state.state.lock().unwrap(),vybe_runtime::value::ContinuationPhase::Done);
    drop(object);
    // Genuine same-Source caught throw + next validates exact exception classes.
    let mut caught=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut caught);caught.run(chunks).unwrap();
    let status=function(&caught,"source_throw_done_status");
    let data=Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new_array(vec![Value::F64(7.),Value::F64(3.)])));
    assert_eq!(caught.invoke(&status,&[data]).unwrap(),Value::String("done".into()));
}
