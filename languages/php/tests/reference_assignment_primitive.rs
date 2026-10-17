use std::path::PathBuf;
use vybe_compiler::{dynamic::RuntimeCompilerService, primitives::platforms};
use vybe_runtime::{VM, Value, capabilities::Capabilities, value::ObjectKind};

fn run(source: &str) -> Value {
    vybe_language_php::register();
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let mut service = RuntimeCompilerService::with_capabilities(&mut vm, Capabilities::all());
    let program = service.compile_source(source, language, PathBuf::from("reference-assignment.php")).unwrap();
    let result = service.run_compiled(program).unwrap();
    drop(service);
    assert!(vm.last_exception.is_none());
    result
}

#[test]
fn array_parameter_keeps_entries_through_shared_assignment() {
    let result = run("<?php function merge_defaults($s,$defaults){$parsed=[];parse_str($s,$parsed);return array_merge($defaults,$parsed);}return merge_defaults('a=1',['b'=>2]);");
    let Value::Object(result) = result else { panic!("missing PHP array") };
    let result = result.lock().unwrap();
    let ObjectKind::Map(entries) = &result.kind else { panic!("PHP array carrier changed") };
    assert_eq!(entries.len(), 2);
    assert_eq!(entries.get(&Value::String("a".into())), Some(&Value::String("1".into())));
    assert_eq!(entries.get(&Value::String("b".into())), Some(&Value::F64(2.0)));
}

#[test]
fn object_parameter_mutation_reaches_original_reference() {
    assert_eq!(run("<?php class ReferenceBox{public $map=[];}function update_box($box){$box->map=['key'=>'value'];}$box=new ReferenceBox;update_box($box);return $box->map['key']==='value';"), Value::Bool(true));
}

#[test]
fn global_element_reference_writes_original_slot() {
    for source in [
        "<?php $GLOBALS['a']=3;$x=&$GLOBALS['a'];$x=8;function g($key){return $GLOBALS[$key];}return g('a');",
        "<?php $GLOBALS['a']=3;$key='a';$x=&$GLOBALS[$key];$x=8;return $GLOBALS['a'];",
        "<?php $GLOBALS['a']=3;function change(){$key='a';$x=&$GLOBALS[$key];$x=8;}change();return $GLOBALS['a'];",
        "<?php $GLOBALS['a']=3;$calls=0;function key_once(){global $calls;$calls++;return 'a';}$x=&$GLOBALS[key_once()];$x=8;return $GLOBALS['a']+$calls-1;",
    ] {
        let direct = vybe_language_php::parse(source).unwrap();
        let peg = vybe_language_php::parse_with_peg(source).unwrap();
        assert_eq!(format!("{direct:?}"), format!("{peg:?}"), "{source}");
        assert_eq!(run(source), Value::F64(8.0), "{source}");
    }
}
