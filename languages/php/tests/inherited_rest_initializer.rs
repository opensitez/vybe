use vybe_compiler::dynamic::RuntimeCompilerService;
use vybe_runtime::{capabilities::Capabilities, Value, VM};

#[test]
fn separately_loaded_parent_constructor_receives_packed_rest_arguments() {
    struct Scratch(std::path::PathBuf);
    impl Drop for Scratch {
        fn drop(&mut self) { let _ = std::fs::remove_dir_all(&self.0); }
    }
    let folder = Scratch(std::env::temp_dir().join(format!(
        "vybe-inherited-rest-{}-{}", std::process::id(),
        std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap().as_nanos()
    )));
    std::fs::create_dir(&folder.0).unwrap();
    std::fs::write(folder.0.join("base.php"), r#"<?php
        class RestBase {
            public $answer;
            function __construct($a=10,$b=20,$c=30) {
                $this->answer = func_num_args() > 3 ? func_get_arg(3) : -1;
            }
        }
        class VariadicBase {
            public $answer;
            function __construct($a, ...$rest) { $this->answer=$rest[0]+$rest[1]; }
        }
        class ArgumentMethods {
            function lookup($key) { return func_num_args()>1 ? func_get_arg(1) : 41; }
        }
    "#).unwrap();
    std::fs::write(folder.0.join("child.php"), r#"<?php
        class RestChild extends RestBase {}
        class VariadicChild extends VariadicBase {}
        class MethodChild extends ArgumentMethods {}
        class MethodWrapper extends ArgumentMethods {
            function lookup($key,$default=null) { return parent::lookup($key,$default); }
            function inherited($key) { return parent::lookup($key); }
        }
    "#).unwrap();
    vybe_language_php::register();
    for (expression, expected) in [
        ("(new RestChild(1,2,3,42))->answer", 42),
        ("(new VariadicChild(1,20,22))->answer", 42),
        ("(new MethodChild())->lookup('key')", 41),
        ("(new MethodChild())->lookup('key',42)", 42),
        ("(new MethodWrapper())->lookup('key',42)", 42),
        ("(new MethodWrapper())->inherited('key')", 41),
    ] {
        let mut vm = VM::new();
        vybe_runtime::init_all_registered(&mut vm, &Capabilities::all());
        vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
        let mut service = RuntimeCompilerService::with_capabilities(&mut vm, Capabilities::all());
        let source = format!("<?php include __DIR__.'/base.php'; include __DIR__.'/child.php'; return {expression};");
        let compiled = service.compile_source_by_name(source, "php", folder.0.join("main.php")).unwrap();
        let result = service.run_compiled(compiled).unwrap();
        assert_eq!(result.as_f64(), expected as f64, "{expression}: {result:?}");
        assert!(!matches!(result, Value::Null | Value::Undefined));
    }
}
