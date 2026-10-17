use std::path::PathBuf;
use vybe_compiler::{dynamic::RuntimeCompilerService, primitives::platforms};
use vybe_runtime::{capabilities::Capabilities, Value, VM};

fn run(source: &str) -> Value {
    vybe_language_php::register();
    let direct = vybe_language_php::parse(source).unwrap();
    let peg = vybe_language_php::parse_with_peg(source).unwrap();
    assert_eq!(format!("{direct:?}"), format!("{peg:?}"));
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let mut service = RuntimeCompilerService::with_capabilities(&mut vm, Capabilities::all());
    let compiled = service
        .compile_source(source, language, PathBuf::from("exceptions.php"))
        .unwrap();
    let value = service.run_compiled(compiled).unwrap();
    drop(service);
    assert!(vm.last_exception.is_none());
    value
}

#[test]
fn builtin_exception_subclasses_share_receiver_methods() {
    for source in [
        "<?php $e=new InvalidArgumentException('bad',17); return $e->getMessage().':'.$e->getCode().':'.$e->getTraceAsString();",
        "<?php $e=new BadMethodCallException('bad',17); $f=[$e,'getMessage']; return $f().':'.$e->getCode().':'.$e->getTraceAsString();",
    ] {
        assert_eq!(run(source), Value::String("bad:17:#0 {main}".into()));
    }
}

#[test]
fn inherited_getters_preserve_previous_identity_and_user_overrides() {
    assert_eq!(run("<?php $p=new Exception('previous'); $e=new PDOException('outer',9,$p); return $e->getPrevious()===$p;"), Value::Bool(true));
    assert_eq!(run("<?php class CustomException extends LogicException {public function __toString(){return 'custom';}} $e=new CustomException('base'); return $e->__toString().':'.$e->getMessage();"), Value::String("custom:base".into()));
}

#[test]
fn error_subclasses_keep_getters_and_throwable_identity() {
    assert_eq!(run("<?php $e=new DivisionByZeroError('math'); return ($e instanceof Throwable) && ($e instanceof Error) && $e->getMessage()==='math' && $e->getLine()>=0 && is_string($e->getFile());"), Value::Bool(true));
}

#[test]
fn builtin_exception_names_admit_their_declarations() {
    for name in [
        "Throwable",
        "Exception",
        "Error",
        "ErrorException",
        "TypeError",
        "ValueError",
        "ArithmeticError",
        "DivisionByZeroError",
        "ArgumentCountError",
        "CompileError",
        "ParseError",
        "AssertionError",
        "UnhandledMatchError",
        "FiberError",
        "RuntimeException",
        "PDOException",
        "LogicException",
        "InvalidArgumentException",
        "DomainException",
        "LengthException",
        "OutOfRangeException",
        "BadFunctionCallException",
        "BadMethodCallException",
        "OutOfBoundsException",
        "RangeException",
        "OverflowException",
        "UnderflowException",
        "UnexpectedValueException",
        "JsonException",
        "__PHP_Incomplete_Class",
    ] {
        let module = vybe_language_php::parse(&format!("<?php new {name}();")).unwrap();
        assert!(module.body.iter().any(|statement| matches!(
            &statement.kind, vybe_compiler::ast::StmtKind::ClassDecl { name: declared, .. } if declared == name
        )), "builtin {name} omitted from bootstrap");
    }
}
