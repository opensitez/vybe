use vybe_compiler::dynamic::RuntimeCompilerService;
use vybe_runtime::{VM,Value};

fn configured_vm() -> VM {
    vybe_language_php::register();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms(
        &mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

#[test]
fn class_constants_resolve_forward_dependencies() {
    let mut vm = configured_vm();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
        "<?php namespace ConstScope; class Defaults { const DATA=['agent'=>'prefix/'.self::VERSION]; const VERSION=self::BASE; const BASE='2.0.9'; public static function read() { return self::DATA['agent']; } } return Defaults::read();",
        language, "<forward-class-constants>.php").unwrap();
    assert!(matches!(&result, Value::String(s) if s.as_ref()=="prefix/2.0.9"), "{result:?}");
}

#[test]
fn unreachable_constant_back_reference_preserves_branch_semantics() {
    let mut vm = configured_vm();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
        "<?php class ConditionalConstant { const A=true ? 'ok' : self::A; } return ConditionalConstant::A;",
        language, "<conditional-class-constants>.php").unwrap();
    assert!(matches!(&result, Value::String(s) if s.as_ref()=="ok"), "{result:?}");
}

#[test]
fn wordpress_requests_class_initializes_without_autoload_scope() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/webroot/wordpress/wp-includes/Requests/src/Requests.php");
    let mut vm = configured_vm();
    RuntimeCompilerService::new(&mut vm).compile_and_run_path(&path).unwrap();
    let mut vm = configured_vm();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let source = format!("<?php function loadRequests() {{ require '{}'; }} loadRequests(); return \\WpOrg\\Requests\\Requests::OPTION_DEFAULTS['useragent'];", path.display());
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
        &source, language, "<included-requests>.php").unwrap();
    assert!(matches!(&result, Value::String(s) if s.as_ref()=="php-requests/2.0.9"), "{result:?}");
}

#[test]
fn superglobal_reads_initialize_declared_storage() {
    let mut vm = configured_vm();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    for name in ["_SERVER", "_GET", "_POST", "_COOKIE", "_REQUEST", "_FILES", "_ENV"] {
        let source = format!("<?php return is_array(${name});");
        let value = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
            &source, language, &format!("<superglobal-{name}>.php")).unwrap();
        assert!(matches!(value,Value::Bool(true)), "${name}: {value:?}");
    }
    let value = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
        "<?php function check() { global $_SERVER; return is_array($_SERVER) && count(array_merge(['PHP_SELF'=>''], $_SERVER)) > 0; } return check();",
        language, "<superglobal-function>.php").unwrap();
    assert!(matches!(value,Value::Bool(true)), "function global $_SERVER: {value:?}");
}

#[test]
fn included_units_preserve_mutated_superglobals() {
    let base = std::env::temp_dir().join(format!("vybe-superglobal-{}",uuid::Uuid::new_v4()));
    std::fs::create_dir_all(&base).unwrap();
    std::fs::write(base.join("main.php"),
        "<?php function writeServer() { global $_SERVER; $_SERVER['custom']='kept'; } writeServer(); require __DIR__ . '/child.php'; return $_SERVER['custom'] . ':' . $_SERVER['child'];").unwrap();
    std::fs::write(base.join("child.php"),
        "<?php if (!is_array($_SERVER)) { throw new Exception('missing server'); } function addChild() { global $_SERVER; $_SERVER['child']='set'; } addChild();").unwrap();
    let mut vm = configured_vm();
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_path(&base.join("main.php"));
    std::fs::remove_dir_all(&base).unwrap();
    let result = result.unwrap();
    assert!(matches!(&result, Value::String(s) if s.as_ref()=="kept:set"), "{result:?}");
}

#[test]
fn static_properties_remain_visible_in_closed_function_scopes() {
    let mut vm = configured_vm();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
        "<?php class Tables { public static $prefixes = ['key'=>'present']; public static function answer() { return 42; } } function readTables() { return Tables::$prefixes['key']; } return Tables::$prefixes['key'] . ':' . readTables() . ':' . Tables::answer();",
        language, "<static-property-scope>.php").unwrap();
    assert!(matches!(&result, Value::String(s) if s.as_ref()=="present:present:42"), "{result:?}");
}

#[test]
fn static_array_properties_preserve_array_shape() {
    let mut vm = configured_vm();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
        "<?php class A { public static $map = ['key'=>'value']; } return count(A::$map) . '|' . A::$map['key'];",
        language, "<static-array-shape>.php").unwrap();
    assert!(matches!(&result, Value::String(s) if s.as_ref()=="1|value"), "{result:?}");
}

#[test]
fn parameters_preserve_object_identity_and_array_value_semantics() {
    let mut vm = configured_vm();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_source(
        "<?php class Handle { public $value=1; } function mutateHandle($h) { $h->value=2; } function mutateArray($a) { $a['key']='changed'; return $a['key']; } $h=new Handle(); mutateHandle($h); $a=['key'=>'kept']; $changed=mutateArray($a); return ($h->value === 2) && ($changed === 'changed') && ($a['key'] === 'kept');",
        language, "<parameter-storage>.php").unwrap();
    assert!(matches!(result, Value::Bool(true)), "{result:?}");
}

#[test]
fn included_class_static_data_is_published_outside_caller_variable_scope() {
    let base = std::env::temp_dir().join(format!("vybe-class-scope-{}", uuid::Uuid::new_v4()));
    std::fs::create_dir_all(&base).unwrap();
    std::fs::write(base.join("tables.php"),
        "<?php class IncludedTables { public static $prefixes = ['key'=>'present']; }").unwrap();
    std::fs::write(base.join("main.php"),
        "<?php function loadTables() { $private='caller'; require __DIR__ . '/tables.php'; return IncludedTables::$prefixes['key'] . ':' . $private; } $inside=loadTables(); return $inside . ':' . IncludedTables::$prefixes['key'];").unwrap();
    let mut vm = configured_vm();
    let result = RuntimeCompilerService::new(&mut vm).compile_and_run_path(&base.join("main.php"));
    std::fs::remove_dir_all(&base).unwrap();
    let result = result.unwrap();
    assert!(matches!(&result, Value::String(s) if s.as_ref()=="present:caller:present"), "{result:?}");
}
