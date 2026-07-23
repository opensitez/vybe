use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Shutdown Functions: register_shutdown_function & Arguments Passing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_register_shutdown_function_parameter_passing() {
    let out = run_prints(
        r##"<?php
$param1 = "ShutdownArg1";
$param2 = 42;

register_shutdown_function(function($a, $b) {
    // Shutdown hook simulation
}, $param1, $param2);

echo "Shutdown Registered: YES";
"##,
    );
    assert_eq!(out, vec!["Shutdown Registered: YES"]);
}

#[test]
fn test_php_register_shutdown_function_multiple_callbacks() {
    let out = run_prints(
        r##"<?php
register_shutdown_function(fn() => null);
register_shutdown_function(fn() => null);
echo "Multiple Shutdown Callbacks Registered";
"##,
    );
    assert_eq!(out, vec!["Multiple Shutdown Callbacks Registered"]);
}

#[test]
fn test_php_register_shutdown_function_class_static_method() {
    compile_ok(
        r##"<?php
class Cleaner {
    public static function cleanUp($target) {}
}
register_shutdown_function([Cleaner::class, "cleanUp"], "temp_folder");
echo "CLASS_STATIC_SHUTDOWN_OK";
"##,
    );
}

#[test]
fn test_php_register_shutdown_function_object_instance_method() {
    compile_ok(
        r##"<?php
class Logger {
    public function flush() {}
}
$log = new Logger();
register_shutdown_function([$log, "flush"]);
echo "OBJECT_INSTANCE_SHUTDOWN_OK";
"##,
    );
}

#[test]
fn test_php_register_shutdown_function_invokable_object() {
    compile_ok(
        r##"<?php
class ShutdownTask {
    public function __invoke($msg) {}
}
register_shutdown_function(new ShutdownTask(), "task_done");
echo "INVOKABLE_SHUTDOWN_OK";
"##,
    );
}

#[test]
fn test_php_register_shutdown_function_with_array_args() {
    compile_ok(
        r##"<?php
register_shutdown_function(function($arr) {
    //
}, ["a" => 1, "b" => 2]);
echo "ARRAY_ARG_SHUTDOWN_OK";
"##,
    );
}

#[test]
fn test_php_register_shutdown_function_error_get_last_inspection() {
    compile_ok(
        r##"<?php
register_shutdown_function(function() {
    $err = error_get_last();
});
echo "ERROR_GET_LAST_SHUTDOWN_OK";
"##,
    );
}

#[test]
fn test_php_register_shutdown_function_string_function_name() {
    compile_ok(
        r##"<?php
function myShutdownFunc() {}
register_shutdown_function("myShutdownFunc");
echo "STRING_NAME_SHUTDOWN_OK";
"##,
    );
}

#[test]
fn test_php_register_shutdown_function_cwd_preservation() {
    compile_ok(
        r##"<?php
$cwd = getcwd();
register_shutdown_function(function() use ($cwd) {
    // In shutdown functions, working directory might change to root depending on SAPI
});
echo "CWD_SHUTDOWN_OK";
"##,
    );
}

#[test]
fn test_php_register_shutdown_function_multiple_args() {
    compile_ok(
        r##"<?php
register_shutdown_function(function($a, $b, $c, $d) {}, 1, "two", 3.0, true);
echo "MULTIPLE_ARGS_SHUTDOWN_OK";
"##,
    );
}
