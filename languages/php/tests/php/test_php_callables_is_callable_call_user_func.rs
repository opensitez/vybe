use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Callables & Dynamic Invocations — is_callable, call_user_func, call_user_func_array, forward_static_call, dynamic function/method calls
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_call_user_func_array_named_parameters() {
    let out = run_prints(
        r#"<?php
function makeGreeting(string $name, string $greeting = "Hello") {
    return "$greeting, $name!";
}

$result = call_user_func_array("makeGreeting", ["name" => "Alice", "greeting" => "Welcome"]);
echo $result;
"#,
    );
    assert_eq!(out, vec!["Welcome, Alice!"]);
}

#[test]
fn test_php_is_callable_string_array_object_syntax() {
    let out = run_prints(
        r#"<?php
class Calculator {
    public function add(int $a, int $b): int { return $a + $b; }
    public static function multiply(int $a, int $b): int { return $a * $b; }
}

$c = new Calculator();

echo is_callable("strlen") ? "1" : "0";
echo is_callable([$c, "add"]) ? "1" : "0";
echo is_callable([Calculator::class, "multiply"]) ? "1" : "0";
echo is_callable("Calculator::multiply") ? "1" : "0";
"#,
    );
    assert_eq!(out, vec!["1111"]);
}

#[test]
fn test_php_dynamic_method_name_invocation() {
    let out = run_prints(
        r#"<?php
class ActionHandler {
    public function handleInit(): string { return "INIT_DONE"; }
    public function handleRun(): string { return "RUN_DONE"; }
}

$handler = new ActionHandler();
$action = "Run";
$method = "handle" . $action;

echo $handler->$method();
"#,
    );
    assert_eq!(out, vec!["RUN_DONE"]);
}

#[test]
fn test_php_invokable_object_call_user_func() {
    let out = run_prints(
        r#"<?php
class InvokableTransformer {
    public function __invoke(string $text): string {
        return str_rot13($text);
    }
}

$transformer = new InvokableTransformer();
echo call_user_func($transformer, "Hello");
"#,
    );
    assert_eq!(out, vec!["Uryyb"]);
}

#[test]
fn test_php_dynamic_function_name_invocation() {
    compile_ok(
        r#"<?php
$fnName = "strtoupper";
echo $fnName("dynamic function call");
"#,
    );
}

#[test]
fn test_php_forward_static_call_array() {
    compile_ok(
        r#"<?php
class BaseFactory {
    public static function create(string $type) {
        return "BaseFactory:$type";
    }
}

class CustomFactory extends BaseFactory {
    public static function create(string $type) {
        return forward_static_call_array([BaseFactory::class, "create"], [$type]);
    }
}

echo CustomFactory::create("widget");
"#,
    );
}

#[test]
fn test_php_is_callable_callable_name_out_param() {
    compile_ok(
        r#"<?php
$callableName = "";
$check = is_callable([stdClass::class, "nonExistent"], syntax_only: true, callable_name: $callableName);
echo "Check=$check Name=$callableName";
"#,
    );
}

#[test]
fn test_php_callable_type_hint_validation() {
    compile_ok(
        r#"<?php
function executeCallback(callable $cb, mixed ...$args) {
    return $cb(...$args);
}

echo executeCallback(fn($x, $y) => $x * $y, 6, 7);
"#,
    );
}

#[test]
fn test_php_first_class_callable_syntax_is_callable() {
    compile_ok(
        r#"<?php
$c = strlen(...);
echo is_callable($c) ? "IS_CALLABLE" : "FAIL";
"#,
    );
}

#[test]
fn test_php_call_user_func_anonymous_closure() {
    compile_ok(
        r#"<?php
$closure = function($a, $b) { return $a - $b; };
echo call_user_func($closure, 100, 30);
"#,
    );
}
