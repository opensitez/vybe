use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Variable Debugging & Inspection — var_dump, var_export, __debugInfo, debug_backtrace, debug_print_backtrace
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_var_export_valid_evaluatable_code() {
    let out = run_prints(
        r#"<?php
$data = ["a" => 1, "b" => [2, 3]];
$code = var_export($data, return: true);
eval('$restored = ' . $code . ';');
echo "a={$restored['a']} b0={$restored['b'][0]}";
"#,
    );
    assert_eq!(out, vec!["a=1 b0=2"]);
}

#[test]
fn test_php_debug_info_magic_method_custom_dump() {
    let out = run_prints(
        r#"<?php
class UserAccount {
    public function __construct(
        public string $username,
        private string $passwordHash
    ) {}

    public function __debugInfo(): array {
        return [
            "username" => $this->username,
            "passwordHash" => "********" // Mask sensitive data
        ];
    }
}

$user = new UserAccount("alice", "secret_hash");
$exported = var_export($user->__debugInfo(), return: true);
echo str_contains($exported, "********") ? "PASSWORD_MASKED" : "UNMASKED";
"#,
    );
    assert_eq!(out, vec!["PASSWORD_MASKED"]);
}

#[test]
fn test_php_debug_backtrace_stack_inspection() {
    let out = run_prints(
        r#"<?php
function foo($a, $b) { bar($a + $b); }
function bar($c) { baz($c); }
function baz($d) {
    $trace = debug_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS);
    $funcs = array_column($trace, "function");
    echo implode(" <- ", $funcs);
}

foo(10, 20);
"#,
    );
    assert_eq!(out, vec!["baz <- bar <- foo"]);
}

#[test]
fn test_php_var_dump_output_buffering_capture() {
    compile_ok(
        r#"<?php
ob_start();
$val = ["hello", 123, true, null];
var_dump($val);
$dump = ob_get_clean();

echo str_contains($dump, "string(5)") && str_contains($dump, "int(123)") ? "VAR_DUMP_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_debug_print_backtrace_output_capture() {
    compile_ok(
        r#"<?php
function traceTest() {
    ob_start();
    debug_print_backtrace();
    $out = ob_get_clean();
    echo str_contains($out, "traceTest") ? "TRACE_PRINT_OK" : "FAIL";
}
traceTest();
"#,
    );
}

#[test]
fn test_php_var_export_stdclass_export() {
    compile_ok(
        r#"<?php
$obj = new stdClass();
$obj->title = "Test";
$exported = var_export($obj, return: true);
echo str_contains($exported, "stdClass::__set_state") ? "SET_STATE_EXPORT" : "ANON_EXPORT";
"#,
    );
}

#[test]
fn test_php_debug_backtrace_limit_parameter() {
    compile_ok(
        r#"<?php
function level3() { return debug_backtrace(limit: 2); }
function level2() { return level3(); }
function level1() { return level2(); }

$frames = level1();
echo count($frames) <= 2 ? "LIMIT_2_OK" : "LIMIT_EXCEEDED";
"#,
    );
}

#[test]
fn test_php_var_dump_object_references_recursion() {
    compile_ok(
        r#"<?php
$node1 = new stdClass();
$node2 = new stdClass();
$node1->next = $node2;
$node2->prev = $node1; // Circular reference

ob_start();
var_dump($node1);
$dump = ob_get_clean();
echo str_contains($dump, "*RECURSION*") ? "RECURSION_DETECTED" : "DUMP_OK";
"#,
    );
}

#[test]
fn test_php_var_export_resource_behavior() {
    compile_ok(
        r#"<?php
$fp = fopen("php://memory", "r");
$exp = var_export($fp, return: true);
fclose($fp);
echo str_contains($exp, "NULL") ? "RESOURCE_EXPORT_NULL" : "EXPORT_OK";
"#,
    );
}

#[test]
fn test_php_debug_backtrace_provide_object_option() {
    compile_ok(
        r#"<?php
class Inspector {
    public function inspect() {
        return debug_backtrace(DEBUG_BACKTRACE_PROVIDE_OBJECT);
    }
}

$i = new Inspector();
$trace = $i->inspect();
echo isset($trace[0]["object"]) && $trace[0]["object"] === $i ? "OBJECT_PROVIDED" : "FAIL";
"#,
    );
}
