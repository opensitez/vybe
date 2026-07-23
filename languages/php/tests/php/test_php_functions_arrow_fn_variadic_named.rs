use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Arrow Functions & Advanced Callables — nested fn(), variadics with references, named args with variadics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_nested_arrow_functions_currying() {
    let out = run_prints(
        r#"<?php
$add = fn($x) => fn($y) => $x + $y;
$add5 = $add(5);
echo $add5(10);
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_php_variadic_parameter_by_reference() {
    let out = run_prints(
        r#"<?php
function doubleAll(&...$numbers) {
    foreach ($numbers as &$n) {
        $n *= 2;
    }
}

$a = 1; $b = 2; $c = 3;
doubleAll($a, $b, $c);
echo "$a-$b-$c";
"#,
    );
    assert_eq!(out, vec!["2-4-6"]);
}

#[test]
fn test_php_named_arguments_with_variadic_args() {
    let out = run_prints(
        r#"<?php
function logMessage(string $level, string ...$messages) {
    return "[$level] " . implode(" ", $messages);
}

echo logMessage(messages: ["System", "booted", "successfully"], level: "INFO");
"#,
    );
    assert_eq!(out, vec!["[INFO] System booted successfully"]);
}

#[test]
fn test_php_arrow_function_returning_array() {
    let out = run_prints(
        r#"<?php
$makePair = fn($key, $val) => [$key => $val];
$pair = $makePair("status", "ok");
echo "status=" . $pair["status"];
"#,
    );
    assert_eq!(out, vec!["status=ok"]);
}

#[test]
fn test_php_arrow_function_in_array_map_pipeline() {
    compile_ok(
        r#"<?php
$users = [
    ["name" => "Alice", "active" => true],
    ["name" => "Bob", "active" => false],
    ["name" => "Charlie", "active" => true],
];

$activeNames = array_map(
    fn($u) => $u["name"],
    array_filter($users, fn($u) => $u["active"])
);

echo implode(",", $activeNames);
"#,
    );
}

#[test]
fn test_php_arrow_function_by_ref_return_forbidden() {
    compile_ok(
        r#"<?php
$val = 100;
$getRef = fn&() => $val; // Arrow function return by reference syntax
"#,
    );
}

#[test]
fn test_php_variadic_type_hint_union_types() {
    compile_ok(
        r#"<?php
function stringify(int|float ...$values): array {
    return array_map(fn($v) => (string)$v, $values);
}

$res = stringify(1, 2.5, 3);
echo implode("-", $res);
"#,
    );
}

#[test]
fn test_php_anonymous_function_use_multiple_variables() {
    compile_ok(
        r#"<?php
$prefix = "LOG";
$suffix = "END";

$log = function(string $msg) use ($prefix, $suffix) {
    return "$prefix: $msg ($suffix)";
};

echo $log("Message body");
"#,
    );
}

#[test]
fn test_php_named_arguments_in_constructor_call() {
    compile_ok(
        r#"<?php
class ServerConfig {
    public function __construct(
        public string $host,
        public int $port = 80,
        public int $timeout = 30
    ) {}
}

$config = new ServerConfig(timeout: 60, host: "127.0.0.1");
echo "{$config->host}:{$config->port} t={$config->timeout}";
"#,
    );
}

#[test]
fn test_php_function_return_type_object_or_null() {
    compile_ok(
        r#"<?php
function findService(string $name): ?object {
    if ($name === "db") return new stdClass();
    return null;
}

echo is_object(findService("db")) ? "FOUND" : "NOT_FOUND";
"#,
    );
}
