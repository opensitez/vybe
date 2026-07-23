use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Functions, Arrow Functions & First-Class Callables — fn() => ..., strlen(...), variadics, named arguments
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_arrow_function_outer_scope_capture() {
    let out = run_prints(
        r#"<?php
$multiplier = 3;
$triple = fn($x) => $x * $multiplier;
echo $triple(5);
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_php81_first_class_callable_syntax() {
    let out = run_prints(
        r#"<?php
$fn = strlen(...);
echo $fn("Hello World");
"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_php_variadic_argument_unpacking() {
    let out = run_prints(
        r#"<?php
function sum(...$numbers) {
    return array_sum($numbers);
}
echo sum(10, 20, 30);
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_php_first_class_callable_method() {
    let out = run_prints(
        r#"<?php
class Calculator {
    public function add(int $a, int $b): int {
        return $a + $b;
    }
}

$calc = new Calculator();
$adder = $calc->add(...);
echo $adder(10, 20);
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_php_named_arguments_out_of_order() {
    let out = run_prints(
        r#"<?php
function formatUser(string $name, string $role = "guest", bool $active = true) {
    return "$name ($role) active=" . ($active ? "1" : "0");
}

echo formatUser(active: false, name: "Bob");
"#,
    );
    assert_eq!(out, vec!["Bob (guest) active=0"]);
}

#[test]
fn test_php_anonymous_function_use_by_reference() {
    let out = run_prints(
        r#"<?php
$counter = 0;
$inc = function() use (&$counter) {
    $counter++;
};
$inc();
$inc();
echo $counter;
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_php_static_anonymous_function() {
    compile_ok(
        r#"<?php
class Container {
    public function getClosure() {
        return static function() {
            return "no_this_binding";
        };
    }
}

$c = new Container();
$fn = $c->getClosure();
echo $fn();
"#,
    );
}

#[test]
fn test_php_first_class_callable_static_method() {
    compile_ok(
        r#"<?php
class Utils {
    public static function format(string $str): string {
        return strtoupper($str);
    }
}

$formatter = Utils::format(...);
echo $formatter("test");
"#,
    );
}

#[test]
fn test_php_variadic_parameter_with_type_hint() {
    compile_ok(
        r#"<?php
function concatenate(string $delim, string ...$words): string {
    return implode($delim, $words);
}

echo concatenate("-", "a", "b", "c");
"#,
    );
}

#[test]
fn test_php_named_arguments_array_unpacking() {
    compile_ok(
        r#"<?php
function configure(string $host, int $port = 8080, bool $ssl = false) {
    return "$host:$port ssl=" . ($ssl ? "yes" : "no");
}

$params = ["ssl" => true, "host" => "localhost"];
echo configure(...$params);
"#,
    );
}
