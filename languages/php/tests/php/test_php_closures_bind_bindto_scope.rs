use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Closure Binding & Scoping — Closure::bind, bindTo, Closure::fromCallable, binding to object/scope
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_closure_bindto_private_property_access() {
    let out = run_prints(
        r#"<?php
class SecretHolder {
    private string $secret = "top_secret_code";
}

$getter = function() {
    return $this->secret;
};

$sh = new SecretHolder();
$boundGetter = $getter->bindTo($sh, SecretHolder::class);
echo $boundGetter();
"#,
    );
    assert_eq!(out, vec!["top_secret_code"]);
}

#[test]
fn test_php_closure_static_bind_class_scope() {
    let out = run_prints(
        r#"<?php
class InternalConfig {
    private static string $key = "internal_api_key";
}

$staticGetter = Closure::bind(function() {
    return InternalConfig::$key;
}, null, InternalConfig::class);

echo $staticGetter();
"#,
    );
    assert_eq!(out, vec!["internal_api_key"]);
}

#[test]
fn test_php81_closure_from_callable_factory() {
    let out = run_prints(
        r#"<?php
class StrHelper {
    public function convert(string $s): string {
        return strtoupper($s);
    }
}

$helper = new StrHelper();
$closure = Closure::fromCallable([$helper, "convert"]);
echo $closure("hello closure");
"#,
    );
    assert_eq!(out, vec!["HELLO CLOSURE"]);
}

#[test]
fn test_php_closure_call_immediate_execution_php7() {
    let out = run_prints(
        r#"<?php
class Person {
    private string $name = "Charlie";
}

$p = new Person();
$greeting = (function(string $prefix) {
    return "$prefix {$this->name}";
})->call($p, "Hello");

echo $greeting;
"#,
    );
    assert_eq!(out, vec!["Hello Charlie"]);
}

#[test]
fn test_php_closure_bindto_null_this_unbinding() {
    compile_ok(
        r#"<?php
class User {
    public function getClosure() {
        return function() { return $this; };
    }
}

$u = new User();
$fn = $u->getClosure();
$unbound = $fn->bindTo(null, null);
echo is_null($unbound()) ? "UNBOUND" : "BOUND";
"#,
    );
}

#[test]
fn test_php_closure_from_callable_function_name() {
    compile_ok(
        r#"<?php
$c = Closure::fromCallable("strlen");
echo $c("test");
"#,
    );
}

#[test]
fn test_php_closure_bind_static_closure_error() {
    compile_ok(
        r#"<?php
$staticFn = static function() { return "static"; };
$bound = @$staticFn->bindTo(new stdClass());
"#,
    );
}

#[test]
fn test_php_closure_reflection_introspection() {
    compile_ok(
        r#"<?php
$fn = function(int $a, string $b = "default"): string {
    return $b . $a;
};

$rc = new ReflectionFunction($fn);
echo count($rc->getParameters());
"#,
    );
}

#[test]
fn test_php_closure_bindto_subclass_scope() {
    compile_ok(
        r#"<?php
class ParentScope {
    protected string $prot = "protected_val";
}
class ChildScope extends ParentScope {}

$fn = function() { return $this->prot; };
$child = new ChildScope();
$bound = $fn->bindTo($child, ChildScope::class);
echo $bound();
"#,
    );
}

#[test]
fn test_php_closure_returning_closure() {
    compile_ok(
        r#"<?php
function makeAdder(int $x) {
    return fn(int $y) => $x + $y;
}

$add5 = makeAdder(5);
echo $add5(10);
"#,
    );
}

#[test]
fn test_php_closure_reference_capture_increments_shared_state() {
    let out = run_prints(
        r#"<?php
$count = 1;
$inc = function() use (&$count) { $count += 2; return $count; };
echo $inc();
echo $inc();
"#,
    );
    assert_eq!(out, vec!["3", "5"]);
}

#[test]
fn test_php_closure_value_capture_uses_snapshot_not_live_value() {
    let out = run_prints(
        r#"<?php
$x = 1;
$snap = function() use ($x) { return $x; };
$x = 9;
echo $snap();
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_php_closure_use_global_access_inside_nested_scope() {
    let out = run_prints(
        r#"<?php
$g = 'outside';
$fn = function() {
    global $g;
    return $g;
};
$g = 'inside';
echo $fn();
"#,
    );
    assert_eq!(out, vec!["inside"]);
}

#[test]
fn test_php_closure_static_forbids_this() {
    let out = run_prints(
        r#"<?php
$fn = static function() { return isset($this); };
echo $fn();
"#,
    );
    assert_eq!(out, vec![""]);
}

#[test]
fn test_php_closure_bindto_private_method_access() {
    let out = run_prints(
        r#"<?php
class Vault {
    private function secret(): string { return "locked"; }
}
$vault = new Vault();
$fn = function() { return $this->secret(); };
$bound = $fn->bindTo($vault, Vault::class);
echo $bound();
"#,
    );
    assert_eq!(out, vec!["locked"]);
}

#[test]
fn test_php_closure_bind_to_namespace_scope() {
    let out = run_prints(
        r#"<?php
namespace ScopeDemo {
    class ScopeClass {
        private const TOKEN = 'ns-token';
        public function token(): string {
            $fn = function() { return self::TOKEN; };
            return $fn->call($this);
        }
    }
    $obj = new ScopeClass();
    echo $obj->token();
}
"#,
    );
    assert_eq!(out, vec!["ns-token"]);
}

#[test]
fn test_php_closure_call_with_bound_parameters_in_runtime() {
    let out = run_prints(
        r#"<?php
function scale(int $x, int $factor): int { return $x * $factor; }
$closure = function(int $x, int $factor) { return $x * $factor; };
$bound = $closure->bindTo(null, null);
echo $bound(2, 4);
"#,
    );
    assert_eq!(out, vec!["8"]);
}
