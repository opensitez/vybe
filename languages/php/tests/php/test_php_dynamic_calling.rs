use super::helpers::run_prints;

fn assert_int(expr: &str, expected: i64) {
    assert_eq!(
        run_prints(&format!("<?php echo {}; ", expr)),
        vec![expected.to_string()]
    );
}

#[test]
fn php_dynamic_calling() {
    for n in 1..=10_i64 {
        let doubled = n * 2;
        let method_result = n * 2;

        assert_int(
            &format!(
                "function dynamic_fn_{n}() {{ return {n}; }}\n$fn = 'dynamic_fn_{n}';\necho $fn();"
            ),
            n,
        );
        assert_int(
            &format!(
                "class DynamicTarget{n} {{ public function value(): int {{ return {n}; }} public function add(int $v): int {{ return $v + {n}; }} }}\n$cls = 'DynamicTarget{n}';\n$obj = new $cls();\n$method = 'add';\necho $obj->$method({n});"
            ),
            doubled,
        );
        assert_int(
            &format!("strlen(call_user_func('trim', '  x{n}  ', ' '));"),
            2 + 1 + n,
        );
        assert_int(
            &format!(
                "class DynamicCaller{n} {{ public static function build(int $n): int {{ return $n * {n}; }} }}\n$call = ['DynamicCaller{n}', 'build'];\necho call_user_func($call, 2);"
            ),
            n * 2,
        );
        assert_int(
            &format!(
                "$target = 'DynamicTarget{n}';\n$method = 'value';\n$reflection = new $target();\necho $reflection->$method();"
            ),
            n,
        );
        assert_int(
            &format!(
                " $obj = new DynamicTarget{n}(); echo call_user_func_array([$obj, 'add'], [{n}]);"
            ),
            method_result,
        );
    }
}

#[test]
fn php_dynamic_calling_variable_methods_and_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynamicAccess {
    public int $answer = 42;
    public function answer(): int { return 42; }
    public function plus(int $v): int { return $v + 1; }
}
$obj = new DynamicAccess();
$method = "answer";
echo $obj->$method();
echo $obj->{"plus"}(4);
"#
        ),
        vec!["421"]
    );

    assert_eq!(
        run_prints(
            r#"<?php
$field = "answer";
$obj = new DynamicAccess();
echo $obj->$field;
"#
        ),
        vec!["42"]
    );
}

#[test]
fn php_dynamic_calling_static_targets() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynamicStatic {
    public static function value(int $base, int $delta): int { return $base + $delta; }
}
$method = 'value';
$out = DynamicStatic::$method(10, 5);
echo $out;
"#
        ),
        vec!["15"]
    );

    assert_eq!(
        run_prints(
            r#"<?php
class DynamicFactory {
    public static function mk(int $v): int { return $v * 2; }
}
$callable = ['DynamicFactory', 'mk'];
echo call_user_func($callable, 7);
"#
        ),
        vec!["14"]
    );
}

#[test]
fn php_dynamic_calling_call_user_func_variants() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = fn(int $x): int => $x * 3;
echo call_user_func($fn, 4);
"#
        ),
        vec!["12"]
    );

    assert_eq!(
        run_prints(
            r#"<?php
class Worker {
    public function double(int $x): int { return $x * 2; }
}
$obj = new Worker();
echo call_user_func([$obj, 'double'], 6);
"#
        ),
        vec!["12"]
    );

    assert_eq!(
        run_prints(
            r#"<?php
$sum = function(int $a, int $b): int { return $a + $b; };
$cb = $sum;
echo call_user_func_array($cb, [3, 4]);
"#
        ),
        vec!["7"]
    );
}

#[test]
fn php_dynamic_calling_forwarded_callable_like() {
    assert_eq!(
        run_prints(
            r#"<?php
class Invokable {
    public function __invoke(string $name): string { return strtoupper($name); }
}
$inv = new Invokable();
echo call_user_func($inv, 'php');
"#
        ),
        vec!["PHP"]
    );

    assert_eq!(
        run_prints(
            r#"<?php
class Invokable2 {
    public function __invoke(int $x): int { return $x + 1; }
}
$inv = fn(Invokable2 $target, int $x): int => $target($x);
$o = new Invokable2();
echo $inv($o, 8);
"#
        ),
        vec!["9"]
    );
}

#[test]
fn php_dynamic_calling_callable_array_and_fallback() {
    assert_eq!(
        run_prints(
            r#"<?php
class CallableHolder {
    public static function marker(): string { return 'static'; }
}
$callable = ['CallableHolder', 'marker'];
echo is_callable($callable) ? 'yes' : 'no';
echo '|';
echo call_user_func($callable);
"#
        ),
        vec!["yes|static"]
    );
}

#[test]
fn php_dynamic_calling_variable_method_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class ChainNode {
    public function level(int $n): string {
        $method = 'suffix';
        return $this->$method($n);
    }
    public function suffix(int $n): string { return 'v'.$n; }
}
$obj = new ChainNode();
echo $obj->{"level"}(9);
"#
        ),
        vec!["v9"]
    );
}

#[test]
fn php_dynamic_calling_magic_call_with_variadics() {
    assert_eq!(
        run_prints(
            r#"<?php
class Handler {
    public function __call(string $name, array $args): mixed {
        if ($name === 'sum') {
            return array_sum($args);
        }
        return null;
    }
}
$handler = new Handler();
$method = 'sum';
echo $handler->$method(1, 2, 3);
"#
        ),
        vec!["6"]
    );
}

#[test]
fn php_dynamic_calling_callable_string_with_namespace() {
    assert_eq!(
        run_prints(
            r#"<?php
namespace DynamicTestNs;

function ns_dynamic_target(): string { return 'ns'; }
echo call_user_func(__NAMESPACE__ . '\\\\ns_dynamic_target');
"#
        ),
        vec!["ns"]
    );
}

#[test]
fn php_dynamic_calling_array_shift_calling_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
class Shifter {
    public function inc(int $n): int { return $n + 1; }
}
$target = [new Shifter(), 'inc'];
$first = array_shift($target);
echo is_object($first) ? 'obj' : 'no';
echo is_string($target[0]) ? 'fn' : 'bad';
"#
        ),
        vec!["objfn"]
    );
}

#[test]
fn php_dynamic_calling_static_member_via_variable_name() {
    assert_eq!(
        run_prints(
            r#"<?php
class ServiceFactory {
    public static function create(int $id): int { return $id + 100; }
}
$klass = ServiceFactory::class;
$method = 'create';
echo $klass::$method(7);
"#
        ),
        vec!["107"]
    );
}

#[test]
fn php_dynamic_calling_nested_variable_method_name() {
    assert_eq!(
        run_prints(
            r#"<?php
class Router {
    public function route(string $path): string {
        return 'route:' . $path;
    }
}
$route_call = ['route'];
$obj = new Router();
$name = $route_call[0];
echo $obj->$name('home');
"#
        ),
        vec!["route:home"]
    );
}

#[test]
fn php_dynamic_calling_callable_check_and_dispatch() {
    assert_eq!(
        run_prints(
            r#"<?php
class Handler {
    public static function enabled(): string { return 'enabled'; }
}
$target = Handler::class;
$method = 'enabled';
echo is_callable([$target, $method]) ? 'ok' : 'no';
echo '|';
echo call_user_func([$target, $method]);
"#
        ),
        vec!["ok|enabled"]
    );
}

#[test]
fn php_dynamic_calling_callable_object_or_string_mix() {
    assert_eq!(
        run_prints(
            r#"<?php
$suffix = fn(string $s): string => $s . '!';
echo is_callable($suffix) ? 'func' : 'no';
echo '|';
echo $suffix('ok');
echo '|';
echo is_callable('strlen') ? 'strlen' : 'non';
"#
        ),
        vec!["func|ok!|strlen"]
    );
}

#[test]
fn php_dynamic_calling_class_string_from_namespace() {
    assert_eq!(
        run_prints(
            r#"<?php
namespace DynNs;

class Api {
    public static function ping(string $name): string { return 'pong:' . $name; }
}

$fqn = __NAMESPACE__ . '\\\\Api';
echo (new $fqn())->ping('x');
"#
        ),
        vec!["pong:x"]
    );
}

#[test]
fn php_dynamic_calling_constructor_from_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynamicCtor {
    public function __construct(public int $n) {}
    public function value(): int { return $this->n * 2; }
}
$klass = 'DynamicCtor';
$obj = new $klass(4);
echo $obj->value();
"#
        ),
        vec!["8"]
    );
}

#[test]
fn php_dynamic_calling_callable_from_fqn_string() {
    assert_eq!(
        run_prints(
            r#"<?php
namespace DynCallNs {
    class Factory {
        public static function create(string $label): string {
            return 'created:' . $label;
        }
    }
    $handler = __NAMESPACE__ . '\\Factory';
    echo is_callable([$handler, 'create']) ? 'yes' : 'no';
    echo '|';
    echo call_user_func([$handler, 'create'], 'item');
}
"#
        ),
        vec!["yes|created:item"]
    );
}

#[test]
fn php_dynamic_calling_call_user_func_array_runtime_unpacking() {
    assert_eq!(
        run_prints(
            r#"<?php
function combine(string $a, string $b, string $c): string {
    return $a . '-' . $b . '-' . $c;
}
$fn = 'combine';
echo call_user_func_array($fn, ['a', 'b', 'c']);
"#
        ),
        vec!["a-b-c"]
    );
}

#[test]
fn php_dynamic_calling_nested_callable_arrays() {
    assert_eq!(
        run_prints(
            r#"<?php
class Stepper {
    public function step(int $n): int { return $n + 1; }
}

function trampoline(callable $cb, int $value): int {
    return $cb($value);
}

$target = [new Stepper(), 'step'];
$callback = fn(int $n): int => trampoline($target, $n);
echo $callback(5);
"#
        ),
        vec!["6"]
    );
}

#[test]
fn php_dynamic_calling_callable_as_array_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$callables = [
    'twice' => function(int $v): int { return $v * 2; },
    'sum' => function(int $a, int $b): int { return $a + $b; },
];
echo $callables['twice'](4);
echo '|';
echo call_user_func_array($callables['sum'], [3, 5]);
"#
        ),
        vec!["8|8"]
    );
}

#[test]
fn php_dynamic_calling_property_callable_on_closure_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class Holder {
    public \Closure $cb;
    public function __construct() {
        $this->cb = function(string $label): string {
            return strtoupper($label);
        };
    }
}

$h = new Holder();
$method = 'cb';
echo $h->$method('ok');
"#
        ),
        vec!["OK"]
    );
}

#[test]
fn php_dynamic_calling_string_callable_in_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$inputs = [' a ', ' b '];
$labels = array_map('trim', $inputs);
echo $labels[0];
echo '|';
echo array_map('strlen', $labels)[1];
"#
        ),
        vec!["a|1"]
    );
}

#[test]
fn php_dynamic_calling_static_scope_callable_string() {
    assert_eq!(
        run_prints(
            r#"<?php
namespace DynStaticCallNs;
function scoped(): string { return 'scoped'; }
echo call_user_func(__NAMESPACE__ . '\\\\scoped');
echo '|';
echo is_callable(__NAMESPACE__ . '\\\\scoped') ? 'callable' : 'no';
"#
        ),
        vec!["scoped|callable"]
    );
}

#[test]
fn php_dynamic_calling_property_held_callable_invocation() {
    assert_eq!(
        run_prints(
            r#"<?php
class HandlerWithCallback {
    public \Closure $cb;
    public function __construct() {
        $this->cb = fn(int $a, int $b): int => $a + $b;
    }
}

$h = new HandlerWithCallback();
$fn = $h->cb;
echo $fn(3, 4);
"#
        ),
        vec!["7"]
    );
}

#[test]
fn php_dynamic_calling_class_string_and_method_string_with_call_user_func() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynamicMethodCarrier {
    public static function fromStatic(string $value): string { return 'static-' . $value; }
}
$class = DynamicMethodCarrier::class;
$method = 'fromStatic';
echo call_user_func([$class, $method], 'ok');
"#
        ),
        vec!["static-ok"]
    );
}

#[test]
fn php_dynamic_calling_method_name_from_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class MethodCarrier {
    public string $method = 'compute';
    public function compute(int $n): int { return $n * 3; }
}
$obj = new MethodCarrier();
$name = $obj->method;
echo $obj->$name(5);
"#
        ),
        vec!["15"]
    );
}

#[test]
fn php_dynamic_calling_instance_callable_property_on_trait() {
    assert_eq!(
        run_prints(
            r#"<?php
trait CallableTrait {
    public \Closure $formatter;
    public function __construct() {
        $this->formatter = fn(string $value): string => strtoupper($value);
    }
}

class CallableTarget {
    use CallableTrait;
    public function formatted(string $value): string {
        $method = 'formatter';
        return ($this->$method)($value);
    }
}

$obj = new CallableTarget();
echo $obj->formatted('ok');
"#,
        ),
        vec!["OK"]
    );
}

#[test]
fn php_dynamic_calling_object_property_storing_static_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
class StaticCarrier {
    public \Closure|string $dispatcher;
    public function __construct() {
        $this->dispatcher = [self::class, 'make'];
    }
    public static function make(int $n): int { return $n + 10; }
    public function run(int $n): int {
        $callable = $this->dispatcher;
        return $callable($n);
    }
}

$obj = new StaticCarrier();
echo $obj->run(5);
"#,
        ),
        vec!["15"]
    );
}

#[test]
fn php_dynamic_calling_call_user_func_array_with_nulls() {
    assert_eq!(
        run_prints(
            r#"<?php
function combine_three(string $a, ?string $b = null, string $c = 'c'): string {
    return $a . '|' . ($b ?? 'none') . '|' . $c;
}

echo call_user_func_array('combine_three', ['a', null, 'z']);
"#,
        ),
        vec!["a|none|z"]
    );
}

#[test]
fn php_dynamic_calling_invokable_object_in_static_context() {
    assert_eq!(
        run_prints(
            r#"<?php
class Invokable {
    public function __invoke(string $name): string { return 'hi:' . $name; }
}
class Caller {
    public static function execute(callable $cb, string $label): string {
        return $cb($label);
    }
}
$obj = new Invokable();
echo Caller::execute($obj, 'php');
"#,
        ),
        vec!["hi:php"]
    );
}

#[test]
fn php_dynamic_calling_callable_type_guard() {
    assert_eq!(
        run_prints(
            r#"<?php
function sink(callable $cb): string {
    return is_callable($cb) ? 'callable' : 'no';
}
echo sink('trim');
echo '|';
echo sink(['StdClass', 'foo']);
"#,
        ),
        vec!["callable|no"]
    );
}

#[test]
fn php_dynamic_calling_variable_function_alias_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
function source(): string { return 'base'; }
function transform(string $s): string { return strtoupper($s); }
$step1 = 'source';
$step2 = 'transform';
echo $step2($step1());
"#,
        ),
        vec!["BASE"]
    );
}

#[test]
fn php_dynamic_calling_call_result_in_arithmetic() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynArithmetic {
    public function base(int $n): int { return $n + 2; }
}
$obj = new DynArithmetic();
$method = 'base';
echo $obj->$method(3) + $obj->$method(4);
"#,
        ),
        vec!["11"]
    );
}

#[test]
fn php_dynamic_calling_method_name_from_array_value() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynResolver {
    public function combine(string $a, string $b): string {
        return $a . '-' . $b;
    }
}

$obj = new DynResolver();
$call = ['name' => 'combine'];
echo $obj->{$call['name']}('x', 'y');
"#,
        ),
        vec!["x-y"]
    );
}

#[test]
fn php_dynamic_calling_callable_array_with_closure_and_string_target() {
    assert_eq!(
        run_prints(
            r#"<?php
function format_label(string $value): string {
    return 'L:' . $value;
}

$callables = [
    'format' => 'format_label',
    'twice' => function(string $value): string { return $value . $value; },
];

echo call_user_func($callables['format'], 'ok');
echo '|';
echo $callables['twice']('ha');
"#,
        ),
        vec!["L:ok|haha"]
    );
}

#[test]
fn php_dynamic_calling_ternary_function_name_resolution() {
    assert_eq!(
        run_prints(
            r#"<?php
$transform = true ? 'strtoupper' : 'strtolower';
echo $transform('php');
"#,
        ),
        vec!["PHP"]
    );
}

#[test]
fn php_dynamic_calling_static_call_with_variable_membership() {
    assert_eq!(
        run_prints(
            r#"<?php
namespace DynAdvanced {
    class Service {
        public static function render(int $n): string {
            return 'v' . $n;
        }
    }

    $klass = __NAMESPACE__ . '\\\\Service';
    $method = 'render';
    echo $klass::$method(9);
}
"#,
        ),
        vec!["v9"]
    );
}

#[test]
fn php_dynamic_calling_callable_chain_and_truthiness() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynBoolCarrier {
    public function value(int $n): int { return $n; }
}

$obj = new DynBoolCarrier();
$method = $obj->value(0) ? 'missing' : 'value';
echo is_string($method) ? 'str' : 'no';
echo '|';
echo $obj->$method(2);
"#,
        ),
        vec!["str|2"]
    );
}
