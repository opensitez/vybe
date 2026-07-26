//! Closure binding, callable validation, call_user_func guards, first-class callables, arrow $this.

crate::php_cases! {
    closure_bind_reads_private_property => {
        r#"<?php
class Vault { private int $secret = 99; }
$open = Closure::bind(function(): int { return $this->secret; }, new Vault(), Vault::class);
echo $open();
"#,
        ["99"]
    };

    closure_bind_wrong_object_scope_silent_miss => {
        r#"<?php
class Alpha { private int $n = 1; }
class Beta {}
$fn = Closure::bind(function() { return $this->n ?? 0; }, new Beta(), Alpha::class);
echo $fn === null ? 'null' : (string)$fn();
"#,
        ["0"]
    };

    closure_bind_static_with_null_object => {
        r#"<?php
class Registry { private static int $count = 7; }
$read = Closure::bind(static function(): int { return static::$count; }, null, Registry::class);
echo $read();
"#,
        ["7"]
    };

    closure_bind_static_closure_no_instance => {
        r#"<?php
class Config { private static string $env = 'prod'; }
$fn = static function(): string { return 'static'; };
$bound = Closure::bind($fn, null, Config::class);
echo $bound();
"#,
        ["static"]
    };

    invoking_bound_closure_after_rebind => {
        r#"<?php
class Node { private string $label; public function __construct(string $l) { $this->label = $l; } }
$get = function(): string { return $this->label; };
$a = Closure::bind($get, new Node('east'), Node::class);
$b = Closure::bind($get, new Node('west'), Node::class);
echo $a() . ',' . $b();
"#,
        ["east,west"]
    };

    bound_closure_called_via_call_user_func => {
        r#"<?php
class Box { private int $size = 5; }
$fn = Closure::bind(function(): int { return $this->size; }, new Box(), Box::class);
echo call_user_func($fn);
"#,
        ["5"]
    };

    bound_closure_called_via_call_user_func_array => {
        r#"<?php
class Pair { private int $a = 2; private int $b = 3; }
$sum = Closure::bind(function(): int { return $this->a + $this->b; }, new Pair(), Pair::class);
echo call_user_func_array($sum, []);
"#,
        ["5"]
    };

    static_closure_invoked_directly => {
        r#"<?php
$inc = static function(int $n): int { return $n + 1; };
echo $inc(41);
"#,
        ["42"]
    };

    static_arrow_invoked_directly => {
        r#"<?php
$double = static fn(int $n): int => $n * 2;
echo $double(21);
"#,
        ["42"]
    };

    closure_bind_derived_object_base_scope => {
        r#"<?php
class Base { private int $v = 10; }
class Child extends Base {}
$read = Closure::bind(function(): int { return $this->v; }, new Child(), Base::class);
echo $read();
"#,
        ["10"]
    };

    closure_bind_rejects_foreign_private_property => {
        r#"<?php
class Owner { private int $id = 1; }
class Stranger {}
$fn = Closure::bind(function() { return $this->id ?? -1; }, new Stranger(), Owner::class);
echo $fn();
"#,
        ["-1"]
    };

    is_callable_matrix_builtin_missing_null_empty => {
        r#"<?php
echo (is_callable('strlen') ? 'S' : '-')
   . (is_callable('missing_fn_xyz') ? 'S' : 'M')
   . (is_callable(null) ? 'S' : 'N')
   . (is_callable('') ? 'S' : 'E');
"#,
        ["SMNE"]
    };

    is_callable_syntax_only_vs_runtime_exists => {
        r#"<?php
echo (is_callable('strlen', false) ? 'syn' : 'no')
   . ':'
   . (is_callable('strlen', true) ? 'rt' : 'no')
   . ':'
   . (is_callable('ghost_fn', false) ? 'syn' : 'no')
   . ':'
   . (is_callable('ghost_fn', true) ? 'rt' : 'no');
"#,
        ["syn:rt:no:rt"]
    };

    is_callable_integer_false => {
        r#"<?php
echo is_callable(42) ? 'yes' : 'no';
"#,
        ["no"]
    };

    is_callable_array_not_callable => {
        r#"<?php
echo is_callable([1, 2]) ? 'yes' : 'no';
"#,
        ["no"]
    };

    is_callable_instance_method_matrix => {
        r#"<?php
class Worker { public function run(): string { return 'ok'; } }
echo (is_callable([new Worker(), 'run']) ? 'Y' : 'N')
   . (is_callable([new Worker(), 'missing']) ? 'Y' : 'N');
"#,
        ["YN"]
    };

    is_callable_static_method_string => {
        r#"<?php
class Math { public static function add(int $a, int $b): int { return $a + $b; } }
echo is_callable('Math::add') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    call_user_func_named_function => {
        r#"<?php
function greet(string $name): string { return 'hi ' . $name; }
echo call_user_func('greet', 'vybe');
"#,
        ["hi vybe"]
    };

    call_user_func_closure => {
        r#"<?php
$mul = function(int $a, int $b): int { return $a * $b; };
echo call_user_func($mul, 6, 7);
"#,
        ["42"]
    };

    call_user_func_guarded_invalid_callable => {
        r#"<?php
$target = 'absent_fn_abc';
echo is_callable($target) ? call_user_func($target) : 'blocked';
"#,
        ["blocked"]
    };

    call_user_func_array_named_function => {
        r#"<?php
function sum3(int $a, int $b, int $c): int { return $a + $b + $c; }
echo call_user_func_array('sum3', [1, 2, 3]);
"#,
        ["6"]
    };

    call_user_func_array_closure => {
        r#"<?php
$join = function(string $a, string $b): string { return $a . $b; };
echo call_user_func_array($join, ['foo', 'bar']);
"#,
        ["foobar"]
    };

    call_user_func_array_guarded_invalid => {
        r#"<?php
$bad = 99;
echo is_callable($bad) ? call_user_func_array($bad, []) : 'invalid';
"#,
        ["invalid"]
    };

    first_class_callable_user_function => {
        r#"<?php
function triple(int $n): int { return $n * 3; }
$fn = triple(...);
echo $fn(4);
"#,
        ["12"]
    };

    first_class_callable_static_method => {
        r#"<?php
class Calc { public static function square(int $n): int { return $n * $n; } }
$sq = Calc::square(...);
echo $sq(5);
"#,
        ["25"]
    };

    first_class_callable_instance_method => {
        r#"<?php
class Scale { public function __construct(private int $factor) {} public function apply(int $n): int { return $n * $this->factor; } }
$fn = (new Scale(2))->apply(...);
echo $fn(11);
"#,
        ["22"]
    };

    first_class_callable_in_array_map => {
        r#"<?php
$fn = strtoupper(...);
echo implode(',', array_map($fn, ['a', 'b']));
"#,
        ["A,B"]
    };

    first_class_callable_in_usort => {
        r#"<?php
function desc(int $a, int $b): int { return $b <=> $a; }
$arr = [1, 3, 2];
usort($arr, desc(...));
echo implode('', $arr);
"#,
        ["321"]
    };

    first_class_callable_match_default_throw_caught => {
        r#"<?php
$pick = fn(int $n) => match ($n) { 1 => 'one', default => throw new RuntimeException('bad') };
try { echo $pick(0); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["bad"]
    };

    first_class_callable_coalesce_throw_caught => {
        r#"<?php
$need = fn(?string $s) => $s ?? throw new RuntimeException('need');
try { $need(null); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["need"]
    };

    arrow_function_this_in_instance_method => {
        r#"<?php
class Host {
    private string $tag = 'inner';
    public function run(): string {
        $f = fn(): string => $this->tag;
        return $f();
    }
}
echo (new Host())->run();
"#,
        ["inner"]
    };

    arrow_function_this_in_nested_method => {
        r#"<?php
class Outer {
    private int $n = 4;
    public function wrap(): int {
        $inner = fn(): int => $this->n * 2;
        return $inner();
    }
}
echo (new Outer())->wrap();
"#,
        ["8"]
    };

    arrow_function_static_has_no_this => {
        r#"<?php
class Demo {
    public static function make(): callable {
        return static fn(): string => 'static';
    }
}
echo Demo::make()();
"#,
        ["static"]
    };

    closure_from_callable_function_name => {
        r#"<?php
function add(int $a, int $b): int { return $a + $b; }
$c = Closure::fromCallable('add');
echo $c(2, 3);
"#,
        ["5"]
    };

    closure_from_callable_array_method => {
        r#"<?php
class Greeter { public function hello(string $n): string { return 'hey ' . $n; } }
$c = Closure::fromCallable([new Greeter(), 'hello']);
echo $c('you');
"#,
        ["hey you"]
    };

    closure_from_callable_invalid_name => {
        r#"<?php
try {
    $c = Closure::fromCallable('no_such_fn');
    echo 'made';
} catch (Throwable $e) {
    echo 'err';
}
"#,
        ["err"]
    };

    invokable_object_is_callable => {
        r#"<?php
class Handler { public function __invoke(string $s): string { return strtoupper($s); } }
$h = new Handler();
echo is_callable($h) ? $h('ok') : 'no';
"#,
        ["OK"]
    };

    bind_then_invoke_in_loop => {
        r#"<?php
class Counter { private int $n = 0; public function bump(): int { return ++$this->n; } }
$inc = Closure::bind(function(): int { return $this->bump(); }, new Counter(), Counter::class);
echo $inc() . $inc();
"#,
        ["12"]
    };

    first_class_callable_parent_method => {
        r#"<?php
class Base { public function tag(): string { return 'base'; } }
class Child extends Base {
    public function getTag(): callable { return parent::tag(...); }
}
echo (new Child())->getTag()();
"#,
        ["base"]
    };

    call_user_func_after_function_exists => {
        r#"<?php
function ping(): string { return 'pong'; }
echo function_exists('ping') && is_callable('ping') ? call_user_func('ping') : 'skip';
"#,
        ["pong"]
    };

    is_callable_on_first_class_result => {
        r#"<?php
function id(int $n): int { return $n; }
$fcc = id(...);
echo is_callable($fcc) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    arrow_throw_in_method_caught => {
        r#"<?php
class Gate {
    public function admit(?string $token): string {
        return $token ?? throw new RuntimeException('denied');
    }
}
try { (new Gate())->admit(null); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["denied"]
    };

    closure_validation_throw_caught => {
        r#"<?php
$assertPos = function(int $n): int {
    return $n > 0 ? $n : throw new RuntimeException('non-positive');
};
try { $assertPos(-1); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["non-positive"]
    };

    call_user_func_returns_null => {
        r#"<?php
function noop(): void {}
echo call_user_func('noop') === null ? 'null' : 'value';
"#,
        ["null"]
    };

    bind_closure_null_scope_static_access => {
        r#"<?php
class State { private static int $v = 3; }
$read = Closure::bind(static function(): int { return static::$v; }, null, State::class);
echo $read();
"#,
        ["3"]
    };

    callable_private_method_in_scope => {
        r#"<?php
class Safe {
    private function secret(): string { return 'hidden'; }
    public function expose(): bool { return is_callable([$this, 'secret']); }
}
echo (new Safe())->expose() ? 'yes' : 'no';
"#,
        ["yes"]
    };

    bind_closure_preserves_parameter_types => {
        r#"<?php
class Meter { private int $base = 10; }
$add = Closure::bind(function(int $x): int { return $this->base + $x; }, new Meter(), Meter::class);
echo $add(5);
"#,
        ["15"]
    };

    call_user_func_variadic_from_array => {
        r#"<?php
function join3(string $a, string $b, string $c): string { return $a . $b . $c; }
echo call_user_func_array('join3', ['x', 'y', 'z']);
"#,
        ["xyz"]
    };

    first_class_callable_reject_invalid_class_method => {
        r#"<?php
class X {}
$ok = false;
try { $fn = X::missing(...); } catch (Throwable $e) { $ok = true; }
echo $ok ? 'caught' : 'end';
"#,
        ["caught"]
    };

    arrow_this_returns_same_instance => {
        r#"<?php
class SelfRef {
    public function capture(): object {
        $f = fn(): object => $this;
        return $f();
    }
}
$o = new SelfRef();
echo $o->capture() === $o ? 'same' : 'diff';
"#,
        ["same"]
    };

    is_callable_on_closure_object => {
        r#"<?php
$fn = function(): int { return 1; };
echo is_callable($fn) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    bind_invalid_scope_still_invokes => {
        r#"<?php
class A { private int $x = 1; }
$fn = function(): int { return $this->x ?? 0; };
$bound = Closure::bind($fn, new A(), 'A');
echo $bound();
"#,
        ["1"]
    };

    call_user_func_with_static_closure => {
        r#"<?php
$fn = static function(): string { return 'static-call'; };
echo call_user_func($fn);
"#,
        ["static-call"]
    };

    first_class_callable_strlen_pipeline => {
        r#"<?php
$len = strlen(...);
$mapped = array_map($len, ['ab', 'cde']);
echo implode('-', $mapped);
"#,
        ["2-3"]
    };

    closure_bind_to_anonymous_instance => {
        r#"<?php
$obj = new class { private string $k = 'anon'; };
$read = Closure::bind(function(): string { return $this->k; }, $obj, $obj::class);
echo $read();
"#,
        ["anon"]
    };

    variable_function_name_invocation => {
        r#"<?php
function dyn_greet(string $n): string { return 'hello ' . $n; }
$name = 'dyn_greet';
echo $name('world');
"#,
        ["hello world"]
    };

    variable_static_method_call => {
        r#"<?php
class DynService {
    public static function status(string $tag): string { return $tag . '-ok'; }
}
$cls = 'DynService';
$method = 'status';
echo $cls::$method('x');
"#,
        ["x-ok"]
    };

    variable_instance_class_and_method_call => {
        r#"<?php
class Node {
    public function label(string $s): string { return 'node:' . $s; }
}
$obj = new Node();
$method = 'label';
echo $obj->$method('A');
"#,
        ["node:A"]
    };

    variable_property_access_and_call => {
        r#"<?php
class Maker {
    private string $name = 'ok';
    public function get(): callable { return [$this, 'render']; }
    public function render(): string { return $this->name; }
}
$m = new Maker();
[$obj, $method] = $m->get();
echo $obj->$method();
"#,
        ["ok"]
    };

    call_user_func_array_on_invokable => {
        r#"<?php
class Handler {
    public function __invoke(int $n): string { return 'value:' . $n; }
}
$h = new Handler();
echo call_user_func_array($h, [7]);
"#,
        ["value:7"]
    };

    variable_static_class_and_method_call => {
        r#"<?php
class Dispatcher {
    public static function name(string $suffix): string { return 'ok:' . $suffix; }
}
$service = 'Dispatcher';
$method = 'name';
echo $service::$method('done');
"#,
        ["ok:done"]
    };

    variable_instance_method_call_with_parentheses => {
        r#"<?php
class Printer {
    public function paint(string $label): string { return "paint:$label"; }
}
$obj = new Printer();
$method = 'paint';
echo $obj->{$method}('blue');
"#,
        ["paint:blue"]
    };

    variable_function_object_property_call => {
        r#"<?php
class Builder {
    public function make(string $value): callable {
        return [new Renderer(), 'run'];
    }
}
class Renderer {
    public function run(string $value): string { return "R:$value"; }
}
$b = new Builder();
$c = $b->make('v');
echo $c($c[0] ? 'z' : 'unused');
"#,
        ["R:z"]
    };

    invocation_of_magic_call_through_is_callable => {
        r#"<?php
class MagicApi {
    private function hidden(string $s): string { return "h:$s"; }
    public function __call(string $name, array $args): string { return "m:$name(" . $args[0] . ")"; }
}
$m = new MagicApi();
echo is_callable([$m, 'dynamic']) ? 'yes' : 'no';
echo '|' . $m->dynamic('x');
"#,
        ["yes|m:dynamic(x)"]
    };

    first_class_callable_of_namespaced_function_via_relative_name => {
        r#"<?php
namespace Ops {
    function scale(int $n): int { return $n * 10; }
}
namespace App {
    use function Ops\scale;
    $f = \Ops\scale(...);
    echo $f(3);
}
"#,
        ["30"]
    };

    call_user_func_array_with_named_parameters => {
        r#"<?php
function format(string $prefix, string $value): string { return $prefix . '-' . $value; }
echo call_user_func_array('format', ['tag' => 'x', 'value' => 'y']);
"#,
        ["x-y"]
    };

    closure_from_callable_on_magic_call => {
        r#"<?php
class Proxy {
    public function __call(string $name, array $args): string { return "proxy:$name"; }
}
$p = new Proxy();
try {
    $call = Closure::fromCallable([$p, 'anything']);
    echo $call();
} catch (Throwable $e) {
    echo 'err';
}
"#,
        ["proxy:anything"]
    };

    callable_on_non_public_instance_method_runtime_context => {
        r#"<?php
class Context {
    private function secret(): string { return 'secret'; }
    public function expose(callable $f): string { return $f($this) . ':' . $this->secret(); }
}
$ctx = new Context();
$f = fn(Context $c): string => 'open';
echo $ctx->expose($f);
"#,
        ["open:secret"]
    };

    is_callable_on_array_with_string_instance => {
        r#"<?php
class Handler {
    public function run(): string { return 'ok'; }
}
echo is_callable(['Handler', 'run']) ? 'yes' : 'no';
echo '|';
echo is_callable([new Handler(), 'run']) ? 'yes' : 'no';
"#,
        ["yes|yes"]
    };
}
