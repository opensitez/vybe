//! Magic-method and undefined-call failure paths (errors, not happy-path hooks).

crate::php_cases! {
    undefined_instance_method_triggers_error => {
        r#"<?php
class Plain {}
try { (new Plain())->missing(); echo 'ok'; }
catch (Error $e) { echo 'undef'; }
"#,
        ["undef"]
    };

    undefined_static_method_triggers_error => {
        r#"<?php
class Plain {}
try { Plain::missing(); echo 'ok'; }
catch (Error $e) { echo 'static'; }
"#,
        ["static"]
    };

    magic_call_routes_to_handler => {
        r#"<?php
class Proxy {
    public function __call(string $name, array $args): string {
        return $name . ':' . count($args);
    }
}
echo (new Proxy())->run(1, 2);
"#,
        ["run:2"]
    };

    magic_call_static_routes_to_handler => {
        r#"<?php
class Proxy {
    public static function __callStatic(string $name, array $args): string {
        return 's:' . $name;
    }
}
echo Proxy::create();
"#,
        ["s:create"]
    };

    magic_call_throws_from_handler => {
        r#"<?php
class Gate {
    public function __call(string $name, array $args): void {
        throw new BadMethodCallException($name);
    }
}
try { (new Gate())->deny(); }
catch (BadMethodCallException $e) { echo $e->getMessage(); }
"#,
        ["deny"]
    };

    magic_get_returns_dynamic_value => {
        r#"<?php
class Bag {
    private array $data = ['x' => 9];
    public function __get(string $k) { return $this->data[$k] ?? null; }
}
$b = new Bag();
echo $b->x;
"#,
        ["9"]
    };

    magic_get_throws_when_missing => {
        r#"<?php
class StrictBag {
    public function __get(string $k): mixed {
        throw new OutOfBoundsException($k);
    }
}
try { echo (new StrictBag())->nope; }
catch (OutOfBoundsException $e) { echo $e->getMessage(); }
"#,
        ["nope"]
    };

    magic_set_stores_dynamic_property => {
        r#"<?php
class Dyn {
    private array $store = [];
    public function __set(string $k, mixed $v): void { $this->store[$k] = $v; }
    public function dump(): string { return implode(',', $this->store); }
}
$d = new Dyn();
$d->name = 'vybe';
echo $d->dump();
"#,
        ["vybe"]
    };

    magic_isset_false_for_missing => {
        r#"<?php
class Dyn {
    public function __isset(string $k): bool { return false; }
}
$d = new Dyn();
echo isset($d->ghost) ? 'yes' : 'no';
"#,
        ["no"]
    };

    magic_unset_clears_dynamic => {
        r#"<?php
class Dyn {
    public array $store = ['a' => 1];
    public function __unset(string $k): void { unset($this->store[$k]); }
}
$d = new Dyn();
unset($d->a);
echo count($d->store);
"#,
        ["0"]
    };

    invoke_with_too_few_args_type_error => {
        r#"<?php
class Fn {
    public function __invoke(int $a, int $b): int { return $a + $b; }
}
try { (new Fn())(1); echo 'ok'; }
catch (ArgumentCountError $e) { echo 'count'; }
"#,
        ["count"]
    };

    invoke_with_correct_args => {
        r#"<?php
class Fn {
    public function __invoke(int $a, int $b): int { return $a + $b; }
}
echo (new Fn())(2, 3);
"#,
        ["5"]
    };

    call_user_func_on_invokable => {
        r#"<?php
class Greeter {
    public function __invoke(string $name): string { return 'hi ' . $name; }
}
echo call_user_func(new Greeter(), 'ana');
"#,
        ["hi ana"]
    };

    magic_sleep_returns_array_of_fields => {
        r#"<?php
class User {
    public function __construct(public string $name, public int $id) {}
    public function __sleep(): array { return ['name']; }
}
$u = new User('bob', 1);
$data = serialize($u);
echo str_contains($data, 'bob') ? 'sleep' : 'no';
"#,
        ["sleep"]
    };

    magic_wakeup_restores_object => {
        r#"<?php
class User {
    public function __construct(public string $name) {}
}
$u = unserialize(serialize(new User('kim')));
echo $u->name;
"#,
        ["kim"]
    };

    magic_clone_creates_copy => {
        r#"<?php
class Cell {
    public function __construct(public int $v) {}
    public function __clone(): void { $this->v++; }
}
$a = new Cell(1);
$b = clone $a;
echo $a->v . ':' . $b->v;
"#,
        ["1:2"]
    };

    magic_to_string_in_concat => {
        r#"<?php
class Tag {
    public function __construct(private string $t) {}
    public function __toString(): string { return $this->t; }
}
echo '{' . new Tag('x') . '}';
"#,
        ["{x}"]
    };

    magic_to_string_throws_propagates => {
        r#"<?php
class BadString {
    public function __toString(): string { throw new RuntimeException('cast'); }
}
try { echo (string)new BadString(); }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["cast"]
    };

    magic_debug_info_exports_array => {
        r#"<?php
class DebugMe {
    private int $secret = 5;
    public function __debugInfo(): array { return ['secret' => $this->secret]; }
}
ob_start();
var_dump(new DebugMe());
$out = ob_get_clean();
echo str_contains($out, 'secret') ? 'debug' : 'no';
"#,
        ["debug"]
    };

    property_hooks_get_without_set => {
        r#"<?php
class RO {
    public string $name {
        get => 'fixed';
    }
}
echo (new RO())->name;
"#,
        ["fixed"]
    };

    property_hooks_get_computes_derived_value => {
        r#"<?php
class Circle {
    public function __construct(private float $radius) {}
    public float $area {
        get => 3 * $this->radius * $this->radius;
    }
}
$c = new Circle(2);
echo $c->area;
"#,
        ["12"]
    };

    readonly_dynamic_property_blocked => {
        r#"<?php
readonly class Box { public int $n = 1; }
$b = new Box();
try { $b->extra = 2; echo 'ok'; }
catch (Error $e) { echo 'dyn'; }
"#,
        ["dyn"]
    };

    accessing_private_property_without_magic => {
        r#"<?php
class Vault { private int $gold = 3; }
$v = new Vault();
try { echo $v->gold; echo 'ok'; }
catch (Error $e) { echo 'private'; }
"#,
        ["private"]
    };

    calling_non_callable_string_errors => {
        r#"<?php
try { ('not_a_func')(); echo 'ok'; }
catch (Error $e) { echo 'call'; }
"#,
        ["call"]
    };

    calling_null_callable_errors => {
        r#"<?php
$fn = null;
try { $fn(); echo 'ok'; }
catch (Error $e) { echo 'null'; }
"#,
        ["null"]
    };

    array_access_on_non_array_object_without_interface => {
        r#"<?php
class NotArray {}
try { (new NotArray())[0]; echo 'ok'; }
catch (Error $e) { echo 'offset'; }
"#,
        ["offset"]
    };

    countable_on_non_countable_without_interface => {
        r#"<?php
class NotCount {}
try { count(new NotCount()); echo 'ok'; }
catch (TypeError $e) { echo 'count'; }
"#,
        ["count"]
    };

    iterator_on_non_traversable_errors => {
        r#"<?php
class NotTraversable {}
try { foreach (new NotTraversable() as $_) { echo 'x'; } echo 'ok'; }
catch (Error $e) { echo 'foreach'; }
"#,
        ["foreach"]
    };

    magic_serialize_on_custom_object => {
        r#"<?php
class Token implements Stringable {
    public function __construct(private string $v) {}
    public function __toString(): string { return $this->v; }
}
echo (string)new Token('t');
"#,
        ["t"]
    };

    magic_call_with_spread_args => {
        r#"<?php
class Varargs {
    public function __call(string $name, array $args): int { return array_sum($args); }
}
echo (new Varargs())->sum(1, 2, 3);
"#,
        ["6"]
    };

    magic_static_call_with_args => {
        r#"<?php
class Factory {
    public static function __callStatic(string $name, array $args): string {
        return $name . '=' . ($args[0] ?? '');
    }
}
echo Factory::build('item');
"#,
        ["build=item"]
    };

    magic_get_set_pair_roundtrip => {
        r#"<?php
class Store {
    private array $d = [];
    public function __get($k) { return $this->d[$k] ?? null; }
    public function __set($k, $v) { $this->d[$k] = $v; }
}
$s = new Store();
$s->key = 'val';
echo $s->key;
"#,
        ["val"]
    };

    magic_isset_after_set_true => {
        r#"<?php
class Store {
    private array $d = [];
    public function __set($k, $v) { $this->d[$k] = $v; }
    public function __isset($k) { return isset($this->d[$k]); }
}
$s = new Store();
$s->a = 1;
echo isset($s->a) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    parent_call_to_missing_parent_method => {
        r#"<?php
class Child extends stdClass {
    public function run(): void { parent::missing(); }
}
try { (new Child())->run(); }
catch (Error $e) { echo 'parent'; }
"#,
        ["parent"]
    };

    abstract_class_cannot_be_instantiated => {
        r#"<?php
abstract class Base { abstract public function run(): string; }
try { new Base(); echo 'made'; }
catch (Error $e) { echo 'abstract'; }
"#,
        ["abstract"]
    };

    trait_collision_requires_insteadof => {
        r#"<?php
trait A { public function talk(): string { return 'a'; } }
trait B { public function talk(): string { return 'b'; } }
class C { use A, B { A::talk insteadof B; } }
echo (new C())->talk();
"#,
        ["a"]
    };

    magic_call_recurses_until_stack_or_handler => {
        r#"<?php
class Loop {
    public function __call($name, $args) { return $this->$name; }
}
try { (new Loop())->loop(); echo 'ok'; }
catch (Error $e) { echo 'loop'; }
"#,
        ["loop"]
    };

    stringable_required_context_cast => {
        r#"<?php
class Label implements Stringable {
    public function __construct(private string $t) {}
    public function __toString(): string { return $this->t; }
}
function needsString(Stringable $s): string { return $s->__toString(); }
echo needsString(new Label('ok'));
"#,
        ["ok"]
    };

    clone_on_uncloneable_with_private_clone => {
        r#"<?php
class Single {
    private function __clone() {}
    public static function make(): self { return new self(); }
}
$s = Single::make();
try { clone $s; echo 'cloned'; }
catch (Error $e) { echo 'clone'; }
"#,
        ["clone"]
    };

    dynamic_call_private_method_from_outside => {
        r#"<?php
class Hidden { private function secret(): string { return 'no'; } }
$h = new Hidden();
try { $h->secret(); echo 'ok'; }
catch (Error $e) { echo 'hidden'; }
"#,
        ["hidden"]
    };

    magic_get_on_uninitialized_typed_with_magic => {
        r#"<?php
class Box {
    public int $n;
    public function __get($k) { return 'magic'; }
}
$b = new Box();
try { echo $b->n; }
catch (Error $e) { echo 'uninit'; }
"#,
        ["uninit"]
    };
}
