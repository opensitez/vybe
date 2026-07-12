use super::helpers::run_prints;

// ── Union type declarations (PHP 8.0) ─────────────────────────

#[test]
fn union_type_int_or_string() {
    assert_eq!(
        run_prints(
            r#"<?php
function identity(int|string $v): int|string { return $v; }
echo identity(42) . ',' . identity('hello');
"#
        ),
        vec!["42,hello"]
    );
}
#[test]
fn union_type_nullable_vs_null_union() {
    assert_eq!(
        run_prints(
            r#"<?php
function maybeNull(?string $s): string { return $s ?? 'none'; }
echo maybeNull('hi') . ',' . maybeNull(null);
"#
        ),
        vec!["hi,none"]
    );
}
#[test]
fn union_type_float_or_int() {
    assert_eq!(
        run_prints(
            r#"<?php
function toNum(int|float $n): string { return is_int($n) ? 'int' : 'float'; }
echo toNum(5) . ',' . toNum(5.5);
"#
        ),
        vec!["int,float"]
    );
}
#[test]
fn union_type_in_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Container { public int|string $value; }
$c = new Container; $c->value = 'text';
echo $c->value;
$c->value = 42;
echo ',' . $c->value;
"#
        ),
        vec!["text,42"]
    );
}
#[test]
fn union_type_null_checked() {
    assert_eq!(
        run_prints(
            r#"<?php
function greet(string|null $name): string { return 'Hello, ' . ($name ?? 'Guest'); }
echo greet('Alice') . ',' . greet(null);
"#
        ),
        vec!["Hello, Alice,Hello, Guest"]
    );
}

// ── Return type declarations ──────────────────────────────────

#[test]
fn return_type_void() {
    assert_eq!(
        run_prints(
            r#"<?php
function logMsg(string $msg): void { echo $msg; }
logMsg('test');
"#
        ),
        vec!["test"]
    );
}
#[test]
fn return_type_self() {
    assert_eq!(
        run_prints(
            r#"<?php
class Builder {
    private array $parts = [];
    public function add(string $p): self { $this->parts[] = $p; return $this; }
    public function build(): string { return implode('-', $this->parts); }
}
echo (new Builder)->add('a')->add('b')->build();
"#
        ),
        vec!["a-b"]
    );
}
#[test]
fn return_type_mixed() {
    assert_eq!(
        run_prints(
            r#"<?php
function anything(bool $b): mixed { return $b ? 42 : 'hello'; }
echo anything(true) . ',' . anything(false);
"#
        ),
        vec!["42,hello"]
    );
}
#[test]
fn return_type_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function getList(): array { return [1,2,3]; }
echo implode(',', getList());
"#
        ),
        vec!["1,2,3"]
    );
}
#[test]
fn return_type_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
function makeAdder(int $n): callable { return fn($x) => $x + $n; }
$add5 = makeAdder(5);
echo $add5(3);
"#
        ),
        vec!["8"]
    );
}

// ── Parameter type declarations ───────────────────────────────

#[test]
fn param_type_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function sumArray(array $a): int { return array_sum($a); }
echo sumArray([1,2,3,4,5]);
"#
        ),
        vec!["15"]
    );
}
#[test]
fn param_type_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Vec2 { public function __construct(public float $x, public float $y) {} }
function length(Vec2 $v): float { return sqrt($v->x**2 + $v->y**2); }
echo length(new Vec2(3.0, 4.0));
"#
        ),
        vec!["5"]
    );
}
#[test]
fn param_type_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Printable { public function asString(): string; }
class Name implements Printable { public function __construct(private string $n) {} public function asString(): string { return $this->n; } }
function display(Printable $p): void { echo $p->asString(); }
display(new Name('World'));
"#
        ),
        vec!["World"]
    );
}

// ── Type widening in PHP 8.x ──────────────────────────────────

#[test]
fn covariant_return_in_override() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal {}
class Dog extends Animal {}
class Base { public function get(): Animal { return new Animal; } }
class Child extends Base { public function get(): Dog { return new Dog; } }
echo get_class((new Child)->get());
"#
        ),
        vec!["Dog"]
    );
}
#[test]
fn contravariant_param_in_override() {
    assert_eq!(
        run_prints(
            r#"<?php
class AnimalFood {}
class DogFood extends AnimalFood {}
interface Handler { public function handle(DogFood $f): void; }
class AnyHandler implements Handler { public function handle(AnimalFood $f): void { echo get_class($f); } }
(new AnyHandler)->handle(new DogFood);
"#
        ),
        vec!["DogFood"]
    );
}

// ── mixed type ────────────────────────────────────────────────

#[test]
fn mixed_accepts_all() {
    assert_eq!(
        run_prints(
            r#"<?php
function id(mixed $v): mixed { return $v; }
echo id(1) . ',' . id('a') . ',' . (id(null) ?? 'null');
"#
        ),
        vec!["1,a,null"]
    );
}

// ── Type juggling with strict_types=0 ────────────────────────

#[test]
fn coerce_string_to_int_without_strict() {
    assert_eq!(
        run_prints(
            r#"<?php
function add(int $a, int $b): int { return $a + $b; }
echo add('3', '4');
"#
        ),
        vec!["7"]
    );
}
#[test]
fn coerce_float_to_int_param() {
    assert_eq!(
        run_prints(
            r#"<?php
function n(int $v): int { return $v; }
echo n(3.9);
"#
        ),
        vec!["3"]
    );
}
