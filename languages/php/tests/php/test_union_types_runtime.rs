//! Union and intersection types at runtime (no violation paths).

crate::php_cases! {
    union_int_or_string_accepts_int => {
        r#"<?php
function f(int|string $v): string { return (string)$v; }
echo f(7);
"#,
        ["7"]
    };

    union_int_or_string_accepts_string => {
        r#"<?php
function f(int|string $v): string { return $v; }
echo f('x');
"#,
        ["x"]
    };

    union_return_nullable_int => {
        r#"<?php
function g(bool $ok): int|null { return $ok ? 1 : null; }
echo g(false) === null ? 'null' : 'n';
"#,
        ["null"]
    };

    union_array_or_false_from_search => {
        r#"<?php
function find(array $a, int $n): array|false { return array_search($n, $a, true) !== false ? ['ok'] : false; }
echo find([1, 2], 3) === false ? 'no' : 'yes';
"#,
        ["no"]
    };

    union_with_false_literal => {
        r#"<?php
function maybe(): string|false { return false; }
echo maybe() === false ? 'f' : 's';
"#,
        ["f"]
    };

    union_promoted_property_string_or_int => {
        r#"<?php
class Box { public function __construct(public int|string $v) {} }
echo (new Box('a'))->v;
"#,
        ["a"]
    };

    union_three_types => {
        r#"<?php
function tag(int|string|bool $v): string { return gettype($v); }
echo tag(true);
"#,
        ["boolean"]
    };

    intersection_iterator_and_countable => {
        r#"<?php
function len(Iterator&Countable $c): int { return count($c); }
$it = new ArrayIterator([1, 2, 3]);
echo len($it);
"#,
        ["3"]
    };

    intersection_traversable_and_arrayaccess => {
        r#"<?php
class C implements Traversable, ArrayAccess, Countable, IteratorAggregate {
    public function __construct(private array $d) {}
    public function getIterator(): Traversable { yield from $this->d; }
    public function offsetExists($k): bool { return isset($this->d[$k]); }
    public function offsetGet($k): mixed { return $this->d[$k]; }
    public function offsetSet($k, $v): void { $this->d[$k] = $v; }
    public function offsetUnset($k): void { unset($this->d[$k]); }
    public function count(): int { return count($this->d); }
}
function read(ArrayAccess&Countable $x): int { return count($x); }
echo read(new C([1]));
"#,
        ["1"]
    };

    union_in_closure_param => {
        r#"<?php
$fn = function (float|int $n): int { return (int)$n; };
echo $fn(3.9);
"#,
        ["3"]
    };

    union_with_null_first => {
        r#"<?php
function opt(?string $s = null): string { return $s ?? 'none'; }
echo opt();
"#,
        ["none"]
    };

    union_enum_and_string => {
        r#"<?php
enum Color { case Red; case Blue; }
function paint(Color|string $c): string { return is_string($c) ? $c : $c->name; }
echo paint('green');
"#,
        ["green"]
    };

    union_enum_case => {
        r#"<?php
enum Color { case Red; }
function paint(Color|string $c): string { return is_string($c) ? $c : $c->name; }
echo paint(Color::Red);
"#,
        ["Red"]
    };

    union_static_return_type => {
        r#"<?php
class U {
    public static function val(): int|string { return 42; }
}
echo U::val();
"#,
        ["42"]
    };

    union_property_set_multiple_types => {
        r#"<?php
class Holder { public int|float $n = 1; }
$h = new Holder();
$h->n = 2.5;
echo (string)$h->n;
"#,
        ["2.5"]
    };

    union_in_interface_implementation => {
        r#"<?php
interface I { public function m(int|string $v): int|string; }
class C implements I { public function m(int|string $v): int|string { return $v; } }
echo (new C())->m(5);
"#,
        ["5"]
    };

    union_with_array_type => {
        r#"<?php
function wrap(array|object $v): string { return is_array($v) ? 'arr' : 'obj'; }
echo wrap([1]);
"#,
        ["arr"]
    };

    union_with_object_type => {
        r#"<?php
function wrap(array|object $v): string { return is_array($v) ? 'arr' : 'obj'; }
echo wrap(new stdClass());
"#,
        ["obj"]
    };

    union_callable_type => {
        r#"<?php
function run(callable $fn): int { return $fn(); }
echo run(fn() => 4);
"#,
        ["4"]
    };

    union_iterable_type => {
        r#"<?php
function sum(iterable $it): int { $s = 0; foreach ($it as $n) $s += $n; return $s; }
echo sum([1, 2]);
"#,
        ["3"]
    };

    union_mixed_explicit => {
        r#"<?php
function id(mixed $v): mixed { return $v; }
echo id('z');
"#,
        ["z"]
    };

    union_false_in_return_position => {
        r#"<?php
function parse(string $s): array|false { return $s === '' ? false : [$s]; }
echo parse('') === false ? 'bad' : 'ok';
"#,
        ["bad"]
    };

    union_true_literal_narrow => {
        r#"<?php
function ok(): true { return true; }
echo ok() ? '1' : '0';
"#,
        ["1"]
    };

    union_never_unreachable_not_used => {
        r#"<?php
function always(): string|int { return 'x'; }
echo always();
"#,
        ["x"]
    };

    union_backed_enum_int => {
        r#"<?php
enum N: int { case One = 1; }
function num(N|int $v): int { return is_int($v) ? $v : $v->value; }
echo num(N::One);
"#,
        ["1"]
    };

    union_self_return_child => {
        r#"<?php
class Node { public function next(): self|static { return $this; } }
echo (new Node())->next() instanceof Node ? 'yes' : 'no';
"#,
        ["yes"]
    };

    union_parent_child_instances => {
        r#"<?php
class A {}
class B extends A {}
function pick(bool $b): A|B { return $b ? new B() : new A(); }
echo pick(true)::class;
"#,
        ["B"]
    };

    intersection_stringable_and_json => {
        r#"<?php
class D implements Stringable, JsonSerializable {
    public function __toString(): string { return 's'; }
    public function jsonSerialize(): string { return 'j'; }
}
function out(Stringable&JsonSerializable $d): string { return $d->__toString() . $d->jsonSerialize(); }
echo out(new D());
"#,
        ["sj"]
    };

    union_null_false_zero_coalesce => {
        r#"<?php
function val(): string|int|null { return null; }
echo val() ?? 'd';
"#,
        ["d"]
    };

    union_match_arm_types => {
        r#"<?php
function label(int|string $v): string {
    return match ($v) { 1 => 'one', 'a' => 'alpha', default => 'other' };
}
echo label('a');
"#,
        ["alpha"]
    };

    union_spread_array_merge => {
        r#"<?php
function merge(array ...$parts): array { return array_merge(...$parts); }
echo count(merge([1], [2, 3]));
"#,
        ["3"]
    };

    union_generator_return => {
        r#"<?php
function gen(): Generator|array { yield 1; }
echo iterator_to_array(gen())[0];
"#,
        ["1"]
    };

    union_resource_or_null => {
        r#"<?php
function open(): mixed { return fopen('php://memory', 'r+'); }
$h = open();
echo is_resource($h) ? 'res' : 'no';
"#,
        ["res"]
    };

    union_literal_string_union => {
        r#"<?php
function axis('x'|'y' $a): string { return $a; }
echo axis('y');
"#,
        ["y"]
    };

    union_false_with_int_count => {
        r#"<?php
function cnt(array $a): int|false { return count($a) > 0 ? count($a) : false; }
echo cnt([]);
"#,
        [""]
    };
}
