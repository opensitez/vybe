//! Anonymous classes: inline `new class` patterns with distinct behaviors.

crate::php_cases! {
    anonymous_class_implements_interface => {
        r#"<?php
interface Greet { public function msg(): string; }
$o = new class implements Greet { public function msg(): string { return 'hi'; } };
echo $o->msg();
"#,
        ["hi"]
    };

    anonymous_class_extends_parent => {
        r#"<?php
class Base { public function n(): int { return 1; } }
$o = new class extends Base { public function n(): int { return parent::n() + 1; } };
echo $o->n();
"#,
        ["2"]
    };

    anonymous_class_constructor_promotion => {
        r#"<?php
$o = new class(42) { public function __construct(public int $v) {} };
echo $o->v;
"#,
        ["42"]
    };

    anonymous_class_readonly_promoted => {
        r#"<?php
$o = new readonly class(7) { public function __construct(public int $n) {} };
echo $o->n;
"#,
        ["7"]
    };

    anonymous_class_with_trait => {
        r#"<?php
trait T { public function t(): int { return 3; } }
$o = new class { use T; };
echo $o->t();
"#,
        ["3"]
    };

    anonymous_class_static_method => {
        r#"<?php
$o = new class { public static function id(): string { return 'anon'; } };
echo $o::id();
"#,
        ["anon"]
    };

    anonymous_class_magic_invoke => {
        r#"<?php
$o = new class { public function __invoke(int $n): int { return $n * 2; } };
echo $o(4);
"#,
        ["8"]
    };

    anonymous_class_serialize_roundtrip => {
        r#"<?php
$o = new class { public int $n = 9; };
$s = serialize($o);
$u = unserialize($s);
echo $u->n;
"#,
        ["9"]
    };

    anonymous_class_nested_inside_method => {
        r#"<?php
class Outer {
    public function make(): object {
        return new class { public function k(): string { return 'inner'; } };
    }
}
echo (new Outer())->make()->k();
"#,
        ["inner"]
    };

    anonymous_class_capture_outer_variable => {
        r#"<?php
$seed = 5;
$o = new class($seed) {
    public function __construct(private int $s) {}
    public function v(): int { return $this->s; }
};
echo $o->v();
"#,
        ["5"]
    };

    anonymous_class_implements_multiple => {
        r#"<?php
interface A { public function a(): int; }
interface B { public function b(): int; }
$o = new class implements A, B {
    public function a(): int { return 1; }
    public function b(): int { return 2; }
};
echo $o->a() + $o->b();
"#,
        ["3"]
    };

    anonymous_class_private_method => {
        r#"<?php
$o = new class {
    private function secret(): string { return 'x'; }
    public function reveal(): string { return $this->secret(); }
};
echo $o->reveal();
"#,
        ["x"]
    };

    anonymous_class_property_defaults => {
        r#"<?php
$o = new class { public string $s = 'def'; };
echo $o->s;
"#,
        ["def"]
    };

    anonymous_class_clone => {
        r#"<?php
$o = new class { public int $n = 1; };
$c = clone $o;
$c->n = 2;
echo $o->n . $c->n;
"#,
        ["12"]
    };

    anonymous_class_get_class => {
        r#"<?php
$o = new class {};
echo str_contains(get_class($o), 'class@anonymous') ? 'anon' : get_class($o);
"#,
        ["anon"]
    };

    anonymous_class_with_attribute => {
        r#"<?php
$o = new #[\AllowDynamicProperties] class {};
$o->dyn = 'ok';
echo $o->dyn;
"#,
        ["ok"]
    };

    anonymous_class_generator_method => {
        r#"<?php
$o = new class {
    public function gen(): Generator { yield 1; yield 2; }
};
echo implode('', iterator_to_array($o->gen()));
"#,
        ["12"]
    };

    anonymous_class_typed_return => {
        r#"<?php
$o = new class {
    public function nums(): array { return [1, 2]; }
};
echo count($o->nums());
"#,
        ["2"]
    };

    anonymous_class_union_param => {
        r#"<?php
$o = new class {
    public function show(int|string $v): string { return (string)$v; }
};
echo $o->show('z');
"#,
        ["z"]
    };

    anonymous_class_nullable_return => {
        r#"<?php
$o = new class {
    public function maybe(bool $ok): ?string { return $ok ? 'y' : null; }
};
echo $o->maybe(false) === null ? 'null' : 'val';
"#,
        ["null"]
    };

    anonymous_class_final_extends_base => {
        r#"<?php
class B { public function v(): int { return 1; } }
$o = new class extends B { final public function v(): int { return 2; } };
echo $o->v();
"#,
        ["2"]
    };

    anonymous_class_array_access_via_interface => {
        r#"<?php
$o = new class implements ArrayAccess {
    private array $d = ['k' => 'v'];
    public function offsetExists($o): bool { return isset($this->d[$o]); }
    public function offsetGet($o): mixed { return $this->d[$o]; }
    public function offsetSet($o, $v): void { $this->d[$o] = $v; }
    public function offsetUnset($o): void { unset($this->d[$o]); }
};
echo $o['k'];
"#,
        ["v"]
    };

    anonymous_class_countable => {
        r#"<?php
$o = new class implements Countable {
    public function count(): int { return 3; }
};
echo count($o);
"#,
        ["3"]
    };

    anonymous_class_json_serializable => {
        r#"<?php
$o = new class implements JsonSerializable {
    public function jsonSerialize(): array { return ['a' => 1]; }
};
echo json_encode($o);
"#,
        ["{\"a\":1}"]
    };

    anonymous_class_stringable => {
        r#"<?php
$o = new class implements Stringable {
    public function __toString(): string { return 'str'; }
};
echo (string)$o;
"#,
        ["str"]
    };

    anonymous_class_iterable_via_iterator => {
        r#"<?php
$o = new class implements IteratorAggregate {
    public function getIterator(): Traversable { yield 1; yield 2; }
};
echo implode('', iterator_to_array($o));
"#,
        ["12"]
    };

    anonymous_class_closure_property => {
        r#"<?php
$o = new class {
    public function __construct(public Closure $fn) {}
};
echo ($o->fn)(3);
"#,
        ["3"]
    };

    anonymous_class_two_instances_different => {
        r#"<?php
$a = new class { public int $n = 1; };
$b = new class { public int $n = 1; };
echo $a === $b ? 'same' : 'diff';
"#,
        ["diff"]
    };

    anonymous_class_static_counter => {
        r#"<?php
$o = new class { public static int $c = 0; public function inc(): int { return ++self::$c; } };
echo $o->inc() . $o->inc();
"#,
        ["12"]
    };

    anonymous_class_destructor_echo => {
        r#"<?php
$o = new class { public function __destruct() { echo 'x'; } };
unset($o);
"#,
        ["x"]
    };

    anonymous_class_match_in_method => {
        r#"<?php
$o = new class {
    public function label(string $c): string {
        return match ($c) { 'r' => 'red', default => 'other' };
    }
};
echo $o->label('r');
"#,
        ["red"]
    };

    anonymous_class_enum_property => {
        r#"<?php
enum E { case A; }
$o = new class { public E $e = E::A; };
echo $o->e->name;
"#,
        ["A"]
    };

    anonymous_class_spread_args => {
        r#"<?php
$o = new class {
    public function sum(int ...$n): int { return array_sum($n); }
};
echo $o->sum(...[1, 2, 3]);
"#,
        ["6"]
    };

    anonymous_class_reference_param => {
        r#"<?php
$o = new class {
    public function bump(int &$n): void { $n++; }
};
$x = 1;
$o->bump($x);
echo $x;
"#,
        ["2"]
    };

    anonymous_class_named_argument_call => {
        r#"<?php
$o = new class {
    public function pair(int $a, int $b): int { return $a + $b; }
};
echo $o->pair(b: 2, a: 3);
"#,
        ["5"]
    };
}
