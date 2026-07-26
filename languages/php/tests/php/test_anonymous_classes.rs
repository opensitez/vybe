use super::helpers::run_prints;

// ── Basic anonymous class ─────────────────────────────────────

#[test]
fn anon_class_basic_instantiation() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class { public function hi(): string { return 'hello'; } };
echo $obj->hi();
"#
        ),
        vec!["hello"]
    );
}
#[test]
fn anon_class_with_constructor_args() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class(42) { public function __construct(private int $n) {} public function get(): int { return $this->n; } };
echo $obj->get();
"#
        ),
        vec!["42"]
    );
}
#[test]
fn anon_class_implements_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Greet { public function greet(): string; }
$obj = new class implements Greet { public function greet(): string { return 'hi'; } };
echo $obj->greet();
"#
        ),
        vec!["hi"]
    );
}
#[test]
fn anon_class_extends_parent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public function name(): string { return 'base'; } }
$obj = new class extends Base {};
echo $obj->name();
"#
        ),
        vec!["base"]
    );
}
#[test]
fn anon_class_extends_and_overrides() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public function val(): int { return 1; } }
$obj = new class extends Base { public function val(): int { return parent::val() * 10; } };
echo $obj->val();
"#
        ),
        vec!["10"]
    );
}

// ── Anonymous class in context ────────────────────────────────

#[test]
fn anon_class_as_return_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function makeLogger(string $prefix) {
    return new class($prefix) {
        public function __construct(private string $p) {}
        public function log(string $m): string { return $this->p . ': ' . $m; }
    };
}
echo makeLogger('[INFO]')->log('started');
"#
        ),
        vec!["[INFO]: started"]
    );
}
#[test]
fn anon_class_captures_outer_scope_via_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
$multiplier = 3;
$obj = new class($multiplier) {
    public function __construct(private int $m) {}
    public function apply(int $n): int { return $n * $this->m; }
};
echo $obj->apply(7);
"#
        ),
        vec!["21"]
    );
}
#[test]
fn anon_class_instanceof_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Countable2 { public function count(): int; }
$obj = new class implements Countable2 { public function count(): int { return 5; } };
echo ($obj instanceof Countable2) ? 'yes' : 'no';
"#
        ),
        vec!["yes"]
    );
}
#[test]
fn anon_class_chaining_methods() {
    assert_eq!(
        run_prints(
            r#"<?php
$b = new class {
    private array $parts = [];
    public function add(string $s): static { $this->parts[] = $s; return $this; }
    public function build(): string { return implode('-', $this->parts); }
};
echo $b->add('a')->add('b')->add('c')->build();
"#
        ),
        vec!["a-b-c"]
    );
}
#[test]
fn anon_class_static_method() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class { public static function create(): string { return 'created'; } };
echo $obj::create();
"#
        ),
        vec!["created"]
    );
}

// ── Anonymous class in array and iteration ────────────────────

#[test]
fn anon_class_in_array_of_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Shape { public function area(): float; }
$shapes = [
    new class implements Shape { public function area(): float { return 4.0; } },
    new class implements Shape { public function area(): float { return 9.0; } },
];
echo array_sum(array_map(fn($s) => $s->area(), $shapes));
"#
        ),
        vec!["13"]
    );
}
#[test]
fn anon_class_get_class_is_anonymous() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class {};
echo str_starts_with(get_class($obj), 'class@anonymous') || str_contains(get_class($obj), '@anonymous') ? 'yes' : 'no';
"#
        ),
        vec!["yes"]
    );
}
#[test]
fn anon_class_with_property_promotion() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class(10, 20) {
    public function __construct(public int $x, public int $y) {}
    public function sum(): int { return $this->x + $this->y; }
};
echo $obj->sum();
"#
        ),
        vec!["30"]
    );
}
#[test]
fn anon_class_inner_class_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
class Outer {
    public function inner(): object {
        return new class($this) {
            public function __construct(private Outer $o) {}
            public function tag(): string { return 'inner-of-' . get_class($this->o); }
        };
    }
}
echo (new Outer)->inner()->tag();
"#
        ),
        vec!["inner-of-Outer"]
    );
}
#[test]
fn anon_class_multiple_interfaces() {
    assert_eq!(
        run_prints(
            r#"<?php
interface A { public function a(): string; }
interface B { public function b(): string; }
$obj = new class implements A, B {
    public function a(): string { return 'A'; }
    public function b(): string { return 'B'; }
};
echo $obj->a() . $obj->b();
"#
        ),
        vec!["AB"]
    );
}

#[test]
fn anon_class_implements_countable() {
    assert_eq!(
        run_prints(
            r#"<?php
$c = new class implements Countable {
    public function count(): int { return 42; }
};
echo count($c);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn anon_class_with_readonly_property() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class('config_value') {
    public function __construct(public readonly string $cfg) {}
};
echo $obj->cfg;
"#
        ),
        vec!["config_value"]
    );
}

#[test]
fn anon_class_inside_trait_method() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Runner {
    public function makeRunner() {
        return new class { public function run(): string { return 'running_from_trait'; } };
    }
}
class Host { use Runner; }
$h = new Host();
echo $h->makeRunner()->run();
"#
        ),
        vec!["running_from_trait"]
    );
}

#[test]
fn anon_class_with_final_method() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class {
    final public function fixed(): string { return 'immutable_method'; }
};
echo $obj->fixed();
"#
        ),
        vec!["immutable_method"]
    );
}

#[test]
fn anon_class_callable_as_callback() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class {
    public function __invoke(int $v): int { return $v + 10; }
};

echo $obj(5);
"#
        ),
        vec!["15"]
    );
}

#[test]
fn anon_class_with_static_property_counter() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class {
    public static int $count = 0;
    public function add(): int {
        return ++self::$count;
    }
};

echo $obj->add() . $obj->add();
"#
        ),
        vec!["12"]
    );
}

#[test]
fn anon_class_with_readonly_property_promotion() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (new class {
    public function __construct(public readonly string $value) {}

    public function reveal(): string { return $this->value; }
}('secret'))->reveal();
"#
        ),
        vec!["secret"]
    );
}
