use super::helpers::{compile_ok, run_prints};

// ── Property hooks: basic get ─────────────────────────────────

#[test]
fn property_hook_get_returns_computed_value() {
    assert_eq!(
        run_prints(
            r#"<?php
class Circle {
    public float $radius = 5.0;
    public float $area {
        get => M_PI * $this->radius ** 2;
    }
}
$c = new Circle();
echo round($c->area, 2);
"#
        ),
        vec!["78.54"]
    );
}

#[test]
fn property_hook_get_derived_from_other_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Name {
    public string $first = '';
    public string $last = '';
    public string $full {
        get => trim($this->first . ' ' . $this->last);
    }
}
$n = new Name();
$n->first = 'John';
$n->last = 'Doe';
echo $n->full;
"#
        ),
        vec!["John Doe"]
    );
}

// ── Property hooks: basic set ─────────────────────────────────

#[test]
fn property_hook_set_validates_before_storing() {
    assert_eq!(
        run_prints(
            r#"<?php
class Age {
    public int $value {
        set(int $v) {
            if ($v < 0 || $v > 150) throw new \InvalidArgumentException("Invalid age");
            $this->value = $v;
        }
    }
}
$a = new Age();
$a->value = 30;
echo $a->value;
"#
        ),
        vec!["30"]
    );
}

#[test]
fn property_hook_set_throws_on_invalid_value() {
    assert_eq!(
        run_prints(
            r#"<?php
class PositiveInt {
    public int $value {
        set(int $v) {
            if ($v <= 0) throw new \RangeException("Must be positive");
            $this->value = $v;
        }
    }
}
$p = new PositiveInt();
try {
    $p->value = -5;
} catch (\RangeException $e) {
    echo $e->getMessage();
}
"#
        ),
        vec!["Must be positive"]
    );
}

// ── Property hooks: both get and set ─────────────────────────

#[test]
fn property_hook_get_and_set_together() {
    assert_eq!(
        run_prints(
            r#"<?php
class Temperature {
    public float $celsius {
        get => $this->celsius;
        set(float $v) { $this->celsius = round($v, 1); }
    }
}
$t = new Temperature();
$t->celsius = 36.66666;
echo $t->celsius;
"#
        ),
        vec!["36.7"]
    );
}

// ── Virtual property (no backing field) ──────────────────────

#[test]
fn virtual_property_computed_from_parts() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function __construct(
        public float $x,
        public float $y,
    ) {}
    public float $distance {
        get => sqrt($this->x ** 2 + $this->y ** 2);
    }
}
$p = new Point(3, 4);
echo $p->distance;
"#
        ),
        vec!["5"]
    );
}

// ── Property hooks with inheritance ──────────────────────────

#[test]
fn property_hook_get_in_child_overrides_parent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public string $label {
        get => "base";
    }
}
class Child extends Base {
    public string $label {
        get => "child";
    }
}
echo (new Child())->label;
"#
        ),
        vec!["child"]
    );
}

// ── Property hooks with readonly ─────────────────────────────

#[test]
fn readonly_property_with_get_hook_accessible() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {
    public readonly string $hash {
        get => strtoupper($this->hash);
    }
    public function __construct(string $hash) {
        $this->hash = $hash;
    }
}
$t = new Token("abc123");
echo $t->hash;
"#
        ),
        vec!["ABC123"]
    );
}

// ── Set hook with type coercion ───────────────────────────────

#[test]
fn property_hook_set_coerces_string_to_int() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public int $count {
        set(int|string $v) { $this->count = (int)$v; }
    }
}
$c = new Counter();
$c->count = "42";
echo $c->count;
"#
        ),
        vec!["42"]
    );
}

// ── Property hook in interface (PHP 8.4) ─────────────────────

#[test]
fn property_hook_declared_in_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface HasName {
    public string $name { get; }
}
class Person implements HasName {
    public string $name {
        get => $this->name;
    }
    public function __construct(string $name) { $this->name = $name; }
}
$p = new Person("Alice");
echo $p->name;
"#
        ),
        vec!["Alice"]
    );
}

// ── Set hook transforms before storage ───────────────────────

#[test]
fn property_hook_set_normalizes_string() {
    assert_eq!(
        run_prints(
            r#"<?php
class Slug {
    public string $value {
        set(string $v) {
            $this->value = strtolower(preg_replace('/\s+/', '-', trim($v)));
        }
    }
}
$s = new Slug();
$s->value = "  Hello World  ";
echo $s->value;
"#
        ),
        vec!["hello-world"]
    );
}

// ── Property hook with short set syntax ──────────────────────

#[test]
fn property_hook_set_short_arrow_syntax() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public int $width {
        set => max(0, $value);
    }
}
$b = new Box();
$b->width = -5;
echo $b->width;
"#
        ),
        vec!["0"]
    );
}

// ── Property hook used in method ─────────────────────────────

#[test]
fn property_hook_used_inside_class_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Rectangle {
    public function __construct(
        public float $width,
        public float $height,
    ) {}
    public float $area { get => $this->width * $this->height; }
    public function describe(): string {
        return "{$this->width}x{$this->height}={$this->area}";
    }
}
echo (new Rectangle(4, 5))->describe();
"#
        ),
        vec!["4x5=20"]
    );
}

// ── Multiple property hooks in same class ─────────────────────

#[test]
fn multiple_properties_with_hooks_in_same_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    public string $host {
        set(string $v) { $this->host = trim($v); }
    }
    public int $port {
        set(int $v) { $this->port = max(1, min(65535, $v)); }
    }
}
$c = new Config();
$c->host = "  localhost  ";
$c->port = 99999;
echo $c->host . ':' . $c->port;
"#
        ),
        vec!["localhost:65535"]
    );
}

// ── Property hook in promoted property ───────────────────────

#[test]
fn property_hook_with_constructor_access() {
    assert_eq!(
        run_prints(
            r#"<?php
class Email {
    public string $address {
        set(string $v) {
            if (!str_contains($v, '@')) throw new \InvalidArgumentException("Invalid email");
            $this->address = strtolower($v);
        }
    }
    public function __construct(string $addr) { $this->address = $addr; }
}
$e = new Email("Alice@Example.COM");
echo $e->address;
"#
        ),
        vec!["alice@example.com"]
    );
}

// ── Get hook format convenience ───────────────────────────────

#[test]
fn property_hook_get_formats_money_cents() {
    assert_eq!(
        run_prints(
            r#"<?php
class Price {
    public function __construct(public int $cents) {}
    public string $formatted {
        get => '$' . number_format($this->cents / 100, 2);
    }
}
$p = new Price(1999);
echo $p->formatted;
"#
        ),
        vec!["$19.99"]
    );
}

// ── Chained property hook reads ──────────────────────────────

#[test]
fn property_hook_get_used_in_another_hook() {
    assert_eq!(
        run_prints(
            r#"<?php
class Measure {
    public function __construct(public float $meters) {}
    public float $cm { get => $this->meters * 100; }
    public float $mm { get => $this->cm * 10; }
}
$m = new Measure(1.5);
echo $m->cm . ',' . $m->mm;
"#
        ),
        vec!["150,1500"]
    );
}

// ── Property hook lazy initialization ─────────────────────────

#[test]
fn property_hook_set_triggers_side_effect() {
    assert_eq!(
        run_prints(
            r#"<?php
class Observable {
    private array $listeners = [];
    public int $value {
        set(int $v) {
            $old = $this->value ?? null;
            $this->value = $v;
            foreach ($this->listeners as $fn) $fn($old, $v);
        }
    }
    public function onChange(callable $fn): void { $this->listeners[] = $fn; }
}
$o = new Observable();
$o->onChange(fn($old, $new) => print("changed: $new\n"));
$o->value = 10;
$o->value = 20;
"#
        ),
        vec!["changed: 10", "changed: 20"]
    );
}

// ── Property hook access from subclass ───────────────────────

#[test]
fn property_hook_accessible_from_subclass_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public string $type { get => "base"; }
}
class Child extends Base {}
echo (new Child())->type;
"#
        ),
        vec!["base"]
    );
}

// ── Compile-only: abstract property hook in interface ─────────

#[test]
fn interface_property_hook_declaration_compiles() {
    compile_ok(
        r#"<?php
interface Shape {
    public float $area { get; }
}
"#,
    );
}

// ── Property hook with static property fallback ───────────────

#[test]
fn property_hook_uses_static_lookup() {
    assert_eq!(
        run_prints(
            r#"<?php
class Registry {
    private static array $data = [];
    public string $key {
        set(string $v) { $this->key = $v; self::$data[$v] = true; }
        get => $this->key;
    }
    public static function has(string $k): bool { return isset(self::$data[$k]); }
}
$r = new Registry();
$r->key = "mykey";
echo Registry::has("mykey") ? 'found' : 'not found';
"#
        ),
        vec!["found"]
    );
}
