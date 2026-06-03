use super::helpers::run_prints;

// ── static:: vs self:: ────────────────────────────────────────

#[test]
fn lsb_static_class_returns_called_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public static function which(): string { return static::class; } }
class Child extends Base {}
echo Child::which();
"#
        ),
        vec!["Child"]
    );
}
#[test]
fn lsb_self_class_returns_defining_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public static function which(): string { return self::class; } }
class Child extends Base {}
echo Child::which();
"#
        ),
        vec!["Base"]
    );
}
#[test]
fn lsb_new_static_creates_child_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public static function create(): static { return new static(); } }
class Child extends Base {}
$obj = Child::create();
echo get_class($obj);
"#
        ),
        vec!["Child"]
    );
}
#[test]
fn lsb_new_self_always_creates_base() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public static function create(): self { return new self(); } }
class Child extends Base {}
$obj = Child::create();
echo get_class($obj);
"#
        ),
        vec!["Base"]
    );
}

// ── Late static binding factory pattern ──────────────────────

#[test]
fn lsb_factory_named_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
class Model {
    protected string $type;
    public static function make(): static {
        $obj = new static();
        $obj->type = static::class;
        return $obj;
    }
    public function getType(): string { return $this->type; }
}
class User extends Model {}
class Post extends Model {}
echo User::make()->getType() . ',' . Post::make()->getType();
"#
        ),
        vec!["User,Post"]
    );
}
#[test]
fn lsb_static_property_per_subclass() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    protected static int $count = 0;
    public static function inc(): void { static::$count++; }
    public static function get(): int { return static::$count; }
}
class A extends Counter {}
class B extends Counter {}
A::inc(); A::inc(); B::inc();
echo A::get() . ',' . B::get();
"#
        ),
        vec!["2,1"]
    );
}
#[test]
fn lsb_static_in_instance_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public function clone(): static { return new static(); }
}
class Extended extends Base {}
$obj = new Extended;
echo get_class($obj->clone());
"#
        ),
        vec!["Extended"]
    );
}
#[test]
fn lsb_get_called_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public static function who(): string { return get_called_class(); } }
class Child extends Base {}
echo Child::who();
"#
        ),
        vec!["Child"]
    );
}

// ── LSB with inheritance chains ───────────────────────────────

#[test]
fn lsb_three_level_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class A { public static function make(): static { return new static(); } }
class B extends A {}
class C extends B {}
echo get_class(C::make());
"#
        ),
        vec!["C"]
    );
}
#[test]
fn lsb_registry_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
class Entity {
    private static array $items = [];
    public static function register(string $name): void { static::$items[static::class][] = $name; }
    public static function all(): array { return static::$items[static::class] ?? []; }
}
class Tag extends Entity {}
Tag::register('php'); Tag::register('rust');
echo implode(',', Tag::all());
"#
        ),
        vec!["php,rust"]
    );
}
#[test]
fn lsb_fluent_builder_with_static_return() {
    assert_eq!(
        run_prints(
            r#"<?php
class Query {
    protected array $conditions = [];
    public function where(string $c): static { $this->conditions[] = $c; return $this; }
    public function build(): string { return implode(' AND ', $this->conditions); }
}
class UserQuery extends Query {}
echo (new UserQuery)->where('age>18')->where('active=1')->build();
"#
        ),
        vec!["age>18 AND active=1"]
    );
}

// ── LSB with abstract base class ─────────────────────────────

#[test]
fn lsb_abstract_factory() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Shape {
    abstract protected function area(): float;
    public static function describe(): string { return static::class . ' is a shape'; }
}
class Circle extends Shape { protected function area(): float { return 3.14; } }
echo Circle::describe();
"#
        ),
        vec!["Circle is a shape"]
    );
}
#[test]
fn lsb_static_const_resolved_in_child() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    const TYPE = 'base';
    public static function type(): string { return static::TYPE; }
}
class Child extends Base { const TYPE = 'child'; }
echo Base::type() . ',' . Child::type();
"#
        ),
        vec!["base,child"]
    );
}
