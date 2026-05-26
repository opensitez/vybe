use super::helpers::run_prints;

// ── Interface basics ──────────────────────────────────────────

#[test] fn interface_forces_implementation() {
    assert_eq!(run_prints(r#"<?php
interface Printable { public function print(): void; }
class Doc implements Printable { public function print(): void { echo 'doc'; } }
(new Doc)->print();
"#), vec!["doc"]);
}
#[test] fn interface_type_hint_accepted() {
    assert_eq!(run_prints(r#"<?php
interface Logger { public function log(string $msg): void; }
class ConsoleLogger implements Logger { public function log(string $msg): void { echo $msg; } }
function doWork(Logger $l): void { $l->log('done'); }
doWork(new ConsoleLogger);
"#), vec!["done"]);
}
#[test] fn interface_multiple_implementations() {
    assert_eq!(run_prints(r#"<?php
interface Area { public function area(): float; }
class Rect implements Area { public function __construct(private float $w, private float $h) {} public function area(): float { return $this->w * $this->h; } }
class Circle implements Area { public function __construct(private float $r) {} public function area(): float { return round(M_PI * $this->r ** 2, 2); } }
echo (new Rect(3,4))->area() . ',' . (new Circle(1))->area();
"#), vec!["12,3.14"]);
}

// ── Interface inheritance ─────────────────────────────────────

#[test] fn interface_extends_interface() {
    assert_eq!(run_prints(r#"<?php
interface Base { public function base(): string; }
interface Extended extends Base { public function extra(): string; }
class Impl implements Extended {
    public function base(): string { return 'base'; }
    public function extra(): string { return 'extra'; }
}
$o = new Impl;
echo $o->base() . ',' . $o->extra();
"#), vec!["base,extra"]);
}
#[test] fn interface_multiple_parents() {
    assert_eq!(run_prints(r#"<?php
interface A { public function a(): int; }
interface B { public function b(): int; }
interface C extends A, B {}
class Impl implements C {
    public function a(): int { return 1; }
    public function b(): int { return 2; }
}
$o = new Impl;
echo $o->a() + $o->b();
"#), vec!["3"]);
}

// ── Class implements multiple interfaces ──────────────────────

#[test] fn class_implements_multiple() {
    assert_eq!(run_prints(r#"<?php
interface Drawable { public function draw(): string; }
interface Resizable { public function resize(float $f): static; }
class Shape implements Drawable, Resizable {
    public function __construct(private float $size = 1.0) {}
    public function draw(): string { return "shape({$this->size})"; }
    public function resize(float $f): static { return new static($this->size * $f); }
}
$s = (new Shape(2.0))->resize(3.0);
echo $s->draw();
"#), vec!["shape(6)"]);
}
#[test] fn instanceof_checks_interface() {
    assert_eq!(run_prints(r#"<?php
interface Tagged { public function tag(): string; }
class Item implements Tagged { public function tag(): string { return 'item'; } }
$o = new Item;
echo ($o instanceof Tagged) ? 'yes' : 'no';
"#), vec!["yes"]);
}

// ── Interface constants ───────────────────────────────────────

#[test] fn interface_constant_accessible_via_class() {
    assert_eq!(run_prints(r#"<?php
interface Config { const VERSION = '1.0'; }
class App implements Config {}
echo App::VERSION;
"#), vec!["1.0"]);
}
#[test] fn interface_constant_accessible_via_interface() {
    assert_eq!(run_prints(r#"<?php
interface Limits { const MAX = 100; const MIN = 0; }
echo Limits::MAX - Limits::MIN;
"#), vec!["100"]);
}

// ── Abstract classes ──────────────────────────────────────────

#[test] fn abstract_class_partial_implementation() {
    assert_eq!(run_prints(r#"<?php
abstract class Animal {
    abstract public function sound(): string;
    public function describe(): string { return get_class($this) . ' says ' . $this->sound(); }
}
class Dog extends Animal { public function sound(): string { return 'woof'; } }
class Cat extends Animal { public function sound(): string { return 'meow'; } }
echo (new Dog)->describe() . ',' . (new Cat)->describe();
"#), vec!["Dog says woof,Cat says meow"]);
}
#[test] fn abstract_class_cannot_instantiate() {
    assert_eq!(run_prints(r#"<?php
abstract class Base { abstract public function run(): void; }
try { new Base; } catch (Error $e) { echo 'cannot'; }
"#), vec!["cannot"]);
}
#[test] fn abstract_class_constructor() {
    assert_eq!(run_prints(r#"<?php
abstract class Service {
    public function __construct(protected string $name) {}
    abstract public function run(): string;
}
class MyService extends Service {
    public function run(): string { return 'running: ' . $this->name; }
}
echo (new MyService('worker'))->run();
"#), vec!["running: worker"]);
}

// ── Interface with default constant in PHP 8.1+ ───────────────

#[test] fn interface_implemented_constant_override() {
    assert_eq!(run_prints(r#"<?php
interface HasVersion { const VERSION = '1.0'; }
class App implements HasVersion { const VERSION = '2.0'; }
echo App::VERSION . ',' . HasVersion::VERSION;
"#), vec!["2.0,1.0"]);
}

// ── Duck typing with interface ────────────────────────────────

#[test] fn interface_used_in_array_of_objects() {
    assert_eq!(run_prints(r#"<?php
interface Worker { public function work(): int; }
class DevWorker implements Worker { public function work(): int { return 8; } }
class OpWorker implements Worker { public function work(): int { return 10; } }
$workers = [new DevWorker, new OpWorker, new DevWorker];
echo array_sum(array_map(fn(Worker $w) => $w->work(), $workers));
"#), vec!["26"]);
}
