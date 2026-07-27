use super::helpers::run_prints;

// ── Interface basics ──────────────────────────────────────────

#[test]
fn interface_forces_implementation() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Printable { public function print(): void; }
class Doc implements Printable { public function print(): void { echo 'doc'; } }
(new Doc)->print();
"#
        ),
        vec!["doc"]
    );
}
#[test]
fn interface_type_hint_accepted() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Logger { public function log(string $msg): void; }
class ConsoleLogger implements Logger { public function log(string $msg): void { echo $msg; } }
function doWork(Logger $l): void { $l->log('done'); }
doWork(new ConsoleLogger);
"#
        ),
        vec!["done"]
    );
}
#[test]
fn interface_multiple_implementations() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Area { public function area(): float; }
class Rect implements Area { public function __construct(private float $w, private float $h) {} public function area(): float { return $this->w * $this->h; } }
class Circle implements Area { public function __construct(private float $r) {} public function area(): float { return round(M_PI * $this->r ** 2, 2); } }
echo (new Rect(3,4))->area() . ',' . (new Circle(1))->area();
"#
        ),
        vec!["12,3.14"]
    );
}

// ── Interface inheritance ─────────────────────────────────────

#[test]
fn interface_extends_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Base { public function base(): string; }
interface Extended extends Base { public function extra(): string; }
class Impl implements Extended {
    public function base(): string { return 'base'; }
    public function extra(): string { return 'extra'; }
}
$o = new Impl;
echo $o->base() . ',' . $o->extra();
"#
        ),
        vec!["base,extra"]
    );
}
#[test]
fn interface_multiple_parents() {
    assert_eq!(
        run_prints(
            r#"<?php
interface A { public function a(): int; }
interface B { public function b(): int; }
interface C extends A, B {}
class Impl implements C {
    public function a(): int { return 1; }
    public function b(): int { return 2; }
}
$o = new Impl;
echo $o->a() + $o->b();
"#
        ),
        vec!["3"]
    );
}

// ── Class implements multiple interfaces ──────────────────────

#[test]
fn class_implements_multiple() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Drawable { public function draw(): string; }
interface Resizable { public function resize(float $f): static; }
class Shape implements Drawable, Resizable {
    public function __construct(private float $size = 1.0) {}
    public function draw(): string { return "shape({$this->size})"; }
    public function resize(float $f): static { return new static($this->size * $f); }
}
$s = (new Shape(2.0))->resize(3.0);
echo $s->draw();
"#
        ),
        vec!["shape(6)"]
    );
}
#[test]
fn instanceof_checks_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Tagged { public function tag(): string; }
class Item implements Tagged { public function tag(): string { return 'item'; } }
$o = new Item;
echo ($o instanceof Tagged) ? 'yes' : 'no';
"#
        ),
        vec!["yes"]
    );
}

// ── Interface constants ───────────────────────────────────────

#[test]
fn interface_constant_accessible_via_class() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Config { const VERSION = '1.0'; }
class App implements Config {}
echo App::VERSION;
"#
        ),
        vec!["1.0"]
    );
}
#[test]
fn interface_constant_accessible_via_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Limits { const MAX = 100; const MIN = 0; }
echo Limits::MAX - Limits::MIN;
"#
        ),
        vec!["100"]
    );
}

// ── Abstract classes ──────────────────────────────────────────

#[test]
fn abstract_class_partial_implementation() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Animal {
    abstract public function sound(): string;
    public function describe(): string { return get_class($this) . ' says ' . $this->sound(); }
}
class Dog extends Animal { public function sound(): string { return 'woof'; } }
class Cat extends Animal { public function sound(): string { return 'meow'; } }
echo (new Dog)->describe() . ',' . (new Cat)->describe();
"#
        ),
        vec!["Dog says woof,Cat says meow"]
    );
}
#[test]
fn abstract_class_cannot_instantiate() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Base { abstract public function run(): void; }
try { new Base; } catch (Error $e) { echo 'cannot'; }
"#
        ),
        vec!["cannot"]
    );
}
#[test]
fn abstract_class_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Service {
    public function __construct(protected string $name) {}
    abstract public function run(): string;
}
class MyService extends Service {
    public function run(): string { return 'running: ' . $this->name; }
}
echo (new MyService('worker'))->run();
"#
        ),
        vec!["running: worker"]
    );
}

// ── Interface with default constant in PHP 8.1+ ───────────────

#[test]
fn interface_implemented_constant_override() {
    assert_eq!(
        run_prints(
            r#"<?php
interface HasVersion { const VERSION = '1.0'; }
class App implements HasVersion { const VERSION = '2.0'; }
echo App::VERSION . ',' . HasVersion::VERSION;
"#
        ),
        vec!["2.0,1.0"]
    );
}

// ── Duck typing with interface ────────────────────────────────

#[test]
fn interface_used_in_array_of_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Worker { public function work(): int; }
class DevWorker implements Worker { public function work(): int { return 8; } }
class OpWorker implements Worker { public function work(): int { return 10; } }
$workers = [new DevWorker, new OpWorker, new DevWorker];
echo array_sum(array_map(fn(Worker $w) => $w->work(), $workers));
"#
        ),
        vec!["26"]
    );
}

#[test]
fn interface_return_type_covariance_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Creator {
    public function build(): object;
}
class A {}
class B extends A {}
class Factory implements Creator {
    public function build(): B {
        return new B();
    }
}
$f = new Factory();
echo $f->build() instanceof B ? 'b' : 'n';
"#,
        ),
        vec!["b"]
    );
}

#[test]
fn interface_static_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Api {
    public static function name(): string;
}
class Service implements Api {
    public static function name(): string { return 'Service'; }
}
echo Service::name();
"#,
        ),
        vec!["Service"]
    );
}

#[test]
fn interface_in_array_sum_via_mapped_typecheck() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Priced {
    public function price(): float;
}
class Product implements Priced { public function __construct(private float $p) {} public function price(): float { return $this->p; } }
$items = [new Product(1.2), new Product(2.8)];
$total = array_sum(array_map(fn(Priced $x) => $x->price(), $items));
echo $total;
"#,
        ),
        vec!["4"]
    );
}

#[test]
fn interface_implements_multiple_via_aliasing_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface A { public function a(): string; }
interface B { public function b(): string; }
class C implements A, B {
    public function a(): string { return 'a'; }
    public function b(): string { return 'b'; }
}
$c = new C();
echo $c->a() . $c->b();
"#,
        ),
        vec!["ab"]
    );
}

#[test]
fn interface_as_function_argument_dispatch_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Handler { public function handle(string $in): string; }
class Upper implements Handler { public function handle(string $in): string { return strtoupper($in); } }
class Lower implements Handler { public function handle(string $in): string { return strtolower($in); } }
function run_handler(Handler $h, string $in): string { return $h->handle($in); }
echo run_handler(new Upper(), 'abc') . '|' . run_handler(new Lower(), 'ABC');
"#,
        ),
        vec!["ABC|abc"]
    );
}

#[test]
fn interface_instanceof_checks_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Marked {}
class Item implements Marked {}
class NoMark {}
echo (new Item) instanceof Marked ? 'yes' : 'no';
echo '|';
echo (new NoMark) instanceof Marked ? 'yes' : 'no';
"#,
        ),
        vec!["yes|no"]
    );
}

#[test]
fn interface_default_method_simulated_by_trait_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Handler {
    public function call(string $v): string;
}
trait DefaultHandler {
    public function toUpper(string $v): string { return strtoupper($v); }
}
class Router implements Handler {
    use DefaultHandler;
    public function call(string $v): string { return $this->toUpper($v); }
}
echo (new Router())->call('ok');
"#,
        ),
        vec!["OK"]
    );
}

#[test]
fn interface_const_aliasing_from_implementer_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Versioned { public const VERSION = '1'; }
class Plugin implements Versioned {
    public const VERSION = '2';
}
echo Plugin::VERSION;
echo '|';
echo Versioned::VERSION;
"#,
        ),
        vec!["2|1"]
    );
}

#[test]
fn interface_with_nullable_return_union_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Finder {
    public function find(string $k): ?string;
}
class Store implements Finder {
    public function find(string $k): ?string { return $k === 'exists' ? 'yes' : null; }
}
$s = new Store();
echo $s->find('exists') ?? 'miss';
echo '|';
echo $s->find('other') ?? 'miss';
"#,
        ),
        vec!["yes|miss"]
    );
}

#[test]
fn interface_inheritance_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Base { public function base(): string; }
interface Mid extends Base { public function mid(): string; }
interface Top extends Mid { public function top(): string; }
class Impl implements Top {
    public function base(): string { return 'b'; }
    public function mid(): string { return 'm'; }
    public function top(): string { return 't'; }
}
$obj = new Impl();
echo $obj->base() . $obj->mid() . $obj->top();
"#,
        ),
        vec!["bmt"]
    );
}

#[test]
fn interface_call_order_and_implementation_graph_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Logger {
    public function emit(string $message): string;
}
interface Timestamped {
    public function emit(string $message): string;
}
class Audit implements Logger, Timestamped {
    public function emit(string $message): string {
        return 'log:' . $message;
    }
}
echo (new Audit)->emit('x');
"#,
        ),
        vec!["log:x"]
    );
}

#[test]
fn interface_default_with_covariant_returns_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Provider {
    public function payload(): iterable;
}
class StringBag implements Provider {
    public function payload(): array {
        return ['a', 'b'];
    }
}
$p = new StringBag();
echo json_encode($p->payload());
"#,
        ),
        vec!["[\"a\",\"b\"]"]
    );
}

#[test]
fn interface_private_constructor_contract_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Buildable {
    public static function build(int $v): static;
}
class Node implements Buildable {
    private function __construct(private int $v) {}
    public static function build(int $v): static {
        return new static($v);
    }
    public function value(): int { return $this->v; }
}
echo (Node::build(8))->value();
"#,
        ),
        vec!["8"]
    );
}

#[test]
fn interface_exists_for_declared_name_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface ExistsDemo {
    public function ping(): string;
}
echo interface_exists('ExistsDemo') ? 'yes' : 'no';
echo '|';
echo interface_exists('DoesNotExist') ? 'yes' : 'no';
"#,
        ),
        vec!["yes|no"]
    );
}

#[test]
fn class_implements_for_object_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Reader {
    public function read(): string;
}
interface Writer {
    public function write(string $v): void;
}
class Logger implements Reader, Writer {
    public function read(): string { return 'r'; }
    public function write(string $v): void { $this->value = $v; }
    private string $value = '';
    public function value(): string { return $this->value; }
}
$logger = new Logger();
$impl = class_implements($logger);
echo (isset($impl[Reader::class]) ? 'R' : '?') . (isset($impl[Writer::class]) ? 'W' : '?');
"#,
        ),
        vec!["RW"]
    );
}

#[test]
fn interface_runtime_graph_with_parent_class_and_interface_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface A { public function a(): string; }
interface B extends A { public function b(): string; }
class ParentSvc {
    public function p(): string { return 'p'; }
}
class ChildSvc extends ParentSvc implements B {
    public function a(): string { return 'a'; }
    public function b(): string { return 'b'; }
}
$o = new ChildSvc();
echo $o->a() . $o->b() . $o->p();
echo '|';
echo ($o instanceof B) ? 'ok' : 'bad';
"#,
        ),
        vec!["abp|ok"]
    );
}

#[test]
fn interface_static_method_called_via_class_implements_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Identifiable {
    public static function label(): string;
}
class Widget implements Identifiable {
    public static function label(): string { return 'widget'; }
}
echo Widget::label();
echo '|';
$ifaces = class_implements(Widget::class);
echo isset($ifaces[Identifiable::class]) ? 'seen' : 'missing';
"#,
        ),
        vec!["widget|seen"]
    );
}

#[test]
fn interface_method_dispatch_with_object_storage_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Strategy {
    public function execute(int $x, int $y): int;
}
class Add implements Strategy {
    public function execute(int $x, int $y): int { return $x + $y; }
}
class Mul implements Strategy {
    public function execute(int $x, int $y): int { return $x * $y; }
}
function run_all(array $items): int {
    $total = 0;
    foreach ($items as $item) { $total += $item->execute(3, 4); }
    return $total;
}
$items = [new Add(), new Mul()];
echo run_all($items);
"#,
        ),
        vec!["19"]
    );
}

#[test]
fn interface_default_like_methods_in_runtime_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Escaper {
    public function escape(string $v): string;
}
class HtmlSafe implements Escaper {
    public function escape(string $v): string {
        return str_replace('<', '&lt;', $v);
    }
}
class NoEscape implements Escaper {
    public function escape(string $v): string {
        return $v;
    }
}
function render(Escaper $e, string $v): string {
    return $e->escape($v);
}
echo render(new HtmlSafe, "<x>") . '|' . render(new NoEscape, "<x>");
"#,
        ),
        vec!["&lt;x>|<x>"]
    );
}
