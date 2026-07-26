use super::helpers::run_prints;

// ── Method chaining and fluent interfaces ─────────────────────

#[test]
fn fluent_validation_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class Validator {
    private array $errors = [];
    private mixed $value;
    public function __construct(mixed $v) { $this->value = $v; }
    public function required(): static { if (empty($this->value)) $this->errors[] = 'required'; return $this; }
    public function minLength(int $n): static { if (strlen($this->value) < $n) $this->errors[] = "min:$n"; return $this; }
    public function isValid(): bool { return empty($this->errors); }
    public function errors(): array { return $this->errors; }
}
$v = new Validator('hi');
$v->required()->minLength(5);
echo $v->isValid() ? 'valid' : implode(',', $v->errors());
"#
        ),
        vec!["min:5"]
    );
}

// ── Object serialization ──────────────────────────────────────

#[test]
fn serialize_and_unserialize_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point { public function __construct(public int $x, public int $y) {} }
$p = new Point(3, 4);
$s = serialize($p);
$p2 = unserialize($s);
echo $p2->x . ',' . $p2->y;
"#
        ),
        vec!["3,4"]
    );
}
#[test]
fn serialize_sleep_wakeup() {
    assert_eq!(
        run_prints(
            r#"<?php
class Conn {
    public string $host = 'localhost';
    private mixed $resource = null;
    public function __sleep(): array { return ['host']; }
    public function __wakeup(): void { $this->resource = 'reconnected'; }
    public function status(): string { return $this->resource ?? 'null'; }
}
$c = new Conn;
$c2 = unserialize(serialize($c));
echo $c2->host . ':' . $c2->status();
"#
        ),
        vec!["localhost:reconnected"]
    );
}

// ── Abstract class patterns ───────────────────────────────────

#[test]
fn template_method_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class DataProcessor {
    final public function process(array $data): array {
        $data = $this->filter($data);
        $data = $this->transform($data);
        return $this->sort($data);
    }
    abstract protected function filter(array $d): array;
    abstract protected function transform(array $d): array;
    protected function sort(array $d): array { sort($d); return $d; }
}
class EvenDoubler extends DataProcessor {
    protected function filter(array $d): array { return array_filter($d, fn($n) => $n % 2 === 0); }
    protected function transform(array $d): array { return array_map(fn($n) => $n * 2, $d); }
}
echo implode(',', (new EvenDoubler)->process([1,2,3,4,5,6]));
"#
        ),
        vec!["4,8,12"]
    );
}

// ── Interface segregation ─────────────────────────────────────

#[test]
fn segregated_interfaces() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Readable { public function read(): string; }
interface Writable { public function write(string $data): void; }
interface ReadWrite extends Readable, Writable {}
class Buffer implements ReadWrite {
    private string $buffer = '';
    public function read(): string { return $this->buffer; }
    public function write(string $data): void { $this->buffer .= $data; }
}
$b = new Buffer;
$b->write('hello');
$b->write(' world');
echo $b->read();
"#
        ),
        vec!["hello world"]
    );
}

// ── Mixin simulation with traits ──────────────────────────────

#[test]
fn mixin_via_trait() {
    assert_eq!(
        run_prints(
            r#"<?php
trait HasUuid {
    private string $uuid;
    public function initUuid(): void { $this->uuid = sprintf('%08x-%04x-%04x', 1, 2, 3); }
    public function getUuid(): string { return $this->uuid; }
}
class Entity { use HasUuid; }
$e = new Entity; $e->initUuid();
echo str_contains($e->getUuid(), '-') ? 'has-uuid' : 'no';
"#
        ),
        vec!["has-uuid"]
    );
}

// ── Polymorphism ──────────────────────────────────────────────

#[test]
fn polymorphic_method_dispatch() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Shape {
    abstract public function area(): float;
    public function describe(): string { return get_class($this) . ':' . $this->area(); }
}
class Rect extends Shape { public function __construct(private float $w, private float $h) {} public function area(): float { return $this->w * $this->h; } }
class Triangle extends Shape { public function __construct(private float $b, private float $h) {} public function area(): float { return 0.5 * $this->b * $this->h; } }
$shapes = [new Rect(4,3), new Triangle(6,4)];
echo implode(',', array_map(fn($s) => $s->describe(), $shapes));
"#
        ),
        vec!["Rect:12,Triangle:12"]
    );
}

// ── Object graph traversal ────────────────────────────────────

#[test]
fn tree_node_recursive() {
    assert_eq!(
        run_prints(
            r#"<?php
class TreeNode {
    public array $children = [];
    public function __construct(public int $value) {}
    public function add(TreeNode $n): void { $this->children[] = $n; }
    public function sum(): int {
        return $this->value + array_sum(array_map(fn($c) => $c->sum(), $this->children));
    }
}
$root = new TreeNode(1);
$root->add(new TreeNode(2));
$right = new TreeNode(3);
$right->add(new TreeNode(4));
$root->add($right);
echo $root->sum();
"#
        ),
        vec!["10"]
    );
}

// ── Interface with static factory ────────────────────────────

#[test]
fn named_constructor_via_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface HasNamedConstructors {
    public static function empty(): static;
}
class Collection implements HasNamedConstructors {
    private function __construct(private array $items = []) {}
    public static function empty(): static { return new static; }
    public function count(): int { return count($this->items); }
}
echo Collection::empty()->count();
"#
        ),
        vec!["0"]
    );
}

// ── Property hooks / accessors via magic ─────────────────────

#[test]
fn computed_property_via_getter() {
    assert_eq!(
        run_prints(
            r#"<?php
class Circle {
    public function __construct(public float $radius) {}
    public function __get(string $name): mixed {
        return match($name) {
            'area' => M_PI * $this->radius ** 2,
            'circumference' => 2 * M_PI * $this->radius,
            default => null,
        };
    }
}
$c = new Circle(5.0);
echo round($c->area, 2) . ',' . round($c->circumference, 2);
"#
        ),
        vec!["78.54,31.42"]
    );
}

// ── Variance in generics simulation ──────────────────────────

#[test]
fn covariant_container_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public function __construct(private mixed $value) {}
    public function get(): mixed { return $this->value; }
    public function map(callable $fn): static { return new static($fn($this->value)); }
}
$result = (new Box(5))->map(fn($n) => $n * 2)->map(fn($n) => "value:$n")->get();
echo $result;
"#
        ),
        vec!["value:10"]
    );
}

#[test]
fn dynamic_dispatch_to_callables_on_trait_object() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Executor {
    public function execute(string $task): string;
}
class SyncExecutor implements Executor {
    public function execute(string $task): string { return "sync:$task"; }
}
class AsyncLikeExecutor implements Executor {
    public function execute(string $task): string { return "async:$task"; }
}
$executors = [new SyncExecutor(), new AsyncLikeExecutor()];
echo $executors[0]->execute('build') . '|' . $executors[1]->execute('build');
"#
        ),
        vec!["sync:build|async:build"]
    );
}

#[test]
fn constructor_visibility_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public function __construct(public string $name) {}
}
class Child extends Base {
    public function __construct(string $name, public int $id) {
        parent::__construct($name);
    }
}
$child = new Child('worker', 7);
echo $child->name . ':' . $child->id;
"#
        ),
        vec!["worker:7"]
    );
}

#[test]
fn static_property_per_class_hierarchy() {
    assert_eq!(
        run_prints(
            r#"<?php
class CounterA {
    public static int $count = 1;
    public static function bump(): void { self::$count += 1; }
}
class CounterB {
    public static int $count = 10;
    public static function bump(): void { self::$count += 10; }
}
CounterA::bump();
CounterB::bump();
echo CounterA::$count . '|' . CounterB::$count;
"#
        ),
        vec!["2|20"]
    );
}

#[test]
fn property_hooks_getter_setter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    private int $count = 0;
    public int $value {
        get { return $this->count; }
        set { $this->count = $value; }
    }
}
$c = new Counter();
$c->value = 9;
echo $c->value;
"#
        ),
        vec!["9"]
    );
}

#[test]
fn object_clone_preserves_immutability_by_design() {
    assert_eq!(
        run_prints(
            r#"<?php
class Snapshot {
    public function __construct(public readonly string $label) {}
    public function __clone(): void {}
}
$s = new Snapshot('v');
$c = clone $s;
echo $s->label . '|' . $c->label;
"#
        ),
        vec!["v|v"]
    );
}

#[test]
fn __call_static_dispatch_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class HandlerRegistry {
    public static function __callStatic(string $name, array $args): mixed {
        return match($name) {
            'make' => 'created:' . ($args[0] ?? 'default'),
            default => null,
        };
    }
}
echo HandlerRegistry::make('report');
"#
        ),
        vec!["created:report"]
    );
}
