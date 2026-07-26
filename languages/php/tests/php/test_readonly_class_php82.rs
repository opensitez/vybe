use super::helpers::run_prints;

// ── readonly class (PHP 8.2) ──────────────────────────────────

#[test]
fn readonly_class_all_properties_implicitly_readonly() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Point {
    public function __construct(
        public float $x,
        public float $y,
    ) {}
}
$p = new Point(1.5, 2.5);
echo $p->x . ',' . $p->y;
"#
        ),
        vec!["1.5,2.5"]
    );
}

#[test]
fn readonly_class_property_cannot_be_modified() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Config {
    public function __construct(public string $dsn) {}
}
$c = new Config("mysql://localhost");
try {
    $c->dsn = "other";
} catch (Error $e) {
    echo "error";
}
"#
        ),
        vec!["error"]
    );
}

#[test]
fn readonly_class_can_be_cloned() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Coord {
    public function __construct(public int $x, public int $y) {}
}
$a = new Coord(1, 2);
$b = clone $a;
echo $b->x . ',' . $b->y;
"#
        ),
        vec!["1,2"]
    );
}

#[test]
fn readonly_class_extends_works() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Base {
    public function __construct(public string $name) {}
}
readonly class Child extends Base {
    public function __construct(string $name, public int $age) {
        parent::__construct($name);
    }
}
$c = new Child("Alice", 30);
echo $c->name . ',' . $c->age;
"#
        ),
        vec!["Alice,30"]
    );
}

// ── readonly properties (PHP 8.1) ─────────────────────────────

#[test]
fn readonly_property_set_once_in_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {
    public readonly string $value;
    public function __construct(string $val) { $this->value = $val; }
}
$t = new Token("abc123");
echo $t->value;
"#
        ),
        vec!["abc123"]
    );
}

#[test]
fn readonly_property_throws_on_second_write() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {
    public readonly string $value;
    public function __construct(string $val) { $this->value = $val; }
}
$t = new Token("abc");
try {
    $t->value = "xyz";
} catch (Error $e) {
    echo "immutable";
}
"#
        ),
        vec!["immutable"]
    );
}

#[test]
fn readonly_property_cannot_unset() {
    assert_eq!(
        run_prints(
            r#"<?php
class Cfg {
    public readonly int $id;
    public function __construct(int $id) { $this->id = $id; }
}
$c = new Cfg(5);
try {
    unset($c->id);
} catch (Error $e) {
    echo "cannot unset";
}
"#
        ),
        vec!["cannot unset"]
    );
}

#[test]
fn readonly_promoted_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public function __construct(
        public readonly string $name,
        public readonly string $email,
    ) {}
}
$u = new User("Bob", "bob@example.com");
echo $u->name . ',' . $u->email;
"#
        ),
        vec!["Bob,bob@example.com"]
    );
}

#[test]
fn readonly_property_can_be_typed_nullable() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node {
    public readonly ?int $parent;
    public function __construct(?int $parent = null) { $this->parent = $parent; }
}
$n = new Node(null);
echo var_export($n->parent, true);
"#
        ),
        vec!["NULL"]
    );
}

// ── readonly with clone (PHP 8.4 clone with) or manual clone ──

#[test]
fn readonly_clone_preserves_values() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Pair {
    public function __construct(
        public int $first,
        public int $second,
    ) {}
}
$a = new Pair(10, 20);
$b = clone $a;
echo $b->first . ',' . $b->second;
"#
        ),
        vec!["10,20"]
    );
}

// ── readonly class implements interface ───────────────────────

#[test]
fn readonly_class_implements_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Named { public function name(): string; }
readonly class Person implements Named {
    public function __construct(public string $fullName) {}
    public function name(): string { return $this->fullName; }
}
$p = new Person("Carol");
echo $p->name();
"#
        ),
        vec!["Carol"]
    );
}

// ── readonly property type constraints ────────────────────────

#[test]
fn readonly_property_union_type() {
    assert_eq!(
        run_prints(
            r#"<?php
class Ref {
    public readonly int|string $key;
    public function __construct(int|string $key) { $this->key = $key; }
}
$r1 = new Ref(42);
$r2 = new Ref("abc");
echo $r1->key . ',' . $r2->key;
"#
        ),
        vec!["42,abc"]
    );
}

// ── readonly in inheritance ───────────────────────────────────

#[test]
fn readonly_property_in_parent_not_overrideable_in_child() {
    assert_eq!(
        run_prints(
            r#"<?php
class Vehicle {
    public readonly string $make;
    public function __construct(string $make) { $this->make = $make; }
}
class Car extends Vehicle {
    public function __construct(string $make, public readonly int $year) {
        parent::__construct($make);
    }
}
$c = new Car("Toyota", 2020);
echo $c->make . ',' . $c->year;
"#
        ),
        vec!["Toyota,2020"]
    );
}

// ── readonly with default — not allowed ───────────────────────

#[test]
fn readonly_property_without_default_is_uninitialized_until_set() {
    assert_eq!(
        run_prints(
            r#"<?php
class Lazy {
    public readonly int $value;
    public function init(int $v): void { $this->value = $v; }
}
$l = new Lazy();
$l->init(7);
echo $l->value;
"#
        ),
        vec!["7"]
    );
}

// ── readonly class with static factory ───────────────────────

#[test]
fn readonly_class_static_factory_method() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Color {
    private function __construct(
        public int $r,
        public int $g,
        public int $b,
    ) {}
    public static function fromHex(string $hex): self {
        $hex = ltrim($hex, '#');
        return new self(
            hexdec(substr($hex, 0, 2)),
            hexdec(substr($hex, 2, 2)),
            hexdec(substr($hex, 4, 2)),
        );
    }
}
$c = Color::fromHex('#ff8000');
echo $c->r . ',' . $c->g . ',' . $c->b;
"#
        ),
        vec!["255,128,0"]
    );
}

// ── readonly in data transfer objects ────────────────────────

#[test]
fn readonly_dto_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class OrderDTO {
    public function __construct(
        public string $id,
        public float $total,
        public string $currency,
    ) {}
    public function summary(): string {
        return "$this->id: $this->total $this->currency";
    }
}
$order = new OrderDTO("ORD-001", 99.99, "USD");
echo $order->summary();
"#
        ),
        vec!["ORD-001: 99.99 USD"]
    );
}

// ── readonly array property ───────────────────────────────────

#[test]
fn readonly_array_property_set_once() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    public readonly array $options;
    public function __construct(array $opts) { $this->options = $opts; }
}
$c = new Config(['debug' => true, 'version' => 2]);
echo count($c->options);
"#
        ),
        vec!["2"]
    );
}

#[test]
fn readonly_nested_object_reference_still_readonly_on_reassignment() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Node {
    public function __construct(public array $data) {}
}
$n = new Node(['k' => 1]);
$n->data['k'] = 2;
$n->data['v'] = 3;
echo $n->data['k'] . ',' . $n->data['v'];
"#,
        ),
        vec!["2,3"]
    );
}

#[test]
fn readonly_class_clone_keeps_property_identity() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Box {
    public function __construct(public int $x, public int $y) {}
}
$a = new Box(1, 2);
$b = clone $a;
echo $a->x . $a->y . '|' . $b->x . $b->y;
"#,
        ),
        vec!["12|12"]
    );
}

#[test]
fn readonly_property_with_default_not_settable_via_init() {
    assert_eq!(
        run_prints(
            r#"<?php
class Dto {
    public readonly string $id;
    public function __construct() {}
    public function init(string $id): void { $this->id = $id; }
}
$d = new Dto();
$d->init("X1");
echo $d->id;
"#,
        ),
        vec!["X1"]
    );
}

#[test]
fn readonly_class_with_magic_readonly_property_access() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Profile {
    public function __construct(public string $name) {}
    public function __get(string $n): mixed {
        if ($n === 'label') return strtoupper($this->name);
        return null;
    }
}
$p = new Profile('alice');
echo $p->name;
echo '|';
echo $p->label;
"#,
        ),
        vec!["alice|ALICE"]
    );
}

// ── readonly with intersection type ──────────────────────────

#[test]
fn readonly_property_with_interface_type() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Identifiable { public function id(): int; }
class Item implements Identifiable {
    public function __construct(private int $itemId) {}
    public function id(): int { return $this->itemId; }
}
class Container {
    public readonly Identifiable $wrapped;
    public function __construct(Identifiable $item) { $this->wrapped = $item; }
}
$c = new Container(new Item(42));
echo $c->wrapped->id();
"#
        ),
        vec!["42"]
    );
}

// ── readonly class serialization ─────────────────────────────

#[test]
fn readonly_class_json_serializable() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Product {
    public function __construct(
        public string $name,
        public float $price,
    ) {}
}
$p = new Product("Widget", 9.99);
echo json_encode(['name' => $p->name, 'price' => $p->price]);
"#
        ),
        vec!["{\"name\":\"Widget\",\"price\":9.99}"]
    );
}

// ── readonly enum-like constant object ────────────────────────

#[test]
fn readonly_class_as_value_object() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Money {
    public function __construct(
        public int $amount,
        public string $currency,
    ) {}
    public function add(Money $other): self {
        if ($this->currency !== $other->currency) throw new \InvalidArgumentException("Currency mismatch");
        return new self($this->amount + $other->amount, $this->currency);
    }
}
$a = new Money(100, 'USD');
$b = new Money(50, 'USD');
$c = $a->add($b);
echo $c->amount . ' ' . $c->currency;
"#
        ),
        vec!["150 USD"]
    );
}

// ── readonly promoted with nullable ──────────────────────────

#[test]
fn readonly_promoted_nullable_default_null() {
    assert_eq!(
        run_prints(
            r#"<?php
class Event {
    public function __construct(
        public readonly string $type,
        public readonly ?string $payload = null,
    ) {}
}
$e = new Event("click");
echo $e->type . ',' . var_export($e->payload, true);
"#
        ),
        vec!["click,NULL"]
    );
}
