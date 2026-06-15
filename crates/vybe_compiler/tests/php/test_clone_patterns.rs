use super::helpers::run_prints;

// ── Basic clone ───────────────────────────────────────────────

#[test]
fn clone_creates_new_object_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box { public int $val = 0; }
$a = new Box();
$a->val = 42;
$b = clone $a;
echo ($a === $b) ? 'same' : 'different';
"#
        ),
        vec!["different"]
    );
}

#[test]
fn clone_shallow_copies_scalar_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point { public function __construct(public int $x, public int $y) {} }
$p = new Point(3, 4);
$q = clone $p;
echo $q->x . ',' . $q->y;
"#
        ),
        vec!["3,4"]
    );
}

#[test]
fn clone_modifying_clone_does_not_affect_original() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config { public string $host = 'localhost'; }
$orig = new Config();
$copy = clone $orig;
$copy->host = 'remote';
echo $orig->host . ',' . $copy->host;
"#
        ),
        vec!["localhost,remote"]
    );
}

// ── Shallow clone — object properties are shared ──────────────

#[test]
fn shallow_clone_shares_nested_object_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
class Inner { public int $val = 0; }
class Outer { public Inner $inner; public function __construct() { $this->inner = new Inner(); } }
$a = new Outer();
$a->inner->val = 10;
$b = clone $a;
$b->inner->val = 99;
echo $a->inner->val;
"#
        ),
        vec!["99"]
    );
}

// ── __clone magic method ──────────────────────────────────────

#[test]
fn clone_magic_method_called_on_clone() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public int $copies = 0;
    public function __clone() { $this->copies++; }
}
$a = new Counter();
$b = clone $a;
echo $b->copies;
"#
        ),
        vec!["1"]
    );
}

#[test]
fn clone_magic_method_deep_clones_nested_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class Address { public string $city; public function __construct(string $city) { $this->city = $city; } }
class Person {
    public Address $address;
    public function __construct(public string $name, Address $addr) { $this->address = $addr; }
    public function __clone() { $this->address = clone $this->address; }
}
$alice = new Person("Alice", new Address("London"));
$bob = clone $alice;
$bob->address->city = "Paris";
echo $alice->address->city . ',' . $bob->address->city;
"#
        ),
        vec!["London,Paris"]
    );
}

#[test]
fn clone_magic_method_resets_id() {
    assert_eq!(
        run_prints(
            r#"<?php
class Entity {
    private static int $nextId = 1;
    public int $id;
    public function __construct() { $this->id = self::$nextId++; }
    public function __clone() { $this->id = self::$nextId++; }
}
$a = new Entity();
$b = clone $a;
echo $a->id . ',' . $b->id;
"#
        ),
        vec!["1,2"]
    );
}

// ── Clone in inheritance ──────────────────────────────────────

#[test]
fn clone_child_class_creates_child_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal { public string $type = 'animal'; }
class Dog extends Animal { public string $breed = 'labrador'; }
$d = new Dog();
$c = clone $d;
echo get_class($c) . ',' . $c->breed;
"#
        ),
        vec!["Dog,labrador"]
    );
}

#[test]
fn parent_clone_called_from_child_clone() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public array $tags = [];
    public function __clone() { $this->tags[] = 'base_cloned'; }
}
class Child extends Base {
    public function __clone() {
        parent::__clone();
        $this->tags[] = 'child_cloned';
    }
}
$c = new Child();
$d = clone $c;
echo implode(',', $d->tags);
"#
        ),
        vec!["base_cloned,child_cloned"]
    );
}

// ── Deep clone pattern ────────────────────────────────────────

#[test]
fn deep_clone_array_of_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
class Item { public function __construct(public int $id) {} }
class Cart {
    public array $items = [];
    public function __clone() {
        $this->items = array_map(fn($i) => clone $i, $this->items);
    }
}
$cart = new Cart();
$cart->items[] = new Item(1);
$cart->items[] = new Item(2);
$copy = clone $cart;
$copy->items[0]->id = 99;
echo $cart->items[0]->id . ',' . $copy->items[0]->id;
"#
        ),
        vec!["1,99"]
    );
}

// ── Clone with readonly (PHP 8.1+) ────────────────────────────

#[test]
fn clone_readonly_property_allowed_in_clone_magic() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {
    public readonly string $value;
    public function __construct(string $v) { $this->value = $v; }
    public function withValue(string $v): static {
        $clone = clone $this;
        return $clone;
    }
}
$t = new Token("abc");
$t2 = $t->withValue("xyz");
echo $t->value;
"#
        ),
        vec!["abc"]
    );
}

// ── Clone in value object pattern ────────────────────────────

#[test]
fn clone_with_modifier_returns_new_value_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class Money {
    public function __construct(private int $cents, private string $currency) {}
    public function add(int $cents): self {
        $new = clone $this;
        $new = new self($this->cents + $cents, $this->currency);
        return $new;
    }
    public function amount(): int { return $this->cents; }
    public function currency(): string { return $this->currency; }
}
$price = new Money(1000, 'USD');
$total = $price->add(500);
echo $price->amount() . ',' . $total->amount();
"#
        ),
        vec!["1000,1500"]
    );
}

// ── Clone preserves private properties ───────────────────────

#[test]
fn clone_copies_private_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class Secret {
    private string $key;
    public function __construct(string $k) { $this->key = $k; }
    public function getKey(): string { return $this->key; }
}
$a = new Secret("mykey");
$b = clone $a;
echo $b->getKey();
"#
        ),
        vec!["mykey"]
    );
}

// ── Clone of object with resource ────────────────────────────

#[test]
fn clone_with_null_resource_reset_in_magic() {
    assert_eq!(
        run_prints(
            r#"<?php
class Connection {
    public ?string $handle = null;
    public function connect(): void { $this->handle = "connected"; }
    public function __clone() { $this->handle = null; }
}
$c = new Connection();
$c->connect();
$d = clone $c;
echo $c->handle . ',' . var_export($d->handle, true);
"#
        ),
        vec!["connected,NULL"]
    );
}

// ── Clone in builder pattern ──────────────────────────────────

#[test]
fn builder_clone_immutable_with_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
class Query {
    private array $filters = [];
    public function where(string $filter): static {
        $new = clone $this;
        $new->filters[] = $filter;
        return $new;
    }
    public function build(): string { return implode(' AND ', $this->filters); }
}
$base = new Query();
$q1 = $base->where('a=1')->where('b=2');
$q2 = $base->where('c=3');
echo $q1->build() . '|' . $q2->build();
"#
        ),
        vec!["a=1 AND b=2|c=3"]
    );
}

// ── Clone verifies object equality ───────────────────────────

#[test]
fn cloned_objects_are_equal_but_not_identical() {
    assert_eq!(
        run_prints(
            r#"<?php
class Tag { public string $name; public function __construct(string $n) { $this->name = $n; } }
$a = new Tag("php");
$b = clone $a;
echo ($a == $b ? 'equal' : 'not equal') . ',' . ($a === $b ? 'identical' : 'not identical');
"#
        ),
        vec!["equal,not identical"]
    );
}

// ── Chained cloning ───────────────────────────────────────────

#[test]
fn chain_of_clones_each_independent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node { public int $val; public function __construct(int $v) { $this->val = $v; } }
$a = new Node(1);
$b = clone $a; $b->val = 2;
$c = clone $b; $c->val = 3;
echo $a->val . ',' . $b->val . ',' . $c->val;
"#
        ),
        vec!["1,2,3"]
    );
}

// ── Clone with static property unchanged ─────────────────────

#[test]
fn static_property_not_cloned_per_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter { public static int $total = 0; public int $id; public function __construct() { $this->id = ++self::$total; } }
$a = new Counter();
$b = clone $a;
echo Counter::$total . ',' . $b->id;
"#
        ),
        vec!["1,1"]
    );
}

// ── Clone array containing objects ───────────────────────────

#[test]
fn array_of_cloned_objects_are_independent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Val { public function __construct(public int $n) {} }
$originals = [new Val(1), new Val(2), new Val(3)];
$clones = array_map(fn($o) => clone $o, $originals);
$clones[0]->n = 99;
echo $originals[0]->n . ',' . $clones[0]->n;
"#
        ),
        vec!["1,99"]
    );
}

// ── Clone with __clone and parent class property ──────────────

#[test]
fn clone_magic_has_access_to_parent_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class BaseModel { protected string $createdAt = '2024-01-01'; }
class User extends BaseModel {
    public string $name;
    public function __construct(string $name) { $this->name = $name; }
    public function __clone() { $this->createdAt = '2024-06-01'; }
    public function getCreated(): string { return $this->createdAt; }
}
$u = new User("Alice");
$v = clone $u;
echo $u->getCreated() . ',' . $v->getCreated();
"#
        ),
        vec!["2024-01-01,2024-06-01"]
    );
}

// ── Serialize then unserialize acts like deep clone ───────────

#[test]
fn serialize_unserialize_produces_independent_copy() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node { public function __construct(public int $val, public ?Node $next = null) {} }
$list = new Node(1, new Node(2, new Node(3)));
$copy = unserialize(serialize($list));
$copy->next->val = 99;
echo $list->next->val . ',' . $copy->next->val;
"#
        ),
        vec!["2,99"]
    );
}
