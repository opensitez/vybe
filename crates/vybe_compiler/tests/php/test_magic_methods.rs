use super::helpers::{compile_ok, run_prints};

// ── __toString ───────────────────────────────────────────────────

#[test]
fn magic_tostring_in_echo() {
    assert_eq!(
        run_prints(
            r#"<?php
class Color {
    public function __construct(private string $name) {}
    public function __toString(): string {
        return $this->name;
    }
}
$c = new Color("red");
echo $c;
"#
        ),
        &["red"]
    );
}

#[test]
fn magic_tostring_in_concat() {
    assert_eq!(
        run_prints(
            r#"<?php
class Name {
    public function __construct(private string $val) {}
    public function __toString(): string { return $this->val; }
}
$n = new Name("world");
echo "hello " . $n;
"#
        ),
        &["hello world"]
    );
}

#[test]
fn magic_tostring_in_interpolation() {
    assert_eq!(
        run_prints(
            r#"<?php
class Version {
    public function __construct(private int $major, private int $minor) {}
    public function __toString(): string {
        return "$this->major.$this->minor";
    }
}
$v = new Version(3, 14);
echo "version: $v";
"#
        ),
        &["version: 3.14"]
    );
}

#[test]
fn magic_tostring_with_cast() {
    assert_eq!(
        run_prints(
            r#"<?php
class Amount {
    public function __construct(private float $value) {}
    public function __toString(): string {
        return number_format($this->value, 2);
    }
}
$a = new Amount(42.5);
echo (string) $a;
"#
        ),
        &["42.50"]
    );
}

#[test]
fn magic_tostring_chained_on_method_return() {
    assert_eq!(
        run_prints(
            r#"<?php
class Tag {
    public function __construct(private string $html) {}
    public function __toString(): string { return $this->html; }
    public function wrap(string $tag): self {
        return new self("<$tag>" . $this->html . "</$tag>");
    }
}
$t = new Tag("hello");
echo $t->wrap("b");
"#
        ),
        &["<b>hello</b>"]
    );
}

// ── __invoke ─────────────────────────────────────────────────────

#[test]
fn magic_invoke_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
class Multiplier {
    public function __construct(private int $factor) {}
    public function __invoke($x) {
        return $x * $this->factor;
    }
}
$double = new Multiplier(2);
echo $double(5);
echo $double(10);
"#
        ),
        &["1020"]
    );
}

#[test]
fn magic_invoke_with_parameters() {
    assert_eq!(
        run_prints(
            r#"<?php
class Formatter {
    public function __invoke(string $template, ...$args): string {
        return vsprintf($template, $args);
    }
}
$fmt = new Formatter();
echo $fmt("Hello, %s! You are %d years old.", "Alice", 30);
"#
        ),
        &["Hello, Alice! You are 30 years old."]
    );
}

#[test]
fn magic_invoke_as_callback() {
    assert_eq!(
        run_prints(
            r#"<?php
class Adder {
    public function __construct(private int $n) {}
    public function __invoke($x) { return $x + $this->n; }
}
$add5 = new Adder(5);
$nums = [1, 2, 3];
$result = array_map($add5, $nums);
echo implode(",", $result);
"#
        ),
        &["6,7,8"]
    );
}

#[test]
fn magic_invoke_stored_in_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
class Greeter {
    public function __invoke(string $name): string {
        return "Hello, $name!";
    }
}
$fn = new Greeter();
$call = $fn;
echo $call("Bob");
"#
        ),
        &["Hello, Bob!"]
    );
}

#[test]
fn magic_invoke_is_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
class Handler {
    public function __invoke($msg) { echo "handled: $msg"; }
}
$h = new Handler();
echo is_callable($h) ? "yes" : "no";
$h("test");
"#
        ),
        &["yeshandled: test"]
    );
}

// ── __get / __set ────────────────────────────────────────────────

#[test]
fn magic_get_returns_dynamic_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynProps {
    private array $data = [];
    public function __get($name) {
        return $this->data[$name] ?? "undefined";
    }
    public function __set($name, $value) {
        $this->data[$name] = $value;
    }
}
$obj = new DynProps();
$obj->foo = "bar";
echo $obj->foo;
echo $obj->missing;
"#
        ),
        &["barundefined"]
    );
}

#[test]
fn magic_set_stores_dynamic_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Bag {
    private array $items = [];
    public function __set($k, $v) { $this->items[$k] = $v; }
    public function __get($k) { return $this->items[$k] ?? null; }
    public function keys(): array { return array_keys($this->items); }
}
$b = new Bag();
$b->x = 10;
$b->y = 20;
$b->z = 30;
echo implode(",", $b->keys());
echo $b->y;
"#
        ),
        &["x,y,z20"]
    );
}

#[test]
fn magic_get_lazy_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class LazyLoader {
    private array $computed = [];
    public function __get($name) {
        if (!isset($this->computed[$name])) {
            $this->computed[$name] = strtoupper($name) . "_VALUE";
        }
        return $this->computed[$name];
    }
}
$l = new LazyLoader();
echo $l->foo;
echo $l->bar;
echo $l->foo;
"#
        ),
        &["FOO_VALUEBAR_VALUEFOO_VALUE"]
    );
}

#[test]
fn magic_set_validation() {
    assert_eq!(
        run_prints(
            r#"<?php
class Validated {
    private array $data = [];
    public function __set($name, $value) {
        if ($name === "age" && !is_int($value)) {
            throw new \InvalidArgumentException("age must be int");
        }
        $this->data[$name] = $value;
    }
    public function __get($name) { return $this->data[$name] ?? null; }
}
$v = new Validated();
$v->name = "Alice";
$v->age = 30;
echo $v->name;
echo $v->age;
try {
    $v->age = "thirty";
} catch (\InvalidArgumentException $e) {
    echo "caught";
}
"#
        ),
        &["Alice30caught"]
    );
}

#[test]
fn magic_get_set_read_back() {
    assert_eq!(
        run_prints(
            r#"<?php
class Registry {
    private array $store = [];
    public function __get($key) { return $this->store[$key] ?? null; }
    public function __set($key, $val) { $this->store[$key] = $val; }
    public function __isset($key) { return isset($this->store[$key]); }
}
$r = new Registry();
$r->a = 10;
$r->b = 20;
echo $r->a + $r->b;
echo isset($r->a) ? "set" : "not set";
echo isset($r->c) ? "set" : "not set";
"#
        ),
        &["30setnot set"]
    );
}

#[test]
fn magic_get_property_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    private array $data = ["db" => ["host" => "localhost"]];
    public function __get($key) {
        $val = $this->data[$key] ?? null;
        if (is_array($val)) {
            $child = new Config();
            foreach ($val as $k => $v) {
                $child->data[$k] = $v;
            }
            return $child;
        }
        return $val;
    }
    public function __set($key, $value) {
        $this->data[$key] = $value;
    }
}
$c = new Config();
echo $c->db->host;
"#
        ),
        &["localhost"]
    );
}

// ── __isset / __unset ────────────────────────────────────────────

#[test]
fn magic_isset_returns_true_for_existing() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    private array $data = ["x" => 1];
    public function __isset($name) {
        return isset($this->data[$name]);
    }
    public function __unset($name) {
        unset($this->data[$name]);
    }
}
$b = new Box();
echo isset($b->x) ? "yes" : "no";
echo isset($b->y) ? "yes" : "no";
unset($b->x);
echo isset($b->x) ? "yes" : "no";
"#
        ),
        &["yesnono"]
    );
}

#[test]
fn magic_isset_returns_false_when_absent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Sparse {
    private array $vals = ["a" => 1, "c" => 3];
    public function __isset($k) { return array_key_exists($k, $this->vals); }
}
$s = new Sparse();
echo isset($s->a) ? "yes" : "no";
echo isset($s->b) ? "yes" : "no";
echo isset($s->c) ? "yes" : "no";
"#
        ),
        &["yesnoyes"]
    );
}

#[test]
fn magic_unset_removes_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Store {
    private array $map = ["k1" => "v1", "k2" => "v2", "k3" => "v3"];
    public function __isset($k) { return isset($this->map[$k]); }
    public function __unset($k) { unset($this->map[$k]); }
    public function count(): int { return count($this->map); }
}
$s = new Store();
echo $s->count();
unset($s->k2);
echo $s->count();
echo isset($s->k2) ? "yes" : "no";
"#
        ),
        &["32no"]
    );
}

// ── __call / __callStatic ─────────────────────────────────────────

#[test]
fn magic_call_intercepts_undefined_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Proxy {
    public function __call($name, $args) {
        echo "called $name with " . count($args) . " args";
    }
}
$p = new Proxy();
$p->hello(1, 2, 3);
"#
        ),
        &["called hello with 3 args"]
    );
}

#[test]
fn magic_call_passes_name_and_args() {
    assert_eq!(
        run_prints(
            r#"<?php
class Logger {
    private array $log = [];
    public function __call($method, $args) {
        $this->log[] = "$method(" . implode(",", $args) . ")";
    }
    public function dump(): string { return implode("|", $this->log); }
}
$l = new Logger();
$l->info("a");
$l->warn("b", "c");
echo $l->dump();
"#
        ),
        &["info(a)|warn(b,c)"]
    );
}

#[test]
fn magic_call_static_intercepts() {
    assert_eq!(
        run_prints(
            r#"<?php
class StaticProxy {
    public static function __callStatic($name, $args) {
        echo "static $name";
    }
}
StaticProxy::anything();
"#
        ),
        &["static anything"]
    );
}

#[test]
fn magic_call_for_method_forwarding() {
    assert_eq!(
        run_prints(
            r#"<?php
class Decorator {
    public function __construct(private object $inner) {}
    public function __call($name, $args) {
        echo "before:$name ";
        $result = $this->inner->$name(...$args);
        echo "after:$name";
        return $result;
    }
}
class Service {
    public function greet(string $name): string {
        return "Hello $name";
    }
}
$d = new Decorator(new Service());
$d->greet("World");
"#
        ),
        &["before:greet after:greet"]
    );
}

#[test]
fn magic_call_returns_value() {
    assert_eq!(
        run_prints(
            r#"<?php
class Accessor {
    private array $data = ["name" => "John", "age" => 30];
    public function __call($method, $args) {
        if (str_starts_with($method, "get")) {
            $prop = strtolower(substr($method, 3));
            return $this->data[$prop] ?? null;
        }
        return null;
    }
}
$a = new Accessor();
echo $a->getName();
echo $a->getAge();
"#
        ),
        &["John30"]
    );
}

#[test]
fn magic_call_fluent_builder() {
    assert_eq!(
        run_prints(
            r#"<?php
class Query {
    private array $parts = [];
    public function __call($name, $args) {
        $this->parts[] = "$name:" . implode(",", $args);
        return $this;
    }
    public function build() {
        return implode(" | ", $this->parts);
    }
}
$q = new Query();
echo $q->select("id", "name")->from("users")->where("active=1")->build();
"#
        ),
        &["select:id,name | from:users | where:active=1"]
    );
}

// ── __clone ───────────────────────────────────────────────────────

#[test]
fn magic_clone_deep_copy() {
    assert_eq!(
        run_prints(
            r#"<?php
class DeepCopy {
    public array $items = [1, 2, 3];
    public function __clone() {
        $this->items = array_map(fn($x) => $x * 10, $this->items);
    }
}
$a = new DeepCopy();
$b = clone $a;
echo implode(",", $a->items);
echo implode(",", $b->items);
"#
        ),
        &["1,2,310,20,30"]
    );
}

#[test]
fn magic_clone_resets_mutable_state() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public int $count = 0;
    public function increment(): void { $this->count++; }
    public function __clone() { $this->count = 0; }
}
$a = new Counter();
$a->increment();
$a->increment();
$b = clone $a;
echo $a->count;
echo $b->count;
"#
        ),
        &["20"]
    );
}

#[test]
fn magic_clone_original_unchanged() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node {
    public function __construct(public string $value, public ?Node $next = null) {}
    public function __clone() {
        if ($this->next !== null) {
            $this->next = clone $this->next;
        }
    }
}
$a = new Node("first", new Node("second"));
$b = clone $a;
$b->value = "modified";
$b->next->value = "also modified";
echo $a->value;
echo $a->next->value;
echo $b->value;
echo $b->next->value;
"#
        ),
        &["firstsecondmodifiedalso modified"]
    );
}

#[test]
fn magic_clone_with_array_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Collection {
    public array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function __clone() {
        $this->items = array_reverse($this->items);
    }
}
$a = new Collection([1, 2, 3]);
$b = clone $a;
$b->items[] = 4;
echo implode(",", $a->items);
echo implode(",", $b->items);
"#
        ),
        &["1,2,33,2,1,4"]
    );
}

// ── __debugInfo ───────────────────────────────────────────────────

#[test]
fn magic_debuginfo_filters_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class Secret {
    public string $name = "visible";
    private string $password = "hidden";
    public function __debugInfo(): array {
        return ["name" => $this->name, "password" => "***"];
    }
}
$s = new Secret();
$info = $s->__debugInfo();
echo $info["name"];
echo $info["password"];
"#
        ),
        &["visible***"]
    );
}

#[test]
fn magic_debuginfo_returns_subset() {
    compile_ok(
        r#"<?php
class Config {
    public string $host = "localhost";
    public int $port = 3306;
    private string $dsn = "mysql:host=localhost;port=3306";
    private string $apiKey = "secret-key-12345";
    public function __debugInfo(): array {
        return [
            "host" => $this->host,
            "port" => $this->port,
            "dsn"  => substr($this->dsn, 0, 10) . "...",
        ];
    }
}
$cfg = new Config();
$info = $cfg->__debugInfo();
echo count($info);
echo isset($info["apiKey"]) ? "exposed" : "hidden";
"#,
    );
}

// ── __serialize / __unserialize ───────────────────────────────────

#[test]
fn magic_serialize_returns_custom_array() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {
    public function __construct(
        private string $value,
        private int $expiry,
        private string $secret = "internal"
    ) {}
    public function __serialize(): array {
        return ["value" => $this->value, "expiry" => $this->expiry];
    }
    public function __unserialize(array $data): void {
        $this->value  = $data["value"];
        $this->expiry = $data["expiry"];
        $this->secret = "restored";
    }
    public function getValue(): string { return $this->value; }
    public function getSecret(): string { return $this->secret; }
}
$t = new Token("abc123", 9999);
$raw = serialize($t);
$t2 = unserialize($raw);
echo $t2->getValue();
echo $t2->getSecret();
"#
        ),
        &["abc123restored"]
    );
}

#[test]
fn magic_serialize_unserialize_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php
class Vector2D {
    public function __construct(public float $x, public float $y) {}
    public function __serialize(): array { return ["x" => $this->x, "y" => $this->y]; }
    public function __unserialize(array $d): void { $this->x = $d["x"]; $this->y = $d["y"]; }
    public function length(): float { return sqrt($this->x ** 2 + $this->y ** 2); }
}
$v = new Vector2D(3.0, 4.0);
$raw = serialize($v);
$v2 = unserialize($raw);
echo $v2->x;
echo $v2->y;
echo $v2->length();
"#
        ),
        &["345"]
    );
}

// ── __sleep / __wakeup ────────────────────────────────────────────

#[test]
fn magic_sleep_returns_property_names() {
    assert_eq!(
        run_prints(
            r#"<?php
class Connection {
    public string $dsn = "mysql:host=localhost";
    public string $status = "connected";
    public function __sleep(): array {
        return ["dsn"];
    }
    public function __wakeup(): void {
        $this->status = "reconnected";
    }
}
$c = new Connection();
$data = serialize($c);
$c2 = unserialize($data);
echo $c2->dsn;
echo $c2->status;
"#
        ),
        &["mysql:host=localhostreconnected"]
    );
}

#[test]
fn magic_wakeup_restores_state() {
    assert_eq!(
        run_prints(
            r#"<?php
class Cache {
    public string $key = "my_key";
    private array $data = [];
    public function set(string $k, mixed $v): void { $this->data[$k] = $v; }
    public function get(string $k): mixed { return $this->data[$k] ?? null; }
    public function __sleep(): array { return ["key"]; }
    public function __wakeup(): void { $this->data = []; }
}
$c = new Cache();
$c->set("a", 42);
$raw = serialize($c);
$c2 = unserialize($raw);
echo $c2->key;
echo $c2->get("a") === null ? "cleared" : "kept";
"#
        ),
        &["my_keycleared"]
    );
}

// ── __set_state ───────────────────────────────────────────────────

#[test]
fn magic_set_state_creates_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function __construct(public int $x, public int $y) {}
    public static function __set_state(array $props): self {
        return new self($props['x'], $props['y']);
    }
}
$p = new Point(3, 4);
echo $p->x;
echo $p->y;
"#
        ),
        &["34"]
    );
}

#[test]
fn magic_set_state_compile_structure() {
    compile_ok(
        r#"<?php
class Rectangle {
    public float $width;
    public float $height;
    public function __construct(float $w, float $h) {
        $this->width  = $w;
        $this->height = $h;
    }
    public static function __set_state(array $props): static {
        return new static($props['width'], $props['height']);
    }
    public function area(): float { return $this->width * $this->height; }
}
$r = Rectangle::__set_state(['width' => 4.0, 'height' => 5.0]);
echo $r->area();
"#,
    );
}

// ── Interaction patterns ──────────────────────────────────────────

#[test]
fn magic_multiple_on_same_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class SmartObj {
    private array $props = [];
    public function __get($k) { return $this->props[$k] ?? null; }
    public function __set($k, $v) { $this->props[$k] = $v; }
    public function __isset($k) { return isset($this->props[$k]); }
    public function __toString(): string { return json_encode($this->props); }
    public function __invoke() { return count($this->props); }
}
$s = new SmartObj();
$s->x = 1;
$s->y = 2;
echo $s->x;
echo isset($s->y) ? "yes" : "no";
echo $s();
echo $s;
"#
        ),
        &["1yes2{\"x\":1,\"y\":2}"]
    );
}

#[test]
fn magic_method_in_abstract_class() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class BaseEntity {
    protected array $attributes = [];
    public function __get($name) { return $this->attributes[$name] ?? null; }
    public function __set($name, $value) { $this->attributes[$name] = $value; }
    abstract public function getType(): string;
}
class User extends BaseEntity {
    public function getType(): string { return "user"; }
}
$u = new User();
$u->name = "Alice";
echo $u->name;
echo $u->getType();
"#
        ),
        &["Aliceuser"]
    );
}

#[test]
fn magic_method_in_trait() {
    assert_eq!(
        run_prints(
            r#"<?php
trait DynamicAttributes {
    private array $attrs = [];
    public function __get($k) { return $this->attrs[$k] ?? null; }
    public function __set($k, $v) { $this->attrs[$k] = $v; }
    public function __isset($k) { return isset($this->attrs[$k]); }
}
class Product {
    use DynamicAttributes;
    public function __construct(public string $name) {}
}
$p = new Product("Widget");
$p->price = 9.99;
$p->stock = 100;
echo $p->name;
echo $p->price;
echo isset($p->stock) ? "in stock" : "out";
"#
        ),
        &["Widget9.99in stock"]
    );
}

#[test]
fn magic_get_returning_object_with_get() {
    assert_eq!(
        run_prints(
            r#"<?php
class Nested {
    private array $children = [];
    public function __construct(private string $name) {}
    public function addChild(string $key, Nested $child): void {
        $this->children[$key] = $child;
    }
    public function __get($key): ?Nested {
        return $this->children[$key] ?? null;
    }
    public function getName(): string { return $this->name; }
}
$root = new Nested("root");
$root->addChild("child", new Nested("child_node"));
echo $root->child->getName();
"#
        ),
        &["child_node"]
    );
}

#[test]
fn magic_call_variadic_forwarding() {
    assert_eq!(
        run_prints(
            r#"<?php
class Wrapper {
    public function __construct(private object $target) {}
    public function __call(string $name, array $args) {
        if (method_exists($this->target, $name)) {
            return $this->target->$name(...$args);
        }
        return null;
    }
}
class Math {
    public function add(int $a, int $b): int { return $a + $b; }
    public function multiply(int $a, int $b, int $c): int { return $a * $b * $c; }
}
$w = new Wrapper(new Math());
echo $w->add(3, 4);
echo $w->multiply(2, 3, 4);
"#
        ),
        &["724"]
    );
}

#[test]
fn magic_invoke_counting_invocations() {
    assert_eq!(
        run_prints(
            r#"<?php
class CallCounter {
    private int $calls = 0;
    public function __invoke(int $x): int {
        $this->calls++;
        return $x * $this->calls;
    }
    public function getCalls(): int { return $this->calls; }
}
$fn = new CallCounter();
echo $fn(10);
echo $fn(10);
echo $fn(10);
echo $fn->getCalls();
"#
        ),
        &["1020303"]
    );
}

#[test]
fn magic_tostring_and_invoke_on_same_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Expression {
    public function __construct(private string $expr, private float $value) {}
    public function __toString(): string { return $this->expr . " = " . $this->value; }
    public function __invoke(float $factor): float { return $this->value * $factor; }
}
$e = new Expression("2+3", 5.0);
echo $e;
echo $e(3);
"#
        ),
        &["2+3 = 515"]
    );
}

#[test]
fn magic_overloading_container_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class PropBag {
    private array $data = [];
    public function __set($k, $v) { $this->data[$k] = $v; }
    public function __get($k) { return $this->data[$k] ?? null; }
    public function __isset($k) { return array_key_exists($k, $this->data); }
    public function __unset($k) { unset($this->data[$k]); }
    public function keys(): array { return array_keys($this->data); }
}
$bag = new PropBag();
$bag->name = "test";
$bag->value = 42;
$bag->extra = "remove me";
unset($bag->extra);
echo implode(",", $bag->keys());
echo $bag->value;
echo isset($bag->extra) ? "yes" : "no";
"#
        ),
        &["name,value42no"]
    );
}

#[test]
fn magic_get_chained() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    private array $data = ["db" => ["host" => "localhost"]];
    public function __get($key) {
        $val = $this->data[$key] ?? null;
        if (is_array($val)) {
            $child = new Config();
            foreach ($val as $k => $v) {
                $child->data[$k] = $v;
            }
            return $child;
        }
        return $val;
    }
    public function __set($key, $value) {
        $this->data[$key] = $value;
    }
}
$c = new Config();
echo $c->db->host;
"#
        ),
        &["localhost"]
    );
}

#[test]
fn magic_callstatic_with_args() {
    assert_eq!(
        run_prints(
            r#"<?php
class FluentStatic {
    private static array $calls = [];
    public static function __callStatic(string $name, array $args): string {
        return $name . "(" . implode(",", $args) . ")";
    }
}
echo FluentStatic::greet("Alice", "Bob");
echo FluentStatic::sum("1", "2", "3");
"#
        ),
        &["greet(Alice,Bob)sum(1,2,3)"]
    );
}
