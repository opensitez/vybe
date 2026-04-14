use super::helpers::run_prints;

#[test]
fn magic_tostring() {
    assert_eq!(run_prints(r#"<?php
class Color {
    public function __construct(private string $name) {}
    public function __toString(): string {
        return $this->name;
    }
}
$c = new Color("red");
echo $c;
"#), &["red"]);
}

#[test]
fn magic_get_set() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["bar", "undefined"]);
}

#[test]
fn magic_isset_unset() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["yes", "no", "no"]);
}

#[test]
fn magic_call() {
    assert_eq!(run_prints(r#"<?php
class Proxy {
    public function __call($name, $args) {
        echo "called $name with " . count($args) . " args";
    }
}
$p = new Proxy();
$p->hello(1, 2, 3);
"#), &["called hello with 3 args"]);
}

#[test]
fn magic_call_static() {
    assert_eq!(run_prints(r#"<?php
class StaticProxy {
    public static function __callStatic($name, $args) {
        echo "static $name";
    }
}
StaticProxy::anything();
"#), &["static anything"]);
}

#[test]
fn magic_invoke() {
    assert_eq!(run_prints(r#"<?php
class Multiplier {
    public function __construct(private int $factor) {}
    public function __invoke($x) {
        return $x * $this->factor;
    }
}
$double = new Multiplier(2);
echo $double(5);
echo $double(10);
"#), &["10", "20"]);
}

#[test]
fn magic_clone() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["1,2,3", "10,20,30"]);
}

#[test]
fn magic_tostring_in_concat() {
    assert_eq!(run_prints(r#"<?php
class Name {
    public function __construct(private string $val) {}
    public function __toString(): string { return $this->val; }
}
$n = new Name("world");
echo "hello " . $n;
"#), &["hello world"]);
}

#[test]
fn magic_get_chained() {
    assert_eq!(run_prints(r#"<?php
class Config {
    private array $data = ["db" => ["host" => "localhost"]];
    public function __get($key) {
        $val = $this->data[$key] ?? null;
        if (is_array($val)) {
            $child = new Config();
            // Store in internal data for child access
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
"#), &["localhost"]);
}

#[test]
fn magic_invoke_as_callback() {
    assert_eq!(run_prints(r#"<?php
class Adder {
    public function __construct(private int $n) {}
    public function __invoke($x) { return $x + $this->n; }
}
$add5 = new Adder(5);
$nums = [1, 2, 3];
$result = array_map($add5, $nums);
echo implode(",", $result);
"#), &["6,7,8"]);
}

#[test]
fn magic_tostring_in_interpolation() {
    assert_eq!(run_prints(r#"<?php
class Version {
    public function __construct(private int $major, private int $minor) {}
    public function __toString(): string {
        return "$this->major.$this->minor";
    }
}
$v = new Version(3, 14);
echo "version: $v";
"#), &["version: 3.14"]);
}

#[test]
fn magic_call_fluent() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["select:id,name | from:users | where:active=1"]);
}

#[test]
fn magic_debuginfo() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["visible", "***"]);
}

#[test]
fn magic_sleep_wakeup() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["mysql:host=localhost", "reconnected"]);
}

#[test]
fn magic_set_state() {
    assert_eq!(run_prints(r#"<?php
class Point {
    public function __construct(public int $x, public int $y) {}
    public static function __set_state(array $props): self {
        return new self($props['x'], $props['y']);
    }
}
$p = new Point(3, 4);
echo $p->x;
echo $p->y;
"#), &["3", "4"]);
}

#[test]
fn magic_get_returns_reference_like() {
    assert_eq!(run_prints(r#"<?php
class Registry {
    private array $store = [];
    public function __get($key) {
        return $this->store[$key] ?? null;
    }
    public function __set($key, $val) {
        $this->store[$key] = $val;
    }
    public function __isset($key) {
        return isset($this->store[$key]);
    }
}
$r = new Registry();
$r->a = 10;
$r->b = 20;
echo $r->a + $r->b;
echo isset($r->a) ? "set" : "not set";
echo isset($r->c) ? "set" : "not set";
"#), &["30", "set", "not set"]);
}

#[test]
fn magic_invoke_is_callable() {
    assert_eq!(run_prints(r#"<?php
class Handler {
    public function __invoke($msg) { echo "handled: $msg"; }
}
$h = new Handler();
echo is_callable($h) ? "yes" : "no";
$h("test");
"#), &["yes", "handled: test"]);
}

#[test]
fn magic_tostring_with_cast() {
    assert_eq!(run_prints(r#"<?php
class Amount {
    public function __construct(private float $value) {}
    public function __toString(): string {
        return number_format($this->value, 2);
    }
}
$a = new Amount(42.5);
echo (string) $a;
"#), &["42.50"]);
}

#[test]
fn magic_call_undefined_method_pattern() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["John", "30"]);
}

#[test]
fn magic_multiple_on_same_class() {
    assert_eq!(run_prints(r#"<?php
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
"#), &["1", "yes", "2", "{\"x\":1,\"y\":2}"]);
}
