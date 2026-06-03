use super::helpers::compile_ok;

// ── Property hooks (PHP 8.4) ──────────────────────────────────

#[test]
fn property_hook_get_basic() {
    compile_ok(
        r#"<?php
class Temperature {
    public float $celsius {
        get { return $this->celsius; }
        set(float $value) { $this->celsius = $value; }
    }
    public float $fahrenheit {
        get { return $this->celsius * 9/5 + 32; }
    }
}
$t = new Temperature();
$t->celsius = 100.0;
echo $t->fahrenheit;
"#,
    );
}

#[test]
fn property_hook_set_validation() {
    compile_ok(
        r#"<?php
class User {
    public string $email {
        set(string $value) {
            if (!str_contains($value, '@')) {
                throw new \InvalidArgumentException("Invalid email: $value");
            }
            $this->email = strtolower($value);
        }
    }
}
$u = new User();
$u->email = 'Alice@Example.COM';
echo $u->email;
"#,
    );
}

#[test]
fn property_hook_computed() {
    compile_ok(
        r#"<?php
class Circle {
    public float $radius = 0.0;
    public float $area {
        get { return M_PI * $this->radius ** 2; }
    }
    public float $circumference {
        get { return 2 * M_PI * $this->radius; }
    }
}
$c = new Circle();
$c->radius = 5.0;
echo round($c->area, 4) . ':' . round($c->circumference, 4);
"#,
    );
}

#[test]
fn property_hook_with_backing() {
    compile_ok(
        r#"<?php
class Product {
    private float $_price = 0.0;
    public float $price {
        get { return $this->_price; }
        set(float $value) {
            if ($value < 0) throw new \RangeException("Price cannot be negative");
            $this->_price = round($value, 2);
        }
    }
}
$p = new Product();
$p->price = 9.999;
echo $p->price;
"#,
    );
}

#[test]
fn property_hook_inherited() {
    compile_ok(
        r#"<?php
class Base {
    public int $value {
        get { return $this->value; }
        set(int $v) { $this->value = max(0, $v); }
    }
}
class Derived extends Base {
    public int $value {
        set(int $v) { $this->value = max(0, min(100, $v)); }
    }
}
$d = new Derived();
$d->value = 150;
echo $d->value;
"#,
    );
}

// ── Asymmetric property visibility (PHP 8.4) ──────────────────

#[test]
fn asymmetric_visibility_public_private_set() {
    compile_ok(
        r#"<?php
class Counter {
    public private(set) int $count = 0;
    public function increment(): void { $this->count++; }
}
$c = new Counter();
$c->increment();
$c->increment();
echo $c->count;
"#,
    );
}

#[test]
fn asymmetric_visibility_public_protected_set() {
    compile_ok(
        r#"<?php
class Entity {
    public protected(set) string $id;
    public function __construct(string $id) { $this->id = $id; }
}
class User extends Entity {
    public function rename(string $id): void { $this->id = $id; }
}
$u = new User('user-1');
echo $u->id;
$u->rename('user-2');
echo $u->id;
"#,
    );
}

#[test]
fn asymmetric_visibility_readonly_like() {
    compile_ok(
        r#"<?php
class Config {
    public private(set) string $env = 'production';
    public private(set) bool $debug = false;
    public function setDev(): void { $this->env = 'dev'; $this->debug = true; }
}
$cfg = new Config();
echo $cfg->env . ':' . ($cfg->debug ? 'debug' : 'prod');
$cfg->setDev();
echo $cfg->env . ':' . ($cfg->debug ? 'debug' : 'prod');
"#,
    );
}

// ── array_find / array_find_key (PHP 8.4) ─────────────────────

#[test]
fn array_find_basic() {
    compile_ok(
        r#"<?php
$users = [
    ['name' => 'Alice', 'age' => 28],
    ['name' => 'Bob',   'age' => 35],
    ['name' => 'Carol', 'age' => 22],
];
$found = array_find($users, fn($u) => $u['age'] > 30);
echo $found['name'];
"#,
    );
}

#[test]
fn array_find_not_found() {
    compile_ok(
        r#"<?php
$numbers = [1, 3, 5, 7, 9];
$even = array_find($numbers, fn($n) => $n % 2 === 0);
echo $even === null ? 'not found' : 'found';
"#,
    );
}

#[test]
fn array_find_key_basic() {
    compile_ok(
        r#"<?php
$items = ['apple' => 1.5, 'banana' => 0.5, 'cherry' => 3.0];
$key = array_find_key($items, fn($price) => $price > 2.0);
echo $key;
"#,
    );
}

#[test]
fn array_find_key_index() {
    compile_ok(
        r#"<?php
$scores = [45, 78, 92, 61, 88];
$idx = array_find_key($scores, fn($s) => $s >= 90);
echo $idx;
"#,
    );
}

// ── array_any / array_all (PHP 8.4) ──────────────────────────

#[test]
fn array_any_basic() {
    compile_ok(
        r#"<?php
$numbers = [1, 3, 5, 7, 8];
echo array_any($numbers, fn($n) => $n % 2 === 0) ? 'has even' : 'all odd';
"#,
    );
}

#[test]
fn array_any_false() {
    compile_ok(
        r#"<?php
$numbers = [1, 3, 5, 7, 9];
echo array_any($numbers, fn($n) => $n % 2 === 0) ? 'has even' : 'all odd';
"#,
    );
}

#[test]
fn array_all_true() {
    compile_ok(
        r#"<?php
$positives = [1, 5, 8, 3, 12];
echo array_all($positives, fn($n) => $n > 0) ? 'all positive' : 'some negative';
"#,
    );
}

#[test]
fn array_all_false() {
    compile_ok(
        r#"<?php
$mixed = [1, 5, -3, 8];
echo array_all($mixed, fn($n) => $n > 0) ? 'all positive' : 'some negative';
"#,
    );
}

// ── #[\Deprecated] attribute (PHP 8.4) ───────────────────────

#[test]
fn deprecated_attribute_function() {
    compile_ok(
        r#"<?php
#[\Deprecated('Use newFunction() instead', since: '2.0')]
function oldFunction(): string { return 'old'; }
function newFunction(): string { return 'new'; }
echo newFunction();
"#,
    );
}

#[test]
fn deprecated_attribute_method() {
    compile_ok(
        r#"<?php
class Api {
    #[\Deprecated('Use v2() instead')]
    public function v1(): string { return 'v1'; }
    public function v2(): string { return 'v2'; }
}
echo (new Api())->v2();
"#,
    );
}

// ── new in class constant initializers (PHP 8.1) ───────────────

#[test]
fn new_in_const_initializer() {
    compile_ok(
        r#"<?php
class Config {
    const DEFAULT_TIMEOUT = new \DateInterval('PT30S');
}
echo Config::DEFAULT_TIMEOUT->s;
"#,
    );
}

// ── Readonly class properties (PHP 8.2) ───────────────────────

#[test]
fn readonly_class_deep() {
    compile_ok(
        r#"<?php
readonly class Coordinate {
    public function __construct(
        public float $lat,
        public float $lon
    ) {}
    public function distanceTo(Coordinate $other): float {
        return sqrt(($this->lat - $other->lat)**2 + ($this->lon - $other->lon)**2);
    }
}
$a = new Coordinate(0.0, 0.0);
$b = new Coordinate(3.0, 4.0);
echo $b->distanceTo($a);
"#,
    );
}

// ── null, true, false as types (PHP 8.2) ─────────────────────

#[test]
fn null_as_standalone_type() {
    compile_ok(
        r#"<?php
function alwaysNull(): null { return null; }
$v = alwaysNull();
var_dump($v);
"#,
    );
}

#[test]
fn true_false_as_types() {
    compile_ok(
        r#"<?php
function succeed(): true  { return true; }
function fail(): false    { return false; }
var_dump(succeed());
var_dump(fail());
"#,
    );
}

// ── DNF types (PHP 8.2) ───────────────────────────────────────

#[test]
fn dnf_type_hint() {
    compile_ok(
        r#"<?php
interface Countable2 { public function count2(): int; }
interface Serializable2 { public function serialize2(): string; }
class Set implements Countable2 {
    private array $items = [];
    public function count2(): int { return count($this->items); }
    public function add(mixed $v): void { $this->items[] = $v; }
}
function describe((Countable2&Serializable2)|null $obj): string {
    if ($obj === null) return 'null';
    return 'count=' . $obj->count2();
}
// Pass null (valid for (C&S)|null)
echo describe(null);
"#,
    );
}

// ── Typed class constants (PHP 8.3) ───────────────────────────

#[test]
fn typed_class_constants() {
    compile_ok(
        r#"<?php
class Config {
    const int    MAX_SIZE     = 1024;
    const string DEFAULT_ENV  = 'production';
    const float  TAX_RATE     = 0.08;
    const bool   DEBUG        = false;
}
echo Config::MAX_SIZE . ':' . Config::DEFAULT_ENV . ':' . Config::DEBUG;
"#,
    );
}

#[test]
fn typed_class_constants_interface() {
    compile_ok(
        r#"<?php
interface HasVersion {
    const string VERSION = '1.0.0';
}
class App implements HasVersion {
    const string VERSION = '2.0.0';
}
echo App::VERSION;
"#,
    );
}

// ── #[Override] attribute (PHP 8.3) ───────────────────────────

#[test]
fn override_attribute() {
    compile_ok(
        r#"<?php
class Base { public function render(): string { return 'base'; } }
class Derived extends Base {
    #[\Override]
    public function render(): string { return 'derived'; }
}
echo (new Derived())->render();
"#,
    );
}

// ── Dynamic class constant fetch (PHP 8.3) ────────────────────

#[test]
fn dynamic_class_const_fetch() {
    compile_ok(
        r#"<?php
class HttpStatus {
    const OK       = 200;
    const NOT_FOUND = 404;
    const ERROR    = 500;
}
$const = 'NOT_FOUND';
echo HttpStatus::{$const};
"#,
    );
}
