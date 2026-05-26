use super::helpers::run_prints;

// ── Constructor promotion ─────────────────────────────────────

#[test] fn constructor_promotion_public() {
    assert_eq!(run_prints(r#"<?php
class Point { public function __construct(public float $x, public float $y) {} }
$p = new Point(3.0, 4.0);
echo $p->x . ',' . $p->y;
"#), vec!["3,4"]);
}
#[test] fn constructor_promotion_defaults() {
    assert_eq!(run_prints(r#"<?php
class Config { public function __construct(public string $host = 'localhost', public int $port = 8080) {} }
$c = new Config(port: 3000);
echo $c->host . ':' . $c->port;
"#), vec!["localhost:3000"]);
}
#[test] fn constructor_promotion_mixed() {
    assert_eq!(run_prints(r#"<?php
class User {
    public string $fullname;
    public function __construct(public string $first, public string $last) {
        $this->fullname = "$first $last";
    }
}
$u = new User('John', 'Doe');
echo $u->fullname;
"#), vec!["John Doe"]);
}

// ── __toString and Stringable ─────────────────────────────────

#[test] fn to_string_magic_method() {
    assert_eq!(run_prints(r#"<?php
class Money {
    public function __construct(private int $amount, private string $currency) {}
    public function __toString(): string { return $this->amount . ' ' . $this->currency; }
}
echo new Money(100, 'USD');
"#), vec!["100 USD"]);
}
#[test] fn stringable_interface_type_hint() {
    assert_eq!(run_prints(r#"<?php
class Label implements Stringable {
    public function __construct(private string $text) {}
    public function __toString(): string { return $this->text; }
}
function print_label(Stringable $s): void { echo (string)$s; }
print_label(new Label('hello'));
"#), vec!["hello"]);
}

// ── __invoke ─────────────────────────────────────────────────

#[test] fn invoke_makes_object_callable() {
    assert_eq!(run_prints(r#"<?php
class Multiplier {
    public function __construct(private int $factor) {}
    public function __invoke(int $n): int { return $n * $this->factor; }
}
$triple = new Multiplier(3);
echo $triple(7);
"#), vec!["21"]);
}
#[test] fn invoke_used_as_callback() {
    assert_eq!(run_prints(r#"<?php
class Adder { public function __construct(private int $n) {} public function __invoke(int $x): int { return $x + $this->n; } }
echo implode(',', array_map(new Adder(10), [1,2,3]));
"#), vec!["11,12,13"]);
}

// ── __get / __set / __isset / __unset ─────────────────────────

#[test] fn magic_get_set() {
    assert_eq!(run_prints(r#"<?php
class DynProps {
    private array $data = [];
    public function __get(string $k): mixed { return $this->data[$k] ?? null; }
    public function __set(string $k, mixed $v): void { $this->data[$k] = $v; }
}
$o = new DynProps;
$o->name = 'Alice';
echo $o->name;
"#), vec!["Alice"]);
}
#[test] fn magic_isset_returns_true() {
    assert_eq!(run_prints(r#"<?php
class Bag {
    private array $d = [];
    public function __set(string $k, mixed $v): void { $this->d[$k] = $v; }
    public function __isset(string $k): bool { return isset($this->d[$k]); }
}
$b = new Bag; $b->x = 1;
echo isset($b->x) ? 'yes' : 'no';
echo isset($b->y) ? 'yes' : 'no';
"#), vec!["yesno"]);
}
#[test] fn magic_unset() {
    assert_eq!(run_prints(r#"<?php
class Store {
    private array $d = ['k' => 'v'];
    public function __isset(string $k): bool { return isset($this->d[$k]); }
    public function __unset(string $k): void { unset($this->d[$k]); }
}
$s = new Store;
echo isset($s->k) ? 'before' : '';
unset($s->k);
echo isset($s->k) ? 'after' : 'gone';
"#), vec!["beforegone"]);
}

// ── __call / __callStatic ─────────────────────────────────────

#[test] fn magic_call_intercepts_missing_method() {
    assert_eq!(run_prints(r#"<?php
class Proxy {
    public function __call(string $name, array $args): mixed {
        return "called $name with " . count($args) . " args";
    }
}
echo (new Proxy)->doSomething(1, 2, 3);
"#), vec!["called doSomething with 3 args"]);
}
#[test] fn magic_call_static() {
    assert_eq!(run_prints(r#"<?php
class Api {
    public static function __callStatic(string $name, array $args): string {
        return "static:$name(" . implode(',', $args) . ")";
    }
}
echo Api::get('/users', 'json');
"#), vec!["static:get(/users,json)"]);
}

// ── Object cloning ────────────────────────────────────────────

#[test] fn clone_shallow_copy() {
    assert_eq!(run_prints(r#"<?php
class Box { public int $val = 1; }
$a = new Box;
$b = clone $a;
$b->val = 99;
echo $a->val . ',' . $b->val;
"#), vec!["1,99"]);
}
#[test] fn clone_deep_via_clone_magic() {
    assert_eq!(run_prints(r#"<?php
class Inner { public int $v = 0; }
class Outer {
    public Inner $inner;
    public function __construct() { $this->inner = new Inner; }
    public function __clone() { $this->inner = clone $this->inner; }
}
$a = new Outer; $a->inner->v = 5;
$b = clone $a; $b->inner->v = 99;
echo $a->inner->v . ',' . $b->inner->v;
"#), vec!["5,99"]);
}

// ── Object comparison ─────────────────────────────────────────

#[test] fn same_instance_triple_equal() {
    assert_eq!(run_prints(r#"<?php class A {} $a = new A; $b = $a; echo ($a === $b) ? 'same' : 'diff'; "#), vec!["same"]);
}
#[test] fn clone_not_identical() {
    assert_eq!(run_prints(r#"<?php class A {} $a = new A; $b = clone $a; echo ($a === $b) ? 'same' : 'diff'; "#), vec!["diff"]);
}
