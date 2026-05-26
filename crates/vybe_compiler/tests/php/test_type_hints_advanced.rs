use super::helpers::run_prints;

// ── Intersection types (PHP 8.1) ──────────────────────────────

#[test] fn intersection_type_narrows_both_interfaces() {
    assert_eq!(run_prints(r#"<?php
interface Serializable2 { public function serialize2(): string; }
interface Loggable { public function log(): void; }
class Service implements Serializable2, Loggable {
    public function serialize2(): string { return 'data'; }
    public function log(): void { echo 'logged'; }
}
function process(Serializable2&Loggable $obj): string { $obj->log(); return $obj->serialize2(); }
echo process(new Service);
"#), vec!["loggeddata"]);
}
#[test] fn never_return_type_exit_path() {
    assert_eq!(run_prints(r#"<?php
function panic(string $msg): never { throw new \RuntimeException($msg); }
try { panic('oops'); } catch (\RuntimeException $e) { echo $e->getMessage(); }
"#), vec!["oops"]);
}
#[test] fn return_static_new_instance() {
    assert_eq!(run_prints(r#"<?php
class Fluent {
    private array $items = [];
    public function push(mixed $v): static { $this->items[] = $v; return $this; }
    public function count(): int { return count($this->items); }
}
$f = (new Fluent)->push(1)->push(2)->push(3);
echo $f->count();
"#), vec!["3"]);
}

// ── Nullsafe with array access ────────────────────────────────

#[test] fn nullsafe_static_method() {
    assert_eq!(run_prints(r#"<?php
class Factory { public static function create(): ?self { return new self; } public function value(): int { return 42; } }
echo (Factory::create())?->value() ?? 0;
"#), vec!["42"]);
}
#[test] fn nullsafe_null_returns_null_not_error() {
    assert_eq!(run_prints(r#"<?php
class A { public function b(): ?B { return null; } }
class B { public function c(): string { return 'c'; } }
$result = (new A)->b()?->c();
echo $result ?? 'null';
"#), vec!["null"]);
}

// ── DNF types (PHP 8.2) ───────────────────────────────────────

#[test] fn dnf_union_of_intersections() {
    assert_eq!(run_prints(r#"<?php
interface X { public function x(): int; }
interface Y { public function y(): int; }
class Both implements X, Y { public function x(): int { return 1; } public function y(): int { return 2; } }
class OnlyX implements X { public function x(): int { return 10; } }
function sum((X&Y)|null $obj): int { return $obj === null ? 0 : $obj->x() + $obj->y(); }
echo sum(new Both) . ',' . sum(null);
"#), vec!["3,0"]);
}

// ── Abstract static method ────────────────────────────────────

#[test] fn static_abstract_method_in_child() {
    assert_eq!(run_prints(r#"<?php
abstract class Registry {
    abstract protected static function tableName(): string;
    public static function all(): string { return 'SELECT * FROM ' . static::tableName(); }
}
class Users extends Registry { protected static function tableName(): string { return 'users'; } }
echo Users::all();
"#), vec!["SELECT * FROM users"]);
}

// ── Constructor in trait with abstract ────────────────────────

#[test] fn trait_with_constructor_assistance() {
    assert_eq!(run_prints(r#"<?php
trait AutoId {
    private static int $next = 0;
    private int $id;
    protected function initId(): void { $this->id = ++self::$next; }
    public function getId(): int { return $this->id; }
}
class Entity { use AutoId; public function __construct() { $this->initId(); } }
$a = new Entity; $b = new Entity; $c = new Entity;
echo $a->getId() . ',' . $b->getId() . ',' . $c->getId();
"#), vec!["1,2,3"]);
}
