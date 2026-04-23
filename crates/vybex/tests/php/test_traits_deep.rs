use super::helpers::compile_ok;

// ── Trait with abstract method requirement ─────────────────────

#[test] fn trait_abstract_requirement() {
    compile_ok(r#"<?php
trait Printable {
    abstract public function toString(): string;
    public function print(): void { echo $this->toString(); }
}
class Color {
    use Printable;
    public function __construct(private string $name) {}
    public function toString(): string { return "Color({$this->name})"; }
}
(new Color('red'))->print();
"#);
}

#[test] fn trait_abstract_template_method() {
    compile_ok(r#"<?php
trait Report {
    abstract protected function gatherData(): array;
    abstract protected function formatRow(array $row): string;
    public function generate(): string {
        $rows = array_map([$this, 'formatRow'], $this->gatherData());
        return implode("\n", $rows);
    }
}
class SalesReport {
    use Report;
    protected function gatherData(): array { return [['item' => 'Widget', 'qty' => 5], ['item' => 'Gadget', 'qty' => 3]]; }
    protected function formatRow(array $row): string { return "{$row['item']}: {$row['qty']}"; }
}
echo (new SalesReport())->generate();
"#);
}

// ── Trait with properties ─────────────────────────────────────

#[test] fn trait_property_default() {
    compile_ok(r#"<?php
trait HasName {
    private string $name = 'unnamed';
    public function getName(): string { return $this->name; }
    public function setName(string $n): void { $this->name = $n; }
}
class Animal { use HasName; }
$a = new Animal();
echo $a->getName();
$a->setName('Rex');
echo $a->getName();
"#);
}

#[test] fn trait_property_multiple_classes() {
    compile_ok(r#"<?php
trait HasId {
    private static int $nextId = 0;
    private int $id;
    public function initId(): void { $this->id = ++self::$nextId; }
    public function getId(): int   { return $this->id; }
}
class UserA { use HasId; }
class UserB { use HasId; }
$a = new UserA(); $a->initId();
$b = new UserB(); $b->initId();
echo $a->getId() . ',' . $b->getId();
"#);
}

// ── Trait with static methods / properties ────────────────────

#[test] fn trait_static_method() {
    compile_ok(r#"<?php
trait Singleton {
    private static ?self $instance = null;
    public static function getInstance(): static {
        if (static::$instance === null) { static::$instance = new static(); }
        return static::$instance;
    }
}
class Config { use Singleton; public string $env = 'production'; }
$c1 = Config::getInstance();
$c2 = Config::getInstance();
$c1->env = 'staging';
echo $c2->env;
"#);
}

#[test] fn trait_static_counter() {
    compile_ok(r#"<?php
trait Counter {
    private static int $count = 0;
    public static function increment(): void { static::$count++; }
    public static function getCount(): int   { return static::$count; }
}
class A { use Counter; }
class B { use Counter; }
A::increment(); A::increment(); A::increment();
B::increment();
echo A::getCount() . ',' . B::getCount();
"#);
}

// ── Trait constants (PHP 8.2+) ────────────────────────────────

#[test] fn trait_constant() {
    compile_ok(r#"<?php
trait Configurable {
    const DEFAULT_TIMEOUT = 30;
    const MAX_RETRIES = 3;
}
class HttpClient { use Configurable; }
echo HttpClient::DEFAULT_TIMEOUT . ',' . HttpClient::MAX_RETRIES;
"#);
}

#[test] fn trait_constant_multiple_uses() {
    compile_ok(r#"<?php
trait StatusCodes {
    const OK    = 200;
    const ERROR = 500;
}
class ApiA { use StatusCodes; }
class ApiB { use StatusCodes; }
echo ApiA::OK . ',' . ApiB::ERROR;
"#);
}

// ── Trait visibility change ───────────────────────────────────

#[test] fn trait_visibility_as() {
    compile_ok(r#"<?php
trait Greetable {
    public function hello(): string { return "Hello!"; }
    public function goodbye(): string { return "Goodbye!"; }
}
class Formal {
    use Greetable { goodbye as protected; }
    public function farewell(): string { return $this->goodbye(); }
}
$f = new Formal();
echo $f->hello();
echo $f->farewell();
"#);
}

#[test] fn trait_alias_rename() {
    compile_ok(r#"<?php
trait Logger {
    public function log(string $msg): void { echo "[LOG] $msg"; }
}
class Service {
    use Logger { log as writeLog; }
    public function process(): void { $this->writeLog("processing"); }
}
(new Service())->process();
"#);
}

// ── Trait insteadof conflict resolution ───────────────────────

#[test] fn trait_insteadof_conflict() {
    compile_ok(r#"<?php
trait A { public function hello(): string { return "A::hello"; } }
trait B { public function hello(): string { return "B::hello"; } }
class C {
    use A, B { A::hello insteadof B; B::hello as helloFromB; }
}
$c = new C();
echo $c->hello() . ',' . $c->helloFromB();
"#);
}

#[test] fn trait_insteadof_three_way() {
    compile_ok(r#"<?php
trait X { public function op(): string { return "X"; } }
trait Y { public function op(): string { return "Y"; } }
trait Z { public function op(): string { return "Z"; } }
class UseXY {
    use X, Y, Z { X::op insteadof Y, Z; Y::op as opY; Z::op as opZ; }
}
$u = new UseXY();
echo $u->op() . $u->opY() . $u->opZ();
"#);
}

// ── Trait in anonymous class ───────────────────────────────────

#[test] fn trait_anonymous_class() {
    compile_ok(r#"<?php
trait Taggable {
    private array $tags = [];
    public function addTag(string $tag): void { $this->tags[] = $tag; }
    public function getTags(): array { return $this->tags; }
}
$obj = new class { use Taggable; };
$obj->addTag('php');
$obj->addTag('oop');
echo implode(',', $obj->getTags());
"#);
}

// ── Trait with interface ───────────────────────────────────────

#[test] fn trait_satisfies_interface() {
    compile_ok(r#"<?php
interface Identifiable { public function getId(): int; }
trait HasId {
    private int $id;
    public function __construct(int $id) { $this->id = $id; }
    public function getId(): int { return $this->id; }
}
class User implements Identifiable { use HasId; }
$u = new User(42);
echo $u->getId();
"#);
}

// ── Chained traits ────────────────────────────────────────────

#[test] fn trait_uses_another_via_class() {
    compile_ok(r#"<?php
trait HasTimestamps {
    private int $createdAt = 0;
    private int $updatedAt = 0;
    public function touch(): void { $this->updatedAt = time(); }
    public function getUpdatedAt(): int { return $this->updatedAt; }
}
trait HasSoftDelete {
    private ?int $deletedAt = null;
    public function delete(): void { $this->deletedAt = time(); }
    public function isDeleted(): bool { return $this->deletedAt !== null; }
}
class Post {
    use HasTimestamps, HasSoftDelete;
    public function __construct(public string $title) {}
}
$p = new Post('Hello World');
$p->delete();
echo $p->isDeleted() ? 'deleted' : 'active';
"#);
}

// ── Trait with constructor requirement ─────────────────────────

#[test] fn trait_with_constructor_use() {
    compile_ok(r#"<?php
trait Validatable {
    private array $errors = [];
    abstract protected function validate(): void;
    public function isValid(): bool { $this->validate(); return empty($this->errors); }
    protected function addError(string $msg): void { $this->errors[] = $msg; }
    public function getErrors(): array { return $this->errors; }
}
class Email {
    use Validatable;
    public function __construct(private string $addr) {}
    protected function validate(): void {
        if (!str_contains($this->addr, '@')) {
            $this->addError("Invalid email: {$this->addr}");
        }
    }
}
$e = new Email('notvalid');
echo $e->isValid() ? 'valid' : 'invalid';
echo ':' . count($e->getErrors());
"#);
}

// ── Trait property type conflict ──────────────────────────────

#[test] fn trait_late_binding_in_static() {
    compile_ok(r#"<?php
trait Registry {
    private static array $items = [];
    public static function register(string $key, mixed $val): void { static::$items[$key] = $val; }
    public static function get(string $key): mixed { return static::$items[$key] ?? null; }
    public static function all(): array { return static::$items; }
}
class ServiceContainer { use Registry; }
ServiceContainer::register('db', 'sqlite');
ServiceContainer::register('cache', 'redis');
echo count(ServiceContainer::all());
"#);
}
