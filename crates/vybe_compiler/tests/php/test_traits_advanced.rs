use super::helpers::run_prints;

// ── Trait basic usage ─────────────────────────────────────────

#[test] fn trait_adds_method() {
    assert_eq!(run_prints(r#"<?php
trait Greet { public function greet(): string { return 'Hello, ' . $this->name; } }
class Person { use Greet; public function __construct(public string $name) {} }
echo (new Person('Alice'))->greet();
"#), vec!["Hello, Alice"]);
}
#[test] fn trait_multiple_classes_share() {
    assert_eq!(run_prints(r#"<?php
trait Timestamped {
    private string $createdAt;
    public function setCreatedAt(string $t): void { $this->createdAt = $t; }
    public function getCreatedAt(): string { return $this->createdAt; }
}
class Post { use Timestamped; }
class Comment { use Timestamped; }
$p = new Post; $p->setCreatedAt('2024-01-01');
$c = new Comment; $c->setCreatedAt('2024-06-15');
echo $p->getCreatedAt() . ',' . $c->getCreatedAt();
"#), vec!["2024-01-01,2024-06-15"]);
}

// ── Trait properties ──────────────────────────────────────────

#[test] fn trait_property_access() {
    assert_eq!(run_prints(r#"<?php
trait HasId { public int $id = 0; }
class Entity { use HasId; }
$e = new Entity; $e->id = 42;
echo $e->id;
"#), vec!["42"]);
}

// ── Trait abstract method ─────────────────────────────────────

#[test] fn trait_abstract_forces_implementation() {
    assert_eq!(run_prints(r#"<?php
trait Validator {
    abstract protected function rules(): array;
    public function validate(array $data): bool {
        foreach ($this->rules() as $field) {
            if (empty($data[$field])) return false;
        }
        return true;
    }
}
class UserValidator {
    use Validator;
    protected function rules(): array { return ['name','email']; }
}
$v = new UserValidator;
echo $v->validate(['name'=>'Al','email'=>'al@x.com']) ? 'ok' : 'fail';
echo $v->validate(['name'=>'Al']) ? 'ok' : 'fail';
"#), vec!["okfail"]);
}

// ── Multiple traits ───────────────────────────────────────────

#[test] fn multiple_traits_composed() {
    assert_eq!(run_prints(r#"<?php
trait Serializable2 { public function serialize(): string { return json_encode((array)$this); } }
trait Loggable { public function log(): void { echo 'log:' . get_class($this); } }
class Item { use Serializable2, Loggable; public function __construct(public string $name) {} }
$item = new Item('test');
$item->log();
echo ',' . json_decode($item->serialize())->name;
"#), vec!["log:Item,test"]);
}
#[test] fn trait_method_conflicts_resolved() {
    assert_eq!(run_prints(r#"<?php
trait A { public function hello(): string { return 'from A'; } }
trait B { public function hello(): string { return 'from B'; } }
class C { use A, B { A::hello insteadof B; B::hello as helloB; } }
$c = new C;
echo $c->hello() . ',' . $c->helloB();
"#), vec!["from A,from B"]);
}

// ── Trait with static members ─────────────────────────────────

#[test] fn trait_static_method() {
    assert_eq!(run_prints(r#"<?php
trait Singleton {
    private static ?self $instance = null;
    public static function getInstance(): static {
        if (static::$instance === null) static::$instance = new static();
        return static::$instance;
    }
}
class Config { use Singleton; public int $value = 42; }
$a = Config::getInstance(); $a->value = 99;
$b = Config::getInstance();
echo $b->value;
"#), vec!["99"]);
}
#[test] fn trait_static_property() {
    assert_eq!(run_prints(r#"<?php
trait Counter {
    private static int $count = 0;
    public static function increment(): void { static::$count++; }
    public static function getCount(): int { return static::$count; }
}
class Foo { use Counter; }
class Bar { use Counter; }
Foo::increment(); Foo::increment(); Bar::increment();
echo Foo::getCount() . ',' . Bar::getCount();
"#), vec!["2,1"]);
}

// ── Trait visibility change ───────────────────────────────────

#[test] fn trait_method_aliased_private() {
    assert_eq!(run_prints(r#"<?php
trait Helper { public function doWork(): string { return 'work'; } }
class Service {
    use Helper { doWork as private internalWork; }
    public function run(): string { return $this->internalWork(); }
}
echo (new Service)->run();
"#), vec!["work"]);
}

// ── Trait in abstract class ───────────────────────────────────

#[test] fn trait_in_abstract_class_used_by_concrete() {
    assert_eq!(run_prints(r#"<?php
trait EventEmitter {
    private array $listeners = [];
    public function on(string $event, callable $cb): void { $this->listeners[$event][] = $cb; }
    public function emit(string $event, mixed ...$args): void {
        foreach ($this->listeners[$event] ?? [] as $cb) $cb(...$args);
    }
}
abstract class Component { use EventEmitter; abstract public function render(): string; }
class Button extends Component {
    public function render(): string { return '<button>'; }
}
$btn = new Button;
$btn->on('click', fn($x) => print("clicked:$x"));
$btn->emit('click', 'left');
"#), vec!["clicked:left"]);
}

// ── Trait requiring interface ─────────────────────────────────

#[test] fn trait_requires_method_from_using_class() {
    assert_eq!(run_prints(r#"<?php
trait Printable2 {
    abstract protected function content(): string;
    public function print(): void { echo '[' . $this->content() . ']'; }
}
class Article {
    use Printable2;
    public function __construct(private string $text) {}
    protected function content(): string { return $this->text; }
}
(new Article('Hello World'))->print();
"#), vec!["[Hello World]"]);
}

// ── Trait constants (PHP 8.2) ─────────────────────────────────

#[test] fn trait_constant_access() {
    assert_eq!(run_prints(r#"<?php
trait HasVersion {
    public function getVersion(): string { return static::VERSION; }
}
class AppV1 {
    use HasVersion;
    const VERSION = '1.0.0';
}
echo (new AppV1)->getVersion();
"#), vec!["1.0.0"]);
}
