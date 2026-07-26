use super::helpers::{compile_ok, run_prints};

// ── Interface extending multiple interfaces ────────────────────

#[test]
fn interface_extends_two() {
    compile_ok(
        r#"<?php
interface Readable  { public function read(): string; }
interface Writable  { public function write(string $data): void; }
interface ReadWrite extends Readable, Writable {}
class File implements ReadWrite {
    private string $buffer = '';
    public function read(): string { return $this->buffer; }
    public function write(string $data): void { $this->buffer .= $data; }
}
$f = new File();
$f->write('hello');
$f->write(' world');
echo $f->read();
"#,
    );
}

#[test]
fn interface_extends_three() {
    compile_ok(
        r#"<?php
interface Named    { public function getName(): string; }
interface Aged     { public function getAge(): int; }
interface Skilled  { public function getSkills(): array; }
interface Person extends Named, Aged, Skilled {}
class Developer implements Person {
    public function __construct(private string $name, private int $age, private array $skills) {}
    public function getName(): string  { return $this->name; }
    public function getAge(): int      { return $this->age; }
    public function getSkills(): array { return $this->skills; }
}
$d = new Developer('Alice', 30, ['PHP', 'Rust']);
echo $d->getName() . ':' . $d->getAge() . ':' . implode(',', $d->getSkills());
"#,
    );
}

#[test]
fn interface_chain_inheritance() {
    compile_ok(
        r#"<?php
interface A { public function a(): string; }
interface B extends A { public function b(): string; }
interface C extends B { public function c(): string; }
class Impl implements C {
    public function a(): string { return 'a'; }
    public function b(): string { return 'b'; }
    public function c(): string { return 'c'; }
}
$obj = new Impl();
echo $obj->a() . $obj->b() . $obj->c();
"#,
    );
}

#[test]
fn interface_instanceof_chain() {
    compile_ok(
        r#"<?php
interface Shape { public function area(): float; }
interface ColoredShape extends Shape { public function color(): string; }
class RedCircle implements ColoredShape {
    public function __construct(private float $r) {}
    public function area(): float { return M_PI * $this->r ** 2; }
    public function color(): string { return 'red'; }
}
$c = new RedCircle(2.0);
echo ($c instanceof Shape) ? 'is Shape' : 'not Shape';
echo ($c instanceof ColoredShape) ? ':is ColoredShape' : ':not ColoredShape';
"#,
    );
}

// ── Typed interface constants (PHP 8.3) ───────────────────────

#[test]
fn interface_typed_constants() {
    compile_ok(
        r#"<?php
interface Versioned {
    const string VERSION = '1.0.0';
    const int    BUILD   = 42;
}
class App implements Versioned {}
echo App::VERSION . ':' . App::BUILD;
"#,
    );
}

#[test]
fn interface_constant_override() {
    compile_ok(
        r#"<?php
interface HasDefault { const string MODE = 'default'; }
class Custom implements HasDefault { const string MODE = 'custom'; }
class Default_ implements HasDefault {}
echo Custom::MODE . ':' . Default_::MODE;
"#,
    );
}

// ── Covariant return types ────────────────────────────────────

#[test]
fn covariant_return_type_basic() {
    compile_ok(
        r#"<?php
class Animal {}
class Dog extends Animal {}
interface AnimalFactory { public function create(): Animal; }
class DogFactory implements AnimalFactory {
    public function create(): Dog { return new Dog(); }
}
$factory = new DogFactory();
$dog = $factory->create();
echo ($dog instanceof Animal) ? 'is Animal' : 'not Animal';
echo ($dog instanceof Dog) ? ':is Dog' : ':not Dog';
"#,
    );
}

#[test]
fn covariant_return_self_static() {
    compile_ok(
        r#"<?php
interface Buildable { public function withName(string $name): static; }
class Widget implements Buildable {
    private string $name = '';
    public function withName(string $name): static {
        $clone = clone $this;
        $clone->name = $name;
        return $clone;
    }
    public function getName(): string { return $this->name; }
}
class Button extends Widget {}
$btn = (new Button())->withName('Submit');
echo $btn->getName();
echo ($btn instanceof Button) ? ':is Button' : ':not Button';
"#,
    );
}

// ── Interface-based polymorphism ─────────────────────────────

#[test]
fn polymorphism_collection() {
    compile_ok(
        r#"<?php
interface Formatter { public function format(mixed $value): string; }
class IntFormatter implements Formatter {
    public function format(mixed $value): string { return number_format((int)$value); }
}
class DateFormatter implements Formatter {
    public function format(mixed $value): string { return date('Y-m-d', (int)$value); }
}
class BoolFormatter implements Formatter {
    public function format(mixed $value): string { return $value ? 'yes' : 'no'; }
}
/** @param Formatter[] $formatters */
function formatAll(array $data, array $formatters): array {
    $result = [];
    foreach ($data as $i => $v) {
        $result[] = ($formatters[$i] ?? $formatters[0])->format($v);
    }
    return $result;
}
$rows = formatAll([1000, true], [new IntFormatter(), new BoolFormatter()]);
echo implode('|', $rows);
"#,
    );
}

#[test]
fn interface_type_hint_accept_any_impl() {
    compile_ok(
        r#"<?php
interface Logger { public function log(string $msg): void; }
class ConsoleLogger implements Logger {
    private array $log = [];
    public function log(string $msg): void { $this->log[] = $msg; }
    public function getLog(): array { return $this->log; }
}
class NullLogger implements Logger { public function log(string $msg): void {} }
function doWork(Logger $logger): void {
    $logger->log('started');
    $logger->log('done');
}
$c = new ConsoleLogger();
doWork($c);
echo count($c->getLog());
doWork(new NullLogger());
echo 'ok';
"#,
    );
}

// ── Abstract class implementing interface ──────────────────────

#[test]
fn abstract_class_partial_interface() {
    compile_ok(
        r#"<?php
interface Lifecycle {
    public function start(): void;
    public function stop(): void;
    public function isRunning(): bool;
}
abstract class BaseService implements Lifecycle {
    protected bool $running = false;
    public function isRunning(): bool { return $this->running; }
    // start() and stop() left to subclasses
}
class HttpService extends BaseService {
    public function start(): void { $this->running = true; }
    public function stop(): void  { $this->running = false; }
}
$svc = new HttpService();
echo $svc->isRunning() ? 'running' : 'stopped';
$svc->start();
echo $svc->isRunning() ? ':running' : ':stopped';
"#,
    );
}

// ── Multiple interface implementation ──────────────────────────

#[test]
fn class_implements_multiple_interfaces() {
    compile_ok(
        r#"<?php
interface Printable  { public function print(): void; }
interface Saveable   { public function save(): bool; }
interface Deletable  { public function delete(): bool; }
class Record implements Printable, Saveable, Deletable {
    public function print(): void  { echo 'printing'; }
    public function save(): bool   { return true; }
    public function delete(): bool { return true; }
}
$r = new Record();
$r->print();
echo $r->save()   ? ':saved'   : ':save failed';
echo $r->delete() ? ':deleted' : ':delete failed';
"#,
    );
}

#[test]
fn interface_multiple_type_checks() {
    compile_ok(
        r#"<?php
interface Countable2  { public function count2(): int; }
interface Iterable2   { public function toArray(): array; }
class Collection implements Countable2, Iterable2 {
    private array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function count2(): int { return count($this->items); }
    public function toArray(): array { return $this->items; }
}
$c = new Collection([1, 2, 3]);
echo ($c instanceof Countable2) ? 'countable' : 'not countable';
echo ($c instanceof Iterable2)  ? ':iterable' : ':not iterable';
echo ':' . $c->count2();
"#,
    );
}

// ── Interface default method workaround ───────────────────────

#[test]
fn interface_with_trait_default() {
    compile_ok(
        r#"<?php
interface Hashable { public function hash(): string; }
trait DefaultHash {
    public function hash(): string { return md5(serialize($this)); }
}
class User implements Hashable {
    use DefaultHash;
    public function __construct(public string $name) {}
}
$u = new User('alice');
echo strlen($u->hash()) === 32 ? 'valid hash' : 'invalid hash';
"#,
    );
}

// ── Interface segregation pattern ──────────────────────────────

#[test]
fn interface_segregation() {
    compile_ok(
        r#"<?php
interface CanRead  { public function read(string $key): mixed; }
interface CanWrite { public function write(string $key, mixed $value): void; }
interface CanDelete { public function delete(string $key): void; }
interface Cache extends CanRead, CanWrite, CanDelete {}
class InMemoryCache implements Cache {
    private array $store = [];
    public function read(string $key): mixed   { return $this->store[$key] ?? null; }
    public function write(string $key, mixed $value): void { $this->store[$key] = $value; }
    public function delete(string $key): void  { unset($this->store[$key]); }
}
function cacheValue(CanWrite $cache, string $key, mixed $value): void {
    $cache->write($key, $value);
}
function readValue(CanRead $cache, string $key): mixed {
    return $cache->read($key);
}
$c = new InMemoryCache();
cacheValue($c, 'name', 'Alice');
echo readValue($c, 'name');
"#,
    );
}

// ── Interface constant visibility (PHP 8.1+) ──────────────────

#[test]
fn interface_constant_public_only() {
    compile_ok(
        r#"<?php
interface Configurable {
    const string DEFAULT_HOST = 'localhost';
    const int    DEFAULT_PORT = 8080;
}
class Server implements Configurable {
    public string $host;
    public int    $port;
    public function __construct() {
        $this->host = self::DEFAULT_HOST;
        $this->port = self::DEFAULT_PORT;
    }
}
$s = new Server();
echo $s->host . ':' . $s->port;
"#,
    );
}

#[test]
fn interface_runtime_dispatch_by_contract() {
    let out = run_prints(
        r#"<?php
interface Logger {
    public function log(string $message): void;
}
class MemoryLogger implements Logger {
    public array $events = [];
    public function log(string $message): void {
        $this->events[] = $message;
    }
}

function audit(Logger $logger): void {
    $logger->log('start');
    $logger->log('end');
}

$logger = new MemoryLogger();
audit($logger);
echo implode(',', $logger->events);
"#,
    );
    assert_eq!(out, vec!["start,end"]);
}

#[test]
fn interface_implements_multiple_dispatched_order() {
    let out = run_prints(
        r#"<?php
interface Reader { public function read(): string; }
interface Writer { public function write(string $v): string; }

class Document implements Reader, Writer {
    public function read(): string { return 'read'; }
    public function write(string $v): string { return 'write:' . $v; }
}

$doc = new Document();
echo $doc->read() . '|' . $doc->write('v');
"#,
    );
    assert_eq!(out, vec!["read|write:v"]);
}

#[test]
fn interface_cast_after_dynamic_binding() {
    let out = run_prints(
        r#"<?php
interface IAnimal { public function kind(): string; }
interface IPet extends IAnimal { public function name(): string; }

class Cat implements IPet {
    public function __construct(private string $nameValue) {}
    public function kind(): string { return 'cat'; }
    public function name(): string { return $this->nameValue; }
}

$animal = new Cat('Misty');
echo ($animal instanceof IAnimal ? 'a' : 'x') . '-' . ($animal instanceof IPet ? 'p' : 'y');
echo '-' . $animal->kind() . ':' . $animal->name();
"#,
    );
    assert_eq!(out, vec!["a-p-cat:Misty"]);
}

#[test]
fn interface_static_constant_from_implementer() {
    let out = run_prints(
        r#"<?php
interface Versioned {
    public const string CHANNEL = 'base';
}
class Adapter implements Versioned {}
echo Adapter::CHANNEL;
"#,
    );
    assert_eq!(out, vec!["base"]);
}

#[test]
fn interface_covariant_return_runtime() {
    let out = run_prints(
        r#"<?php
class Base {
    public function __construct(public string $n) {}
}
class Child extends Base {}

interface Maker {
    public function make(string $value): Base;
}
class ChildMaker implements Maker {
    public function make(string $value): Child {
        return new Child($value);
    }
}
$maker = new ChildMaker();
$obj = $maker->make('x');
echo ($obj instanceof Child ? 'child' : 'other') . '|' . $obj->n;
"#,
    );
    assert_eq!(out, vec!["child|x"]);
}

#[test]
fn interface_anonymous_class_runtime() {
    let out = run_prints(
        r#"<?php
interface Notifier {
    public function notify(string $message): string;
}
$notifier = new class implements Notifier {
    public function notify(string $message): string {
        return 'notify:' . $message;
    }
};
echo $notifier->notify('ping');
"#,
    );
    assert_eq!(out, vec!["notify:ping"]);
}

#[test]
fn interface_implements_and_iterator_pattern_runtime() {
    let out = run_prints(
        r#"<?php
interface Seq {
    public function next(): ?int;
}
class RangeSeq implements Seq {
    private int $current;
    public function __construct(private int $end, int $start = 1) {
        $this->current = $start;
    }
    public function next(): ?int {
        if ($this->current > $this->end) {
            return null;
        }
        return $this->current++;
    }
}
function takeThree(Seq $seq): array {
    $out = [];
    for ($i = 0; $i < 3; $i++) {
        $value = $seq->next();
        if ($value === null) {
            break;
        }
        $out[] = $value;
    }
    return $out;
}
$seq = new RangeSeq(4, 2);
echo implode(',', takeThree($seq));
"#,
    );
    assert_eq!(out, vec!["2,3,4"]);
}

#[test]
fn interface_union_type_accepts_contracts_runtime() {
    let out = run_prints(
        r#"<?php
interface Readable { public function read(): string; }
interface Writable { public function write(string $value): void; }

class Buffer implements Readable, Writable {
    private string $value = '';
    public function read(): string { return $this->value; }
    public function write(string $value): void { $this->value = $value; }
}

function appendSuffix(Readable|Writable $target): string {
    if ($target instanceof Writable) {
        $target->write('ok');
    }
    if ($target instanceof Readable) {
        return $target->read();
    }
    return '';
}
$target = new Buffer();
echo appendSuffix($target);
// output reflects the write done above
echo '|' . $target->read();
"#,
    );
    assert_eq!(out, vec!["ok|ok"]);
}

#[test]
fn interface_static_method_dispatch_inheritance_runtime() {
    let out = run_prints(
        r#"<?php
interface Identifiable {
    public static function kind(): string;
}
class Base implements Identifiable {
    public static function kind(): string {
        return 'base';
    }
}
class Child extends Base {
    public static function kind(): string {
        return 'child';
    }
}
function typeKind(string $name): string {
    return $name::kind();
}
echo typeKind(Base::class) . '|' . typeKind(Child::class);
"#,
    );
    assert_eq!(out, vec!["base|child"]);
}

#[test]
fn interface_casting_and_fallback_runtime() {
    let out = run_prints(
        r#"<?php
interface MarkerA { public function marker(): string; }
interface MarkerB { public function marker(): string; }

class ItemA implements MarkerA { public function marker(): string { return 'A'; } }
class ItemB implements MarkerB { public function marker(): string { return 'B'; } }

$list = [new ItemA(), new ItemB()];
$out = [];
foreach ($list as $item) {
    if ($item instanceof MarkerA) {
        $out[] = 'A';
    } elseif ($item instanceof MarkerB) {
        $out[] = 'B';
    } else {
        $out[] = 'U';
    }
}
echo implode('', $out);
"#,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn interface_class_implements_runtime() {
    let out = run_prints(
        r#"<?php
interface A {}
interface B {}
class C implements A, B {}
echo array_key_exists('A', class_implements(C::class)) ? 'A' : 'X';
echo '|';
echo array_key_exists('B', class_implements(C::class)) ? 'B' : 'X';
"#,
    );
    assert_eq!(out, vec!["A|B"]);
}
