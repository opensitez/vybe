use super::helpers::{compile_ok, run_prints};

// ── Intersection type — parameter must satisfy both interfaces

#[test] fn intersection_type_parameter_accepted() {
    assert_eq!(run_prints(r#"<?php
interface Serializable2 { public function serialize(): string; }
interface Loggable { public function log(): void; }
class Payload implements Serializable2, Loggable {
    public function serialize(): string { return "data"; }
    public function log(): void { echo "logged"; }
}
function process(Serializable2&Loggable $obj): string {
    $obj->log();
    return $obj->serialize();
}
echo process(new Payload());
"#), vec!["logged", "data"]);
}

#[test] fn intersection_type_three_interfaces() {
    assert_eq!(run_prints(r#"<?php
interface A { public function a(): int; }
interface B { public function b(): int; }
interface C { public function c(): int; }
class ABC implements A, B, C {
    public function a(): int { return 1; }
    public function b(): int { return 2; }
    public function c(): int { return 3; }
}
function sum(A&B&C $obj): int { return $obj->a() + $obj->b() + $obj->c(); }
echo sum(new ABC());
"#), vec!["6"]);
}

// ── Intersection in return type ───────────────────────────────

#[test] fn intersection_return_type_both_methods_callable() {
    assert_eq!(run_prints(r#"<?php
interface Countable3 { public function count(): int; }
interface Iterable2 { public function toArray(): array; }
class Collection implements Countable3, Iterable2 {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function count(): int { return count($this->data); }
    public function toArray(): array { return $this->data; }
}
function getCollection(): Countable3&Iterable2 { return new Collection([1,2,3]); }
$c = getCollection();
echo $c->count() . ',' . implode(',', $c->toArray());
"#), vec!["3,1,2,3"]);
}

// ── DNF types (A&B)|C ────────────────────────────────────────

#[test] fn dnf_type_union_of_intersection_and_class() {
    assert_eq!(run_prints(r#"<?php
interface Readable { public function read(): string; }
interface Writable { public function write(string $s): void; }
class Stream implements Readable, Writable {
    private string $buf = '';
    public function read(): string { return $this->buf; }
    public function write(string $s): void { $this->buf .= $s; }
}
class NullStream {
    public function write(string $s): void {}
}
function writeIfPossible((Readable&Writable)|NullStream $s, string $data): void {
    $s->write($data);
}
$s = new Stream();
writeIfPossible($s, "hello");
echo $s->read();
"#), vec!["hello"]);
}

#[test] fn dnf_nullable_intersection_type() {
    assert_eq!(run_prints(r#"<?php
interface Identifiable { public function id(): int; }
interface Named { public function name(): string; }
class User implements Identifiable, Named {
    public function __construct(private int $id, private string $name) {}
    public function id(): int { return $this->id; }
    public function name(): string { return $this->name; }
}
function describe((Identifiable&Named)|null $entity): string {
    if ($entity === null) return 'none';
    return $entity->id() . ':' . $entity->name();
}
echo describe(new User(1, 'Alice')) . ',' . describe(null);
"#), vec!["1:Alice,none"]);
}

// ── Intersection type with abstract class satisfaction ────────

#[test] fn intersection_requires_both_implementations() {
    assert_eq!(run_prints(r#"<?php
interface Printable { public function print(): void; }
interface Saveable { public function save(): bool; }
class Document implements Printable, Saveable {
    public function print(): void { echo "printing"; }
    public function save(): bool { echo " saving"; return true; }
}
function process(Printable&Saveable $doc): void { $doc->print(); $doc->save(); }
process(new Document());
"#), vec!["printing saving"]);
}

// ── Intersection type in class property ───────────────────────

#[test] fn intersection_type_class_property() {
    compile_ok(r#"<?php
interface Closeable { public function close(): void; }
interface Flushable { public function flush(): void; }
class Buffer implements Closeable, Flushable {
    private string $data = '';
    public function write(string $s): void { $this->data .= $s; }
    public function flush(): void { echo $this->data; $this->data = ''; }
    public function close(): void { $this->flush(); }
}
class Writer {
    public Closeable&Flushable $buffer;
    public function __construct(Closeable&Flushable $buf) { $this->buffer = $buf; }
}
$w = new Writer(new Buffer());
"#);
}

// ── instanceof with intersection ──────────────────────────────

#[test] fn instanceof_checks_each_interface_separately() {
    assert_eq!(run_prints(r#"<?php
interface Shape { public function area(): float; }
interface Drawable { public function draw(): void; }
class Circle implements Shape, Drawable {
    public function area(): float { return 3.14; }
    public function draw(): void {}
}
$c = new Circle();
echo ($c instanceof Shape ? '1' : '0') . ($c instanceof Drawable ? '1' : '0');
"#), vec!["11"]);
}

// ── DNF in foreach with type-narrowing ────────────────────────

#[test] fn dnf_type_used_in_function_dispatch() {
    assert_eq!(run_prints(r#"<?php
interface Sizeable { public function size(): int; }
interface Nameable { public function name(): string; }
class File implements Sizeable, Nameable {
    public function __construct(private string $n, private int $s) {}
    public function size(): int { return $this->s; }
    public function name(): string { return $this->n; }
}
class Unknown {}
function describe((Sizeable&Nameable)|Unknown $item): string {
    if ($item instanceof Unknown) return 'unknown';
    return $item->name() . ':' . $item->size();
}
$items = [new File('a.txt', 100), new Unknown(), new File('b.php', 200)];
$results = array_map('describe', $items);
echo implode(',', $results);
"#), vec!["a.txt:100,unknown,b.php:200"]);
}

// ── Intersection in generic-style container ───────────────────

#[test] fn intersection_typed_collection_add_retrieve() {
    assert_eq!(run_prints(r#"<?php
interface Hashable { public function hash(): string; }
interface Comparable2 { public function compareTo(mixed $other): int; }
class SortedKey implements Hashable, Comparable2 {
    public function __construct(private string $key) {}
    public function hash(): string { return md5($this->key); }
    public function compareTo(mixed $other): int { return strcmp($this->key, $other->key); }
}
class TypedSet {
    private array $items = [];
    public function add(Hashable&Comparable2 $item): void { $this->items[$item->hash()] = $item; }
    public function count(): int { return count($this->items); }
}
$set = new TypedSet();
$set->add(new SortedKey('foo'));
$set->add(new SortedKey('bar'));
echo $set->count();
"#), vec!["2"]);
}

// ── Compile-only: intersection type in closure parameter ──────

#[test] fn intersection_type_in_closure_parameter() {
    compile_ok(r#"<?php
interface Readable2 { public function read(): string; }
interface Seekable { public function seek(int $pos): void; }
$process = function(Readable2&Seekable $stream): string {
    $stream->seek(0);
    return $stream->read();
};
"#);
}

// ── Intersection type with null coalescing ────────────────────

#[test] fn dnf_type_with_null_coalescing() {
    assert_eq!(run_prints(r#"<?php
interface HasId { public function id(): int; }
interface HasName { public function name(): string; }
class Entity implements HasId, HasName {
    public function __construct(private int $id, private string $name) {}
    public function id(): int { return $this->id; }
    public function name(): string { return $this->name; }
}
function display((HasId&HasName)|null $e): string {
    return $e?->name() ?? 'anonymous';
}
echo display(new Entity(1, 'Alice')) . ',' . display(null);
"#), vec!["Alice,anonymous"]);
}

// ── Multiple DNF arms ─────────────────────────────────────────

#[test] fn dnf_multiple_intersection_groups_in_union() {
    compile_ok(r#"<?php
interface A2 { public function a(): void; }
interface B2 { public function b(): void; }
interface C2 { public function c(): void; }
interface D2 { public function d(): void; }
function process((A2&B2)|(C2&D2) $obj): void {}
"#);
}

// ── Intersection type with static methods ─────────────────────

#[test] fn intersection_interface_static_not_part_of_intersection() {
    assert_eq!(run_prints(r#"<?php
interface Activatable { public function activate(): void; }
interface Deactivatable { public function deactivate(): void; }
class Toggle implements Activatable, Deactivatable {
    private bool $on = false;
    public function activate(): void { $this->on = true; }
    public function deactivate(): void { $this->on = false; }
    public function isOn(): bool { return $this->on; }
}
function toggle(Activatable&Deactivatable $t): void {
    $t->activate();
    $t->deactivate();
}
$t = new Toggle();
toggle($t);
echo $t->isOn() ? 'on' : 'off';
"#), vec!["off"]);
}

// ── Type checking at runtime for intersection ─────────────────

#[test] fn runtime_interface_check_simulates_intersection() {
    assert_eq!(run_prints(r#"<?php
interface Loggable2 { public function log(): string; }
interface Auditable { public function audit(): string; }
class Event implements Loggable2, Auditable {
    public function log(): string { return "log"; }
    public function audit(): string { return "audit"; }
}
function verify(object $obj): string {
    if (!($obj instanceof Loggable2 && $obj instanceof Auditable)) return 'invalid';
    return $obj->log() . '+' . $obj->audit();
}
echo verify(new Event());
"#), vec!["log+audit"]);
}

// ── Intersection in abstract class ───────────────────────────

#[test] fn abstract_method_with_intersection_parameter() {
    assert_eq!(run_prints(r#"<?php
interface Displayable { public function display(): string; }
interface Exportable { public function export(): array; }
abstract class Processor {
    abstract public function handle(Displayable&Exportable $obj): string;
}
class ConcreteProcessor extends Processor {
    public function handle(Displayable&Exportable $obj): string {
        return $obj->display() . ':' . count($obj->export());
    }
}
class Widget implements Displayable, Exportable {
    public function display(): string { return "widget"; }
    public function export(): array { return ['a', 'b']; }
}
echo (new ConcreteProcessor())->handle(new Widget());
"#), vec!["widget:2"]);
}

// ── Intersection type in interface method ─────────────────────

#[test] fn interface_method_intersection_return_type() {
    assert_eq!(run_prints(r#"<?php
interface Source { public function source(): string; }
interface Sink { public function sink(string $s): void; }
interface Pipe {
    public function getTransformer(): Source&Sink;
}
class PassThrough implements Source, Sink {
    private string $buf = '';
    public function source(): string { return $this->buf; }
    public function sink(string $s): void { $this->buf = strtoupper($s); }
}
class Pipeline implements Pipe {
    private PassThrough $t;
    public function __construct() { $this->t = new PassThrough(); }
    public function getTransformer(): Source&Sink { return $this->t; }
}
$p = new Pipeline();
$t = $p->getTransformer();
$t->sink("hello");
echo $t->source();
"#), vec!["HELLO"]);
}

// ── DNF in first-class callable ──────────────────────────────

#[test] fn dnf_type_enforced_via_instanceof_chain() {
    assert_eq!(run_prints(r#"<?php
interface Readable3 { public function read(): string; }
interface Closeable2 { public function close(): void; }
class FileHandle implements Readable3, Closeable2 {
    private bool $closed = false;
    public function read(): string { return $this->closed ? '' : "data"; }
    public function close(): void { $this->closed = true; }
}
function readAndClose((Readable3&Closeable2)|null $handle): string {
    if ($handle === null) return 'null';
    $data = $handle->read();
    $handle->close();
    return $data;
}
echo readAndClose(new FileHandle()) . ',' . readAndClose(null);
"#), vec!["data,null"]);
}

// ── Intersection type variance ────────────────────────────────

#[test] fn intersection_type_narrower_is_subtype_of_single_interface() {
    assert_eq!(run_prints(r#"<?php
interface Worker { public function work(): string; }
interface Reporter { public function report(): string; }
class FullEmployee implements Worker, Reporter {
    public function work(): string { return "working"; }
    public function report(): string { return "reporting"; }
}
function getWorker(): Worker { return new FullEmployee(); }
$w = getWorker();
echo $w->work();
"#), vec!["working"]);
}

// ── Intersection type accepted by union ───────────────────────

#[test] fn class_satisfying_intersection_accepted_as_single_type() {
    assert_eq!(run_prints(r#"<?php
interface Printable2 { public function print2(): string; }
interface Saveable2 { public function save2(): string; }
class Doc implements Printable2, Saveable2 {
    public function print2(): string { return "print"; }
    public function save2(): string { return "save"; }
}
function useDoc(Printable2 $p): string { return $p->print2(); }
$doc = new Doc();
echo useDoc($doc);
"#), vec!["print"]);
}

// ── PHP 8.2 null as standalone type in union ──────────────────

#[test] fn null_standalone_type_in_union() {
    assert_eq!(run_prints(r#"<?php
function maybeNull(bool $returnNull): null|string {
    return $returnNull ? null : "value";
}
echo maybeNull(false) . ',' . var_export(maybeNull(true), true);
"#), vec!["value,NULL"]);
}
