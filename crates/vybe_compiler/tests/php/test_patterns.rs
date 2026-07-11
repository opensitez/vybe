use super::helpers::{compile_ok, run_prints};

// ── Singleton ────────────────────────────────────────────────────

#[test]
fn singleton_same_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    private static $instance = null;
    public $data = [];
    private function __construct() {}
    public static function getInstance() {
        if (Config::$instance === null) {
            Config::$instance = new Config();
        }
        return Config::$instance;
    }
    public function set($k, $v) { $this->data[$k] = $v; }
    public function get($k) { return $this->data[$k] ?? null; }
}
$a = Config::getInstance();
$a->set('env', 'prod');
$b = Config::getInstance();
echo $b->get('env');
echo ($a === $b) ? 'same' : 'different';
"#
        ),
        vec!["prodsame"]
    );
}

// ── Factory method ───────────────────────────────────────────────

#[test]
fn factory_method_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Transport {
    abstract public function getType(): string;
}
class Truck extends Transport {
    public function getType(): string { return 'truck'; }
}
class Ship extends Transport {
    public function getType(): string { return 'ship'; }
}
abstract class Logistics {
    abstract public function createTransport(): Transport;
    public function plan(): string { return $this->createTransport()->getType(); }
}
class RoadLogistics extends Logistics {
    public function createTransport(): Transport { return new Truck(); }
}
class SeaLogistics extends Logistics {
    public function createTransport(): Transport { return new Ship(); }
}
echo (new RoadLogistics())->plan();
echo (new SeaLogistics())->plan();
"#
        ),
        vec!["truckship"]
    );
}

// ── Abstract factory ─────────────────────────────────────────────

#[test]
fn abstract_factory_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Button { public function render(): string; }
interface Checkbox { public function check(): string; }
class WinButton implements Button { public function render(): string { return 'win-button'; } }
class WinCheckbox implements Checkbox { public function check(): string { return 'win-check'; } }
class MacButton implements Button { public function render(): string { return 'mac-button'; } }
class MacCheckbox implements Checkbox { public function check(): string { return 'mac-check'; } }
interface GUIFactory {
    public function createButton(): Button;
    public function createCheckbox(): Checkbox;
}
class WinFactory implements GUIFactory {
    public function createButton(): Button { return new WinButton(); }
    public function createCheckbox(): Checkbox { return new WinCheckbox(); }
}
class MacFactory implements GUIFactory {
    public function createButton(): Button { return new MacButton(); }
    public function createCheckbox(): Checkbox { return new MacCheckbox(); }
}
function buildUI(GUIFactory $f) {
    echo $f->createButton()->render();
    echo $f->createCheckbox()->check();
}
buildUI(new WinFactory());
buildUI(new MacFactory());
"#
        ),
        vec!["win-buttonwin-checkmac-buttonmac-check"]
    );
}

// ── Builder ──────────────────────────────────────────────────────

#[test]
fn builder_fluent_build() {
    assert_eq!(
        run_prints(
            r#"<?php
class Pizza {
    public $size = '';
    public $toppings = [];
    public $crust = '';
}
class PizzaBuilder {
    private $pizza;
    public function __construct() { $this->pizza = new Pizza(); }
    public function size(string $s): self { $this->pizza->size = $s; return $this; }
    public function crust(string $c): self { $this->pizza->crust = $c; return $this; }
    public function topping(string $t): self { $this->pizza->toppings[] = $t; return $this; }
    public function build(): Pizza { return $this->pizza; }
}
$p = (new PizzaBuilder())
    ->size('large')
    ->crust('thin')
    ->topping('mozzarella')
    ->topping('pepperoni')
    ->build();
echo $p->size;
echo $p->crust;
echo implode(',', $p->toppings);
"#
        ),
        vec!["largethinmozzarella,pepperoni"]
    );
}

// ── Prototype ────────────────────────────────────────────────────

#[test]
fn prototype_clone_variation() {
    assert_eq!(
        run_prints(
            r#"<?php
class Shape {
    public $color;
    public $x;
    public $y;
    public function __construct(string $color, int $x, int $y) {
        $this->color = $color;
        $this->x = $x;
        $this->y = $y;
    }
    public function move(int $dx, int $dy): self {
        $clone = clone $this;
        $clone->x += $dx;
        $clone->y += $dy;
        return $clone;
    }
}
$s1 = new Shape('red', 0, 0);
$s2 = $s1->move(5, 10);
echo $s1->x . ',' . $s1->y;
echo $s2->x . ',' . $s2->y;
echo $s2->color;
"#
        ),
        vec!["0,05,10red"]
    );
}

// ── Adapter ──────────────────────────────────────────────────────

#[test]
fn adapter_wraps_incompatible_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
class LegacyPrinter {
    public function printText(string $text): void { echo 'Legacy: ' . $text; }
}
interface ModernPrinter {
    public function print(string $text): void;
}
class PrinterAdapter implements ModernPrinter {
    private $legacy;
    public function __construct(LegacyPrinter $l) { $this->legacy = $l; }
    public function print(string $text): void { $this->legacy->printText($text); }
}
function usePrinter(ModernPrinter $p, string $text): void { $p->print($text); }
usePrinter(new PrinterAdapter(new LegacyPrinter()), 'hello');
"#
        ),
        vec!["Legacy: hello"]
    );
}

// ── Decorator ────────────────────────────────────────────────────

#[test]
fn decorator_wraps_same_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Logger {
    public function log(string $msg): void;
}
class ConsoleLogger implements Logger {
    public function log(string $msg): void { echo $msg; }
}
class TimestampDecorator implements Logger {
    private $inner;
    public function __construct(Logger $l) { $this->inner = $l; }
    public function log(string $msg): void { $this->inner->log('[ts] ' . $msg); }
}
class PrefixDecorator implements Logger {
    private $inner;
    private $prefix;
    public function __construct(Logger $l, string $p) { $this->inner = $l; $this->prefix = $p; }
    public function log(string $msg): void { $this->inner->log($this->prefix . ': ' . $msg); }
}
$log = new PrefixDecorator(new TimestampDecorator(new ConsoleLogger()), 'APP');
$log->log('started');
"#
        ),
        vec!["APP: [ts] started"]
    );
}

// ── Composite ────────────────────────────────────────────────────

#[test]
fn composite_tree_sum() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Component {
    public function price(): int;
}
class Leaf implements Component {
    private $cost;
    public function __construct(int $cost) { $this->cost = $cost; }
    public function price(): int { return $this->cost; }
}
class Composite implements Component {
    private $children = [];
    public function add(Component $c): void { $this->children[] = $c; }
    public function price(): int {
        return array_sum(array_map(fn($c) => $c->price(), $this->children));
    }
}
$box = new Composite();
$box->add(new Leaf(10));
$inner = new Composite();
$inner->add(new Leaf(5));
$inner->add(new Leaf(15));
$box->add($inner);
echo $box->price();
"#
        ),
        vec!["30"]
    );
}

// ── Facade ───────────────────────────────────────────────────────

#[test]
fn facade_simple_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
class VideoDecoder {
    public function decode(string $f): string { return 'decoded:' . $f; }
}
class AudioMixer {
    public function mix(string $a): string { return 'mixed:' . $a; }
}
class VideoFacade {
    private $decoder;
    private $mixer;
    public function __construct() {
        $this->decoder = new VideoDecoder();
        $this->mixer = new AudioMixer();
    }
    public function process(string $file): string {
        $v = $this->decoder->decode($file);
        $a = $this->mixer->mix($file);
        return $v . '|' . $a;
    }
}
echo (new VideoFacade())->process('movie.mp4');
"#
        ),
        vec!["decoded:movie.mp4|mixed:movie.mp4"]
    );
}

// ── Proxy ────────────────────────────────────────────────────────

#[test]
fn proxy_lazy_loading() {
    assert_eq!(
        run_prints(
            r#"<?php
interface ImageInterface {
    public function display(): string;
}
class RealImage implements ImageInterface {
    private $filename;
    public function __construct(string $f) {
        $this->filename = $f;
        echo 'loaded:' . $f;
    }
    public function display(): string { return 'showing:' . $this->filename; }
}
class ImageProxy implements ImageInterface {
    private $filename;
    private $real = null;
    public function __construct(string $f) { $this->filename = $f; }
    public function display(): string {
        if ($this->real === null) {
            $this->real = new RealImage($this->filename);
        }
        return $this->real->display();
    }
}
$img = new ImageProxy('photo.jpg');
echo 'proxy created';
echo $img->display();
"#
        ),
        vec!["proxy createdloaded:photo.jpgshowing:photo.jpg"]
    );
}

// ── Observer ─────────────────────────────────────────────────────

#[test]
fn observer_notifies_all() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Observer {
    public function update(string $event, $data): void;
}
class EventBus {
    private $listeners = [];
    public function subscribe(string $event, Observer $o): void {
        $this->listeners[$event][] = $o;
    }
    public function emit(string $event, $data): void {
        foreach ($this->listeners[$event] ?? [] as $o) {
            $o->update($event, $data);
        }
    }
}
class LogObserver implements Observer {
    private $name;
    public function __construct(string $n) { $this->name = $n; }
    public function update(string $event, $data): void { echo $this->name . ':' . $data; }
}
$bus = new EventBus();
$bus->subscribe('login', new LogObserver('A'));
$bus->subscribe('login', new LogObserver('B'));
$bus->emit('login', 'alice');
"#
        ),
        vec!["A:aliceB:alice"]
    );
}

// ── Strategy ─────────────────────────────────────────────────────

#[test]
fn strategy_interchangeable_sorter() {
    assert_eq!(
        run_prints(
            r#"<?php
interface SortStrategy {
    public function sort(array $data): array;
}
class AscendingSort implements SortStrategy {
    public function sort(array $data): array { sort($data); return $data; }
}
class DescendingSort implements SortStrategy {
    public function sort(array $data): array { rsort($data); return $data; }
}
class Sorter {
    private $strategy;
    public function __construct(SortStrategy $s) { $this->strategy = $s; }
    public function sort(array $data): array { return $this->strategy->sort($data); }
}
$s = new Sorter(new AscendingSort());
echo implode(',', $s->sort([3, 1, 4, 1, 5]));
$s2 = new Sorter(new DescendingSort());
echo implode(',', $s2->sort([3, 1, 4, 1, 5]));
"#
        ),
        vec!["1,1,3,4,55,4,3,1,1"]
    );
}

// ── Template method ──────────────────────────────────────────────

#[test]
fn template_method_skeleton() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class DataExporter {
    final public function export(): string {
        $data = $this->fetchData();
        $formatted = $this->format($data);
        return $this->output($formatted);
    }
    abstract protected function fetchData(): array;
    abstract protected function format(array $data): string;
    protected function output(string $s): string { return 'OUT:' . $s; }
}
class CsvExporter extends DataExporter {
    protected function fetchData(): array { return [1, 2, 3]; }
    protected function format(array $data): string { return implode(',', $data); }
}
class JsonExporter extends DataExporter {
    protected function fetchData(): array { return ['a' => 1]; }
    protected function format(array $data): string { return json_encode($data); }
}
echo (new CsvExporter())->export();
echo (new JsonExporter())->export();
"#
        ),
        vec!["OUT:1,2,3OUT:{\"a\":1}"]
    );
}

// ── Command ──────────────────────────────────────────────────────

#[test]
fn command_undo_support() {
    assert_eq!(
        run_prints(
            r#"<?php
class TextEditor {
    public $text = '';
    public function append(string $s): void { $this->text .= $s; }
    public function deleteLast(int $n): void { $this->text = substr($this->text, 0, strlen($this->text) - $n); }
}
interface Command {
    public function execute(): void;
    public function undo(): void;
}
class AppendCommand implements Command {
    private $editor;
    private $text;
    public function __construct(TextEditor $e, string $t) { $this->editor = $e; $this->text = $t; }
    public function execute(): void { $this->editor->append($this->text); }
    public function undo(): void { $this->editor->deleteLast(strlen($this->text)); }
}
$editor = new TextEditor();
$history = [];
$c1 = new AppendCommand($editor, 'Hello');
$c1->execute();
$history[] = $c1;
$c2 = new AppendCommand($editor, ' World');
$c2->execute();
$history[] = $c2;
echo $editor->text;
array_pop($history)->undo();
echo $editor->text;
"#
        ),
        vec!["Hello WorldHello"]
    );
}

// ── Iterator ─────────────────────────────────────────────────────

#[test]
fn custom_iterator_traversal() {
    assert_eq!(
        run_prints(
            r#"<?php
class NumberRange implements Iterator {
    private $current;
    public function __construct(private int $start, private int $end) {
        $this->current = $start;
    }
    public function current(): int { return $this->current; }
    public function key(): int { return $this->current - $this->start; }
    public function next(): void { $this->current++; }
    public function rewind(): void { $this->current = $this->start; }
    public function valid(): bool { return $this->current <= $this->end; }
}
$range = new NumberRange(1, 5);
foreach ($range as $n) { echo $n; }
"#
        ),
        vec!["12345"]
    );
}

// ── State machine ────────────────────────────────────────────────

#[test]
fn state_machine_transitions() {
    assert_eq!(
        run_prints(
            r#"<?php
interface TrafficState {
    public function next(): TrafficState;
    public function color(): string;
}
class Red implements TrafficState {
    public function next(): TrafficState { return new Green(); }
    public function color(): string { return 'red'; }
}
class Green implements TrafficState {
    public function next(): TrafficState { return new Yellow(); }
    public function color(): string { return 'green'; }
}
class Yellow implements TrafficState {
    public function next(): TrafficState { return new Red(); }
    public function color(): string { return 'yellow'; }
}
$state = new Red();
for ($i = 0; $i < 4; $i++) {
    echo $state->color();
    $state = $state->next();
}
"#
        ),
        vec!["redgreenyellowred"]
    );
}

// ── Chain of responsibility ──────────────────────────────────────

#[test]
fn chain_of_responsibility_passes_along() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Handler {
    private $next = null;
    public function setNext(Handler $h): Handler { $this->next = $h; return $h; }
    public function handle(int $req): string {
        if ($this->next !== null) return $this->next->handle($req);
        return 'unhandled';
    }
}
class SmallHandler extends Handler {
    public function handle(int $req): string {
        if ($req < 10) return 'small:' . $req;
        return parent::handle($req);
    }
}
class MediumHandler extends Handler {
    public function handle(int $req): string {
        if ($req < 100) return 'medium:' . $req;
        return parent::handle($req);
    }
}
class LargeHandler extends Handler {
    public function handle(int $req): string { return 'large:' . $req; }
}
$small = new SmallHandler();
$medium = new MediumHandler();
$large = new LargeHandler();
$small->setNext($medium)->setNext($large);
echo $small->handle(5);
echo $small->handle(50);
echo $small->handle(500);
"#
        ),
        vec!["small:5medium:50large:500"]
    );
}

// ── Mediator ─────────────────────────────────────────────────────

#[test]
fn mediator_centralized_communication() {
    assert_eq!(
        run_prints(
            r#"<?php
class ChatRoom {
    private $log = [];
    public function send(string $from, string $to, string $msg): void {
        $this->log[] = "$from->$to:$msg";
    }
    public function getLog(): array { return $this->log; }
}
class User {
    public function __construct(private string $name, private ChatRoom $room) {}
    public function send(string $to, string $msg): void { $this->room->send($this->name, $to, $msg); }
}
$room = new ChatRoom();
$alice = new User('Alice', $room);
$bob = new User('Bob', $room);
$alice->send('Bob', 'hi');
$bob->send('Alice', 'hello');
foreach ($room->getLog() as $entry) { echo $entry; }
"#
        ),
        vec!["Alice->Bob:hiBob->Alice:hello"]
    );
}

// ── Memento ──────────────────────────────────────────────────────

#[test]
fn memento_capture_restore() {
    assert_eq!(
        run_prints(
            r#"<?php
class EditorMemento {
    public function __construct(public readonly string $content) {}
}
class Editor {
    public $content = '';
    public function save(): EditorMemento { return new EditorMemento($this->content); }
    public function restore(EditorMemento $m): void { $this->content = $m->content; }
}
$e = new Editor();
$e->content = 'version1';
$snap = $e->save();
$e->content = 'version2';
echo $e->content;
$e->restore($snap);
echo $e->content;
"#
        ),
        vec!["version2version1"]
    );
}

// ── Visitor ──────────────────────────────────────────────────────

#[test]
fn visitor_double_dispatch() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Visitor {
    public function visitCircle(Circle $c): string;
    public function visitRect(Rect $r): string;
}
interface Shape {
    public function accept(Visitor $v): string;
}
class Circle implements Shape {
    public function __construct(public float $r) {}
    public function accept(Visitor $v): string { return $v->visitCircle($this); }
}
class Rect implements Shape {
    public function __construct(public float $w, public float $h) {}
    public function accept(Visitor $v): string { return $v->visitRect($this); }
}
class AreaVisitor implements Visitor {
    public function visitCircle(Circle $c): string { return (string)round(M_PI * $c->r * $c->r, 2); }
    public function visitRect(Rect $r): string { return (string)($r->w * $r->h); }
}
$v = new AreaVisitor();
echo (new Circle(2.0))->accept($v);
echo (new Rect(3.0, 4.0))->accept($v);
"#
        ),
        vec!["12.5712"]
    );
}

// ── Repository ───────────────────────────────────────────────────

#[test]
fn repository_abstracts_storage() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public function __construct(public int $id, public string $name) {}
}
class UserRepository {
    private $store = [];
    public function save(User $u): void { $this->store[$u->id] = $u; }
    public function find(int $id): ?User { return $this->store[$id] ?? null; }
    public function findAll(): array { return array_values($this->store); }
}
$repo = new UserRepository();
$repo->save(new User(1, 'Alice'));
$repo->save(new User(2, 'Bob'));
echo $repo->find(1)->name;
echo count($repo->findAll());
"#
        ),
        vec!["Alice2"]
    );
}

// ── Service locator ──────────────────────────────────────────────

#[test]
fn service_locator_resolves_by_name() {
    assert_eq!(
        run_prints(
            r#"<?php
class ServiceLocator {
    private static $services = [];
    public static function register(string $name, $service): void { self::$services[$name] = $service; }
    public static function get(string $name) { return self::$services[$name] ?? null; }
}
class Mailer {
    public function send(string $msg): void { echo 'mail:' . $msg; }
}
ServiceLocator::register('mailer', new Mailer());
$m = ServiceLocator::get('mailer');
$m->send('hello');
"#
        ),
        vec!["mail:hello"]
    );
}

// ── Dependency injection ─────────────────────────────────────────

#[test]
fn dependency_injection_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Storage {
    public function write(string $key, string $val): void;
    public function read(string $key): ?string;
}
class MemoryStorage implements Storage {
    private $data = [];
    public function write(string $key, string $val): void { $this->data[$key] = $val; }
    public function read(string $key): ?string { return $this->data[$key] ?? null; }
}
class Cache {
    public function __construct(private Storage $storage) {}
    public function put(string $k, string $v): void { $this->storage->write($k, $v); }
    public function get(string $k): ?string { return $this->storage->read($k); }
}
$cache = new Cache(new MemoryStorage());
$cache->put('key1', 'value1');
echo $cache->get('key1');
echo $cache->get('missing') ?? 'null';
"#
        ),
        vec!["value1null"]
    );
}

// ── Event dispatcher ─────────────────────────────────────────────

#[test]
fn event_dispatcher_emit_subscribe() {
    assert_eq!(
        run_prints(
            r#"<?php
class Dispatcher {
    private $handlers = [];
    public function on(string $event, callable $fn): void { $this->handlers[$event][] = $fn; }
    public function emit(string $event, $payload = null): void {
        foreach ($this->handlers[$event] ?? [] as $fn) { $fn($payload); }
    }
}
$d = new Dispatcher();
$d->on('data', fn($v) => print("got:$v\n"));
$d->on('data', fn($v) => print("also:$v\n"));
$d->emit('data', 42);
"#
        ),
        vec!["got:42", "also:42"]
    );
}

// ── Pipeline ─────────────────────────────────────────────────────

#[test]
fn pipeline_chain_callables() {
    assert_eq!(
        run_prints(
            r#"<?php
class Pipeline {
    private $stages = [];
    public function pipe(callable $fn): self { $this->stages[] = $fn; return $this; }
    public function process($payload) {
        return array_reduce($this->stages, fn($carry, $fn) => $fn($carry), $payload);
    }
}
$result = (new Pipeline())
    ->pipe(fn($s) => trim($s))
    ->pipe(fn($s) => strtoupper($s))
    ->pipe(fn($s) => str_replace(' ', '_', $s))
    ->process('  hello world  ');
echo $result;
"#
        ),
        vec!["HELLO_WORLD"]
    );
}

// ── Specification pattern ────────────────────────────────────────

#[test]
fn specification_combinable_rules() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Specification {
    public function isSatisfiedBy($candidate): bool;
}
class AndSpec implements Specification {
    public function __construct(private Specification $a, private Specification $b) {}
    public function isSatisfiedBy($c): bool { return $this->a->isSatisfiedBy($c) && $this->b->isSatisfiedBy($c); }
}
class MinAgeSpec implements Specification {
    public function __construct(private int $min) {}
    public function isSatisfiedBy($c): bool { return $c['age'] >= $this->min; }
}
class ActiveSpec implements Specification {
    public function isSatisfiedBy($c): bool { return $c['active'] === true; }
}
$spec = new AndSpec(new MinAgeSpec(18), new ActiveSpec());
$users = [
    ['age' => 25, 'active' => true],
    ['age' => 15, 'active' => true],
    ['age' => 30, 'active' => false],
];
$count = count(array_filter($users, fn($u) => $spec->isSatisfiedBy($u)));
echo $count;
"#
        ),
        vec!["1"]
    );
}

// ── Value object ─────────────────────────────────────────────────

#[test]
fn value_object_equality_by_value() {
    assert_eq!(
        run_prints(
            r#"<?php
final class Money {
    public function __construct(private int $amount, private string $currency) {}
    public function equals(Money $other): bool {
        return $this->amount === $other->amount && $this->currency === $other->currency;
    }
    public function add(Money $other): Money {
        if ($this->currency !== $other->currency) throw new \Exception('currency mismatch');
        return new Money($this->amount + $other->amount, $this->currency);
    }
    public function __toString(): string { return $this->amount . ' ' . $this->currency; }
}
$a = new Money(100, 'USD');
$b = new Money(100, 'USD');
$c = new Money(50, 'USD');
echo $a->equals($b) ? 'equal' : 'diff';
echo $a->equals($c) ? 'equal' : 'diff';
echo $a->add($c);
"#
        ),
        // Bare echo emits no newline: the three outputs concatenate (spec PHP).
        vec!["equaldiff150 USD"]
    );
}

// ── DTO ──────────────────────────────────────────────────────────

#[test]
fn dto_plain_data_carrier() {
    assert_eq!(
        run_prints(
            r#"<?php
class UserDTO {
    public function __construct(
        public readonly int $id,
        public readonly string $name,
        public readonly string $email
    ) {}
    public function toArray(): array {
        return ['id' => $this->id, 'name' => $this->name, 'email' => $this->email];
    }
}
$dto = new UserDTO(1, 'Alice', 'alice@example.com');
echo $dto->name;
echo $dto->toArray()['email'];
"#
        ),
        vec!["Alicealice@example.com"]
    );
}

// ── Null object ──────────────────────────────────────────────────

#[test]
fn null_object_avoids_null_checks() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Logger {
    public function log(string $msg): void;
}
class ConsoleLogger implements Logger {
    public function log(string $msg): void { echo $msg; }
}
class NullLogger implements Logger {
    public function log(string $msg): void {}
}
function processData(array $data, Logger $logger): int {
    $logger->log('processing');
    return count($data);
}
echo processData([1, 2, 3], new ConsoleLogger());
echo processData([1, 2], new NullLogger());
"#
        ),
        vec!["processing32"]
    );
}

// ── Flyweight ────────────────────────────────────────────────────

#[test]
fn flyweight_shared_state() {
    assert_eq!(
        run_prints(
            r#"<?php
class TreeType {
    public function __construct(public string $name, public string $color) {}
    public function draw(int $x, int $y): string { return "{$this->name}@{$x},{$y}"; }
}
class TreeFactory {
    private static $types = [];
    public static function get(string $name, string $color): TreeType {
        $key = "$name-$color";
        if (!isset(self::$types[$key])) {
            self::$types[$key] = new TreeType($name, $color);
        }
        return self::$types[$key];
    }
    public static function count(): int { return count(self::$types); }
}
$t1 = TreeFactory::get('oak', 'green');
$t2 = TreeFactory::get('oak', 'green');
$t3 = TreeFactory::get('pine', 'dark-green');
echo ($t1 === $t2) ? 'shared' : 'different';
echo TreeFactory::count();
echo $t1->draw(1, 2);
"#
        ),
        vec!["shared2oak@1,2"]
    );
}

// ── Registry ─────────────────────────────────────────────────────

#[test]
fn registry_global_named_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
class Registry {
    private static $instances = [];
    public static function set(string $key, $obj): void { self::$instances[$key] = $obj; }
    public static function get(string $key) { return self::$instances[$key] ?? null; }
    public static function has(string $key): bool { return isset(self::$instances[$key]); }
}
Registry::set('db', (object)['host' => 'localhost']);
echo Registry::has('db') ? 'found' : 'missing';
echo Registry::get('db')->host;
echo Registry::has('cache') ? 'found' : 'missing';
"#
        ),
        vec!["foundlocalhostmissing"]
    );
}

// ── Multiton ─────────────────────────────────────────────────────

#[test]
fn multiton_named_singletons() {
    assert_eq!(
        run_prints(
            r#"<?php
class Connection {
    private static $pool = [];
    private function __construct(public string $name) {}
    public static function getInstance(string $name): self {
        if (!isset(self::$pool[$name])) {
            self::$pool[$name] = new self($name);
        }
        return self::$pool[$name];
    }
}
$a = Connection::getInstance('primary');
$b = Connection::getInstance('primary');
$c = Connection::getInstance('replica');
echo ($a === $b) ? 'same' : 'diff';
echo ($a === $c) ? 'same' : 'diff';
echo $c->name;
"#
        ),
        vec!["samediffreplica"]
    );
}

// ── Result type ──────────────────────────────────────────────────

#[test]
fn result_type_ok_err() {
    assert_eq!(
        run_prints(
            r#"<?php
class Result {
    private function __construct(private bool $ok, private $value, private string $error = '') {}
    public static function ok($value): self { return new self(true, $value); }
    public static function err(string $error): self { return new self(false, null, $error); }
    public function isOk(): bool { return $this->ok; }
    public function unwrap() { if (!$this->ok) throw new \Exception($this->error); return $this->value; }
    public function error(): string { return $this->error; }
}
function divide(int $a, int $b): Result {
    if ($b === 0) return Result::err('division by zero');
    return Result::ok($a / $b);
}
$r1 = divide(10, 2);
echo $r1->isOk() ? 'ok' : 'err';
echo $r1->unwrap();
$r2 = divide(5, 0);
echo $r2->isOk() ? 'ok' : 'err';
echo $r2->error();
"#
        ),
        vec!["ok5errdivision by zero"]
    );
}

// ── Option/Maybe monad ───────────────────────────────────────────

#[test]
fn option_maybe_some_none() {
    assert_eq!(
        run_prints(
            r#"<?php
class Option {
    private function __construct(private bool $hasValue, private $value = null) {}
    public static function some($v): self { return new self(true, $v); }
    public static function none(): self { return new self(false); }
    public function isSome(): bool { return $this->hasValue; }
    public function get() { return $this->value; }
    public function map(callable $fn): self {
        if (!$this->hasValue) return self::none();
        return self::some($fn($this->value));
    }
    public function getOrElse($default) { return $this->hasValue ? $this->value : $default; }
}
$opt = Option::some(10)->map(fn($x) => $x * 2);
echo $opt->isSome() ? 'some' : 'none';
echo $opt->get();
$empty = Option::none()->map(fn($x) => $x * 2);
echo $empty->getOrElse(99);
"#
        ),
        vec!["some2099"]
    );
}

// ── Event sourcing stub ──────────────────────────────────────────

#[test]
fn event_sourcing_append_only_log() {
    assert_eq!(
        run_prints(
            r#"<?php
class EventStore {
    private $events = [];
    public function append(string $type, array $payload): void {
        $this->events[] = ['type' => $type, 'payload' => $payload];
    }
    public function replay(callable $reducer, $initial) {
        return array_reduce($this->events, fn($state, $e) => $reducer($state, $e), $initial);
    }
    public function count(): int { return count($this->events); }
}
$store = new EventStore();
$store->append('deposit', ['amount' => 100]);
$store->append('deposit', ['amount' => 50]);
$store->append('withdraw', ['amount' => 30]);
$balance = $store->replay(function($bal, $e) {
    if ($e['type'] === 'deposit') return $bal + $e['payload']['amount'];
    if ($e['type'] === 'withdraw') return $bal - $e['payload']['amount'];
    return $bal;
}, 0);
echo $balance;
echo $store->count();
"#
        ),
        vec!["1203"]
    );
}

// ── Active record skeleton ───────────────────────────────────────

#[test]
fn active_record_save_find() {
    compile_ok(
        r#"<?php
class Model {
    protected static $table = 'models';
    protected static $records = [];
    protected $attributes = [];
    public function __construct(array $attrs = []) { $this->attributes = $attrs; }
    public function __get($key) { return $this->attributes[$key] ?? null; }
    public function __set($key, $val) { $this->attributes[$key] = $val; }
    public function save(): void {
        $id = $this->attributes['id'] ?? count(static::$records) + 1;
        $this->attributes['id'] = $id;
        static::$records[$id] = $this;
    }
    public static function find(int $id): ?static {
        return static::$records[$id] ?? null;
    }
}
class Post extends Model {
    protected static $table = 'posts';
    protected static $records = [];
}
$p = new Post(['title' => 'Hello', 'body' => 'World']);
$p->save();
echo Post::find(1)->title;
"#,
    );
}

// ── Unit of work ─────────────────────────────────────────────────

#[test]
fn unit_of_work_track_commit() {
    assert_eq!(
        run_prints(
            r#"<?php
class UnitOfWork {
    private $new = [];
    private $dirty = [];
    private $deleted = [];
    public function registerNew(object $entity): void { $this->new[] = $entity; }
    public function registerDirty(object $entity): void { $this->dirty[] = $entity; }
    public function registerDeleted(object $entity): void { $this->deleted[] = $entity; }
    public function commit(): array {
        return [
            'inserted' => count($this->new),
            'updated' => count($this->dirty),
            'deleted' => count($this->deleted),
        ];
    }
}
$uow = new UnitOfWork();
$uow->registerNew((object)['id' => 1]);
$uow->registerNew((object)['id' => 2]);
$uow->registerDirty((object)['id' => 3]);
$uow->registerDeleted((object)['id' => 4]);
$result = $uow->commit();
echo $result['inserted'];
echo $result['updated'];
echo $result['deleted'];
"#
        ),
        vec!["211"]
    );
}

// ── Identity map ─────────────────────────────────────────────────

#[test]
fn identity_map_cache_by_id() {
    assert_eq!(
        run_prints(
            r#"<?php
class IdentityMap {
    private $map = [];
    private $loads = 0;
    public function get(string $type, int $id, callable $loader) {
        $key = "$type:$id";
        if (!isset($this->map[$key])) {
            $this->loads++;
            $this->map[$key] = $loader($id);
        }
        return $this->map[$key];
    }
    public function loads(): int { return $this->loads; }
}
$idmap = new IdentityMap();
$u1 = $idmap->get('user', 1, fn($id) => (object)['id' => $id, 'name' => 'Alice']);
$u2 = $idmap->get('user', 1, fn($id) => (object)['id' => $id, 'name' => 'SHOULDNOTRUN']);
$u3 = $idmap->get('user', 2, fn($id) => (object)['id' => $id, 'name' => 'Bob']);
echo $u1->name;
echo ($u1 === $u2) ? 'cached' : 'dup';
echo $idmap->loads();
"#
        ),
        // Bare echo emits no newline: outputs concatenate (spec PHP).
        vec!["Alicecached2"]
    );
}

// ── Two-step view ────────────────────────────────────────────────

#[test]
fn two_step_view_gather_then_render() {
    assert_eq!(
        run_prints(
            r#"<?php
class ViewModel {
    public array $data = [];
    public function set(string $k, $v): void { $this->data[$k] = $v; }
}
function gatherData(): ViewModel {
    $vm = new ViewModel();
    $vm->set('title', 'My Page');
    $vm->set('items', ['a', 'b', 'c']);
    return $vm;
}
function renderView(ViewModel $vm): string {
    $html = '<h1>' . $vm->data['title'] . '</h1>';
    $html .= '<ul>';
    foreach ($vm->data['items'] as $item) {
        $html .= '<li>' . $item . '</li>';
    }
    $html .= '</ul>';
    return $html;
}
$vm = gatherData();
echo $vm->data['title'];
echo count($vm->data['items']);
echo renderView($vm);
"#
        ),
        vec!["My Page3<h1>My Page</h1><ul><li>a</li><li>b</li><li>c</li></ul>"]
    );
}

// ── Interceptor ──────────────────────────────────────────────────

#[test]
fn interceptor_wrap_method_calls() {
    assert_eq!(
        run_prints(
            r#"<?php
class ServiceProxy {
    private $service;
    private $callLog = [];
    public function __construct(object $service) { $this->service = $service; }
    public function __call(string $method, array $args) {
        $this->callLog[] = $method;
        echo 'before:' . $method;
        $result = $this->service->$method(...$args);
        echo 'after:' . $method;
        return $result;
    }
    public function getCallLog(): array { return $this->callLog; }
}
class RealService {
    public function doWork(string $task): string { echo 'working:' . $task; return 'done'; }
}
$proxy = new ServiceProxy(new RealService());
$proxy->doWork('task1');
echo implode(',', $proxy->getCallLog());
"#
        ),
        vec!["before:doWorkworking:task1after:doWorkdoWork"]
    );
}

// ── CQRS command object ──────────────────────────────────────────

#[test]
fn cqrs_command_encapsulates_write() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Command {}
class CreateUserCommand implements Command {
    public function __construct(public readonly string $name, public readonly string $email) {}
}
class CommandBus {
    private $handlers = [];
    public function register(string $commandClass, callable $handler): void {
        $this->handlers[$commandClass] = $handler;
    }
    public function dispatch(Command $cmd): void {
        $class = get_class($cmd);
        if (!isset($this->handlers[$class])) throw new \Exception("no handler for $class");
        ($this->handlers[$class])($cmd);
    }
}
$bus = new CommandBus();
$bus->register(CreateUserCommand::class, function(CreateUserCommand $cmd) {
    echo 'created:' . $cmd->name . ':' . $cmd->email;
});
$bus->dispatch(new CreateUserCommand('Alice', 'alice@example.com'));
"#
        ),
        vec!["created:Alice:alice@example.com"]
    );
}
