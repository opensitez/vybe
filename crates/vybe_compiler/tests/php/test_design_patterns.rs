use super::helpers::run_prints;

// ── Builder pattern ───────────────────────────────────────────

#[test] fn builder_fluent_interface() {
    assert_eq!(run_prints(r#"<?php
class QueryBuilder {
    private string $table = '';
    private array $conditions = [];
    private ?int $limit = null;
    public function from(string $t): static { $this->table = $t; return $this; }
    public function where(string $c): static { $this->conditions[] = $c; return $this; }
    public function limit(int $n): static { $this->limit = $n; return $this; }
    public function build(): string {
        $sql = "SELECT * FROM $this->table";
        if ($this->conditions) $sql .= ' WHERE ' . implode(' AND ', $this->conditions);
        if ($this->limit) $sql .= " LIMIT $this->limit";
        return $sql;
    }
}
echo (new QueryBuilder)->from('users')->where('age>18')->where('active=1')->limit(10)->build();
"#), vec!["SELECT * FROM users WHERE age>18 AND active=1 LIMIT 10"]);
}

// ── Observer pattern ──────────────────────────────────────────

#[test] fn observer_pattern() {
    assert_eq!(run_prints(r#"<?php
interface Observer { public function update(string $event, mixed $data): void; }
class EventEmitter {
    private array $observers = [];
    public function subscribe(string $event, Observer $o): void { $this->observers[$event][] = $o; }
    public function emit(string $event, mixed $data = null): void {
        foreach ($this->observers[$event] ?? [] as $o) $o->update($event, $data);
    }
}
class Logger implements Observer {
    public array $log = [];
    public function update(string $e, mixed $d): void { $this->log[] = "$e:$d"; }
}
$emitter = new EventEmitter;
$logger = new Logger;
$emitter->subscribe('login', $logger);
$emitter->emit('login', 'Alice');
$emitter->emit('login', 'Bob');
echo implode(',', $logger->log);
"#), vec!["login:Alice,login:Bob"]);
}

// ── Strategy pattern ──────────────────────────────────────────

#[test] fn strategy_pattern() {
    assert_eq!(run_prints(r#"<?php
interface SortStrategy { public function sort(array &$data): void; }
class BubbleSort implements SortStrategy {
    public function sort(array &$data): void { sort($data); }
}
class Sorter {
    public function __construct(private SortStrategy $strategy) {}
    public function sort(array $data): array { $this->strategy->sort($data); return $data; }
}
echo implode(',', (new Sorter(new BubbleSort))->sort([3,1,2]));
"#), vec!["1,2,3"]);
}

// ── Decorator pattern ─────────────────────────────────────────

#[test] fn decorator_pattern() {
    assert_eq!(run_prints(r#"<?php
interface TextProcessor { public function process(string $t): string; }
class BaseProcessor implements TextProcessor { public function process(string $t): string { return $t; } }
class TrimDecorator implements TextProcessor {
    public function __construct(private TextProcessor $inner) {}
    public function process(string $t): string { return trim($this->inner->process($t)); }
}
class UpperDecorator implements TextProcessor {
    public function __construct(private TextProcessor $inner) {}
    public function process(string $t): string { return strtoupper($this->inner->process($t)); }
}
$proc = new UpperDecorator(new TrimDecorator(new BaseProcessor));
echo $proc->process('  hello world  ');
"#), vec!["HELLO WORLD"]);
}

// ── Factory method pattern ────────────────────────────────────

#[test] fn factory_method() {
    assert_eq!(run_prints(r#"<?php
abstract class Notification {
    abstract public function send(string $msg): string;
    public static function create(string $type): self {
        return match($type) {
            'email' => new EmailNotification,
            'sms'   => new SmsNotification,
            default => throw new InvalidArgumentException("Unknown: $type"),
        };
    }
}
class EmailNotification extends Notification { public function send(string $m): string { return "email:$m"; } }
class SmsNotification extends Notification { public function send(string $m): string { return "sms:$m"; } }
echo Notification::create('email')->send('hello') . ',' . Notification::create('sms')->send('hi');
"#), vec!["email:hello,sms:hi"]);
}

// ── Proxy pattern ─────────────────────────────────────────────

#[test] fn proxy_pattern_lazy_load() {
    assert_eq!(run_prints(r#"<?php
interface DataStore { public function get(string $key): mixed; }
class RealStore implements DataStore {
    private array $data = ['name' => 'Alice'];
    public function get(string $key): mixed { echo 'fetched,'; return $this->data[$key] ?? null; }
}
class CachedProxy implements DataStore {
    private array $cache = [];
    private DataStore $store;
    public function __construct() { $this->store = new RealStore; }
    public function get(string $key): mixed {
        if (!isset($this->cache[$key])) $this->cache[$key] = $this->store->get($key);
        return $this->cache[$key];
    }
}
$proxy = new CachedProxy;
echo $proxy->get('name') . ',';
echo $proxy->get('name');
"#), vec!["fetched,Alice,Alice"]);
}

// ── Value object pattern ──────────────────────────────────────

#[test] fn value_object_immutability() {
    assert_eq!(run_prints(r#"<?php
final class Money {
    public function __construct(private readonly int $amount, private readonly string $currency) {}
    public function add(Money $other): self {
        if ($this->currency !== $other->currency) throw new \InvalidArgumentException('Currency mismatch');
        return new self($this->amount + $other->amount, $this->currency);
    }
    public function __toString(): string { return $this->amount . ' ' . $this->currency; }
}
$a = new Money(100, 'USD');
$b = new Money(50, 'USD');
echo $a->add($b);
"#), vec!["150 USD"]);
}

// ── Command pattern ───────────────────────────────────────────

#[test] fn command_pattern_undo() {
    assert_eq!(run_prints(r#"<?php
interface Command { public function execute(): void; public function undo(): void; }
class Stack {
    private array $stack = [];
    private array $history = [];
    public function execute(Command $cmd): void { $cmd->execute(); $this->history[] = $cmd; }
    public function undo(): void { if ($c = array_pop($this->history)) $c->undo(); }
}
class PushCommand implements Command {
    public function __construct(private array &$list, private int $val) {}
    public function execute(): void { $this->list[] = $this->val; }
    public function undo(): void { array_pop($this->list); }
}
$list = [];
$stack = new Stack;
$stack->execute(new PushCommand($list, 1));
$stack->execute(new PushCommand($list, 2));
$stack->execute(new PushCommand($list, 3));
echo implode(',', $list) . ',';
$stack->undo();
echo implode(',', $list);
"#), vec!["1,2,3,1,2"]);
}

// ── Repository pattern ────────────────────────────────────────

#[test] fn repository_pattern() {
    assert_eq!(run_prints(r#"<?php
class User { public function __construct(public readonly int $id, public readonly string $name) {} }
class UserRepository {
    private array $store = [];
    public function save(User $u): void { $this->store[$u->id] = $u; }
    public function find(int $id): ?User { return $this->store[$id] ?? null; }
    public function findAll(): array { return array_values($this->store); }
}
$repo = new UserRepository;
$repo->save(new User(1, 'Alice'));
$repo->save(new User(2, 'Bob'));
echo $repo->find(2)?->name . ':' . count($repo->findAll());
"#), vec!["Bob:2"]);
}
