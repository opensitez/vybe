//! Runtime behavior for WeakReference and WeakMap (working VM subset).

crate::php_cases! {
    weak_reference_basic_id => {
        r#"<?php
class Node { public int $id; public function __construct(int $id) { $this->id = $id; } }
$obj = new Node(42);
$ref = WeakReference::create($obj);
echo $ref->get()->id;
"#,
        ["42"]
    };

    weak_reference_alive_while_object_lives => {
        r#"<?php
class Resource { public string $name; public function __construct(string $n) { $this->name = $n; } }
$res = new Resource('db_conn');
$weak = WeakReference::create($res);
$alive = $weak->get();
echo ($alive !== null ? 'alive' : 'collected') . ':' . $weak->get()->name;
"#,
        ["alive:db_conn"]
    };

    weak_reference_cache_lookup => {
        r#"<?php
class ExpensiveObject {
    public function __construct(public readonly int $id) {}
}
$cache = [];
$obj1 = new ExpensiveObject(1);
$cache[1] = WeakReference::create($obj1);
$retrieved = $cache[1]->get();
echo $retrieved?->id ?? 'not found';
"#,
        ["1"]
    };

    weak_reference_multiple_all_alive => {
        r#"<?php
class Item { public function __construct(public string $label) {} }
$items = [new Item('a'), new Item('b'), new Item('c')];
$refs = array_map(fn($i) => WeakReference::create($i), $items);
$count = 0;
foreach ($refs as $ref) { if ($ref->get() !== null) $count++; }
echo $count;
"#,
        ["3"]
    };

    weak_map_basic_store => {
        r#"<?php
$map = new WeakMap();
$obj = new stdClass();
$obj->name = 'test';
$map[$obj] = 'associated data';
echo $map[$obj];
"#,
        ["associated data"]
    };

    weak_map_event_emitter_callback => {
        r#"<?php
class EventEmitter {
    private WeakMap $listeners;
    public function __construct() { $this->listeners = new WeakMap(); }
    public function on(object $target, callable $cb): void { $this->listeners[$target] = $cb; }
    public function emit(object $target, string $event): void {
        $cb = $this->listeners[$target] ?? null;
        if ($cb) $cb($event);
    }
}
$emitter = new EventEmitter();
$btn = new stdClass();
$emitter->on($btn, fn($e) => print("Button: $e"));
$emitter->emit($btn, 'click');
"#,
        ["Button: click"]
    };

    weak_reference_singleton_factory => {
        r#"<?php
class Singleton {
    private static ?Singleton $instance = null;
    public static function getInstance(): static {
        if (static::$instance === null) {
            static::$instance = new static();
        }
        return static::$instance;
    }
    public static function createWeak(): WeakReference {
        return WeakReference::create(static::getInstance());
    }
}
$ref = Singleton::createWeak();
echo ($ref->get() instanceof Singleton) ? 'ok' : 'fail';
"#,
        ["ok"]
    };
}
