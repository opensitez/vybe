use super::helpers::compile_ok;

// ── WeakReference (PHP 8.0+) ──────────────────────────────────

#[test] fn weak_reference_basic() {
    compile_ok(r#"<?php
class Node { public int $id; public function __construct(int $id) { $this->id = $id; } }
$obj = new Node(42);
$ref = WeakReference::create($obj);
echo $ref->get()->id;
"#);
}

#[test] fn weak_reference_get_while_alive() {
    compile_ok(r#"<?php
class Resource { public string $name; public function __construct(string $n) { $this->name = $n; } }
$res = new Resource('db_conn');
$weak = WeakReference::create($res);
$alive = $weak->get();
echo $alive !== null ? 'alive' : 'collected';
echo ':' . $weak->get()->name;
"#);
}

#[test] fn weak_reference_null_after_collect() {
    compile_ok(r#"<?php
class Temp {}
$weak = null;
{
    $obj = new Temp();
    $weak = WeakReference::create($obj);
    unset($obj);
}
// After unset, get() may return null
$result = $weak->get();
echo $result === null ? 'null' : 'alive';
"#);
}

#[test] fn weak_reference_in_cache() {
    compile_ok(r#"<?php
class ExpensiveObject {
    public function __construct(public readonly int $id) {}
}
$cache = [];
$obj1 = new ExpensiveObject(1);
$cache[1] = WeakReference::create($obj1);
$retrieved = $cache[1]->get();
echo $retrieved?->id ?? 'not found';
"#);
}

#[test] fn weak_reference_multiple() {
    compile_ok(r#"<?php
class Item { public function __construct(public string $label) {} }
$items = [new Item('a'), new Item('b'), new Item('c')];
$refs = array_map(fn($i) => WeakReference::create($i), $items);
$count = 0;
foreach ($refs as $ref) { if ($ref->get() !== null) $count++; }
echo $count;
"#);
}

// ── WeakMap (PHP 8.0+) ───────────────────────────────────────

#[test] fn weak_map_basic() {
    compile_ok(r#"<?php
$map = new WeakMap();
$obj = new stdClass();
$obj->name = 'test';
$map[$obj] = 'associated data';
echo $map[$obj];
"#);
}

#[test] fn weak_map_count() {
    compile_ok(r#"<?php
$map = new WeakMap();
$a = new stdClass();
$b = new stdClass();
$c = new stdClass();
$map[$a] = 1;
$map[$b] = 2;
$map[$c] = 3;
echo count($map);
"#);
}

#[test] fn weak_map_isset_unset() {
    compile_ok(r#"<?php
$map = new WeakMap();
$obj = new stdClass();
echo isset($map[$obj]) ? 'set' : 'not set';
$map[$obj] = 'value';
echo isset($map[$obj]) ? 'set' : 'not set';
unset($map[$obj]);
echo isset($map[$obj]) ? 'set' : 'not set';
"#);
}

#[test] fn weak_map_metadata() {
    compile_ok(r#"<?php
class Connection {
    public function __construct(public readonly string $dsn) {}
}
$map = new WeakMap();
$conn1 = new Connection('sqlite::memory:');
$conn2 = new Connection('mysql://localhost');
$map[$conn1] = ['created' => time(), 'queries' => 0];
$map[$conn2] = ['created' => time(), 'queries' => 0];
$map[$conn1]['queries']++;
echo $map[$conn1]['queries'];
echo count($map);
"#);
}

#[test] fn weak_map_object_cache() {
    compile_ok(r#"<?php
class User { public function __construct(public int $id, public string $name) {} }
$computed = new WeakMap();
function getDisplayName(User $user, WeakMap $cache): string {
    if (!isset($cache[$user])) {
        $cache[$user] = strtoupper($user->name) . '#' . $user->id;
    }
    return $cache[$user];
}
$u = new User(1, 'alice');
echo getDisplayName($u, $computed);
echo getDisplayName($u, $computed);
"#);
}

#[test] fn weak_map_iterate() {
    compile_ok(r#"<?php
$map = new WeakMap();
$objs = [];
for ($i = 0; $i < 3; $i++) {
    $obj = new stdClass();
    $obj->n = $i;
    $objs[] = $obj;
    $map[$obj] = "value_$i";
}
$vals = [];
foreach ($map as $k => $v) { $vals[] = $v; }
sort($vals);
echo implode(',', $vals);
"#);
}

#[test] fn weak_map_as_event_listener_table() {
    compile_ok(r#"<?php
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
"#);
}

#[test] fn weak_map_type_checking() {
    compile_ok(r#"<?php
$map = new WeakMap();
echo ($map instanceof WeakMap) ? 'is WeakMap' : 'not WeakMap';
"#);
}

#[test] fn weak_reference_create_static() {
    compile_ok(r#"<?php
class Singleton {
    private static ?Singleton $instance = null;
    private static ?WeakReference $weakRef = null;
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
"#);
}
