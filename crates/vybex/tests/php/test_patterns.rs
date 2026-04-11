use super::helpers::compile_ok;

// ── Design patterns / real-world PHP ────────────────────────

#[test]
fn singleton() { compile_ok(r#"<?php
class Database {
    public static $instance = null;
    public $connected = false;
    public static function getInstance() {
        if (Database::$instance === null) {
            Database::$instance = new Database();
        }
        return Database::$instance;
    }
    public function connect() { $this->connected = true; }
}
$db = Database::getInstance();
$db->connect();
"#); }

#[test]
fn builder_pattern() { compile_ok(r#"<?php
class QueryBuilder {
    public $table = '';
    public $conditions = [];
    public $limit = 0;
    public function from($table) { $this->table = $table; return $this; }
    public function where($cond) { array_push($this->conditions, $cond); return $this; }
    public function limit($n) { $this->limit = $n; return $this; }
    public function build() {
        $sql = 'SELECT * FROM ' . $this->table;
        if (count($this->conditions) > 0) {
            $sql .= ' WHERE ' . implode(' AND ', $this->conditions);
        }
        if ($this->limit > 0) {
            $sql .= ' LIMIT ' . $this->limit;
        }
        return $sql;
    }
}
$query = (new QueryBuilder())
    ->from('users')
    ->where('age > 18')
    ->where('active = 1')
    ->limit(10)
    ->build();
echo $query;
"#); }

#[test]
fn strategy_pattern() { compile_ok(r#"<?php
interface Formatter {
    public function format($data): string;
}
class JsonFormatter implements Formatter {
    public function format($data): string { return json_encode($data); }
}
class CsvFormatter implements Formatter {
    public function format($data): string { return implode(',', $data); }
}
function export($data, $formatter) {
    return $formatter->format($data);
}
echo export(['a', 'b', 'c'], new CsvFormatter());
"#); }

#[test]
fn observer_pattern() { compile_ok(r#"<?php
class EventEmitter {
    public $listeners = [];
    public function on($event, $callback) {
        if (!isset($this->listeners[$event])) {
            $this->listeners[$event] = [];
        }
        array_push($this->listeners[$event], $callback);
    }
}
$emitter = new EventEmitter();
$emitter->on('click', fn($data) => 'clicked: ' . $data);
"#); }

#[test]
fn collection_pipeline() { compile_ok(r#"<?php
$numbers = range(1, 20);
$result = array_filter($numbers, fn($n) => $n % 2 == 0);
$result = array_map(fn($n) => $n * $n, $result);
$sum = array_reduce($result, fn($carry, $item) => $carry + $item, 0);
echo $sum;
"#); }

#[test]
fn recursive_tree() { compile_ok(r#"<?php
class TreeNode {
    public $value;
    public $left;
    public $right;
    public function __construct($val, $left = null, $right = null) {
        $this->value = $val;
        $this->left = $left;
        $this->right = $right;
    }
}
function treeSum($node) {
    if ($node === null) return 0;
    return $node->value + treeSum($node->left) + treeSum($node->right);
}
$tree = new TreeNode(1,
    new TreeNode(2, new TreeNode(4), new TreeNode(5)),
    new TreeNode(3)
);
echo treeSum($tree);
"#); }

#[test]
fn linked_list() { compile_ok(r#"<?php
class Node {
    public $value;
    public $next;
    public function __construct($val, $next = null) {
        $this->value = $val;
        $this->next = $next;
    }
}
function toArray($head) {
    $result = [];
    $current = $head;
    while ($current !== null) {
        array_push($result, $current->value);
        $current = $current->next;
    }
    return $result;
}
$list = new Node(1, new Node(2, new Node(3)));
$arr = toArray($list);
echo implode('->', $arr);
"#); }

#[test]
fn memoization() { compile_ok(r#"<?php
function memoize($fn) {
    $cache = [];
    return function() use ($fn, $cache) {
        $key = json_encode([]);
        if (!array_key_exists($key, $cache)) {
            $cache[$key] = $fn();
        }
        return $cache[$key];
    };
}
$expensive = memoize(function() { return 42; });
echo $expensive();
"#); }

#[test]
fn enum_state_machine() { compile_ok(r#"<?php
enum OrderStatus {
    case Pending;
    case Processing;
    case Shipped;
    case Delivered;
    case Cancelled;
}
class Order {
    public $status;
    public function __construct() { $this->status = OrderStatus::Pending; }
    public function process() { $this->status = OrderStatus::Processing; return $this; }
    public function ship() { $this->status = OrderStatus::Shipped; return $this; }
    public function deliver() { $this->status = OrderStatus::Delivered; return $this; }
}
$order = new Order();
$order->process()->ship()->deliver();
echo $order->status->name;
"#); }

#[test]
fn middleware_chain() { compile_ok(r#"<?php
function pipeline($value, $fns) {
    return array_reduce($fns, fn($carry, $fn) => $fn($carry), $value);
}
$result = pipeline('hello world', [
    fn($s) => strtoupper($s),
    fn($s) => trim($s),
    fn($s) => str_replace(' ', '-', $s),
]);
echo $result;
"#); }

#[test]
fn generic_collection() { compile_ok(r#"<?php
class Collection {
    public $items = [];
    public function add($item) { array_push($this->items, $item); return $this; }
    public function map($fn) {
        $new = new Collection();
        $new->items = array_map($fn, $this->items);
        return $new;
    }
    public function filter($fn) {
        $new = new Collection();
        $new->items = array_filter($this->items, $fn);
        return $new;
    }
    public function count() { return count($this->items); }
    public function first() { return $this->items[0] ?? null; }
    public function toArray() { return $this->items; }
}
$c = new Collection();
$c->add(1)->add(2)->add(3)->add(4)->add(5);
$evens = $c->filter(fn($n) => $n % 2 == 0);
echo $evens->count();
"#); }
