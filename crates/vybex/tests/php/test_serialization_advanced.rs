use super::helpers::{compile_ok, run_prints};

// ── serialize / unserialize primitives ───────────────────────

#[test] fn serialize_int() {
    compile_ok(r#"<?php
$s = serialize(42);
$v = unserialize($s);
echo $v;
"#);
}

#[test] fn serialize_float() {
    compile_ok(r#"<?php
$s = serialize(3.14);
$v = unserialize($s);
echo $v;
"#);
}

#[test] fn serialize_string() {
    compile_ok(r#"<?php
$s = serialize("hello world");
$v = unserialize($s);
echo $v;
"#);
}

#[test] fn serialize_bool() {
    compile_ok(r#"<?php
$t = serialize(true);
$f = serialize(false);
echo unserialize($t) ? 'true' : 'false';
echo unserialize($f) ? 'true' : 'false';
"#);
}

#[test] fn serialize_null() {
    compile_ok(r#"<?php
$s = serialize(null);
$v = unserialize($s);
var_dump($v);
"#);
}

// ── serialize arrays ──────────────────────────────────────────

#[test] fn serialize_indexed_array() {
    compile_ok(r#"<?php
$arr = [1, 2, 3, 4, 5];
$s = serialize($arr);
$v = unserialize($s);
echo implode(',', $v);
"#);
}

#[test] fn serialize_assoc_array() {
    compile_ok(r#"<?php
$data = ['name' => 'Alice', 'age' => 30, 'active' => true];
$s = serialize($data);
$v = unserialize($s);
echo $v['name'] . ':' . $v['age'];
"#);
}

#[test] fn serialize_nested_array() {
    compile_ok(r#"<?php
$nested = ['users' => [['id' => 1, 'name' => 'Alice'], ['id' => 2, 'name' => 'Bob']]];
$s = serialize($nested);
$v = unserialize($s);
echo $v['users'][0]['name'] . ',' . $v['users'][1]['name'];
"#);
}

// ── serialize objects ─────────────────────────────────────────

#[test] fn serialize_object_basic() {
    compile_ok(r#"<?php
class Point { public function __construct(public int $x, public int $y) {} }
$p = new Point(3, 7);
$s = serialize($p);
$p2 = unserialize($s);
echo $p2->x . ',' . $p2->y;
"#);
}

#[test] fn serialize_object_with_private() {
    compile_ok(r#"<?php
class Secret {
    private string $password;
    public function __construct(string $pw) { $this->password = $pw; }
    public function getPassword(): string { return $this->password; }
}
$s = new Secret('abc123');
$ser = serialize($s);
$s2 = unserialize($ser);
echo $s2->getPassword();
"#);
}

// ── __sleep / __wakeup ────────────────────────────────────────

#[test] fn sleep_wakeup_basic() {
    assert_eq!(run_prints(r#"<?php
class Cached {
    public string $data;
    private bool $loaded = false;
    public function __construct(string $data) { $this->data = $data; $this->loaded = true; }
    public function __sleep(): array { return ['data']; }
    public function __wakeup(): void { $this->loaded = true; }
    public function isLoaded(): bool { return $this->loaded; }
}
$c = new Cached("important");
$s = serialize($c);
$c2 = unserialize($s);
echo $c2->data . ':' . ($c2->isLoaded() ? 'ready' : 'not');
"#), &["important:ready"]);
}

#[test] fn sleep_selective_properties() {
    assert_eq!(run_prints(r#"<?php
class DBConnection {
    private string $host;
    private int $port;
    private mixed $connection = null;
    public function __construct(string $host, int $port) {
        $this->host = $host; $this->port = $port;
    }
    public function __sleep(): array { return ['host', 'port']; }
    public function __wakeup(): void { $this->connection = null; }
    public function getHost(): string { return $this->host; }
}
$db = new DBConnection('localhost', 5432);
$s = serialize($db);
$db2 = unserialize($s);
echo $db2->getHost();
"#), &["localhost"]);
}

// ── __serialize / __unserialize (PHP 7.4+) ────────────────────

#[test] fn serialize_unserialize_magic() {
    assert_eq!(run_prints(r#"<?php
class DateRange {
    public function __construct(
        private \DateTimeImmutable $start,
        private \DateTimeImmutable $end
    ) {}
    public function __serialize(): array {
        return ['start' => $this->start->format('Y-m-d'), 'end' => $this->end->format('Y-m-d')];
    }
    public function __unserialize(array $data): void {
        $this->start = new \DateTimeImmutable($data['start']);
        $this->end   = new \DateTimeImmutable($data['end']);
    }
    public function days(): int {
        return (int)$this->start->diff($this->end)->days;
    }
}
$range = new DateRange(new \DateTimeImmutable('2024-01-01'), new \DateTimeImmutable('2024-01-31'));
$s = serialize($range);
$r2 = unserialize($s);
echo $r2->days();
"#), &["30"]);
}

#[test] fn serialize_custom_representation() {
    assert_eq!(run_prints(r#"<?php
class Vector {
    public function __construct(public float $x, public float $y, public float $z) {}
    public function __serialize(): array { return [$this->x, $this->y, $this->z]; }
    public function __unserialize(array $data): void {
        [$this->x, $this->y, $this->z] = $data;
    }
    public function length(): float { return sqrt($this->x**2 + $this->y**2 + $this->z**2); }
}
$v = new Vector(1.0, 0.0, 0.0);
$s = serialize($v);
$v2 = unserialize($s);
echo round($v2->length(), 4);
"#), &["1"]);
}

// ── JSON encode options ───────────────────────────────────────

#[test] fn json_encode_pretty_print() {
    compile_ok(r#"<?php
$data = ['name' => 'Alice', 'age' => 30];
$json = json_encode($data, JSON_PRETTY_PRINT);
echo strlen($json) > strlen(json_encode($data)) ? 'pretty' : 'compact';
"#);
}

#[test] fn json_encode_unescaped_unicode() {
    compile_ok(r#"<?php
$data = ['greeting' => 'Héllo'];
$escaped = json_encode($data);
$unescaped = json_encode($data, JSON_UNESCAPED_UNICODE);
echo str_contains($unescaped, 'Héllo') ? 'unescaped' : 'escaped';
"#);
}

#[test] fn json_encode_unescaped_slashes() {
    compile_ok(r#"<?php
$data = ['url' => 'https://example.com/path'];
$default = json_encode($data);
$noslash = json_encode($data, JSON_UNESCAPED_SLASHES);
echo str_contains($noslash, 'https://example.com/path') ? 'ok' : 'fail';
"#);
}

#[test] fn json_encode_throw_on_error() {
    compile_ok(r#"<?php
$valid = ['key' => 'value'];
try {
    $json = json_encode($valid, JSON_THROW_ON_ERROR);
    echo 'ok';
} catch (\JsonException $e) {
    echo 'error: ' . $e->getMessage();
}
"#);
}

#[test] fn json_decode_associative() {
    compile_ok(r#"<?php
$json = '{"name":"Bob","scores":[1,2,3]}';
$obj = json_decode($json);
$arr = json_decode($json, true);
echo $obj->name . ':' . implode(',', $arr['scores']);
"#);
}

#[test] fn json_decode_depth() {
    compile_ok(r#"<?php
$deep = '{"a":{"b":{"c":{"d":1}}}}';
$v = json_decode($deep, true, 512);
echo $v['a']['b']['c']['d'];
"#);
}

#[test] fn json_decode_throw_on_error() {
    compile_ok(r#"<?php
try {
    $v = json_decode('invalid json', true, 512, JSON_THROW_ON_ERROR);
} catch (\JsonException $e) {
    echo 'caught: ' . $e->getMessage();
}
"#);
}

// ── json_validate (PHP 8.3) ───────────────────────────────────

#[test] fn json_validate_valid() {
    compile_ok(r#"<?php
echo json_validate('{"key": "value"}') ? 'valid' : 'invalid';
echo json_validate('[1, 2, 3]') ? 'valid' : 'invalid';
"#);
}

#[test] fn json_validate_invalid() {
    compile_ok(r#"<?php
echo json_validate('not json') ? 'valid' : 'invalid';
echo json_validate('{bad: "json"}') ? 'valid' : 'invalid';
"#);
}

// ── Round-trip integrity ──────────────────────────────────────

#[test] fn serialize_roundtrip_complex() {
    compile_ok(r#"<?php
class Tree {
    public array $children = [];
    public function __construct(public string $label) {}
    public function addChild(Tree $child): void { $this->children[] = $child; }
}
$root = new Tree('root');
$root->addChild(new Tree('child1'));
$root->addChild(new Tree('child2'));
$root->children[0]->addChild(new Tree('grandchild'));
$s = serialize($root);
$r = unserialize($s);
echo $r->label . ':' . count($r->children) . ':' . $r->children[0]->label;
"#);
}

#[test] fn json_roundtrip_array_of_objects() {
    compile_ok(r#"<?php
$users = [
    ['id' => 1, 'name' => 'Alice', 'active' => true],
    ['id' => 2, 'name' => 'Bob',   'active' => false],
];
$json = json_encode($users);
$decoded = json_decode($json, true);
echo $decoded[0]['name'] . ',' . $decoded[1]['name'];
"#);
}
