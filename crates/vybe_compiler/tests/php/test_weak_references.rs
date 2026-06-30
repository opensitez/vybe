use super::helpers::compile_ok;

// Compile-time coverage for WeakReference / WeakMap APIs. Runtime behavior
// is asserted in `test_weak_references_runtime.rs`.

#[test]
fn weak_reference_null_after_collect() {
    compile_ok(
        r#"<?php
class Temp {}
$weak = null;
{
    $obj = new Temp();
    $weak = WeakReference::create($obj);
    unset($obj);
}
$result = $weak->get();
echo $result === null ? 'null' : 'alive';
"#,
    );
}

#[test]
fn weak_map_count() {
    compile_ok(
        r#"<?php
$map = new WeakMap();
$a = new stdClass();
$b = new stdClass();
$c = new stdClass();
$map[$a] = 1;
$map[$b] = 2;
$map[$c] = 3;
echo count($map);
"#,
    );
}

#[test]
fn weak_map_isset_unset() {
    compile_ok(
        r#"<?php
$map = new WeakMap();
$obj = new stdClass();
echo isset($map[$obj]) ? 'set' : 'not set';
$map[$obj] = 'value';
echo isset($map[$obj]) ? 'set' : 'not set';
unset($map[$obj]);
echo isset($map[$obj]) ? 'set' : 'not set';
"#,
    );
}

#[test]
fn weak_map_metadata() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

#[test]
fn weak_map_object_cache() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

#[test]
fn weak_map_iterate() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

#[test]
fn weak_map_type_checking() {
    compile_ok(
        r#"<?php
$map = new WeakMap();
echo ($map instanceof WeakMap) ? 'is WeakMap' : 'not WeakMap';
"#,
    );
}
