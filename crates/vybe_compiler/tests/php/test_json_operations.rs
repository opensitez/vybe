use super::helpers::{compile_ok, run_prints};

// ── json_encode basic indexed array ──────────────────────────────
#[test]
fn json_encode_basic_array() {
    compile_ok(
        r#"<?php
$arr = [1, 2, 3];
echo json_encode($arr);
"#,
    );
}

// ── json_encode associative array ────────────────────────────────
#[test]
fn json_encode_assoc_array() {
    compile_ok(
        r#"<?php
$data = ['name' => 'Alice', 'age' => 30, 'active' => true];
echo json_encode($data);
"#,
    );
}

// ── json_encode with JSON_PRETTY_PRINT flag ───────────────────────
#[test]
fn json_encode_pretty_print() {
    compile_ok(
        r#"<?php
$data = ['key' => 'value', 'num' => 42];
echo json_encode($data, JSON_PRETTY_PRINT);
"#,
    );
}

// ── json_encode with JSON_UNESCAPED_SLASHES flag ──────────────────
#[test]
fn json_encode_unescaped_slashes() {
    compile_ok(
        r#"<?php
$data = ['url' => 'https://example.com/path/to/resource'];
echo json_encode($data, JSON_UNESCAPED_SLASHES);
"#,
    );
}

// ── json_encode with JSON_UNESCAPED_UNICODE flag ──────────────────
#[test]
fn json_encode_unescaped_unicode() {
    compile_ok(
        r#"<?php
$data = ['greeting' => 'Bonjour'];
echo json_encode($data, JSON_UNESCAPED_UNICODE);
"#,
    );
}

// ── json_encode with JSON_FORCE_OBJECT on indexed array ──────────
#[test]
fn json_encode_force_object() {
    compile_ok(
        r#"<?php
$arr = ['apple', 'banana', 'cherry'];
echo json_encode($arr, JSON_FORCE_OBJECT);
"#,
    );
}

// ── json_encode with JSON_NUMERIC_CHECK (string numbers) ─────────
#[test]
fn json_encode_numeric_check() {
    compile_ok(
        r#"<?php
$data = ['count' => '42', 'price' => '9.99'];
echo json_encode($data, JSON_NUMERIC_CHECK);
"#,
    );
}

// ── json_encode null value ────────────────────────────────────────
#[test]
fn json_encode_null() {
    compile_ok(
        r#"<?php
echo json_encode(null);
"#,
    );
}

// ── json_encode boolean values ────────────────────────────────────
#[test]
fn json_encode_booleans() {
    compile_ok(
        r#"<?php
echo json_encode(true);
echo json_encode(false);
echo json_encode(['flag' => true, 'other' => false]);
"#,
    );
}

// ── json_encode nested objects/arrays ────────────────────────────
#[test]
fn json_encode_nested() {
    compile_ok(
        r#"<?php
$data = [
    'user' => [
        'name' => 'Bob',
        'roles' => ['admin', 'editor'],
        'meta' => ['score' => 99]
    ]
];
echo json_encode($data);
"#,
    );
}

// ── json_decode to associative array (true flag) ──────────────────
#[test]
fn json_decode_to_array() {
    compile_ok(
        r#"<?php
$json = '{"name":"Alice","age":30}';
$arr = json_decode($json, true);
echo $arr['name'];
echo $arr['age'];
"#,
    );
}

// ── json_decode to stdClass object ───────────────────────────────
#[test]
fn json_decode_to_stdclass() {
    compile_ok(
        r#"<?php
$json = '{"title":"Hello","count":5}';
$obj = json_decode($json);
echo $obj->title;
echo $obj->count;
"#,
    );
}

// ── json_decode nested structure access ──────────────────────────
#[test]
fn json_decode_nested_access() {
    compile_ok(
        r#"<?php
$json = '{"user":{"name":"Carol","scores":[10,20,30]}}';
$data = json_decode($json, true);
echo $data['user']['name'];
echo $data['user']['scores'][1];
"#,
    );
}

// ── json_decode returns null on invalid input ─────────────────────
#[test]
fn json_decode_invalid_returns_null() {
    compile_ok(
        r#"<?php
$result = json_decode('not valid json', true);
var_dump($result);
"#,
    );
}

// ── json_last_error after failed decode ──────────────────────────
#[test]
fn json_last_error_after_failure() {
    compile_ok(
        r#"<?php
json_decode('{bad json}');
$err = json_last_error();
echo $err !== JSON_ERROR_NONE ? 'error' : 'ok';
"#,
    );
}

// ── json_last_error_msg ───────────────────────────────────────────
#[test]
fn json_last_error_msg() {
    compile_ok(
        r#"<?php
json_decode('{bad json}');
$msg = json_last_error_msg();
echo is_string($msg) ? 'string' : 'not-string';
"#,
    );
}

// ── json_encode integer (numeric) keys ───────────────────────────
#[test]
fn json_encode_integer_keys() {
    compile_ok(
        r#"<?php
$map = [0 => 'zero', 1 => 'one', 2 => 'two'];
echo json_encode($map);
"#,
    );
}

// ── json_encode deeply nested structure ──────────────────────────
#[test]
fn json_encode_deeply_nested() {
    std::thread::Builder::new()
        .stack_size(8 * 1024 * 1024)
        .spawn(|| {
            compile_ok(
                r#"<?php
$deep = ['a' => ['b' => ['c' => ['d' => 'leaf']]]];
echo json_encode($deep);
"#,
            )
        })
        .unwrap()
        .join()
        .unwrap();
}

// ── json_decode with depth limit (exceeding depth returns null) ───
#[test]
fn json_decode_depth_limit() {
    compile_ok(
        r#"<?php
$json = '{"a":{"b":{"c":{"d":"deep"}}}}';
$shallow = json_decode($json, true, 2);
$deep    = json_decode($json, true, 512);
echo is_null($shallow) ? 'null' : 'ok';
echo is_array($deep)   ? 'array' : 'not-array';
"#,
    );
}

// ── json_encode empty array vs empty object ───────────────────────
#[test]
fn json_encode_empty_array_vs_object() {
    compile_ok(
        r#"<?php
$arr = [];
$obj = new stdClass();
echo json_encode($arr);
echo json_encode($obj);
"#,
    );
}
