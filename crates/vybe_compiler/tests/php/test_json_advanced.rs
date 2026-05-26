use super::helpers::run_prints;

// ── json_encode basics ────────────────────────────────────────

#[test] fn json_encode_array() {
    assert_eq!(run_prints(r#"<?php echo json_encode([1,2,3]); "#), vec!["[1,2,3]"]);
}
#[test] fn json_encode_assoc() {
    assert_eq!(run_prints(r#"<?php echo json_encode(['name'=>'Alice','age'=>30]); "#), vec![r#"{"name":"Alice","age":30}"#]);
}
#[test] fn json_encode_unicode_escaped() {
    assert_eq!(run_prints(r#"<?php echo json_encode(['emoji'=>'<>&']); "#), vec![r#"{"emoji":"<>&"}"#]);
}
#[test] fn json_encode_unicode_unescaped() {
    assert_eq!(run_prints(r#"<?php echo json_encode('héllo', JSON_UNESCAPED_UNICODE); "#), vec![r#""héllo""#]);
}
#[test] fn json_encode_pretty_print() {
    assert_eq!(run_prints(r#"<?php
$s = json_encode(['a'=>1,'b'=>2], JSON_PRETTY_PRINT);
echo str_contains($s, "\n") ? 'multiline' : 'single';
"#), vec!["multiline"]);
}
#[test] fn json_encode_nested() {
    assert_eq!(run_prints(r#"<?php echo json_encode(['user'=>['name'=>'Bob','scores'=>[10,20,30]]]); "#), vec![r#"{"user":{"name":"Bob","scores":[10,20,30]}}"#]);
}

// ── json_decode ───────────────────────────────────────────────

#[test] fn json_decode_to_object() {
    assert_eq!(run_prints(r#"<?php
$o = json_decode('{"name":"Alice","age":30}');
echo $o->name . ':' . $o->age;
"#), vec!["Alice:30"]);
}
#[test] fn json_decode_to_assoc_array() {
    assert_eq!(run_prints(r#"<?php
$a = json_decode('{"x":1,"y":2}', true);
echo $a['x'] + $a['y'];
"#), vec!["3"]);
}
#[test] fn json_decode_array() {
    assert_eq!(run_prints(r#"<?php
$a = json_decode('[1,2,3,4,5]', true);
echo array_sum($a);
"#), vec!["15"]);
}
#[test] fn json_decode_nested() {
    assert_eq!(run_prints(r#"<?php
$d = json_decode('{"user":{"name":"Bob","tags":["php","rust"]}}', true);
echo $d['user']['name'] . ':' . implode(',', $d['user']['tags']);
"#), vec!["Bob:php,rust"]);
}
#[test] fn json_decode_returns_null_on_error() {
    assert_eq!(run_prints(r#"<?php $r = json_decode('invalid json'); echo $r === null ? 'null' : 'not'; "#), vec!["null"]);
}

// ── json_last_error ───────────────────────────────────────────

#[test] fn json_last_error_no_error() {
    assert_eq!(run_prints(r#"<?php json_encode([1,2,3]); echo json_last_error(); "#), vec!["0"]);
}
#[test] fn json_last_error_on_invalid() {
    assert_eq!(run_prints(r#"<?php json_decode('{invalid}'); echo json_last_error() !== JSON_ERROR_NONE ? 'error' : 'ok'; "#), vec!["error"]);
}
#[test] fn json_last_error_msg() {
    assert_eq!(run_prints(r#"<?php json_decode('{bad}'); echo strlen(json_last_error_msg()) > 0 ? 'has_msg' : 'empty'; "#), vec!["has_msg"]);
}

// ── JSON encode flags ─────────────────────────────────────────

#[test] fn json_encode_unescaped_slashes() {
    assert_eq!(run_prints(r#"<?php echo json_encode('a/b/c', JSON_UNESCAPED_SLASHES); "#), vec![r#""a/b/c""#]);
}
#[test] fn json_encode_numeric_check() {
    assert_eq!(run_prints(r#"<?php echo json_encode(['n'=>'1.5'], JSON_NUMERIC_CHECK); "#), vec![r#"{"n":1.5}"#]);
}
#[test] fn json_encode_force_object() {
    assert_eq!(run_prints(r#"<?php echo json_encode([1,2,3], JSON_FORCE_OBJECT); "#), vec![r#"{"0":1,"1":2,"2":3}"#]);
}

// ── Round-trip ────────────────────────────────────────────────

#[test] fn json_roundtrip_types() {
    assert_eq!(run_prints(r#"<?php
$data = ['int'=>42,'float'=>3.14,'bool'=>true,'null'=>null,'str'=>'hello'];
$decoded = json_decode(json_encode($data), true);
echo $decoded['int'] . ',' . $decoded['float'] . ',' . ($decoded['bool'] ? 't' : 'f') . ',' . ($decoded['null'] === null ? 'n' : 'x') . ',' . $decoded['str'];
"#), vec!["42,3.14,t,n,hello"]);
}
#[test] fn json_roundtrip_nested_array() {
    assert_eq!(run_prints(r#"<?php
$orig = [['a',1],['b',2],['c',3]];
$back = json_decode(json_encode($orig), true);
echo $back[1][0] . ':' . $back[1][1];
"#), vec!["b:2"]);
}

// ── json_validate (PHP 8.3) ───────────────────────────────────

#[test] fn json_validate_valid_json() {
    assert_eq!(run_prints(r#"<?php
if (function_exists('json_validate')) {
    echo json_validate('{"a":1}') ? 'valid' : 'invalid';
} else {
    json_decode('{"a":1}');
    echo json_last_error() === JSON_ERROR_NONE ? 'valid' : 'invalid';
}
"#), vec!["valid"]);
}
#[test] fn json_validate_invalid_json() {
    assert_eq!(run_prints(r#"<?php
if (function_exists('json_validate')) {
    echo json_validate('{bad}') ? 'valid' : 'invalid';
} else {
    json_decode('{bad}');
    echo json_last_error() !== JSON_ERROR_NONE ? 'invalid' : 'valid';
}
"#), vec!["invalid"]);
}
