use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: JSON Encode & Decode Flags — JSON_PRETTY_PRINT, JSON_FORCE_OBJECT, JSON_UNESCAPED_SLASHES, JSON_UNESCAPED_UNICODE, JSON_PRESERVE_ZERO_FRACTION
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_json_encode_force_object_flag() {
    let out = run_prints(
        r#"<?php
$indexed = ["a", "b", "c"];
$encoded = json_encode($indexed, JSON_FORCE_OBJECT);
echo $encoded;
"#,
    );
    assert_eq!(out, vec![r#"{"0":"a","1":"b","2":"c"}"#]);
}

#[test]
fn test_php_json_encode_preserve_zero_fraction() {
    let out = run_prints(
        r#"<?php
$data = ["val" => 10.0];
$encoded = json_encode($data, JSON_PRESERVE_ZERO_FRACTION);
echo $encoded;
"#,
    );
    assert_eq!(out, vec![r#"{"val":10.0}"#]);
}

#[test]
fn test_php_json_decode_bigint_as_string() {
    let out = run_prints(
        r#"<?php
$json = '{"big_int": 9223372036854775807}';
$data = json_decode($json, true, flags: JSON_BIGINT_AS_STRING);
echo gettype($data["big_int"]) . "=" . $data["big_int"];
"#,
    );
    assert_eq!(out, vec!["string=9223372036854775807"]);
}

#[test]
fn test_php_json_encode_empty_object_and_array() {
    let out = run_prints(
        r#"<?php
$data = [
    "empty_arr" => [],
    "empty_obj" => new stdClass()
];
echo json_encode($data);
"#,
    );
    assert_eq!(out, vec![r#"{"empty_arr":[],"empty_obj":{}}"#]);
}

#[test]
fn test_php_json_encode_numeric_keys_preservation() {
    compile_ok(
        r#"<?php
$assoc = [1 => "one", 2 => "two"];
echo json_encode($assoc);
"#,
    );
}

#[test]
fn test_php_json_decode_into_existing_class() {
    compile_ok(
        r#"<?php
class UserDto {
    public string $name;
    public int $age;
}

$json = '{"name":"Alice","age":30}';
$dto = json_decode($json);
echo "$dto->name is $dto->age";
"#,
    );
}

#[test]
fn test_php_json_encode_hex_tag_amp_apos_quot_flags() {
    compile_ok(
        r#"<?php
$html = '<a href="test.php?a=1&b=2">O\'Reilly</a>';
$encoded = json_encode($html, JSON_HEX_TAG | JSON_HEX_AMP | JSON_HEX_APOS | JSON_HEX_QUOT);
echo $encoded;
"#,
    );
}

#[test]
fn test_php_json_encode_partial_output_on_error() {
    compile_ok(
        r#"<?php
$invalidUtf8 = ["valid" => "text", "invalid" => "\xB1\x31"];
$json = @json_encode($invalidUtf8, JSON_PARTIAL_OUTPUT_ON_ERROR);
echo is_string($json) ? "PARTIAL_JSON" : "FAIL";
"#,
    );
}

#[test]
fn test_php_json_decode_max_depth_exceeded_error() {
    compile_ok(
        r#"<?php
$nestedJson = '{"a":{"b":{"c":{"d":1}}}}';
$res = json_decode($nestedJson, depth: 3);
if ($res === null && json_last_error() === JSON_ERROR_DEPTH) {
    echo "DEPTH_EXCEEDED";
}
"#,
    );
}

#[test]
fn test_php_json_exception_code_and_message() {
    compile_ok(
        r#"<?php
try {
    json_decode("{malformed}", flags: JSON_THROW_ON_ERROR);
} catch (JsonException $e) {
    echo "Code=" . $e->getCode() . " Msg=" . $e->getMessage();
}
"#,
    );
}
