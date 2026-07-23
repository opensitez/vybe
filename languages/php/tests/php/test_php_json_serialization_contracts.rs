use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: JSON Serialization & Contracts — json_encode, json_decode, JsonSerializable, JSON_THROW_ON_ERROR, flags
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_json_serializable_interface_implementation() {
    let out = run_prints(
        r#"<?php
class ApiResponse implements JsonSerializable {
    public function __construct(
        private string $status,
        private array $data
    ) {}

    public function jsonSerialize(): mixed {
        return [
            "success" => $this->status === "ok",
            "payload" => $this->data,
        ];
    }
}

$res = new ApiResponse("ok", ["user_id" => 42]);
echo json_encode($res);
"#,
    );
    assert_eq!(out, vec![r#"{"success":true,"payload":{"user_id":42}}"#]);
}

#[test]
fn test_php_json_decode_associative_array() {
    let out = run_prints(
        r#"<?php
$json = '{"name":"Alice","skills":["PHP","Rust"]}';
$data = json_decode($json, true);
echo $data["name"] . " -> " . implode(",", $data["skills"]);
"#,
    );
    assert_eq!(out, vec!["Alice -> PHP,Rust"]);
}

#[test]
fn test_php_json_throw_on_error_exception() {
    let out = run_prints(
        r#"<?php
try {
    json_decode("{invalid json}", flags: JSON_THROW_ON_ERROR);
} catch (JsonException $e) {
    echo "JsonException: " . $e->getMessage();
}
"#,
    );
    assert_eq!(out, vec!["JsonException: Syntax error"]);
}

#[test]
fn test_php_json_encode_flags_pretty_print_unescaped() {
    let out = run_prints(
        r#"<?php
$data = ["url" => "https://example.com/api", "title" => "Home & About"];
$json = json_encode($data, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
echo $json;
"#,
    );
    assert_eq!(
        out,
        vec![r#"{"url":"https://example.com/api","title":"Home & About"}"#]
    );
}

#[test]
fn test_php_json_encode_stdclass_cast() {
    let out = run_prints(
        r#"<?php
$obj = (object)["key" => "value", "id" => 123];
echo json_encode($obj);
"#,
    );
    assert_eq!(out, vec![r#"{"key":"value","id":123}"#]);
}

#[test]
fn test_php_json_validate_function_php83() {
    compile_ok(
        r#"<?php
if (function_exists('json_validate')) {
    echo json_validate('{"valid": true}') ? "VALID" : "INVALID";
} else {
    echo "VALID";
}
"#,
    );
}

#[test]
fn test_php_json_last_error_and_msg() {
    compile_ok(
        r#"<?php
$result = json_decode("{bad json}");
if (json_last_error() !== JSON_ERROR_NONE) {
    echo "Error code: " . json_last_error() . " Msg: " . json_last_error_msg();
}
"#,
    );
}

#[test]
fn test_php_json_encode_depth_limit() {
    compile_ok(
        r#"<?php
$nested = [[[["deep"]]]];
$json = json_encode($nested, depth: 2);
if ($json === false && json_last_error() === JSON_ERROR_DEPTH) {
    echo "Exceeded maximum depth";
}
"#,
    );
}

#[test]
fn test_php_json_encode_numeric_check_flag() {
    compile_ok(
        r#"<?php
$data = ["id" => "123", "score" => "98.6", "name" => "Alice"];
echo json_encode($data, JSON_NUMERIC_CHECK);
"#,
    );
}

#[test]
fn test_php_json_decode_max_depth() {
    compile_ok(
        r#"<?php
$json = '{"a":{"b":{"c":1}}}';
$obj = json_decode($json, associative: false, depth: 512);
echo $obj->a->b->c;
"#,
    );
}
