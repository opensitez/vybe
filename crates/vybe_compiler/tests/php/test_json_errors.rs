//! JSON encode/decode failure modes with distinct payloads (not duplicate invalid-json catch).

crate::php_cases! {
    json_encode_nan_with_throw_flag => {
        r#"<?php
try { json_encode(NAN, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'nan'; }
"#,
        ["nan"]
    };

    json_encode_inf_with_throw_flag => {
        r#"<?php
try { json_encode(INF, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'inf'; }
"#,
        ["inf"]
    };

    json_encode_negative_inf_with_throw_flag => {
        r#"<?php
try { json_encode(-INF, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'ninf'; }
"#,
        ["ninf"]
    };

    json_encode_resource_rejected => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
try { json_encode($fp, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'res'; }
finally { fclose($fp); }
"#,
        ["res"]
    };

    json_encode_recursive_array_detected => {
        r#"<?php
$a = [];
$a['self'] = &$a;
try { json_encode($a, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'recurse'; }
"#,
        ["recurse"]
    };

    json_encode_object_with_private_cycle => {
        r#"<?php
class Node { public Node $next; }
$n = new Node();
$n->next = $n;
try { json_encode($n, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'cycle'; }
"#,
        ["cycle"]
    };

    json_encode_valid_object_without_throw => {
        r#"<?php
$obj = new stdClass();
$obj->x = 1;
echo json_encode($obj, JSON_THROW_ON_ERROR);
"#,
        ["{\"x\":1}"]
    };

    json_decode_trailing_garbage_with_throw => {
        r#"<?php
try { json_decode('{} junk', false, 512, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'trail'; }
"#,
        ["trail"]
    };

    json_decode_single_quoted_string_invalid => {
        r#"<?php
try { json_decode("{'a':1}", false, 512, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'quotes'; }
"#,
        ["quotes"]
    };

    json_decode_depth_limit_exceeded => {
        r#"<?php
$deep = str_repeat('{"a":', 100) . '1' . str_repeat('}', 100);
try { json_decode($deep, false, 5, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'depth'; }
"#,
        ["depth"]
    };

    json_decode_true_literal_returns_bool => {
        r#"<?php
$v = json_decode('true', false, 512, JSON_THROW_ON_ERROR);
echo $v === true ? 'bool' : 'other';
"#,
        ["bool"]
    };

    json_decode_assoc_true_returns_array => {
        r#"<?php
$data = json_decode('{"k":2}', true, 512, JSON_THROW_ON_ERROR);
echo is_array($data) && $data['k'] === 2 ? 'assoc' : 'no';
"#,
        ["assoc"]
    };

    json_decode_object_default_returns_stdclass => {
        r#"<?php
$data = json_decode('{"k":2}', false, 512, JSON_THROW_ON_ERROR);
echo $data instanceof stdClass ? 'obj' : 'no';
"#,
        ["obj"]
    };

    json_last_error_ok_after_successful_encode => {
        r#"<?php
json_encode([1]);
echo json_last_error() === JSON_ERROR_NONE ? 'ok' : 'err';
"#,
        ["ok"]
    };

    json_last_error_msg_after_success => {
        r#"<?php
json_encode('x');
$msg = json_last_error_msg();
echo $msg === 'No error' ? 'clean' : $msg;
"#,
        ["clean"]
    };

    json_encode_pretty_print_inserts_newlines => {
        r#"<?php
$out = json_encode(['a' => 1], JSON_THROW_ON_ERROR | JSON_PRETTY_PRINT);
echo str_contains($out, "\n") ? 'pretty' : 'flat';
"#,
        ["pretty"]
    };

    json_encode_unicode_unescaped_preserves_utf8 => {
        r#"<?php
$out = json_encode(['msg' => 'café'], JSON_THROW_ON_ERROR | JSON_UNESCAPED_UNICODE);
echo str_contains($out, 'café') ? 'utf8' : 'esc';
"#,
        ["utf8"]
    };

    json_encode_numeric_check_stringifies_numbers => {
        r#"<?php
$out = json_encode(['id' => '42'], JSON_THROW_ON_ERROR | JSON_NUMERIC_CHECK);
echo str_contains($out, '42') && !str_contains($out, '"42"') ? 'num' : 'str';
"#,
        ["num"]
    };

    json_encode_force_object_on_array => {
        r#"<?php
$out = json_encode([1, 2], JSON_THROW_ON_ERROR | JSON_FORCE_OBJECT);
echo str_starts_with($out, '{') ? 'object' : 'array';
"#,
        ["object"]
    };

    json_decode_bigint_as_string => {
        r#"<?php
$out = json_decode('{"n":12345678901234567890}', true, 512, JSON_THROW_ON_ERROR | JSON_BIGINT_AS_STRING);
echo is_string($out['n']) ? 'string' : 'int';
"#,
        ["string"]
    };

    json_encode_empty_array => {
        r#"<?php
echo json_encode([], JSON_THROW_ON_ERROR);
"#,
        ["[]"]
    };

    json_encode_empty_object_via_cast => {
        r#"<?php
echo json_encode((object)[], JSON_THROW_ON_ERROR);
"#,
        ["{}"]
    };

    json_validate_matrix_invalid_and_valid => {
        r#"<?php
if (!function_exists('json_validate')) { echo 'skip'; }
else {
    $bad = json_validate('{') ? 'B' : 'b';
    $good = json_validate('{"a":1}') ? 'G' : 'g';
    echo $bad . $good;
}
"#,
        ["bG"]
    };

    json_decode_null_literal => {
        r#"<?php
$v = json_decode('null', false, 512, JSON_THROW_ON_ERROR);
echo $v === null ? 'null' : 'other';
"#,
        ["null"]
    };

    json_decode_boolean_literals => {
        r#"<?php
$t = json_decode('true', false, 512, JSON_THROW_ON_ERROR);
$f = json_decode('false', false, 512, JSON_THROW_ON_ERROR);
echo ($t ? 't' : 'f') . ($f ? 't' : 'f');
"#,
        ["tf"]
    };

    json_decode_number_as_float => {
        r#"<?php
$v = json_decode('1.5', false, 512, JSON_THROW_ON_ERROR);
echo is_float($v) ? 'float' : 'other';
"#,
        ["float"]
    };

    json_encode_flags_combine_throw_and_unicode => {
        r#"<?php
$out = json_encode(['€' => 1], JSON_THROW_ON_ERROR | JSON_UNESCAPED_UNICODE);
echo str_contains($out, '€') ? 'euro' : 'no';
"#,
        ["euro"]
    };

    json_decode_invalid_utf8_sequence => {
        r#"<?php
$bad = "\"\xB1\x31\"";
try { json_decode($bad, false, 512, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'utf8'; }
"#,
        ["utf8"]
    };

    json_encode_nan_without_throw_returns_false => {
        r#"<?php
$ok = json_encode(NAN);
echo $ok === false ? 'false' : 'value';
"#,
        ["false"]
    };

    json_last_error_after_failed_encode_without_throw => {
        r#"<?php
json_encode(NAN);
echo json_last_error() !== JSON_ERROR_NONE ? 'err' : 'ok';
"#,
        ["err"]
    };

    json_roundtrip_nested_user_count => {
        r#"<?php
$data = ['users' => [['id' => 1], ['id' => 2]]];
$out = json_encode($data, JSON_THROW_ON_ERROR);
$back = json_decode($out, true, 512, JSON_THROW_ON_ERROR);
echo count($back['users']);
"#,
        ["2"]
    };

    json_decode_to_object_nested_property => {
        r#"<?php
$obj = json_decode('{"outer":{"inner":3}}', false, 512, JSON_THROW_ON_ERROR);
echo $obj->outer->inner;
"#,
        ["3"]
    };

    json_encode_replaces_invalid_utf8_when_substitute_set => {
        r#"<?php
$out = json_encode("\xB1\x31", JSON_THROW_ON_ERROR | JSON_INVALID_UTF8_SUBSTITUTE);
echo is_string($out) ? 'sub' : 'fail';
"#,
        ["sub"]
    };

    json_decode_extra_comma_in_array_fails => {
        r#"<?php
try { json_decode('[1,]', false, 512, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'comma'; }
"#,
        ["comma"]
    };

    json_decode_duplicate_keys_last_wins => {
        r#"<?php
$obj = json_decode('{"k":1,"k":2}', true, 512, JSON_THROW_ON_ERROR);
echo $obj['k'];
"#,
        ["2"]
    };

    json_encode_nan_in_array_fails_with_throw => {
        r#"<?php
try { json_encode([1, NAN, 3], JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'arr-nan'; }
"#,
        ["arr-nan"]
    };

    json_encode_preserve_zero_fraction => {
        r#"<?php
$out = json_encode(1.0, JSON_THROW_ON_ERROR | JSON_PRESERVE_ZERO_FRACTION);
echo str_contains($out, '.0') ? 'zero' : 'intish';
"#,
        ["zero"]
    };

    json_decode_large_int_within_int_range => {
        r#"<?php
$v = json_decode('12345', false, 512, JSON_THROW_ON_ERROR);
echo $v === 12345 ? 'int' : 'other';
"#,
        ["int"]
    };

    json_encode_hex_tag_escapes_angle_brackets => {
        r#"<?php
$out = json_encode('<tag>', JSON_THROW_ON_ERROR | JSON_HEX_TAG);
echo str_contains($out, '\\u003C') ? 'hex' : 'plain';
"#,
        ["hex"]
    };

    json_decode_control_char_in_string_with_throw => {
        r#"<?php
try { json_decode("\"a\nb\"", false, 512, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'ctrl'; }
"#,
        ["ctrl"]
    };

    json_encode_object_with_jsonserializable => {
        r#"<?php
class Payload implements JsonSerializable {
    public function jsonSerialize(): array { return ['v' => 7]; }
}
echo json_encode(new Payload(), JSON_THROW_ON_ERROR);
"#,
        ["{\"v\":7}"]
    };

    json_decode_malformed_unicode_escape => {
        r#"<?php
try { json_decode('"\uZZZZ"', false, 512, JSON_THROW_ON_ERROR); echo 'ok'; }
catch (JsonException $e) { echo 'unicode'; }
"#,
        ["unicode"]
    };

    json_encode_max_depth_nested_arrays => {
        r#"<?php
$a = []; $cursor = &$a;
for ($i = 0; $i < 3; $i++) { $cursor['n'] = []; $cursor = &$cursor['n']; }
echo json_encode($a, JSON_THROW_ON_ERROR);
"#,
        ["{\"n\":{\"n\":{\"n\":[]}}}"]
    };

    json_decode_stream_of_array_values => {
        r#"<?php
$vals = json_decode('[true,false,null,1,"s"]', true, 512, JSON_THROW_ON_ERROR);
echo count($vals);
"#,
        ["5"]
    };

    json_encode_slash_escaping_default => {
        r#"<?php
$out = json_encode(['path' => 'a/b'], JSON_THROW_ON_ERROR);
echo str_contains($out, '\\/') ? 'escaped' : 'raw';
"#,
        ["escaped"]
    };

    json_decode_scientific_notation_number => {
        r#"<?php
$v = json_decode('1e2', false, 512, JSON_THROW_ON_ERROR);
echo $v == 100 ? 'sci' : 'no';
"#,
        ["sci"]
    };
}
