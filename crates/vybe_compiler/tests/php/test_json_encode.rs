//! `json_encode` / `json_decode` success paths and flag combinations.

crate::php_cases! {
    json_encode_indexed_array => {
        r#"<?php
echo json_encode([1, 2, 3]);
"#,
        ["[1,2,3]"]
    };

    json_encode_assoc_object_keys => {
        r#"<?php
echo json_encode(['name' => 'ada', 'n' => 1]);
"#,
        ["{\"name\":\"ada\",\"n\":1}"]
    };

    json_encode_true_false_null => {
        r#"<?php
echo json_encode([true, false, null]);
"#,
        ["[true,false,null]"]
    };

    json_encode_pretty_print_inserts_newlines => {
        r#"<?php
$out = json_encode(['k' => 1], JSON_PRETTY_PRINT);
echo str_contains($out, "\n") ? 'pretty' : 'flat';
"#,
        ["pretty"]
    };

    json_encode_unescaped_slashes => {
        r#"<?php
echo json_encode(['u' => 'https://ex.com/a/b'], JSON_UNESCAPED_SLASHES);
"#,
        ["{\"u\":\"https://ex.com/a/b\"}"]
    };

    json_encode_force_object_on_list => {
        r#"<?php
echo json_encode([1, 2], JSON_FORCE_OBJECT);
"#,
        ["{\"0\":1,\"1\":2}"]
    };

    json_decode_assoc_array => {
        r#"<?php
$d = json_decode('{"a":1,"b":2}', true);
echo $d['a'] . ':' . $d['b'];
"#,
        ["1:2"]
    };

    json_decode_to_object_property_access => {
        r#"<?php
$o = json_decode('{"x":"y"}');
echo $o->x;
"#,
        ["y"]
    };

    json_decode_empty_array => {
        r#"<?php
echo json_encode(json_decode('[]', true));
"#,
        ["[]"]
    };

    json_last_error_none_after_success => {
        r#"<?php
json_decode('{}');
echo json_last_error() === JSON_ERROR_NONE ? 'ok' : 'err';
"#,
        ["ok"]
    };

    json_encode_numeric_string_preserved_with_flag => {
        r#"<?php
echo json_encode(['id' => '42'], JSON_NUMERIC_CHECK);
"#,
        ["{\"id\":42}"]
    };

    json_encode_hex_tag_escapes_angle_brackets => {
        r#"<?php
echo json_encode('<tag>', JSON_HEX_TAG);
"#,
        ["\"\\u003Ctag\\u003E\""]
    };

    json_decode_nested_structure => {
        r#"<?php
$d = json_decode('{"user":{"id":3}}', true);
echo $d['user']['id'];
"#,
        ["3"]
    };

    json_encode_empty_string => {
        r#"<?php
echo json_encode('');
"#,
        ["\"\""]
    };

    json_encode_unicode_unescaped => {
        r#"<?php
echo json_encode(['m' => 'café'], JSON_UNESCAPED_UNICODE);
"#,
        ["{\"m\":\"café\"}"]
    };

    json_decode_invalid_returns_null_without_throw => {
        r#"<?php
echo json_decode('{bad') === null ? 'null' : 'val';
"#,
        ["null"]
    };

    json_encode_float_rounded => {
        r#"<?php
echo json_encode(1.5);
"#,
        ["1.5"]
    };

    json_decode_boolean_literals => {
        r#"<?php
$b = json_decode('true');
echo $b ? 'true' : 'false';
"#,
        ["true"]
    };

    json_encode_object_with_public_props => {
        r#"<?php
$o = new stdClass();
$o->k = 'v';
echo json_encode($o);
"#,
        ["{\"k\":\"v\"}"]
    };

    json_decode_depth_within_limit => {
        r#"<?php
echo json_decode('{"a":{"b":1}}', true)['a']['b'];
"#,
        ["1"]
    };
}
