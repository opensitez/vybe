
crate::php_cases! {
    array_filter_use_key => {
        r#"<?php
$arr = ['a' => 1, 'b' => 2, 'c' => 3];
$res = array_filter($arr, function($k) { return $k == 'b'; }, ARRAY_FILTER_USE_KEY);
echo implode(',', array_values($res));
"#,
        ["2"]
    };

    array_filter_use_both => {
        r#"<?php
$arr = ['a' => 1, 'b' => 2, 'c' => 3];
$res = array_filter($arr, function($v, $k) {
    return $k == 'a' || $v == 3;
}, ARRAY_FILTER_USE_BOTH);
echo implode(',', array_keys($res));
"#,
        ["a,c"]
    };

    array_filter_default_uses_value_and_keeps_truthy_only => {
        r#"<?php
$arr = [0, 1, '', 2, false, 3];
$res = array_filter($arr);
echo count($res) . '|' . implode('', $res);
"#,
        ["3|123"]
    };

    array_filter_does_not_reindex_without_callback => {
        r#"<?php
$arr = ['x' => 0, 'y' => 2, 'z' => 0, 'w' => 4];
$res = array_filter($arr);
echo implode('|', array_keys($res));
"#,
        ["y|w"]
    };

    array_filter_use_key_with_keys_out_of_order => {
        r#"<?php
$arr = ['z' => 3, 'x' => 2, 'y' => 1];
$res = array_filter($arr, function($v, $k) {
    return $k === 'x' || $v > 2;
}, ARRAY_FILTER_USE_BOTH);
echo implode('|', array_keys($res));
"#,
        ["z|x"]
    };

    array_filter_with_lambda_returns_only_true_indexes => {
        r#"<?php
$arr = [0 => false, 1 => true, 2 => false, 3 => true];
$res = array_filter($arr, fn($v) => $v);
echo json_encode(array_values(array_keys($res)));
"#,
        ["[1,3]"]
    };

    array_filter_use_key_numeric_string_coercion => {
        r#"<?php
$arr = ['0' => 0, 1 => 1, 2 => 2];
$res = array_filter($arr, function($k) {
    return $k == 1;
}, ARRAY_FILTER_USE_KEY);
echo implode(',', array_keys($res)) . '|' . implode(',', array_values($res));
"#,
        ["1|1"]
    };

    array_filter_use_both_with_strict_key_value_match => {
        r#"<?php
$arr = ['a' => 1, 'b' => 2, 'c' => '2'];
$res = array_filter($arr, fn($v, $k) => $k === 'a' || $v === 2, ARRAY_FILTER_USE_BOTH);
ksort($res);
echo implode('|', array_keys($res)) . '|' . implode(',', array_values($res));
"#,
        ["a,b|1,2"]
    };

    array_filter_use_key_when_values_negative => {
        r#"<?php
$arr = ['a' => -1, 'b' => 0, 'c' => -2];
$res = array_filter($arr, fn($v, $k) => $v < 0, ARRAY_FILTER_USE_BOTH);
echo implode('|', array_keys($res));
"#,
        ["a|c"]
    };
}
