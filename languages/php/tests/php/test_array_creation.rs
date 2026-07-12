//! `array_fill`, `array_pad`, `range`, `array_merge_recursive`, and `array_replace_recursive`.

crate::php_cases! {
    array_fill_creates_repeated_values => {
        r#"<?php
echo implode(',', array_fill(0, 3, 'x'));
"#,
        ["x,x,x"]
    };

    array_fill_keys_uses_keys_list => {
        r#"<?php
echo json_encode(array_fill_keys(['a', 'b'], 0));
"#,
        ["{\"a\":0,\"b\":0}"]
    };

    array_pad_appends_pad_value => {
        r#"<?php
echo implode(',', array_pad([1], 4, 0));
"#,
        ["1,0,0,0"]
    };

    array_pad_negative_prepends => {
        r#"<?php
echo implode(',', array_pad([9], -3, 0));
"#,
        ["0,0,9"]
    };

    range_inclusive_integers => {
        r#"<?php
echo implode(',', range(2, 5));
"#,
        ["2,3,4,5"]
    };

    range_with_step => {
        r#"<?php
echo implode(',', range(0, 6, 2));
"#,
        ["0,2,4,6"]
    };

    array_merge_recursive_combines_nested => {
        r#"<?php
$a = ['x' => ['a' => 1]];
$b = ['x' => ['b' => 2]];
echo json_encode(array_merge_recursive($a, $b));
"#,
        ["{\"x\":{\"a\":1,\"b\":2}}"]
    };

    array_replace_recursive_overwrites_nested => {
        r#"<?php
$a = ['k' => ['a' => 1, 'b' => 2]];
$b = ['k' => ['b' => 9]];
echo json_encode(array_replace_recursive($a, $b));
"#,
        ["{\"k\":{\"a\":1,\"b\":9}}"]
    };

    array_change_key_case_lower => {
        r#"<?php
echo json_encode(array_change_key_case(['Foo' => 1], CASE_LOWER));
"#,
        ["{\"foo\":1}"]
    };

    array_rand_picks_key_from_list => {
        r#"<?php
$keys = ['only' => 1];
$k = array_rand($keys);
echo $k;
"#,
        ["only"]
    };

    array_sum_numeric_list => {
        r#"<?php
echo array_sum([1, 2, 3, 4]);
"#,
        ["10"]
    };

    array_product_multiplies => {
        r#"<?php
echo array_product([2, 3, 4]);
"#,
        ["24"]
    };

    array_count_values_frequency_map => {
        r#"<?php
echo json_encode(array_count_values(['a', 'b', 'a']));
"#,
        ["{\"a\":2,\"b\":1}"]
    };

    array_reverse_preserves_keys_by_default => {
        r#"<?php
echo json_encode(array_reverse(['a' => 1, 'b' => 2]));
"#,
        ["{\"b\":2,\"a\":1}"]
    };

    array_is_list_detects_sequential => {
        r#"<?php
echo array_is_list([1, 2, 3]) ? 'list' : 'map';
"#,
        ["list"]
    };
}
