use super::helpers::run_prints;

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
}
