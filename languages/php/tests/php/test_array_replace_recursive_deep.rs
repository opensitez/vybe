use super::helpers::run_prints;

crate::php_cases! {
    array_replace_recursive_deep => {
        r#"<?php
$base = ['citrus' => ['orange'], 'berries' => ['blackberry', 'raspberry'], 'others' => 'banana'];
$replacements = ['citrus' => 'pineapple', 'berries' => ['blueberry'], 'others' => ['litchi']];

$basket = array_replace_recursive($base, $replacements);
echo $basket['citrus'] . "|" . $basket['berries'][0] . "|" . $basket['berries'][1] . "|" . $basket['others'][0];
"#,
        ["pineapple|blueberry|raspberry|litchi"]
    };

    array_replace_recursive_multiple_arrays => {
        r#"<?php
$a1 = ['a' => ['b' => 1]];
$a2 = ['a' => ['c' => 2]];
$a3 = ['a' => ['b' => 3]];

$res = array_replace_recursive($a1, $a2, $a3);
echo $res['a']['b'] . "|" . $res['a']['c'];
"#,
        ["3|2"]
    };
}
