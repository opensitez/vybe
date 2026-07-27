crate::php_cases! {
    array_intersect_ukey_basic => {
        r#"<?php
$array1 = ['blue'  => 1, 'red'  => 2, 'green'  => 3, 'purple' => 4];
$array2 = ['green' => 5, 'blue' => 6, 'yellow' => 7, 'cyan'   => 8];

$result = array_intersect_ukey($array1, $array2, function ($key1, $key2) {
    if ($key1 == $key2) return 0;
    else if ($key1 > $key2) return 1;
    else return -1;
});

echo implode(',', array_keys($result));
"#,
        ["blue,green"]
    };

    array_intersect_ukey_custom_comparison => {
        r#"<?php
$array1 = ['apple' => 1, 'banana' => 2];
$array2 = ['APPLE' => 3, 'Orange' => 4];

$result = array_intersect_ukey($array1, $array2, 'strcasecmp');
echo implode(',', array_keys($result));
"#,
        ["apple"]
    };

    array_intersect_ukey_preserves_input_order => {
        r#"<?php
$a = ['z' => 1, 'a' => 2, 'm' => 3];
$b = ['A' => 9, 'm' => 8];
$r = array_intersect_ukey($a, $b, 'strcasecmp');
echo implode('|', array_keys($r));
"#,
        ["a|m"]
    };

    array_intersect_ukey_empty_right => {
        r#"<?php
$a = ['x' => 1, 'y' => 2];
$r = array_intersect_ukey($a, [], 'strcmp');
echo count($r) . '|' . (empty($r) ? 'empty' : 'not');
"#,
        ["0|empty"]
    };
}
