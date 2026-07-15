use super::helpers::run_prints;

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
}
