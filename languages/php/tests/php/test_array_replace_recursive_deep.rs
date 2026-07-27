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

    array_replace_recursive_empty_replacement_keeps_base => {
        r#"<?php
$base = ['x' => ['y' => 1, 'z' => 2]];
$res = array_replace_recursive($base, []);
echo ($res === $base ? "same" : "diff") . "|" . $res['x']['y'];
"#,
        ["same|1"]
    };

    array_replace_recursive_adds_new_nested_arrays => {
        r#"<?php
$base = ['a' => ['one' => 1]];
$patch = ['a' => ['two' => 2], 'b' => ['three' => 3]];
$res = array_replace_recursive($base, $patch);
echo $res['a']['one'] . "|" . $res['a']['two'] . "|" . $res['b']['three'];
"#,
        ["1|2|3"]
    };

    array_replace_recursive_mixed_depth_override => {
        r#"<?php
$base = ['cfg' => ['feature' => ['enabled' => true, 'mode' => 'auto'], 'limit' => 10], 'other' => 1];
$patch = ['cfg' => ['feature' => 'off', 'limit' => 20]];
$res = array_replace_recursive($base, $patch);
echo is_string($res['cfg']['feature']) ? $res['cfg']['feature'] : 'arr';
echo "|" . $res['cfg']['limit'];
echo "|" . $res['other'];
"#,
        ["off|20|1"]
    };

    array_replace_recursive_with_numeric_string_keys => {
        r#"<?php
$base = ["1" => ["a" => 1], "2" => ["b" => 2]];
$patch = [1 => ["a" => 9], 2 => ["c" => 3]];
$res = array_replace_recursive($base, $patch);
echo $res[1]["a"] . "|" . $res[2]["b"] . "|" . $res[2]["c"];
"#,
        ["9|2|3"]
    };

    array_replace_recursive_keeps_scalar_non_array => {
        r#"<?php
$base = ['a' => ['x' => 1], 'b' => ['y' => 2]];
$patch = ['a' => 9];
$res = array_replace_recursive($base, $patch);
echo is_array($res['a']) ? 'arr' : 'scalar';
echo "|" . $res['a'];
"#,
        ["scalar|9"]
    };

    array_replace_recursive_removes_nesting_on_override => {
        r#"<?php
$base = ['cfg' => ['mode' => ['safe' => true], 'level' => 1]];
$patch = ['cfg' => ['mode' => 'off']];
$res = array_replace_recursive($base, $patch);
echo is_array($res['cfg']['mode']) ? 'arr' : 'scalar';
echo "|" . $res['cfg']['mode'];
"#,
        ["scalar|off"]
    };

    array_replace_recursive_multiple_patches_merge => {
        r#"<?php
$a = ['root' => ['a' => 1], 'x' => ['y' => 2]];
$b = ['root' => ['b' => 3]];
$c = ['root' => ['c' => 4], 'new' => 5];
$res = array_replace_recursive($a, $b, $c);
echo $res['root']['a'] . "|" . $res['root']['b'] . "|" . $res['root']['c'] . "|" . $res['new'];
"#,
        ["1|3|4|5"]
    };

    array_replace_recursive_empty_base_with_patch => {
        r#"<?php
$res = array_replace_recursive([], ['a' => 1, 'b' => ['c' => 2]]);
echo $res['a'];
echo "|" . $res['b']['c'];
"#,
        ["1|2"]
    };

    array_replace_recursive_with_list_vs_assoc_merge => {
        r#"<?php
$base = [0 => ['id' => 1], 1 => ['id' => 2]];
$patch = [1 => ['name' => 'x'], 2 => ['name' => 'y']];
$res = array_replace_recursive($base, $patch);
echo count($res);
echo "|" . count($res[1]);
echo "|" . ($res[0]['id'] ?? 'none') . "|" . ($res[1]['id'] ?? 'none') . "|" . ($res[1]['name'] ?? 'none');
"#,
        ["3|2|1|2|x"]
    };
}
