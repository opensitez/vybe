use super::helpers::run_prints;

crate::php_cases! {
    array_walk_recursive_by_reference => {
        r#"<?php
$sweet = ['a' => 'apple', 'b' => 'banana'];
$fruits = ['sweet' => $sweet, 'sour' => 'lemon'];

function test_print(&$item, $key, $prefix) {
    $item = "$prefix: $item";
}

array_walk_recursive($fruits, 'test_print', 'fruit');

echo $fruits['sweet']['a'] . "|" . $fruits['sour'];
"#,
        ["fruit: apple|fruit: lemon"]
    };

    array_walk_recursive_objects => {
        r#"<?php
class Obj { public $val = 1; }
$arr = [new Obj(), [new Obj()]];

array_walk_recursive($arr, function($v, $k) {
    $v->val += 10;
});
echo $arr[0]->val . "|" . $arr[1][0]->val;
"#,
        ["11|11"]
    };
}
