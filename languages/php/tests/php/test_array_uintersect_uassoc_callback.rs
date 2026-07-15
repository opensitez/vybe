use super::helpers::run_prints;

crate::php_cases! {
    array_uintersect_uassoc_basic => {
        r#"<?php
$array1 = ["a" => "green", "b" => "brown", "c" => "blue", "red"];
$array2 = ["a" => "GREEN", "B" => "brown", "yellow", "red"];

$result = array_uintersect_uassoc($array1, $array2, "strcasecmp", "strcasecmp");
ksort($result);
echo implode(',', array_keys($result)) . "|" . implode(',', array_values($result));
"#,
        ["a,b|green,brown"]
    };

    array_uintersect_uassoc_closure => {
        r#"<?php
$arr1 = [1 => 10, 2 => 20];
$arr2 = ["1" => "10", 3 => 30];

$res = array_uintersect_uassoc(
    $arr1, 
    $arr2, 
    function($a, $b) { return $a <=> $b; }, 
    function($a, $b) { return (int)$a <=> (int)$b; }
);
echo $res[1];
"#,
        ["10"]
    };
}
