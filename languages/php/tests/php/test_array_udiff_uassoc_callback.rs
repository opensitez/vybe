use super::helpers::run_prints;

crate::php_cases! {
    array_udiff_uassoc_basic => {
        r#"<?php
$array1 = array("a" => "green", "b" => "brown", "c" => "blue", "red");
$array2 = array("a" => "GREEN", "B" => "brown", "yellow", "red");

$result = array_udiff_uassoc($array1, $array2, "strcasecmp", "strcasecmp");
ksort($result);
echo implode(',', array_keys($result));
"#,
        ["0,c"]
    };

    array_udiff_uassoc_objects => {
        r#"<?php
class Item {
    public function __construct(public int $id) {}
}
$a1 = ['x' => new Item(1), 'y' => new Item(2)];
$a2 = ['X' => new Item(1)];

$res = array_udiff_uassoc($a1, $a2, 
    function($a, $b) { return $a->id <=> $b->id; },
    function($a, $b) { return strcasecmp($a, $b); }
);
echo count($res) . "|" . array_keys($res)[0];
"#,
        ["1|y"]
    };
}
