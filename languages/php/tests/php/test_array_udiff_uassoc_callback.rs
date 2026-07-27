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

    array_udiff_uassoc_empty_second => {
        r#"<?php
$a = ["a" => "One", "b" => "Two"];
$r = array_udiff_uassoc($a, [], "strcasecmp", "strcasecmp");
ksort($r);
echo implode(',', array_keys($r));
"#,
        ["a,b"]
    };

    array_udiff_uassoc_key_compare_case_insensitive => {
        r#"<?php
$a = ["A" => 1, "b" => 2, "C" => 3];
$b = ["a" => 1, "B" => 2];
$r = array_udiff_uassoc(
    $a,
    $b,
    fn($v1, $v2) => ($v1 <=> $v2),
    "strcasecmp"
);
ksort($r);
echo implode(',', array_keys($r));
"#,
        ["C"]
    };

    array_udiff_uassoc_with_callback_zero => {
        r#"<?php
$a = ["x" => 1, "y" => 2];
$b = ["x" => 1];
$r = array_udiff_uassoc(
    $a,
    $b,
    fn($v1, $v2) => ($v1 <=> $v2),
    fn($k1, $k2) => strcmp($k1, $k2)
);
echo implode(',', array_keys($r));
"#,
        ["y"]
    };

    array_udiff_uassoc_empty_key_set => {
        r#"<?php
$a = ["a" => "A", "b" => "B"];
$r = array_udiff_uassoc($a, ["A" => "A"], fn($v1,$v2) => strcmp($v1, $v2), "strcasecmp");
ksort($r);
echo count($r) . "|" . implode(',', array_keys($r));
"#,
        ["1|b"]
    };

    array_udiff_uassoc_value_mismatch_only => {
        r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3];
$b = ["a" => 1, "b" => 9];
$r = array_udiff_uassoc($a, $b, fn($v1,$v2)=>$v1<=>$v2, "strcmp");
echo implode(',', array_keys($r)) . '|' . $r['c'];
"#,
        ["c|3"]
    };

    array_udiff_uassoc_key_casting => {
        r#"<?php
$a = ["01" => 1, 1 => 2, 2.2 => 3];
$b = [1 => 1, 2 => 3];
$r = array_udiff_uassoc($a, $b, fn($v1,$v2)=>$v1<=>$v2, fn($k1,$k2)=> (string)$k1 <=> (string)$k2);
ksort($r);
echo count($r) . '|' . implode(',', array_keys($r));
"#,
        ["2|01,1"]
    };

    array_udiff_uassoc_empty_first_array => {
        r#"<?php
$a = [];
$b = ["a" => 1];
$r = array_udiff_uassoc($a, $b, fn($v1, $v2) => $v1 <=> $v2, fn($k1, $k2) => strcmp($k1, $k2));
echo is_array($r) ? 'array' : 'not-array';
echo '|';
echo count($r);
"#,
        ["array|0"]
    };

    array_udiff_uassoc_empty_second_array => {
        r#"<?php
$a = ["a" => 1, "b" => "2"];
$r = array_udiff_uassoc($a, [], fn($v1, $v2) => $v1 <=> $v2, fn($k1, $k2) => strcmp($k1, $k2));
ksort($r);
echo implode(',', array_keys($r));
"#,
        ["a,b"]
    };

    array_udiff_uassoc_duplicate_values_different_keys => {
        r#"<?php
$a = ["a" => 1, "b" => 1];
$b = ["x" => 1];
$r = array_udiff_uassoc($a, $b, fn($v1, $v2) => $v1 <=> $v2, fn($k1, $k2) => strcmp($k1, $k2));
ksort($r);
echo implode(',', array_keys($r));
"#,
        ["a,b"]
    };

    array_udiff_uassoc_strict_numeric_match => {
        r#"<?php
$a = ["a" => 1, "b" => 1];
$b = ["a" => 1];
$r = array_udiff_uassoc($a, $b, fn($v1, $v2) => $v1 <=> $v2, fn($k1, $k2) => strcmp($k1, $k2));
echo implode(',', array_keys($r));
"#,
        ["b"]
    };

    array_udiff_uassoc_locale_key_compare => {
        r#"<?php
$a = ["A" => 1, "b" => 2, "C" => 3];
$b = ["a" => 1, "B" => 2];
$r = array_udiff_uassoc(
    $a,
    $b,
    fn($v1, $v2) => $v1 <=> $v2,
    fn($k1, $k2) => strtolower($k1) <=> strtolower($k2)
);
ksort($r);
echo implode(',', array_keys($r));
"#,
        ["C"]
    };

    array_udiff_uassoc_exception_in_value_callback => {
        r#"<?php
$a = ["a" => 1, "b" => 2];
$b = ["c" => 3];
try {
    array_udiff_uassoc(
        $a,
        $b,
        function($v1, $v2) {
            throw new RuntimeException('diff-failed');
        },
        fn($k1, $k2) => strcmp($k1, $k2)
    );
    echo 'no-exception';
} catch (Throwable $e) {
    echo $e->getMessage();
}
"#,
        ["diff-failed"]
    };
}
