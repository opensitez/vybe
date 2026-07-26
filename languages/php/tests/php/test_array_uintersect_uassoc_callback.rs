
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

    array_uintersect_uassoc_empty_other => {
        r#"<?php
$a = ["a" => 1, "b" => 2];
$r = array_uintersect_uassoc($a, [], "strcasecmp", "strcasecmp");
echo count($r);
"#,
        ["0"]
    };

    array_uintersect_uassoc_key_as_numeric_string => {
        r#"<?php
$a = ["1" => "A", "2" => "B", "x" => "X"];
$b = [1 => "a", "2" => "b"];
$r = array_uintersect_uassoc($a, $b, fn($v1,$v2)=>strcmp($v1,$v2), "strcasecmp");
ksort($r);
echo implode('|', array_keys($r)) . ":" . $r["1"];
"#,
        ["1|2:A"]
    };

    array_uintersect_uassoc_partial_match_by_value => {
        r#"<?php
$a = ["a" => "Alpha", "b" => "Beta"];
$b = ["a" => "ALPHA", "c" => "Gamma"];
$r = array_uintersect_uassoc($a, $b, "strcasecmp", "strcmp");
echo implode(',', array_keys($r)) . ":" . implode(',', array_values($r));
"#,
        ["a:Alpha"]
    };

    array_uintersect_uassoc_three_arrays => {
        r#"<?php
$a = ["a" => "X", "b" => "Y"];
$b = ["a" => "x", "c" => "z"];
$c = ["a" => "x", "b" => "y"];
$r = array_uintersect_uassoc($a, $b, "strcasecmp", "strcasecmp");
$r = array_uintersect_uassoc($r, $c, "strcasecmp", "strcasecmp");
echo count($r) . "|" . implode(',', array_keys($r));
"#,
        ["1|a"]
    };

    array_uintersect_uassoc_key_numeric_compare => {
        r#"<?php
$a = ["01" => "x", 1 => "y", "2" => "z"];
$b = [1 => "Y", "2" => "Z"];
$r = array_uintersect_uassoc($a, $b, "strcasecmp", function($k1, $k2) { return (string)$k1 <=> (string)$k2; });
ksort($r);
echo implode('|', array_keys($r));
"#,
        ["1|2"]
    };

    array_uintersect_uassoc_first_empty_returns_empty => {
        r#"<?php
$a = [];
$b = ["a" => "1"];
$r = array_uintersect_uassoc($a, $b, fn($v1, $v2) => strcmp((string)$v1, (string)$v2), fn($k1, $k2) => strcmp((string)$k1, (string)$k2));
echo count($r);
"#,
        ["0"]
    };

    array_uintersect_uassoc_empty_second_keeps_empty => {
        r#"<?php
$a = ["a" => "One", "b" => "Two"];
$r = array_uintersect_uassoc($a, [], "strcasecmp", "strcasecmp");
echo is_array($r) ? 'array' : 'not-array';
echo '|';
echo count($r);
"#,
        ["array|0"]
    };

    array_uintersect_uassoc_duplicate_value_single_key => {
        r#"<?php
$a = ["x" => "same", "y" => "same", "z" => "other"];
$b = ["m" => "SAME", "n" => "other"];
$r = array_uintersect_uassoc($a, $b, "strcasecmp", "strcasecmp");
ksort($r);
echo implode('|', array_keys($r));
"#,
        ["x,y,z"]
    };

    array_uintersect_uassoc_numeric_key_casting => {
        r#"<?php
$a = ["1" => "A", 1 => "B"];
$b = ["01" => "a", "1" => "a"];
$r = array_uintersect_uassoc($a, $b, fn($v1,$v2)=>strcasecmp((string)$v1, (string)$v2), function($k1, $k2){ return (string)$k1 <=> (string)$k2; });
ksort($r);
echo implode('|', array_keys($r));
"#,
        ["1"]
    };

    array_uintersect_uassoc_value_callback_exception => {
        r#"<?php
$a = ["a" => "A", "b" => "B"];
$b = ["a" => "a"];
try {
    array_uintersect_uassoc(
        $a,
        $b,
        function($v1, $v2) {
            throw new RuntimeException('intersect-failed');
        },
        fn($k1, $k2) => strcmp((string)$k1, (string)$k2)
    );
    echo 'no-exception';
} catch (Throwable $e) {
    echo $e->getMessage();
}
"#,
        ["intersect-failed"]
    };

    array_uintersect_uassoc_key_callback_exception => {
        r#"<?php
$a = ["a" => "A"];
$b = ["a" => "A"];
try {
    array_uintersect_uassoc(
        $a,
        $b,
        fn($v1, $v2) => strcmp((string)$v1, (string)$v2),
        function($k1, $k2) {
            throw new RuntimeException('key-compare-failed');
        }
    );
    echo 'no-exception';
} catch (Throwable $e) {
    echo $e->getMessage();
}
"#,
        ["key-compare-failed"]
    };
}
