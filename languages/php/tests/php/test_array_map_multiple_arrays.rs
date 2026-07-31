crate::php_cases! {
    array_map_multiple_arrays => {
        r#"<?php
$a = [1, 2, 3];
$b = ['uno', 'dos', 'tres'];
$c = ['I', 'II', 'III'];

$res = array_map(function($x, $y, $z) {
    return "$x-$y-$z";
}, $a, $b, $c);

echo implode('|', $res);
"#,
        ["1-uno-I|2-dos-II|3-tres-III"]
    };

    array_map_multiple_arrays_different_lengths => {
        r#"<?php
$a = [1, 2];
$b = ['A', 'B', 'C'];

$res = array_map(null, $a, $b);
echo count($res) . "|";
echo ($res[2][0] ?? 'null') . "-" . $res[2][1];
"#,
        ["3|null-C"]
    };

    array_map_multiple_arrays_null_callback_zip => {
        r#"<?php
$nums = [10, 20, 30];
$labels = ['x', 'y', 'z'];
$zipped = array_map(null, $nums, $labels);

echo count($zipped[0]) . ":" . $zipped[0][0] . ":" . $zipped[0][1];
echo "|" . ($zipped[1][0] ?? 'null') . ":" . ($zipped[1][1] ?? 'null');
"#,
        ["2:10:x|20:y"]
    };

    array_map_multiple_arrays_unequal_lengths_with_callback => {
        r#"<?php
$nums = [1, 2, 3, 4];
$tags = ['a', 'b'];
$flags = [true];

$res = array_map(fn($n, $tag, $flag) => "$n:$tag:" . ($flag ? '1' : '0'), $nums, $tags, $flags);
echo implode('|', $res);
"#,
        ["1:a:1|2:b:0|3::0|4::0"]
    };

    array_map_multiple_arrays_returns_nested_arrays => {
        r#"<?php
$a = [1, 2];
$b = ['u', 'v'];
$res = array_map(fn($x, $y) => ['v' => $x, 'k' => $y], $a, $b);
echo json_encode($res[0]) . "|" . json_encode($res[1]);
"#,
        ["{\"v\":1,\"k\":\"u\"}|{\"v\":2,\"k\":\"v\"}"]
    };

    array_map_multiple_arrays_preserves_longest_length => {
        r#"<?php
$a = [10, 20, 30];
$b = ['a' => 'x', 'b' => 'y'];
$res = array_map(function($n, $s) { return $n . ':' . ($s ?? 'miss'); }, $a, $b);
echo implode('|', $res) . '|' . json_encode(array_keys($res));
"#,
        ["10:x|20:y|30:miss|[0,1,2]"]
    };

    array_map_multiple_arrays_empty_and_nulls => {
        r#"<?php
$a = [1, 2];
$b = [];
$res = array_map(null, $a, $b);
echo count($res) . '|' . ($res[1][0] ?? 'z') . '|' . ($res[1][1] ?? 'z');
"#,
        ["2|2|z"]
    };

    array_map_multiple_arrays_empty_first_array => {
        r#"<?php
$a = [];
$b = [1, 2, 3];
$res = array_map(fn($x, $y) => "$x-$y", $a, $b);
echo json_encode($res);
"#,
        // array_map walks the LONGEST input and pads the short ones with null,
        // so an empty FIRST array does not truncate the result — same rule the
        // unequal-lengths case above asserts.
        ["[\"-1\",\"-2\",\"-3\"]"]
    };

    array_map_multiple_arrays_string_like_numeric_keys => {
        r#"<?php
$a = ['first' => 1, 'second' => 3];
$b = ['a', 'b', 'c', 'd'];
$res = array_map(fn($x, $y) => "$x:$y", $a, $b);
echo count($res) . '|' . $res[0] . '|' . ($res[1] ?? 'null') . '|' . ($res[2] ?? 'null');
"#,
        ["3|1:a|3:b|:c"]
    };

    array_map_multiple_arrays_with_four_arrays_padding => {
        r#"<?php
$a = [1, 2, 3];
$b = ['x', 'y', 'z', 'w'];
$c = [10];
$d = [true, false];
$res = array_map(fn($n, $s, $i, $f) => $n . $s . ':' . $i . ':' . ($f ? '1' : '0'), $a, $b, $c, $d);
echo implode('|', $res);
"#,
        ["1x:10:1|2y::0|3z::0"]
    };

    array_map_multiple_arrays_nested_results_count => {
        r#"<?php
$nums = [1, 2, 3, 4];
$labels = ['a', 'b'];
$zipped = array_map(null, $nums, $labels);
echo count($zipped) . '|' . count($zipped[2]);
"#,
        ["4|2"]
    };

    array_map_multiple_arrays_callable_string_notation => {
        r#"<?php
function pair_label($num, $tag) {
    return $num . ':' . $tag;
}
$numbers = [5, 6, 7];
$tags = ['u', 'v', 'w'];
$res = array_map('pair_label', $numbers, $tags);
echo implode('|', $res);
"#,
        ["5:u|6:v|7:w"]
    };
}
