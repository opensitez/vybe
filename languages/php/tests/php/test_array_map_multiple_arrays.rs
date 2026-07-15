use super::helpers::run_prints;

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
}
