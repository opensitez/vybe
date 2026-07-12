//! `usort`, `uasort`, `uksort`, and `array_multisort` with callbacks.

crate::php_cases! {
    usort_orders_integers_ascending => {
        r#"<?php
$a = [3, 1, 2];
usort($a, fn($x, $y) => $x <=> $y);
echo implode(',', $a);
"#,
        ["1,2,3"]
    };

    usort_orders_strings_descending => {
        r#"<?php
$a = ['b', 'a', 'c'];
usort($a, fn($x, $y) => $y <=> $x);
echo implode(',', $a);
"#,
        ["c,b,a"]
    };

    uasort_preserves_keys => {
        r#"<?php
$a = ['x' => 3, 'y' => 1];
uasort($a, fn($a, $b) => $a <=> $b);
echo implode(',', array_keys($a));
"#,
        ["y,x"]
    };

    uksort_sorts_by_key_string => {
        r#"<?php
$a = ['b' => 1, 'a' => 2];
uksort($a, fn($ka, $kb) => $ka <=> $kb);
echo implode(',', array_keys($a));
"#,
        ["a,b"]
    };

    array_multisort_sorts_parallel_arrays => {
        r#"<?php
$nums = [3, 1, 2];
$labels = ['c', 'a', 'b'];
array_multisort($nums, $labels);
echo implode('-', $nums) . ':' . implode('-', $labels);
"#,
        ["1-2-3:a-b-c"]
    };

    usort_stable_like_behavior_with_tiebreaker => {
        r#"<?php
$rows = [['k' => 2, 'n' => 'b'], ['k' => 1, 'n' => 'a'], ['k' => 2, 'n' => 'c']];
usort($rows, function ($a, $b) {
    return $a['k'] <=> $b['k'] ?: $a['n'] <=> $b['n'];
});
echo $rows[0]['n'] . $rows[2]['n'];
"#,
        ["ac"]
    };

    sort_default_ascending => {
        r#"<?php
$a = [3, 1, 2];
sort($a);
echo implode(',', $a);
"#,
        ["1,2,3"]
    };

    rsort_descending => {
        r#"<?php
$a = [3, 1, 2];
rsort($a);
echo implode(',', $a);
"#,
        ["3,2,1"]
    };

    asort_maintains_key_value_assoc => {
        r#"<?php
$a = ['b' => 2, 'a' => 1];
asort($a);
echo implode(',', array_keys($a));
"#,
        ["a,b"]
    };

    ksort_sorts_keys => {
        r#"<?php
$a = ['z' => 1, 'a' => 2];
ksort($a);
echo implode(',', array_keys($a));
"#,
        ["a,z"]
    };

    natsort_natural_order => {
        r#"<?php
$a = ['img12', 'img2', 'img1'];
natsort($a);
echo implode(',', array_values($a));
"#,
        ["img1,img2,img12"]
    };

    usort_with_spaceship_on_objects => {
        r#"<?php
class Box { public function __construct(public int $v) {} }
$a = [new Box(2), new Box(1)];
usort($a, fn($x, $y) => $x->v <=> $y->v);
echo $a[0]->v;
"#,
        ["1"]
    };
}
