//! Pass-by-reference, reference assignment, and `foreach` by-ref — runtime output.

crate::php_cases! {
    pass_by_reference_increments_caller_variable => {
        r#"<?php
function inc(int &$n): void { $n++; }
$x = 5;
inc($x);
echo $x;
"#,
        ["6"]
    };

    pass_by_reference_swap_exchanges_values => {
        r#"<?php
function swap(string &$a, string &$b): void { [$a, $b] = [$b, $a]; }
$x = 'hi'; $y = 'bye';
swap($x, $y);
echo $x . ':' . $y;
"#,
        ["bye:hi"]
    };

    pass_by_reference_array_appends_in_place => {
        r#"<?php
function push(array &$a, mixed $v): void { $a[] = $v; }
$list = [1];
push($list, 2);
echo implode(',', $list);
"#,
        ["1,2"]
    };

    foreach_by_reference_doubles_values => {
        r#"<?php
$a = [1, 2, 3];
foreach ($a as &$v) { $v *= 2; }
unset($v);
echo implode(',', $a);
"#,
        ["2,4,6"]
    };

    reference_assignment_aliases_array => {
        r#"<?php
$orig = [1];
$alias = &$orig;
$alias[0] = 9;
echo $orig[0];
"#,
        ["9"]
    };

    reference_to_reference_chain => {
        r#"<?php
$x = 1;
$a = &$x;
$b = &$a;
$b = 7;
echo $x;
"#,
        ["7"]
    };

    array_element_reference_update => {
        r#"<?php
$a = [10, 20];
$r = &$a[1];
$r = 99;
echo $a[1];
"#,
        ["99"]
    };

    unset_reference_leaves_other_alias => {
        r#"<?php
$x = 1;
$y = &$x;
unset($y);
$x = 3;
echo $x;
"#,
        ["3"]
    };

    function_returns_reference_to_static => {
        r#"<?php
function &counter(): int {
    static $n = 0;
    $n++;
    return $n;
}
counter();
echo counter();
"#,
        ["2"]
    };

    array_walk_by_reference_modifies_original => {
        r#"<?php
$a = ['a' => 1, 'b' => 2];
array_walk($a, function (&$v, $k) { $v = $k . $v; });
echo $a['a'] . $a['b'];
"#,
        ["a1b2"]
    };
}
