use super::helpers::run_prints;

// ── usort ─────────────────────────────────────────────────────

#[test]
fn usort_ascending_numeric() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [3,1,4,1,5,9,2,6];
usort($a, fn($x,$y) => $x <=> $y);
echo implode(',', $a);
"#
        ),
        vec!["1,1,2,3,4,5,6,9"]
    );
}
#[test]
fn usort_descending() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [5,2,8,1,9,3];
usort($a, fn($x,$y) => $y <=> $x);
echo implode(',', $a);
"#
        ),
        vec!["9,8,5,3,2,1"]
    );
}
#[test]
fn usort_by_string_length() {
    assert_eq!(
        run_prints(
            r#"<?php
$words = ['banana','fig','apple','kiwi'];
usort($words, fn($a,$b) => strlen($a) <=> strlen($b));
echo implode(',', $words);
"#
        ),
        vec!["fig,kiwi,apple,banana"]
    );
}
#[test]
fn usort_objects_by_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Item { public function __construct(public string $name, public int $price) {} }
$items = [new Item('c',30), new Item('a',10), new Item('b',20)];
usort($items, fn($a,$b) => $a->price <=> $b->price);
echo implode(',', array_map(fn($i) => $i->name, $items));
"#
        ),
        vec!["a,b,c"]
    );
}

// ── uasort — preserves keys ───────────────────────────────────

#[test]
fn uasort_preserves_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['b' => 2, 'a' => 1, 'c' => 3];
uasort($a, fn($x,$y) => $x <=> $y);
echo implode(',', array_keys($a));
"#
        ),
        vec!["a,b,c"]
    );
}
#[test]
fn uasort_values_correct_after_sort() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x' => 30, 'y' => 10, 'z' => 20];
uasort($a, fn($a,$b) => $a <=> $b);
echo implode(',', $a);
"#
        ),
        vec!["10,20,30"]
    );
}

// ── uksort — sort by keys ─────────────────────────────────────

#[test]
fn uksort_sorts_by_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['banana' => 2, 'apple' => 1, 'cherry' => 3];
uksort($a, fn($a,$b) => strcmp($a,$b));
echo implode(',', array_keys($a));
"#
        ),
        vec!["apple,banana,cherry"]
    );
}
#[test]
fn uksort_by_key_length() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['longest' => 1, 'mid' => 2, 'x' => 3];
uksort($a, fn($a,$b) => strlen($a) <=> strlen($b));
echo implode(',', array_keys($a));
"#
        ),
        vec!["x,mid,longest"]
    );
}

// ── Stable sort PHP 8.0+ ──────────────────────────────────────

#[test]
fn sort_stable_equal_elements_preserve_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [['n'=>'b','v'=>2],['n'=>'a','v'=>2],['n'=>'c','v'=>1]];
usort($items, fn($a,$b) => $a['v'] <=> $b['v']);
echo $items[0]['n'] . ',' . $items[1]['n'] . ',' . $items[2]['n'];
"#
        ),
        vec!["c,b,a"]
    );
}

// ── array_multisort ───────────────────────────────────────────

#[test]
fn array_multisort_primary_secondary() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [3,1,3,1,2];
$b = ['e','d','c','b','a'];
array_multisort($a, SORT_ASC, $b, SORT_ASC);
echo implode(',', $a) . '|' . implode(',', $b);
"#
        ),
        vec!["1,1,2,3,3|b,d,a,c,e"]
    );
}

// ── arsort / krsort ───────────────────────────────────────────

#[test]
fn arsort_preserves_keys_descending() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['b'=>2,'a'=>1,'c'=>3];
arsort($a);
echo implode(',', array_keys($a));
"#
        ),
        vec!["c,b,a"]
    );
}
#[test]
fn krsort_sorts_keys_descending() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['apple'=>1,'cherry'=>3,'banana'=>2];
krsort($a);
echo implode(',', array_keys($a));
"#
        ),
        vec!["cherry,banana,apple"]
    );
}

// ── natsort / natcasesort ─────────────────────────────────────

#[test]
fn natsort_natural_string_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$files = ['file10.txt','file2.txt','file1.txt'];
natsort($files);
echo implode(',', $files);
"#
        ),
        vec!["file1.txt,file2.txt,file10.txt"]
    );
}
#[test]
fn natcasesort_case_insensitive() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['IMG10','img2','IMG1'];
natcasesort($a);
echo implode(',', $a);
"#
        ),
        vec!["IMG1,img2,IMG10"]
    );
}

// ── sort flags ────────────────────────────────────────────────

#[test]
fn sort_flag_string() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['10','9','100'];
sort($a, SORT_STRING);
echo implode(',', $a);
"#
        ),
        vec!["10,100,9"]
    );
}
#[test]
fn sort_flag_numeric() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['10','9','100'];
sort($a, SORT_NUMERIC);
echo implode(',', $a);
"#
        ),
        vec!["9,10,100"]
    );
}
#[test]
fn sort_flag_natural() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['item10','item2','item1'];
sort($a, SORT_NATURAL);
echo implode(',', $a);
"#
        ),
        vec!["item1,item2,item10"]
    );
}

// ── array_column sort idiom ───────────────────────────────────

#[test]
fn sort_by_column_using_array_column() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [['name'=>'Charlie','age'=>30],['name'=>'Alice','age'=>25],['name'=>'Bob','age'=>28]];
$names = array_column($rows, 'name');
array_multisort($names, SORT_ASC, $rows);
echo implode(',', array_column($rows, 'name'));
"#
        ),
        vec!["Alice,Bob,Charlie"]
    );
}

#[test]
fn usort_stable_behavior_with_equal_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [
    ['id' => 1, 'group' => 'a'],
    ['id' => 2, 'group' => 'a'],
    ['id' => 3, 'group' => 'b'],
];
usort($rows, fn($x, $y) => strcmp($x['group'], $y['group']));
echo $rows[0]['id'] . ':' . $rows[1]['id'] . ':' . $rows[2]['id'];
"#,
        ),
        vec!["1:2:3"]
    );
}

#[test]
fn array_multisort_secondary_sort_by_name() {
    assert_eq!(
        run_prints(
            r#"<?php
$scores = [10, 10, 10, 5];
$names = ['d','a','c','b'];
array_multisort($scores, SORT_DESC, $names, SORT_ASC);
echo implode(',', $scores) . '|' . implode(',', $names);
"#,
        ),
        vec!["10,10,10,5|a,c,d,b"]
    );
}

#[test]
fn ksort_with_numeric_string_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['2' => 'b', 10 => 'a', '1' => 'c'];
ksort($a, SORT_STRING);
echo implode('|', array_keys($a));
"#,
        ),
        vec!["1|10|2"]
    );
}

#[test]
fn sort_regular_keeps_numeric_string_ordering() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['2', 1, '10', 3];
sort($a, SORT_REGULAR);
echo implode(',', $a);
"#,
        ),
        vec!["1,2,3,10"]
    );
}

#[test]
fn natsort_with_hyphenated_strings() {
    assert_eq!(
        run_prints(
            r#"<?php
$labels = ['item-2a', 'item-10b', 'item-1c'];
natsort($labels);
echo implode(',', $labels);
"#,
        ),
        vec!["item-1c,item-2a,item-10b"]
    );
}

#[test]
fn sort_empty_array_stays_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [];
sort($a);
echo count($a);
echo '|';
echo is_array($a) ? 'array' : 'no';
"#,
        ),
        vec!["0|array"]
    );
}

#[test]
fn asort_with_float_keys_and_values_preserves_pairs() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['2' => 'b', '10' => 'a', '1' => 'c'];
asort($a, SORT_STRING);
echo implode(',', array_keys($a)) . '|' . implode(',', $a);
"#,
        ),
        vec!["10,2,1|a,b,c"]
    );
}

#[test]
fn rsort_regulates_reverse_numeric() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['2', 10, 1];
rsort($a, SORT_NUMERIC);
echo implode(',', $a);
"#,
        ),
        vec!["10,2,1"]
    );
}

#[test]
fn usort_string_compare_ascii_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['z', 'a', 'b'];
usort($a, fn($x, $y) => strcmp($x, $y));
echo implode(',', $a);
"#,
        ),
        vec!["a,b,z"]
    );
}

#[test]
fn arsort_with_ties_still_stable() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['first' => 1, 'second' => 1, 'third' => 2];
arsort($a);
echo array_key_first($a) . '|' . array_key_last($a);
"#,
        ),
        vec!["third|second"]
    );
}

#[test]
fn usort_desc_callback_uses_callback() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [3,1,2];
usort($a, fn($x,$y) => $y <=> $x);
echo implode(',', $a);
"#,
        ),
        vec!["3,2,1"]
    );
}

#[test]
fn array_multisort_with_same_values_and_payload_reordered_by_secondary() {
    assert_eq!(
        run_prints(
            r#"<?php
$scores = [1,1,1,2];
$labels = ['d','a','c','b'];
array_multisort($scores, SORT_ASC, SORT_NUMERIC, $labels, SORT_ASC, SORT_STRING);
echo implode(',', $labels);
"#,
        ),
        vec!["a,c,d,b"]
    );
}
