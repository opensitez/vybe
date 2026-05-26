use super::helpers::run_prints;

// ── usort ─────────────────────────────────────────────────────

#[test] fn usort_ascending_numeric() {
    assert_eq!(run_prints(r#"<?php
$a = [3,1,4,1,5,9,2,6];
usort($a, fn($x,$y) => $x <=> $y);
echo implode(',', $a);
"#), vec!["1,1,2,3,4,5,6,9"]);
}
#[test] fn usort_descending() {
    assert_eq!(run_prints(r#"<?php
$a = [5,2,8,1,9,3];
usort($a, fn($x,$y) => $y <=> $x);
echo implode(',', $a);
"#), vec!["9,8,5,3,2,1"]);
}
#[test] fn usort_by_string_length() {
    assert_eq!(run_prints(r#"<?php
$words = ['banana','fig','apple','kiwi'];
usort($words, fn($a,$b) => strlen($a) <=> strlen($b));
echo implode(',', $words);
"#), vec!["fig,kiwi,apple,banana"]);
}
#[test] fn usort_objects_by_property() {
    assert_eq!(run_prints(r#"<?php
class Item { public function __construct(public string $name, public int $price) {} }
$items = [new Item('c',30), new Item('a',10), new Item('b',20)];
usort($items, fn($a,$b) => $a->price <=> $b->price);
echo implode(',', array_map(fn($i) => $i->name, $items));
"#), vec!["a,b,c"]);
}

// ── uasort — preserves keys ───────────────────────────────────

#[test] fn uasort_preserves_keys() {
    assert_eq!(run_prints(r#"<?php
$a = ['b' => 2, 'a' => 1, 'c' => 3];
uasort($a, fn($x,$y) => $x <=> $y);
echo implode(',', array_keys($a));
"#), vec!["a,b,c"]);
}
#[test] fn uasort_values_correct_after_sort() {
    assert_eq!(run_prints(r#"<?php
$a = ['x' => 30, 'y' => 10, 'z' => 20];
uasort($a, fn($a,$b) => $a <=> $b);
echo implode(',', $a);
"#), vec!["10,20,30"]);
}

// ── uksort — sort by keys ─────────────────────────────────────

#[test] fn uksort_sorts_by_key() {
    assert_eq!(run_prints(r#"<?php
$a = ['banana' => 2, 'apple' => 1, 'cherry' => 3];
uksort($a, fn($a,$b) => strcmp($a,$b));
echo implode(',', array_keys($a));
"#), vec!["apple,banana,cherry"]);
}
#[test] fn uksort_by_key_length() {
    assert_eq!(run_prints(r#"<?php
$a = ['longest' => 1, 'mid' => 2, 'x' => 3];
uksort($a, fn($a,$b) => strlen($a) <=> strlen($b));
echo implode(',', array_keys($a));
"#), vec!["x,mid,longest"]);
}

// ── Stable sort PHP 8.0+ ──────────────────────────────────────

#[test] fn sort_stable_equal_elements_preserve_order() {
    assert_eq!(run_prints(r#"<?php
$items = [['n'=>'b','v'=>2],['n'=>'a','v'=>2],['n'=>'c','v'=>1]];
usort($items, fn($a,$b) => $a['v'] <=> $b['v']);
echo $items[0]['n'] . ',' . $items[1]['n'] . ',' . $items[2]['n'];
"#), vec!["c,b,a"]);
}

// ── array_multisort ───────────────────────────────────────────

#[test] fn array_multisort_primary_secondary() {
    assert_eq!(run_prints(r#"<?php
$a = [3,1,3,1,2];
$b = ['e','d','c','b','a'];
array_multisort($a, SORT_ASC, $b, SORT_ASC);
echo implode(',', $a) . '|' . implode(',', $b);
"#), vec!["1,1,2,3,3|b,d,a,c,e"]);
}

// ── arsort / krsort ───────────────────────────────────────────

#[test] fn arsort_preserves_keys_descending() {
    assert_eq!(run_prints(r#"<?php
$a = ['b'=>2,'a'=>1,'c'=>3];
arsort($a);
echo implode(',', array_keys($a));
"#), vec!["c,b,a"]);
}
#[test] fn krsort_sorts_keys_descending() {
    assert_eq!(run_prints(r#"<?php
$a = ['apple'=>1,'cherry'=>3,'banana'=>2];
krsort($a);
echo implode(',', array_keys($a));
"#), vec!["cherry,banana,apple"]);
}

// ── natsort / natcasesort ─────────────────────────────────────

#[test] fn natsort_natural_string_order() {
    assert_eq!(run_prints(r#"<?php
$files = ['file10.txt','file2.txt','file1.txt'];
natsort($files);
echo implode(',', $files);
"#), vec!["file1.txt,file2.txt,file10.txt"]);
}
#[test] fn natcasesort_case_insensitive() {
    assert_eq!(run_prints(r#"<?php
$a = ['IMG10','img2','IMG1'];
natcasesort($a);
echo implode(',', $a);
"#), vec!["IMG1,img2,IMG10"]);
}

// ── sort flags ────────────────────────────────────────────────

#[test] fn sort_flag_string() {
    assert_eq!(run_prints(r#"<?php
$a = ['10','9','100'];
sort($a, SORT_STRING);
echo implode(',', $a);
"#), vec!["10,100,9"]);
}
#[test] fn sort_flag_numeric() {
    assert_eq!(run_prints(r#"<?php
$a = ['10','9','100'];
sort($a, SORT_NUMERIC);
echo implode(',', $a);
"#), vec!["9,10,100"]);
}
#[test] fn sort_flag_natural() {
    assert_eq!(run_prints(r#"<?php
$a = ['item10','item2','item1'];
sort($a, SORT_NATURAL);
echo implode(',', $a);
"#), vec!["item1,item2,item10"]);
}

// ── array_column sort idiom ───────────────────────────────────

#[test] fn sort_by_column_using_array_column() {
    assert_eq!(run_prints(r#"<?php
$rows = [['name'=>'Charlie','age'=>30],['name'=>'Alice','age'=>25],['name'=>'Bob','age'=>28]];
$names = array_column($rows, 'name');
array_multisort($names, SORT_ASC, $rows);
echo implode(',', array_column($rows, 'name'));
"#), vec!["Alice,Bob,Charlie"]);
}
