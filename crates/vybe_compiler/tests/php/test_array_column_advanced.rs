use super::helpers::run_prints;

// ── array_column basics ───────────────────────────────────────

#[test] fn array_column_extract_single() {
    assert_eq!(run_prints(r#"<?php
$rows = [['id'=>1,'name'=>'Alice'],['id'=>2,'name'=>'Bob'],['id'=>3,'name'=>'Charlie']];
echo implode(',', array_column($rows, 'name'));
"#), vec!["Alice,Bob,Charlie"]);
}
#[test] fn array_column_with_index() {
    assert_eq!(run_prints(r#"<?php
$rows = [['id'=>1,'name'=>'Alice'],['id'=>2,'name'=>'Bob']];
$indexed = array_column($rows, 'name', 'id');
echo $indexed[1] . ',' . $indexed[2];
"#), vec!["Alice,Bob"]);
}
#[test] fn array_column_null_value_returns_all() {
    assert_eq!(run_prints(r#"<?php
$rows = [['id'=>1,'name'=>'A'],['id'=>2,'name'=>'B']];
$byId = array_column($rows, null, 'id');
echo $byId[2]['name'];
"#), vec!["B"]);
}
#[test] fn array_column_from_objects() {
    assert_eq!(run_prints(r#"<?php
class User { public function __construct(public int $id, public string $name) {} }
$users = [new User(1,'Alice'), new User(2,'Bob')];
echo implode(',', array_column($users, 'name'));
"#), vec!["Alice,Bob"]);
}
#[test] fn array_column_build_lookup_map() {
    assert_eq!(run_prints(r#"<?php
$products = [
    ['sku'=>'A001','price'=>9.99],
    ['sku'=>'B002','price'=>14.99],
    ['sku'=>'C003','price'=>4.99],
];
$prices = array_column($products, 'price', 'sku');
echo $prices['B002'];
"#), vec!["14.99"]);
}

// ── array_combine patterns ────────────────────────────────────

#[test] fn array_combine_zip_to_map() {
    assert_eq!(run_prints(r#"<?php
$keys = ['a','b','c'];
$vals = [1,2,3];
$map = array_combine($keys, $vals);
echo $map['b'];
"#), vec!["2"]);
}
#[test] fn array_combine_transpose_headers() {
    assert_eq!(run_prints(r#"<?php
$headers = ['id','name','score'];
$row = [42,'Alice',98];
$record = array_combine($headers, $row);
echo $record['name'] . ':' . $record['score'];
"#), vec!["Alice:98"]);
}

// ── array_chunk ───────────────────────────────────────────────

#[test] fn array_chunk_basic() {
    assert_eq!(run_prints(r#"<?php
$chunks = array_chunk([1,2,3,4,5], 2);
echo count($chunks) . ':' . count($chunks[2]);
"#), vec!["3:1"]);
}
#[test] fn array_chunk_preserve_keys() {
    assert_eq!(run_prints(r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3,'d'=>4];
$chunks = array_chunk($a, 2, true);
echo implode(',', array_keys($chunks[0]));
"#), vec!["a,b"]);
}
#[test] fn array_chunk_pagination() {
    assert_eq!(run_prints(r#"<?php
$items = range(1, 10);
$page = 2;
$perPage = 3;
$pages = array_chunk($items, $perPage);
echo implode(',', $pages[$page - 1]);
"#), vec!["4,5,6"]);
}

// ── array_splice ──────────────────────────────────────────────

#[test] fn array_splice_remove() {
    assert_eq!(run_prints(r#"<?php
$a = [1,2,3,4,5];
$removed = array_splice($a, 1, 2);
echo implode(',', $a) . '|' . implode(',', $removed);
"#), vec!["1,4,5|2,3"]);
}
#[test] fn array_splice_replace() {
    assert_eq!(run_prints(r#"<?php
$a = [1,2,3,4,5];
array_splice($a, 2, 1, [10, 11]);
echo implode(',', $a);
"#), vec!["1,2,10,11,4,5"]);
}
#[test] fn array_splice_insert() {
    assert_eq!(run_prints(r#"<?php
$a = [1,2,5];
array_splice($a, 2, 0, [3,4]);
echo implode(',', $a);
"#), vec!["1,2,3,4,5"]);
}
#[test] fn array_splice_negative_offset() {
    assert_eq!(run_prints(r#"<?php
$a = ['a','b','c','d'];
array_splice($a, -2, 1, ['X']);
echo implode(',', $a);
"#), vec!["a,b,X,d"]);
}

// ── array_push / array_pop / array_shift / array_unshift ──────

#[test] fn array_push_multiple() {
    assert_eq!(run_prints(r#"<?php $a = [1]; array_push($a, 2, 3, 4); echo implode(',', $a); "#), vec!["1,2,3,4"]);
}
#[test] fn array_unshift_multiple() {
    assert_eq!(run_prints(r#"<?php $a = [3,4]; array_unshift($a, 1, 2); echo implode(',', $a); "#), vec!["1,2,3,4"]);
}
#[test] fn array_pop_removes_last() {
    assert_eq!(run_prints(r#"<?php $a = [1,2,3]; $v = array_pop($a); echo $v . ':' . implode(',', $a); "#), vec!["3:1,2"]);
}
#[test] fn array_shift_removes_first() {
    assert_eq!(run_prints(r#"<?php $a = [1,2,3]; $v = array_shift($a); echo $v . ':' . implode(',', $a); "#), vec!["1:2,3"]);
}

// ── Nested array operations ───────────────────────────────────

#[test] fn nested_array_access_deep() {
    assert_eq!(run_prints(r#"<?php
$data = ['a' => ['b' => ['c' => ['d' => 42]]]];
echo $data['a']['b']['c']['d'];
"#), vec!["42"]);
}
#[test] fn recursive_array_flatten() {
    assert_eq!(run_prints(r#"<?php
function flatten(array $arr): array {
    $result = [];
    array_walk_recursive($arr, function($v) use (&$result) { $result[] = $v; });
    return $result;
}
echo implode(',', flatten([[1,[2,3]],[4,[5,[6]]]]));
"#), vec!["1,2,3,4,5,6"]);
}
