use super::helpers::run_prints;

// ── array_column basics ───────────────────────────────────────

#[test]
fn array_column_extract_single() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [['id'=>1,'name'=>'Alice'],['id'=>2,'name'=>'Bob'],['id'=>3,'name'=>'Charlie']];
echo implode(',', array_column($rows, 'name'));
"#
        ),
        vec!["Alice,Bob,Charlie"]
    );
}
#[test]
fn array_column_with_index() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [['id'=>1,'name'=>'Alice'],['id'=>2,'name'=>'Bob']];
$indexed = array_column($rows, 'name', 'id');
echo $indexed[1] . ',' . $indexed[2];
"#
        ),
        vec!["Alice,Bob"]
    );
}
#[test]
fn array_column_null_value_returns_all() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [['id'=>1,'name'=>'A'],['id'=>2,'name'=>'B']];
$byId = array_column($rows, null, 'id');
echo $byId[2]['name'];
"#
        ),
        vec!["B"]
    );
}
#[test]
fn array_column_from_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
class User { public function __construct(public int $id, public string $name) {} }
$users = [new User(1,'Alice'), new User(2,'Bob')];
echo implode(',', array_column($users, 'name'));
"#
        ),
        vec!["Alice,Bob"]
    );
}
#[test]
fn array_column_build_lookup_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$products = [
    ['sku'=>'A001','price'=>9.99],
    ['sku'=>'B002','price'=>14.99],
    ['sku'=>'C003','price'=>4.99],
];
$prices = array_column($products, 'price', 'sku');
echo $prices['B002'];
"#
        ),
        vec!["14.99"]
    );
}

// ── array_combine patterns ────────────────────────────────────

#[test]
fn array_combine_zip_to_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$keys = ['a','b','c'];
$vals = [1,2,3];
$map = array_combine($keys, $vals);
echo $map['b'];
"#
        ),
        vec!["2"]
    );
}
#[test]
fn array_combine_transpose_headers() {
    assert_eq!(
        run_prints(
            r#"<?php
$headers = ['id','name','score'];
$row = [42,'Alice',98];
$record = array_combine($headers, $row);
echo $record['name'] . ':' . $record['score'];
"#
        ),
        vec!["Alice:98"]
    );
}

// ── array_chunk ───────────────────────────────────────────────

#[test]
fn array_chunk_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$chunks = array_chunk([1,2,3,4,5], 2);
echo count($chunks) . ':' . count($chunks[2]);
"#
        ),
        vec!["3:1"]
    );
}
#[test]
fn array_chunk_preserve_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3,'d'=>4];
$chunks = array_chunk($a, 2, true);
echo implode(',', array_keys($chunks[0]));
"#
        ),
        vec!["a,b"]
    );
}
#[test]
fn array_chunk_pagination() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = range(1, 10);
$page = 2;
$perPage = 3;
$pages = array_chunk($items, $perPage);
echo implode(',', $pages[$page - 1]);
"#
        ),
        vec!["4,5,6"]
    );
}

// ── array_splice ──────────────────────────────────────────────

#[test]
fn array_splice_remove() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1,2,3,4,5];
$removed = array_splice($a, 1, 2);
echo implode(',', $a) . '|' . implode(',', $removed);
"#
        ),
        vec!["1,4,5|2,3"]
    );
}
#[test]
fn array_splice_replace() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1,2,3,4,5];
array_splice($a, 2, 1, [10, 11]);
echo implode(',', $a);
"#
        ),
        vec!["1,2,10,11,4,5"]
    );
}
#[test]
fn array_splice_insert() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1,2,5];
array_splice($a, 2, 0, [3,4]);
echo implode(',', $a);
"#
        ),
        vec!["1,2,3,4,5"]
    );
}
#[test]
fn array_splice_negative_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a','b','c','d'];
array_splice($a, -2, 1, ['X']);
echo implode(',', $a);
"#
        ),
        vec!["a,b,X,d"]
    );
}

// ── array_push / array_pop / array_shift / array_unshift ──────

#[test]
fn array_push_multiple() {
    assert_eq!(
        run_prints(r#"<?php $a = [1]; array_push($a, 2, 3, 4); echo implode(',', $a); "#),
        vec!["1,2,3,4"]
    );
}
#[test]
fn array_unshift_multiple() {
    assert_eq!(
        run_prints(r#"<?php $a = [3,4]; array_unshift($a, 1, 2); echo implode(',', $a); "#),
        vec!["1,2,3,4"]
    );
}
#[test]
fn array_pop_removes_last() {
    assert_eq!(
        run_prints(r#"<?php $a = [1,2,3]; $v = array_pop($a); echo $v . ':' . implode(',', $a); "#),
        vec!["3:1,2"]
    );
}
#[test]
fn array_shift_removes_first() {
    assert_eq!(
        run_prints(
            r#"<?php $a = [1,2,3]; $v = array_shift($a); echo $v . ':' . implode(',', $a); "#
        ),
        vec!["1:2,3"]
    );
}

// ── Nested array operations ───────────────────────────────────

#[test]
fn nested_array_access_deep() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['a' => ['b' => ['c' => ['d' => 42]]]];
echo $data['a']['b']['c']['d'];
"#
        ),
        vec!["42"]
    );
}
#[test]
fn recursive_array_flatten() {
    assert_eq!(
        run_prints(
            r#"<?php
function flatten(array $arr): array {
    $result = [];
    array_walk_recursive($arr, function($v) use (&$result) { $result[] = $v; });
    return $result;
}
echo implode(',', flatten([[1,[2,3]],[4,[5,[6]]]]));
"#
        ),
        vec!["1,2,3,4,5,6"]
    );
}

#[test]
fn array_column_with_missing_key_is_omitted() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [
    ['id' => 1, 'name' => 'Alice'],
    ['id' => 2],
    ['id' => 3, 'name' => null],
];
$names = array_column($rows, 'name');
echo implode('|', $names);
"#,
        ),
        vec!["Alice||"]
    );
}

#[test]
fn array_column_with_non_string_index_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [
    ['id' => 0, 'title' => 'zero'],
    ['id' => 1, 'title' => 'one'],
];
$map = array_column($rows, 'title', 0);
echo $map[0] . '|' . $map[1];
"#,
        ),
        vec!["zero|one"]
    );
}

#[test]
fn array_combine_detects_length_mismatch() {
    assert_eq!(
        run_prints(
            r#"<?php
$keys = ['a', 'b', 'c'];
$vals = [1, 2];
try {
    array_combine($keys, $vals);
    echo 'no_error';
} catch (\ValueError $e) {
    echo 'error';
}
"#,
        ),
        vec!["error"]
    );
}

#[test]
fn array_splice_with_negative_length_keeps_all() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1,2,3,4,5];
array_splice($arr, 2, -1, [9,9]);
echo implode(',', $arr);
"#,
        ),
        vec!["1,2,9,9,5"]
    );
}

#[test]
fn array_push_pop_shift_unshift_full_cycle() {
    assert_eq!(
        run_prints(
            r#"<?php
$q = [2];
array_unshift($q, 1);
array_push($q, 3);
$x = array_shift($q);
$y = array_pop($q);
echo $x . $y . '|' . implode(',', $q);
"#,
        ),
        vec!["13|2"]
    );
}

#[test]
fn array_push_returns_new_length() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [1];
$len = array_push($items, 2, 3, 4);
echo $len . '|' . implode(',', $items);
"#,
        ),
        vec!["4|1,2,3,4"]
    );
}

#[test]
fn array_unshift_returns_new_length() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [3, 4];
$len = array_unshift($items, 1, 2);
echo $len . '|' . implode(',', $items);
"#,
        ),
        vec!["4|1,2,3,4"]
    );
}

#[test]
fn array_pop_returns_last_element() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [10, 20, 30];
$last = array_pop($items);
echo $last . '|' . implode(',', $items);
"#,
        ),
        vec!["30|10,20"]
    );
}

#[test]
fn array_shift_returns_first_element() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [10, 20, 30];
$first = array_shift($items);
echo $first . '|' . implode(',', $items);
"#,
        ),
        vec!["10|20,30"]
    );
}

#[test]
fn array_column_empty_input_returns_empty_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$names = array_column([], 'name');
echo is_array($names) ? 'array' : 'not-array';
echo '|';
echo count($names);
"#,
        ),
        vec!["array|0"]
    );
}

#[test]
fn array_column_with_empty_index_column_returns_values() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [
    ['id' => 'a', 'name' => 'Alice'],
    ['id' => 'b', 'name' => 'Bob'],
];
$vals = array_column($rows, 'name', null);
echo implode(',', $vals);
"#,
        ),
        vec!["Alice,Bob"]
    );
}

#[test]
fn array_splice_replaces_with_more_items_and_return_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [1, 2, 3, 4];
$removed = array_splice($items, 1, 1, [9, 9, 9]);
echo count($removed) . '|' . implode(',', $items);
"#,
        ),
        vec!["1|1,9,9,9,3,4"]
    );
}

#[test]
fn array_chunk_preserve_keys_false_reindexes_from_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
$chunks = array_chunk(['a' => 1, 'b' => 2, 'c' => 3, 'd' => 4], 3);
echo count($chunks) . '|' . array_key_first($chunks[0]) . '|' . implode(',', $chunks[1]);
"#,
        ),
        vec!["2|0|3,4"]
    );
}

#[test]
fn array_fill_keys_with_empty_source_is_empty_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$vals = array_fill_keys([], 'x');
echo is_array($vals) ? 'array' : 'non-array';
echo '|';
echo count($vals);
"#,
        ),
        vec!["array|0"]
    );
}

#[test]
fn array_key_exists_with_null_value_and_missing_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$row = ['a' => null];
echo array_key_exists('a', $row) ? 'a_yes' : 'a_no';
echo '|';
echo array_key_exists('b', $row) ? 'b_yes' : 'b_no';
"#,
        ),
        vec!["a_yes|b_no"]
    );
}

#[test]
fn array_search_with_offset_skips_earlier_matches() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = ['x', 'y', 'x', 'y', 'x'];
echo array_search('y', array_slice($items, 2));
"#,
        ),
        vec!["1"]
    );
}

#[test]
fn array_search_starts_at_zero_ignores_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = ['x', 'y', 'x', 'y'];
echo array_search('x', $items);
"#,
        ),
        vec!["0"]
    );
}

#[test]
fn array_flip_non_scalar_value_throws() {
    assert_eq!(
        run_prints(
            r#"<?php
$input = ['a' => ['nested'], 'b' => 'scalar'];
try {
    array_flip($input);
    echo 'ok';
} catch (Throwable $e) {
    echo 'error';
}
"#,
        ),
        vec!["error"]
    );
}

#[test]
fn array_reduce_short_circuit_identity_and_type() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = [1, 2, 3];
$total = array_reduce($parts, fn($carry, $item) => $carry . $item, '');
echo $total;
"#,
        ),
        vec!["123"]
    );
}
