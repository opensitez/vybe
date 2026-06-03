use super::helpers::run_prints;

// ── array_splice ──────────────────────────────────────────────────
#[test]
fn array_splice_remove() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3, 4, 5];
$removed = array_splice($a, 1, 2);
echo implode(",", $a);
echo implode(",", $removed);
"#
        ),
        &["1,4,5", "2,3"]
    );
}

#[test]
fn array_splice_insert() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 5, 6];
array_splice($a, 2, 0, [3, 4]);
echo implode(",", $a);
"#
        ),
        &["1,2,3,4,5,6"]
    );
}

#[test]
fn array_splice_replace() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a", "b", "c", "d"];
array_splice($a, 1, 2, ["x", "y", "z"]);
echo implode(",", $a);
"#
        ),
        &["a,x,y,z,d"]
    );
}

// ── array_chunk ───────────────────────────────────────────────────
#[test]
fn array_chunk_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3, 4, 5];
$chunks = array_chunk($a, 2);
foreach ($chunks as $chunk) {
    echo implode(",", $chunk);
}
"#
        ),
        &["1,2", "3,4", "5"]
    );
}

#[test]
fn array_chunk_preserve_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3, "d" => 4];
$chunks = array_chunk($a, 3, true);
echo count($chunks);
echo implode(",", array_keys($chunks[0]));
"#
        ),
        &["2", "a,b,c"]
    );
}

// ── array_column ──────────────────────────────────────────────────
#[test]
fn array_column_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$records = [
    ["id" => 1, "name" => "Alice"],
    ["id" => 2, "name" => "Bob"],
    ["id" => 3, "name" => "Charlie"],
];
$names = array_column($records, "name");
echo implode(",", $names);
"#
        ),
        &["Alice,Bob,Charlie"]
    );
}

#[test]
fn array_column_with_index() {
    assert_eq!(
        run_prints(
            r#"<?php
$records = [
    ["id" => 10, "name" => "Alice"],
    ["id" => 20, "name" => "Bob"],
];
$indexed = array_column($records, "name", "id");
echo $indexed[10];
echo $indexed[20];
"#
        ),
        &["Alice", "Bob"]
    );
}

// ── array_unique / array_flip / array_combine ─────────────────────
#[test]
fn array_unique_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 2, 3, 3, 3, 4];
$u = array_unique($a);
echo implode(",", $u);
"#
        ),
        &["1,2,3,4"]
    );
}

#[test]
fn array_flip_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3];
$f = array_flip($a);
echo $f[1];
echo $f[2];
echo $f[3];
"#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn array_combine_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$keys = ["name", "age", "city"];
$vals = ["Alice", 30, "NYC"];
$combined = array_combine($keys, $vals);
echo $combined["name"];
echo $combined["age"];
echo $combined["city"];
"#
        ),
        &["Alice", "30", "NYC"]
    );
}

// ── array_diff / array_intersect ─────────────────────────────────
#[test]
fn array_diff_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3, 4, 5];
$b = [3, 4, 5, 6, 7];
$diff = array_diff($a, $b);
echo implode(",", $diff);
"#
        ),
        &["1,2"]
    );
}

#[test]
fn array_intersect_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3, 4, 5];
$b = [3, 4, 5, 6, 7];
$common = array_intersect($a, $b);
echo implode(",", $common);
"#
        ),
        &["3,4,5"]
    );
}

#[test]
fn array_diff_assoc() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3];
$b = ["a" => 1, "b" => 99, "d" => 4];
$diff = array_diff_assoc($a, $b);
echo implode(",", array_keys($diff));
"#
        ),
        &["b,c"]
    );
}

#[test]
fn array_intersect_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3];
$b = ["a" => 10, "c" => 30, "d" => 40];
$result = array_intersect_key($a, $b);
echo implode(",", array_keys($result));
echo implode(",", $result);
"#
        ),
        &["a,c", "1,3"]
    );
}

#[test]
fn array_diff_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3];
$b = ["a" => 10, "c" => 30, "d" => 40];
$result = array_diff_key($a, $b);
echo implode(",", array_keys($result));
echo implode(",", $result);
"#
        ),
        &["b", "2"]
    );
}

// ── array_fill / array_pad / array_product ───────────────────────
#[test]
fn array_fill_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_fill(0, 5, "x");
echo implode(",", $a);
echo count($a);
"#
        ),
        &["x,x,x,x,x", "5"]
    );
}

#[test]
fn array_pad_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3];
$padded = array_pad($a, 6, 0);
echo implode(",", $padded);
$left = array_pad($a, -6, 0);
echo implode(",", $left);
"#
        ),
        &["1,2,3,0,0,0", "0,0,0,1,2,3"]
    );
}

#[test]
fn array_product_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [2, 3, 4];
echo array_product($a);
"#
        ),
        &["24"]
    );
}

#[test]
fn array_count_values() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["apple", "banana", "apple", "cherry", "banana", "apple"];
$counts = array_count_values($a);
echo $counts["apple"];
echo $counts["banana"];
echo $counts["cherry"];
"#
        ),
        &["3", "2", "1"]
    );
}

// ── Sorting variants ─────────────────────────────────────────────
#[test]
fn rsort_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [3, 1, 4, 1, 5];
rsort($a);
echo implode(",", $a);
"#
        ),
        &["5,4,3,1,1"]
    );
}

#[test]
fn asort_preserves_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["c" => 3, "a" => 1, "b" => 2];
asort($a);
echo implode(",", array_keys($a));
echo implode(",", $a);
"#
        ),
        &["a,b,c", "1,2,3"]
    );
}

#[test]
fn ksort_by_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["banana" => 2, "apple" => 1, "cherry" => 3];
ksort($a);
echo implode(",", array_keys($a));
"#
        ),
        &["apple,banana,cherry"]
    );
}

#[test]
fn krsort_reverse_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "c" => 3, "b" => 2];
krsort($a);
echo implode(",", array_keys($a));
"#
        ),
        &["c,b,a"]
    );
}

#[test]
fn uasort_custom() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["x" => 3, "y" => 1, "z" => 2];
uasort($a, fn($a, $b) => $a - $b);
echo implode(",", $a);
echo implode(",", array_keys($a));
"#
        ),
        &["1,2,3", "y,z,x"]
    );
}

// ── array_walk_recursive ─────────────────────────────────────────
#[test]
fn array_walk_recursive_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, [2, 3], [4, [5]]];
$sum = 0;
array_walk_recursive($a, function($val) use (&$sum) {
    $sum += $val;
});
echo $sum;
"#
        ),
        &["15"]
    );
}

// ── array_multisort ──────────────────────────────────────────────
#[test]
fn array_replace_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ["a" => 1, "b" => 2, "c" => 3];
$replace = ["b" => 20, "d" => 40];
$result = array_replace($base, $replace);
echo $result["a"];
echo $result["b"];
echo $result["c"];
echo $result["d"];
"#
        ),
        &["1", "20", "3", "40"]
    );
}

// ── Nested array operations ──────────────────────────────────────
#[test]
fn nested_array_access() {
    assert_eq!(
        run_prints(
            r#"<?php
$matrix = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
echo $matrix[0][0];
echo $matrix[1][1];
echo $matrix[2][2];
"#
        ),
        &["1", "5", "9"]
    );
}

#[test]
fn array_map_with_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$prices = ["apple" => 1.50, "banana" => 0.75, "cherry" => 2.00];
$doubled = array_map(fn($p) => $p * 2, $prices);
echo $doubled["apple"];
echo $doubled["banana"];
echo $doubled["cherry"];
"#
        ),
        &["3", "1.5", "4"]
    );
}

#[test]
fn array_filter_with_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["" => 1, "hello" => 2, "world" => 3, "" => 4];
$filtered = array_filter($a, fn($key) => strlen($key) > 0, ARRAY_FILTER_USE_KEY);
echo count($filtered);
"#
        ),
        &["2"]
    );
}

#[test]
fn compact_and_extract() {
    assert_eq!(
        run_prints(
            r#"<?php
$name = "Alice";
$age = 30;
$city = "NYC";
$data = compact("name", "age", "city");
echo $data["name"];
echo $data["age"];
extract(["x" => 100, "y" => 200]);
echo $x + $y;
"#
        ),
        &["Alice", "30", "300"]
    );
}

#[test]
fn array_reduce_complex() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [
    ["name" => "apple", "price" => 1.5, "qty" => 3],
    ["name" => "banana", "price" => 0.75, "qty" => 6],
    ["name" => "cherry", "price" => 2.0, "qty" => 2],
];
$total = array_reduce($items, function($carry, $item) {
    return $carry + ($item["price"] * $item["qty"]);
}, 0);
echo $total;
"#
        ),
        &["13"]
    );
}

#[test]
fn array_key_first_last() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["x" => 10, "y" => 20, "z" => 30];
echo array_key_first($a);
echo array_key_last($a);
"#
        ),
        &["x", "z"]
    );
}

#[test]
fn list_nested_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
$coords = [[1, 2], [3, 4], [5, 6]];
foreach ($coords as [$x, $y]) {
    echo $x + $y;
}
"#
        ),
        &["3", "7", "11"]
    );
}

#[test]
fn list_with_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$person = ["name" => "Alice", "age" => 30];
["name" => $name, "age" => $age] = $person;
echo $name;
echo $age;
"#
        ),
        &["Alice", "30"]
    );
}

// ── array_fill_keys ──────────────────────────────────────────────
#[test]
fn array_fill_keys_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$keys = ["foo", "bar", "baz"];
$a = array_fill_keys($keys, 0);
echo $a["foo"];
echo $a["bar"];
echo $a["baz"];
echo count($a);
"#
        ),
        &["0", "0", "0", "3"]
    );
}

// ── array_diff_assoc (extended) ──────────────────────────────────
#[test]
fn array_intersect_assoc_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3];
$b = ["a" => 1, "b" => 99, "c" => 3];
$result = array_intersect_assoc($a, $b);
echo implode(",", array_keys($result));
echo implode(",", $result);
"#
        ),
        &["a,c", "1,3"]
    );
}

// ── in_array strict mode ─────────────────────────────────────────
#[test]
fn in_array_strict_type_check() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3, "4", "5"];
echo in_array("1", $a, true)  ? "yes" : "no";
echo in_array(1,   $a, true)  ? "yes" : "no";
echo in_array("4", $a, true)  ? "yes" : "no";
"#
        ),
        &["no", "yes", "yes"]
    );
}

// ── array_search strict ──────────────────────────────────────────
#[test]
fn array_search_strict_mode() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["10", "20", "30"];
$k = array_search(20, $a, true);
echo ($k === false) ? "not found" : $k;
$k2 = array_search("20", $a, true);
echo $k2;
"#
        ),
        &["not found", "1"]
    );
}

// ── range with float step ────────────────────────────────────────
#[test]
fn range_float_step() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = range(0, 1, 0.25);
echo count($r);
echo $r[0];
echo $r[4];
"#
        ),
        &["5", "0", "1"]
    );
}

// ── natsort natural ordering ─────────────────────────────────────
#[test]
fn natsort_natural_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["img12", "img10", "img2", "img1"];
natsort($a);
echo implode(",", $a);
"#
        ),
        &["img1,img2,img10,img12"]
    );
}

// ── uksort by key ────────────────────────────────────────────────
#[test]
fn uksort_custom_key_compare() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["banana" => 1, "apple" => 2, "cherry" => 3];
uksort($a, fn($a, $b) => strcmp($a, $b));
echo implode(",", array_keys($a));
"#
        ),
        &["apple,banana,cherry"]
    );
}

// ── array_map with null callback (zip) ──────────────────────────
#[test]
fn array_map_null_callback_zip() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3];
$b = ["a", "b", "c"];
$zipped = array_map(null, $a, $b);
foreach ($zipped as $pair) {
    echo $pair[0] . $pair[1];
}
"#
        ),
        &["1a", "2b", "3c"]
    );
}

// ── array_keys with search value ─────────────────────────────────
#[test]
fn array_keys_with_search_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 1, "d" => 3, "e" => 1];
$keys = array_keys($a, 1);
echo implode(",", $keys);
"#
        ),
        &["a,c,e"]
    );
}

// ── array_merge vs + operator ────────────────────────────────────
#[test]
fn array_merge_vs_plus_operator() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["a" => 1, "b" => 2];
$b = ["b" => 20, "c" => 30];
$merged = array_merge($a, $b);
echo $merged["b"];
$union = $a + $b;
echo $union["b"];
echo $union["c"];
"#
        ),
        &["20", "2", "30"]
    );
}

// ── multidimensional usort ───────────────────────────────────────
#[test]
fn usort_by_nested_field() {
    assert_eq!(
        run_prints(
            r#"<?php
$people = [
    ["name" => "Charlie", "age" => 30],
    ["name" => "Alice",   "age" => 25],
    ["name" => "Bob",     "age" => 28],
];
usort($people, fn($a, $b) => $a["age"] - $b["age"]);
foreach ($people as $p) {
    echo $p["name"];
}
"#
        ),
        &["Alice", "Bob", "Charlie"]
    );
}

// ── array_reduce to string ───────────────────────────────────────
#[test]
fn array_reduce_to_string() {
    assert_eq!(
        run_prints(
            r#"<?php
$words = ["Hello", "beautiful", "world"];
$sentence = array_reduce($words, fn($carry, $w) => $carry === "" ? $w : "$carry $w", "");
echo $sentence;
"#
        ),
        &["Hello beautiful world"]
    );
}

// ── list() skip elements ─────────────────────────────────────────
#[test]
fn list_skip_elements() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [10, 20, 30, 40];
[0 => $first, 3 => $last] = $a;
echo $first;
echo $last;
"#
        ),
        &["10", "40"]
    );
}

// ── array_column with null value column (index map) ──────────────
#[test]
fn array_column_null_value_index_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$records = [
    ["id" => "u1", "name" => "Alice", "score" => 90],
    ["id" => "u2", "name" => "Bob",   "score" => 85],
];
$indexed = array_column($records, null, "id");
echo $indexed["u1"]["name"];
echo $indexed["u2"]["score"];
"#
        ),
        &["Alice", "85"]
    );
}

// ── arsort preserves keys ────────────────────────────────────────
#[test]
fn arsort_preserves_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["x" => 5, "y" => 1, "z" => 3];
arsort($a);
echo implode(",", array_keys($a));
echo implode(",", $a);
"#
        ),
        &["x,z,y", "5,3,1"]
    );
}

// ── array_replace replaces recursively via array_replace_recursive
#[test]
fn array_replace_recursive_deep() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ["a" => ["x" => 1, "y" => 2], "b" => 10];
$over = ["a" => ["y" => 99, "z" => 3]];
$result = array_replace_recursive($base, $over);
echo $result["a"]["x"];
echo $result["a"]["y"];
echo $result["a"]["z"];
echo $result["b"];
"#
        ),
        &["1", "99", "3", "10"]
    );
}
