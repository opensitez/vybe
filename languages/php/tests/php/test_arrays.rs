use super::helpers::{compile_ok, run_prints};

// ── Runtime validation — Phase D "all PHP arrays use vybe:js-array" ──
#[test]
fn array_literal_runtime() {
    let out = run_prints(
        "<?php\n$a = [10, 20, 30];\necho $a[0], \"\\n\"; echo $a[1], \"\\n\"; echo $a[2], \"\\n\";\n",
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn array_append_runtime() {
    let out = run_prints(
        "<?php\n$a = [1, 2];\n$a[] = 3;\n$a[] = 4;\necho count($a), \"\\n\"; echo $a[2], \"\\n\"; echo $a[3], \"\\n\";\n",
    );
    assert_eq!(out, vec!["4", "3", "4"]);
}

#[test]
fn array_push_pop_runtime() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3];\narray_push($a, 4, 5);\n$v = array_pop($a);\necho count($a), \"\\n\"; echo $v, \"\\n\";\n",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn array_length_count_runtime() {
    let out = run_prints("<?php\n$a = [10, 20, 30];\necho count($a);\n");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_length_sizeof_runtime() {
    let out = run_prints("<?php\n$a = ['x' => 1, 'y' => 2];\necho sizeof($a);\n");
    assert_eq!(out, vec!["2"]);
}

// ── Array creation ──────────────────────────────────────────
#[test]
fn array_indexed() {
    compile_ok("<?php $a = [1, 2, 3]; echo $a[0];");
}
#[test]
fn array_assoc() {
    compile_ok("<?php $a = ['name' => 'John', 'age' => 30]; echo $a['name'];");
}
#[test]
fn array_empty() {
    compile_ok("<?php $a = []; echo count($a);");
}
#[test]
fn array_nested() {
    compile_ok("<?php $a = [[1,2],[3,4]]; echo $a[0][1];");
}
#[test]
fn array_mixed_keys() {
    compile_ok("<?php $a = [0 => 'a', 'key' => 'b']; echo $a[0]; echo $a['key'];");
}
#[test]
fn array_function() {
    compile_ok("<?php $a = array(1, 2, 3); echo count($a);");
}
#[test]
fn array_trailing_comma() {
    compile_ok("<?php $a = [1, 2, 3,]; echo count($a);");
}

// ── Array modification ──────────────────────────────────────
#[test]
fn array_assign_index() {
    compile_ok("<?php $a = [1, 2]; $a[0] = 99; echo $a[0];");
}
#[test]
fn array_assign_key() {
    compile_ok("<?php $a = []; $a['x'] = 'hello'; echo $a['x'];");
}
#[test]
fn array_push() {
    compile_ok("<?php $a = [1]; array_push($a, 2); echo count($a);");
}
#[test]
fn array_pop() {
    compile_ok("<?php $a = [1, 2, 3]; $v = array_pop($a); echo $v;");
}
#[test]
fn array_shift() {
    compile_ok("<?php $a = [1, 2, 3]; $v = array_shift($a); echo $v;");
}

// ── Array query ─────────────────────────────────────────────
#[test]
fn in_array() {
    compile_ok("<?php echo in_array(2, [1, 2, 3]);");
}
#[test]
fn array_search() {
    compile_ok("<?php echo array_search('b', ['a', 'b', 'c']);");
}
#[test]
fn array_key_exists() {
    compile_ok("<?php echo array_key_exists('name', ['name' => 'John']);");
}
#[test]
fn count_arr() {
    compile_ok("<?php echo count([1, 2, 3]);");
}

// ── Array transform ─────────────────────────────────────────
#[test]
fn array_merge() {
    compile_ok("<?php $x = array_merge([1, 2], [3, 4]); echo count($x);");
}
#[test]
fn array_slice() {
    compile_ok("<?php $x = array_slice([1, 2, 3, 4, 5], 1, 3); echo count($x);");
}
#[test]
fn array_reverse() {
    compile_ok("<?php $x = array_reverse([1, 2, 3]);");
}
#[test]
fn array_keys() {
    compile_ok("<?php $x = array_keys(['a' => 1, 'b' => 2]);");
}
#[test]
fn array_values() {
    compile_ok("<?php $x = array_values(['a' => 1, 'b' => 2]);");
}
#[test]
fn sort_array() {
    compile_ok("<?php $a = [3, 1, 2]; sort($a);");
}
#[test]
fn range_func() {
    compile_ok("<?php $x = range(1, 10);");
}
#[test]
fn array_sum() {
    compile_ok("<?php echo array_sum([1, 2, 3, 4]);");
}

// ── Callback array ops ──────────────────────────────────────
#[test]
fn array_map() {
    compile_ok("<?php $x = array_map(fn($n) => $n * 2, [1, 2, 3]);");
}
#[test]
fn array_filter() {
    compile_ok("<?php $x = array_filter([0, 1, '', 'a', null], fn($v) => $v);");
}
#[test]
fn array_filter_no_cb() {
    compile_ok("<?php $x = array_filter([0, 1, 2, null, 3]);");
}
#[test]
fn array_reduce() {
    compile_ok("<?php $sum = array_reduce([1,2,3], fn($carry, $item) => $carry + $item, 0);");
}
#[test]
fn array_walk() {
    compile_ok("<?php $a = [1,2,3]; array_walk($a, fn($v, $k) => $v);");
}
#[test]
fn usort_array() {
    compile_ok("<?php $a = [3,1,2]; usort($a);");
}

// ── Array destructuring ─────────────────────────────────────
#[test]
fn list_basic() {
    compile_ok("<?php list($a, $b) = [1, 2]; echo $a;");
}
#[test]
fn short_list() {
    compile_ok("<?php [$x, $y, $z] = [10, 20, 30]; echo $y;");
}
#[test]
fn compact_extract() {
    compile_ok("<?php $a = 1; $b = 2; $arr = compact('a', 'b');");
}

// ── Foreach iteration ───────────────────────────────────────
#[test]
fn foreach_indexed() {
    compile_ok("<?php foreach ([1,2,3] as $v) { echo $v; }");
}
#[test]
fn foreach_assoc() {
    compile_ok("<?php foreach (['a'=>1, 'b'=>2] as $k => $v) { echo $k . ': ' . $v; }");
}
#[test]
fn foreach_nested() {
    compile_ok("<?php foreach ([[1,2],[3,4]] as $row) { foreach ($row as $v) { echo $v; } }");
}

// ── Spread operator ─────────────────────────────────────────
#[test]
fn spread_in_array() {
    compile_ok("<?php $a = [1, 2]; $b = [...$a, 3, 4];");
}
#[test]
fn spread_in_call() {
    compile_ok("<?php function sum(...$nums) { return 0; } sum(...[1,2,3]);");
}

#[test]
fn array_fill_runtime() {
    let out = run_prints(
        "<?php\n$a = array_fill(2, 3, 7);\n$b = array_fill_keys(['a', 'b'], 9);\necho implode(',', $a), \"\\n\";\nksort($b);\necho json_encode($b);\n",
    );
    assert_eq!(out, vec!["7,7,7", "{\"a\":9,\"b\":9}"]);
}

#[test]
fn array_chunk_runtime() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3, 4, 5];\n$chunks = array_chunk($a, 2);\n$joined = [];\nforeach ($chunks as $chunk) { $joined[] = implode('-', $chunk); }\necho implode('|', $joined);\n",
    );
    assert_eq!(out, vec!["1-2|3-4|5"]);
}

#[test]
fn array_slice_and_splice_runtime() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3, 4, 5];\n$s = array_slice($a, 1, 2);\n$re = array_splice($a, 2, 2, [8, 9]);\necho count($s), \"\\n\";\necho implode(',', $s), \"\\n\";\necho implode(',', $a);\n",
    );
    assert_eq!(out, vec!["2", "2,3", "1,2,8,9,5"]);
}

#[test]
fn array_merge_and_replace_runtime() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3];\n$b = [2 => 'x', 5 => 'y'];\n$m = array_merge($a, $b);\n$r = array_replace($a, $b);\necho json_encode($m), \"\\n\";\necho json_encode($r);\n",
    );
    assert_eq!(out, vec!["[1,2,3,\"x\",\"y\"]", "{\"0\":1,\"1\":2,\"2\":\"x\",\"5\":\"y\"}"]);
}

#[test]
fn array_diff_and_intersect_runtime() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3, 4];\n$b = [2, 4];\n$d = array_diff($a, $b);\n$i = array_intersect($a, $b);\n$items = [];\nforeach ($d as $v) { $items[] = $v; }\nforeach ($i as $v) { $items[] = \"i$v\"; }\necho implode('-', $items);\n",
    );
    assert_eq!(out, vec!["1-3-i2-i4"]);
}

#[test]
fn array_search_and_flip_runtime() {
    let out = run_prints(
        "<?php\n$a = ['a' => 10, 'b' => 20, 'c' => 10];\necho array_search(10, $a), \"\\n\";\n$f = array_flip(['x' => 1, 'y' => 2]);\nksort($f);\necho json_encode($f);\n",
    );
    assert_eq!(out, vec!["a", "{\"1\":\"x\",\"2\":\"y\"}"]);
}

#[test]
fn array_combine_and_map_runtime() {
    let out = run_prints(
        "<?php\n$k = ['a', 'b', 'c'];\n$v = [1, 2, 3];\n$c = array_combine($k, $v);\n$m = array_map(fn($x) => $x * $x, $c);\necho $m['a'], \"\\n\";\necho $m['b'], \"\\n\";\necho $m['c'];\n",
    );
    assert_eq!(out, vec!["1", "4", "9"]);
}

#[test]
fn array_keycase_runtime() {
    let out = run_prints(
        "<?php\n$a = ['One' => 1, 'Two' => 2];\n$lower = array_change_key_case($a, CASE_LOWER);\n$upper = array_change_key_case($a, CASE_UPPER);\necho isset($lower['one']) ? '1' : '0';\necho isset($upper['TWO']) ? '|1' : '|0';\n",
    );
    assert_eq!(out, vec!["1|1"]);
}

#[test]
fn array_reference_and_copy_behaviors() {
    let out = run_prints(
        "<?php\n$source = [1, 2, 3];\n$copy = $source;\n$source[0] = 9;\necho $copy[0] . \"\\n\";\n$alias = &$source;\n$alias[1] = 11;\necho $source[1];\n",
    );
    assert_eq!(out, vec!["1", "11"]);
}

#[test]
fn array_unpack_default_overrides() {
    let out = run_prints(
        "<?php\n$base = [1, 2];\n$merged = [...$base, ...[2 => 30, 3 => 40]];\necho count($merged) . \"\\n\";\necho $merged[0] . ',' . $merged[2] . ',' . $merged[3];\n",
    );
    assert_eq!(out, vec!["4", "1,30,40"]);
}

#[test]
fn array_column_nested_keys() {
    let out = run_prints(
        "<?php\n$rows = [\n    ['id' => 1, 'tag' => 'x'],\n    ['id' => 2, 'tag' => 'y'],\n];\n$ids = array_column($rows, 'tag', 'id');\nksort($ids);\necho implode(',', array_values($ids));\n",
    );
    assert_eq!(out, vec!["x,y"]);
}

#[test]
fn array_access_in_expression_chain() {
    let out = run_prints(
        "<?php\n$matrix = [[10, 20], [30, 40]];\n$sum = 0;\nfor ($i = 0; $i < count($matrix); $i++) {\n    for ($j = 0; $j < 2; $j++) {\n        $sum += $matrix[$i][$j];\n    }\n}\necho $sum;\n",
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn array_multidim_unpack_with_list() {
    let out = run_prints(
        "<?php\n$packed = [\n    ['x' => 1, 'y' => 2],\n    ['x' => 3, 'y' => 4],\n];\n[$first, $second] = $packed;\necho $first['y'] + $second['x'];\n",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn array_filter_keep_boolean_false_like_values() {
    let out = run_prints(
        "<?php\n$values = [0, '0', '', null, false, 2, 3];\n$a = array_filter($values, fn($v) => $v !== null && $v !== false);\n$b = array_filter($values, fn($v) => is_int($v));\necho count($a) . '|';\necho count($b);\n",
    );
    assert_eq!(out, vec!["5|3"]);
}

#[test]
fn array_intersection_by_key_values() {
    let out = run_prints(
        "<?php\n$a = ['a' => 1, 'b' => 2, 'c' => 3];\n$b = ['a' => 99, 'c' => 33];\n$both = array_intersect_key($a, $b);\n$all = array_intersect_assoc($a, array_merge($b, ['c' => 3]));\nksort($both);\nksort($all);\necho implode(',', array_keys($both)) . '|';\necho implode(',', $all);\n",
    );
    assert_eq!(out, vec!["a,c|c"]);
}

#[test]
fn array_walk_by_value_modifies_external_sum() {
    let out = run_prints(
        "<?php\n$items = [1, 2, 3];\n$total = 0;\narray_walk($items, function(int $v) use (&$total): void { $total += $v; });\necho $total;\n",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_list_semantics() {
    let out = run_prints(
        "<?php\n$a = [10, 20, 30];\necho array_is_list($a) ? 'list' : 'map';\n",
    );
    assert_eq!(out, vec!["list"]);
}

#[test]
fn array_non_list_semantics() {
    let out = run_prints(
        "<?php\n$a = ['x' => 10, 1 => 20, 2 => 30];\necho array_is_list($a) ? 'list' : 'map';\n",
    );
    assert_eq!(out, vec!["map"]);
}

#[test]
fn array_unshift_push_pop_order() {
    let out = run_prints(
        "<?php\n$a = [3, 4];\n$len1 = array_unshift($a, 1, 2);\n$len2 = array_push($a, 5);\n$tail = array_pop($a);\necho $len1 . '|' . $len2 . '|' . $tail . '|';\necho implode(',', $a);\n",
    );
    assert_eq!(out, vec!["4|5|5|1,2,3,4"]);
}

#[test]
fn array_search_with_strict_flag() {
    let out = run_prints(
        "<?php\n$a = [1, '2', 3];\necho in_array('2', $a, true) ? 'strict-true' : 'strict-false';\necho '|';\necho in_array(2, $a, true) ? 'strict-true-2' : 'strict-false-2';\n",
    );
    assert_eq!(out, vec!["strict-true|strict-false-2"]);
}

#[test]
fn array_pad_grows_both_sides() {
    let out = run_prints(
        "<?php\n$a = [1, 2];\n$left = array_pad($a, -4, 0);\n$right = array_pad($a, 4, 9);\nksort($left);\nksort($right);\necho implode(',', $left) . '|';\necho implode(',', $right);\n",
    );
    assert_eq!(out, vec!["0,0,1,2|1,2,9,9"]);
}

#[test]
fn array_unique_keeps_first_occurrence() {
    let out = run_prints(
        "<?php\n$a = [1, '1', 2, 2, 3, 1];\n$b = array_unique($a);\n$flags = array_values($b);\necho implode('-', $flags);\n",
    );
    assert_eq!(out, vec!["1-2-3"]);
}

#[test]
fn array_replace_recursive_nested() {
    let out = run_prints(
        "<?php\n$a = ['x' => ['a' => 1], 'y' => 2];\n$b = ['x' => ['b' => 3], 'y' => ['z' => 4]];\n$r = array_replace_recursive($a, $b);\nksort($r);\nksort($r['x']);\necho json_encode($r['x']) . '|';\necho json_encode($r['y']) . '|';\necho $r['x']['b'];\n",
    );
    assert_eq!(out, vec!["{\"a\":1,\"b\":3}|{\"z\":4}|3"]);
}

#[test]
fn array_multisort_numeric_and_sort_flags() {
    let out = run_prints(
        "<?php\n$nums = [3, 1, 2];\n$letters = ['c', 'a', 'b'];\narray_multisort($nums, SORT_ASC, SORT_NUMERIC, $letters, SORT_ASC, SORT_STRING);\n$items = [];\nfor ($i = 0; $i < count($nums); $i++) {\n    $items[] = $nums[$i] . ':' . $letters[$i];\n}\necho implode('|', $items);\n",
    );
    assert_eq!(out, vec!["1:a|2:b|3:c"]);
}

#[test]
fn array_fill_keys_rejects_non_list_values() {
    let out = run_prints(
        "<?php\n$keys = ['a', 'b', 'a'];\n$vals = [1, 2];\n$result = [];\nif (count($keys) !== count($vals)) {\n    $result = ['error' => 1];\n} else {\n    $result = array_fill_keys($keys, 0);\n}\n$hasA = isset($result['a']) ? 'A' : 'NA';\necho $hasA . '|' . count($result);\n",
    );
    assert_eq!(out, vec!["NA|1"]);
}

#[test]
fn array_values_and_keys_reindexing() {
    let out = run_prints(
        "<?php\n$a = ['x' => 1, 5 => 2, 9 => 3];\n$vals = array_values($a);\n$keys = array_keys($a);\necho implode(',', $vals) . '|';\necho implode(',', $keys);\n",
    );
    assert_eq!(out, vec!["1,2,3|x,5,9"]);
}

#[test]
fn array_column_null_and_default_index() {
    let out = run_prints(
        "<?php\n$rows = [\n    ['id' => 1, 'name' => 'A'],\n    ['name' => 'B'],\n    ['id' => 3, 'name' => 'C', 'score' => 0],\n];\n$names = array_column($rows, 'name');\n$ids = array_column($rows, 'id', 'name');\nif (array_key_exists('B', $ids)) {\n    echo 'missing-id|';\n} else {\n    echo 'no-key|';\n}\necho count($names) . '|' . count($ids) . '|';\necho $ids['A'];\n",
    );
    assert_eq!(out, vec!["no-key|3|2|1"]);
}

#[test]
fn array_fill_keys_ordered_numeric_and_string_keys() {
    let out = run_prints(
        "<?php\n$map = array_fill_keys([0, 1, 'x', 'y'], 7);\nksort($map);\necho isset($map[0]) ? '0' : 'x';\necho isset($map['x']) ? '|x' : '|no';\necho count($map);\n",
    );
    assert_eq!(out, vec!["0|x4"]);
}

#[test]
fn array_udiff_numeric_and_assoc() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3, 4];\n$b = [2, 4, 6];\n$r = array_udiff($a, $b, fn($x, $y) => $x <=> $y);\nksort($r);\necho json_encode(array_values($r));\n",
    );
    assert_eq!(out, vec!["[1,3]"]);
}

#[test]
fn array_uintersect_closure_case() {
    let out = run_prints(
        "<?php\n$a = ['a', 'B', 'C'];\n$b = ['a', 'b', 'd'];\n$r = array_uintersect($a, $b, fn($x, $y) => strcasecmp($x, $y));\nksort($r);\necho implode(',', $r);\n",
    );
    assert_eq!(out, vec!["a,B"]);
}

#[test]
fn array_splice_inserts_and_removes() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3, 4];\n$removed = array_splice($a, 1, 2, [8, 9]);\nksort($a);\nksort($removed);\necho implode(',', $a) . '|';\necho implode(',', $removed);\n",
    );
    assert_eq!(out, vec!["1,8,9,4|2,3"]);
}

#[test]
fn array_sort_flags_string_and_natural() {
    let out = run_prints(
        "<?php\n$files = ['file10.txt', 'file2.txt', 'file1.txt'];\nnatcasesort($files);\necho implode('|', $files);\n",
    );
    assert_eq!(out, vec!["file1.txt|file2.txt|file10.txt"]);
}

#[test]
fn array_partition_by_callback_runtime() {
    let out = run_prints(
        "<?php\n$items = [1, 2, 3, 4, 5];\n$even = array_filter($items, fn($n) => $n % 2 === 0);\n$odd = array_filter($items, fn($n) => $n % 2 === 1);\narray_values($even);\narray_values($odd);\necho implode(',', $even) . '|';\necho implode(',', $odd);\n",
    );
    assert_eq!(out, vec!["2,4|1,3,5"]);
}

#[test]
fn array_chunk_with_preserve_keys() {
    let out = run_prints(
        "<?php\n$a = [10 => 'a', 11 => 'b', 12 => 'c', 13 => 'd'];\n$chunks = array_chunk($a, 2, true);\n$first = array_keys($chunks[0]);\n$second = array_keys($chunks[1]);\necho $first[0] . ',' . $first[1] . '|';\necho $second[0] . ',' . $second[1];\n",
    );
    assert_eq!(out, vec!["10,11|12,13"]);
}

#[test]
fn array_key_cast_preserves_string_numeric_keys() {
    let out = run_prints(
        "<?php\n$a = ['01' => 'a', 1 => 'b', '2x' => 'c'];\necho array_key_exists('1', $a) ? 'one' : 'no';\necho '|';\necho array_key_exists(1, $a) ? 'num1' : 'no';\n",
    );
    assert_eq!(out, vec!["one|no"]);
}

#[test]
fn array_offset_set_on_string_numeric_index() {
    let out = run_prints(
        "<?php\n$a = ['01' => 'a'];\n$a[1] = 'b';\necho json_encode($a);\n",
    );
    assert_eq!(out, vec!["{\"01\":\"a\",\"1\":\"b\"}"]);
}

#[test]
fn array_reference_foreach_updates_original_array() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3];\nforeach ($a as &$v) { $v = $v + 10; }\nforeach ($a as $x) { echo $x; }\n",
    );
    assert_eq!(out, vec!["111213"]);
}

#[test]
fn array_forget_reference_after_foreach() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3];\nforeach ($a as &$v) {}\nunset($v);\n$a[] = 4;\nforeach ($a as $x) { echo $x; }\n",
    );
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn array_destructure_nested_and_default() {
    let out = run_prints(
        "<?php\n$data = [['a', 'b'], ['c']];\n[$left, $right] = $data;\n[$x, $y] = $left;\n[$x2, $y2 = 'zz'] = $right;\necho $x . $y . $x2 . $y2;\n",
    );
    assert_eq!(out, vec!["abczz"]);
}

#[test]
fn array_walk_recursive_collect_path() {
    let out = run_prints(
        "<?php\n$tree = ['a' => ['x' => 1], 'b' => ['y' => 2]];\n$out = [];\narray_walk_recursive($tree, function($v, $k) use (&$out) { $out[] = \"$k=$v\"; });\necho implode('|', $out);\n",
    );
    assert_eq!(out, vec!["x=1|y=2"]);
}

#[test]
fn array_reduce_with_initial_value_nonzero() {
    let out = run_prints(
        "<?php\n$nums = [1, 2, 3];\n$s = array_reduce($nums, fn($carry, $item): int => $carry + $item, 10);\necho $s;\n",
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn array_change_key_case_false_keeps_key_order() {
    let out = run_prints(
        "<?php\n$a = ['B' => 1, 'a' => 2, 'C' => 3];\n$b = array_change_key_case($a, CASE_UPPER);\n$keys = array_keys($b);\necho implode(',', $keys);\n",
    );
    assert_eq!(out, vec!["B,A,C"]);
}

#[test]
fn array_slice_negative_length() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3, 4, 5];\n$s = array_slice($a, 1, -1);\necho implode(',', $s);\n",
    );
    assert_eq!(out, vec!["2,3,4"]);
}

#[test]
fn array_splice_returns_removed_indexed_from_zero() {
    let out = run_prints(
        "<?php\n$a = ['k' => 'v1', 'm' => 'v2', 'n' => 'v3'];\n$r = array_splice($a, 1, 1, ['x' => 'vv']);\necho isset($r[0]) ? $r[0] : 'none';\necho ':' . count($r);\n",
    );
    assert_eq!(out, vec!["v2:1"]);
}

#[test]
fn array_filter_preserve_default_truthy() {
    let out = run_prints(
        "<?php\n$values = [0, 1, '', 'ok', null, 2];\n$r = array_filter($values);\necho implode('|', $r);\n",
    );
    assert_eq!(out, vec!["1|ok|2"]);
}

#[test]
fn array_index_with_false_and_empty_string_key() {
    let out = run_prints(
        "<?php\n$a = [false => 'f', '' => 'e', 0 => 'z'];\necho $a[false] . '|';\necho $a[''];\n",
    );
    assert_eq!(out, vec!["f|e"]);
}

#[test]
fn array_multi_assign_with_numeric_like_keys() {
    let out = run_prints(
        "<?php\n$a = [];\n$a[] = 'first';\n$a['1'] = 'string1';\n$a[1.9] = 'float1';\necho count($a) . '|';\necho $a[1] . '|';\necho $a[1.9];\n",
    );
    assert_eq!(out, vec!["2|string1|string1"]);
}

#[test]
fn array_search_and_replace_with_missing_key() {
    let out = run_prints(
        "<?php\n$a = ['a' => 1, 'b' => 2];\necho array_key_exists('z', $a) ? 'yes' : 'no';\necho '|';\necho array_search(3, $a, true) === false ? 'missing' : 'found';\n",
    );
    assert_eq!(out, vec!["no|missing"]);
}

#[test]
fn array_flip_duplicate_values_keeps_last() {
    let out = run_prints(
        "<?php\n$a = ['a' => 1, 'b' => 2, 'c' => 1];\n$f = array_flip($a);\necho isset($f['1']) ? $f['1'] : 'none';\n",
    );
    assert_eq!(out, vec!["c"]);
}

#[test]
fn array_merge_vs_plus_operator_behavior_runtime() {
    let out = run_prints(
        "<?php\n$a = ['x' => 1, 2 => 3];\n$b = ['x' => 9, 2 => 8, 3 => 7];\n$m = array_merge($a, $b);\n$u = $a + $b;\necho $m['x'] . '|';\necho $m[2] . '|';\necho $m[3] . '|';\necho $u[2];\n",
    );
    assert_eq!(out, vec!["1|3|7|3"]);
}

#[test]
fn array_key_preservation_in_spread_assignments() {
    let out = run_prints(
        "<?php\n$left = ['a' => 1, 2 => 2];\n$right = [...$left, 'b' => 3, 4 => 4];\nksort($right);\necho json_encode($right);\n",
    );
    assert_eq!(out, vec!["{\"a\":1,\"2\":2,\"b\":3,\"4\":4}"]);
}

#[test]
fn array_filter_retain_zero_string_and_boolean_false() {
    let out = run_prints(
        "<?php\n$values = [0, 0.0, '0', false, true, 'ok', ''];\n$strict = array_filter($values, fn($v) => $v !== null);\n$loose = array_filter($values);\necho count($strict) . '|';\necho count($loose);\n",
    );
    assert_eq!(out, vec!["7|2"]);
}

#[test]
fn array_map_with_keys_not_modified() {
    let out = run_prints(
        "<?php\n$map = ['a' => 1, 'b' => 2];\n$out = array_map(fn($v) => $v * 2, $map);\n$first = key($out);\necho $first . '|' . $out['a'];\n",
    );
    assert_eq!(out, vec!["a|2"]);
}

#[test]
fn array_combine_invalid_count_fails() {
    let out = run_prints(
        "<?php\n$ok = true;\ntry {\n    array_combine(['a', 'b'], [1]);\n    $ok = false;\n} catch (ValueError $e) {\n    echo 'value_error';\n}\nif ($ok) { echo 'unexpected'; }\n",
    );
    assert_eq!(out, vec!["value_errorunexpected"]);
}

#[test]
fn array_replace_preserves_associative_precedence() {
    let out = run_prints(
        "<?php\n$base = ['x' => 1, 'y' => ['k' => 'left', 'z' => 0]];\n$patch = ['y' => ['z' => 9]];\n$out = array_replace_recursive($base, $patch);\necho $out['y']['k'] . '|' . $out['y']['z'];\n",
    );
    assert_eq!(out, vec!["left|9"]);
}

#[test]
fn array_comparison_using_spaceship_for_keys() {
    let out = run_prints(
        "<?php\n$a = ['x' => 1, 'y' => 2];\n$b = ['x' => 1, 'y' => 3];\necho ($a == $b ? 'eq' : 'neq') . '|';\necho ($a <=> $b);\n",
    );
    assert_eq!(out, vec!["neq|-1"]);
}

#[test]
fn array_pad_when_padding_with_negative_size() {
    let out = run_prints(
        "<?php\n$a = [1, 2, 3];\n$trimmed = array_pad($a, -2, 9);\necho implode(',', $trimmed);\n",
    );
    assert_eq!(out, vec!["1,2,3"]);
}

#[test]
fn array_fill_preserves_type_when_using_strings() {
    let out = run_prints(
        "<?php\n$a = array_fill(0, 3, 'x');\n$a[1] = 2;\necho implode('', $a) . '|' . gettype($a[0]);\n",
    );
    assert_eq!(out, vec!["x2x|string"]);
}

#[test]
fn array_column_with_no_match_returns_null() {
    let out = run_prints(
        "<?php\n$rows = [['id' => 1, 'name' => 'A'], ['id' => 2]];\n$vals = array_column($rows, 'name');\n$lookup = array_column($rows, 'name', 'id');\necho count($vals) . '|' . (array_key_exists(3, $lookup) ? 'has' : 'no');\n",
    );
    assert_eq!(out, vec!["1|no"]);
}

#[test]
fn array_fill_keys_with_boolean_keys() {
    let out = run_prints(
        "<?php\n$map = array_fill_keys([true, false], 1);\necho (array_key_exists(1, $map) ? 't' : 'n');\necho '|';\necho (array_key_exists(0, $map) ? 'f' : 'n');\n",
    );
    assert_eq!(out, vec!["t|f"]);
}

#[test]
fn array_intersect_assoc_keeps_pairs() {
    let out = run_prints(
        "<?php\n$a = ['a' => 1, 'b' => 2, 'c' => 3];\n$b = ['a' => 1, 'b' => 4, 'c' => 3];\n$r = array_intersect_assoc($a, $b);\nksort($r);\necho implode(',', array_keys($r));\n",
    );
    assert_eq!(out, vec!["a,c"]);
}

#[test]
fn array_merge_with_empty_operands() {
    let out = run_prints(
        "<?php\n$a = [1, 2];\n$b = array_merge($a, []);\n$c = array_merge([], $a);\necho count($b) . '|';\necho count($c) . '|';\necho $b[0] . $c[1];\n",
    );
    assert_eq!(out, vec!["2|2|22"]);
}

#[test]
fn array_plus_keeps_left_numeric_keys() {
    let out = run_prints(
        "<?php\n$a = [0 => 'left', 1 => 'a'];\n$b = [1 => 'right', 2 => 'b'];\n$m = $a + $b;\necho $m[0] . '|' . $m[1] . '|' . $m[2];\n",
    );
    assert_eq!(out, vec!["left|a|b"]);
}

#[test]
fn array_pop_empty_returns_null() {
    let out = run_prints(
        "<?php\n$a = [];\n$v = array_pop($a);\necho is_null($v) ? 'null' : 'not';\n",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn array_shift_empty_returns_null() {
    let out = run_prints(
        "<?php\n$a = [];\n$v = array_shift($a);\necho is_null($v) ? 'null' : 'not';\n",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn array_unshift_count_with_multiple_values() {
    let out = run_prints(
        "<?php\n$a = [3, 4];\n$len = array_unshift($a, 1, 2);\necho $len . '|' . implode(',', $a);\n",
    );
    assert_eq!(out, vec!["4|1,2,3,4"]);
}

#[test]
fn array_key_cast_float_key_to_int() {
    let out = run_prints(
        "<?php\n$a = [];\n$a[1.1] = 'x';\n$a[2.9] = 'y';\necho isset($a[1]) ? '1' : '0';\necho '|';\necho isset($a[2]) ? '2' : '0';\n",
    );
    assert_eq!(out, vec!["1|2"]);
}

#[test]
fn array_key_collision_between_false_null_and_empty_string_runtime() {
    let out = run_prints(
        "<?php\n$a = [];\n$a[false] = 'false';\n$a[null] = 'null';\n$a[''] = 'empty';\necho $a[false];\necho '|';\necho $a[''];\necho '|';\necho count($a);\n",
    );
    assert_eq!(out, vec!["empty|empty|1"]);
}

#[test]
fn array_destructure_with_skipped_slot_runtime() {
    let out = run_prints(
        "<?php\n$tuple = [10, 20, 30, 40];\n[$first, , $third, $fourth] = $tuple;\necho $first . '|' . $third . '|' . $fourth;\n",
    );
    assert_eq!(out, vec!["10|30|40"]);
}

#[test]
fn array_walk_recursive_with_reference_update_runtime() {
    let out = run_prints(
        "<?php\n$tree = [\n    ['a' => 1],\n    ['b' => 2],\n];\narray_walk_recursive($tree, function (&$value) { $value = $value * 10; });\necho $tree[0]['a'] . '|' . $tree[1]['b'];\n",
    );
    assert_eq!(out, vec!["10|20"]);
}

#[test]
fn array_search_with_numeric_like_string_keys_runtime() {
    let out = run_prints(
        "<?php\n$a = ['01' => 'left', '1' => 'right', 2 => 'third'];\necho $a['1'];\necho '|';\necho $a['01'];\necho '|';\necho array_key_exists(1, $a) ? 'one' : 'no';\n",
    );
    assert_eq!(out, vec!["right|left|one"]);
}
