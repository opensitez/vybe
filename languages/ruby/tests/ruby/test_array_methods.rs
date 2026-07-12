use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── each_with_index ───────────────────────────────────────────────────────────

#[test]
fn arr_each_with_index() {
    compile_ok("[10, 20, 30].each_with_index { |v, i| puts i }\n");
}

#[test]
fn arr_each_with_index_runtime() {
    let out = run_ruby("[10, 20, 30].each_with_index { |v, i| puts i }\n");
    assert_eq!(out, vec!["0", "1", "2"]);
}

// ── each_with_object ─────────────────────────────────────────────────────────

#[test]
fn arr_each_with_object() {
    compile_ok("[1, 2, 3].each_with_object([]) { |x, acc| acc.push(x * 2) }\n");
}

// ── flat_map ──────────────────────────────────────────────────────────────────

#[test]
fn arr_flat_map_one_level() {
    compile_ok("x = [[1, 2], [3, 4]].flat_map { |a| a }\n");
}

#[test]
fn arr_flat_map_runtime() {
    assert_eq!(
        run_ruby_one("puts [[1, 2], [3]].flat_map { |a| a }.length\n"),
        "3"
    );
}

// ── zip ───────────────────────────────────────────────────────────────────────

#[test]
fn arr_zip_combine_arrays() {
    compile_ok("x = [1, 2, 3].zip([4, 5, 6])\n");
}

// ── take ──────────────────────────────────────────────────────────────────────

#[test]
fn arr_take_first_n() {
    compile_ok("x = [1, 2, 3, 4, 5].take(3)\n");
}

#[test]
fn arr_take_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3, 4, 5].take(3).length\n"), "3");
}

// ── take_while ────────────────────────────────────────────────────────────────

#[test]
fn arr_take_while_condition() {
    compile_ok("x = [1, 2, 3, 4, 5].take_while { |n| n < 4 }\n");
}

// ── drop ──────────────────────────────────────────────────────────────────────

#[test]
fn arr_drop_first_n() {
    compile_ok("x = [1, 2, 3, 4, 5].drop(2)\n");
}

#[test]
fn arr_drop_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3, 4, 5].drop(2).length\n"), "3");
}

// ── drop_while ────────────────────────────────────────────────────────────────

#[test]
fn arr_drop_while_condition() {
    compile_ok("x = [1, 2, 3, 4, 5].drop_while { |n| n < 3 }\n");
}

// ── find / detect ─────────────────────────────────────────────────────────────

#[test]
fn arr_find_first_matching() {
    compile_ok("x = [1, 2, 3, 4].find { |n| n > 2 }\n");
}

#[test]
fn arr_find_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3, 4].find { |n| n > 2 }\n"), "3");
}

#[test]
fn arr_detect_alias() {
    compile_ok("x = [1, 2, 3].detect { |n| n.even? }\n");
}

// ── find_index / index ────────────────────────────────────────────────────────

#[test]
fn arr_find_index_of_match() {
    compile_ok("x = [10, 20, 30].find_index { |n| n == 20 }\n");
}

#[test]
fn arr_find_index_runtime() {
    assert_eq!(
        run_ruby_one("puts [10, 20, 30].find_index { |n| n == 20 }\n"),
        "1"
    );
}

// ── count with block ──────────────────────────────────────────────────────────

#[test]
fn arr_count_with_block() {
    compile_ok("x = [1, 2, 3, 4].count { |n| n.even? }\n");
}

#[test]
fn arr_count_block_runtime() {
    assert_eq!(
        run_ruby_one("puts [1, 2, 3, 4].count { |n| n.even? }\n"),
        "2"
    );
}

// ── any? ──────────────────────────────────────────────────────────────────────

#[test]
fn arr_any_predicate() {
    compile_ok("x = [1, 2, 3].any? { |n| n > 2 }\n");
}

#[test]
fn arr_any_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3].any? { |n| n > 2 }\n"), "true");
}

// ── all? ──────────────────────────────────────────────────────────────────────

#[test]
fn arr_all_predicate() {
    compile_ok("x = [2, 4, 6].all? { |n| n.even? }\n");
}

#[test]
fn arr_all_runtime() {
    assert_eq!(
        run_ruby_one("puts [2, 4, 6].all? { |n| n.even? }\n"),
        "true"
    );
}

// ── none? ─────────────────────────────────────────────────────────────────────

#[test]
fn arr_none_predicate() {
    compile_ok("x = [1, 3, 5].none? { |n| n.even? }\n");
}

#[test]
fn arr_none_runtime() {
    assert_eq!(
        run_ruby_one("puts [1, 3, 5].none? { |n| n.even? }\n"),
        "true"
    );
}

// ── one? ──────────────────────────────────────────────────────────────────────

#[test]
fn arr_one_predicate() {
    compile_ok("x = [1, 2, 3].one? { |n| n.even? }\n");
}

// ── include? ──────────────────────────────────────────────────────────────────

#[test]
fn arr_include_membership() {
    compile_ok("x = [1, 2, 3].include?(2)\n");
}

#[test]
fn arr_include_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3].include?(2)\n"), "true");
}

// ── sort_by ───────────────────────────────────────────────────────────────────

#[test]
fn arr_sort_by_block_value() {
    compile_ok("x = ['banana', 'apple', 'cherry'].sort_by { |s| s.length }\n");
}

#[test]
fn arr_sort_by_runtime() {
    assert_eq!(
        run_ruby_one("puts ['banana', 'apple', 'cherry'].sort_by { |s| s.length }.first\n"),
        "apple"
    );
}

// ── group_by ──────────────────────────────────────────────────────────────────

#[test]
fn arr_group_by_block() {
    compile_ok("x = [1, 2, 3, 4].group_by { |n| n.even? }\n");
}

// ── tally ─────────────────────────────────────────────────────────────────────

#[test]
fn arr_tally_count_occurrences() {
    compile_ok("x = ['a', 'b', 'a', 'c', 'b', 'a'].tally\n");
}

// ── min_by ────────────────────────────────────────────────────────────────────

#[test]
fn arr_min_by_block_value() {
    compile_ok("x = ['banana', 'apple', 'cherry'].min_by { |s| s.length }\n");
}

#[test]
fn arr_min_by_runtime() {
    assert_eq!(
        run_ruby_one("puts ['banana', 'apple', 'cherry'].min_by { |s| s.length }\n"),
        "apple"
    );
}

// ── max_by ────────────────────────────────────────────────────────────────────

#[test]
fn arr_max_by_block_value() {
    compile_ok("x = ['banana', 'ap', 'cherry'].max_by { |s| s.length }\n");
}

// ── minmax ────────────────────────────────────────────────────────────────────

#[test]
fn arr_minmax_returns_pair() {
    compile_ok("x = [3, 1, 4, 1, 5].minmax\n");
}

// ── minmax_by ─────────────────────────────────────────────────────────────────

#[test]
fn arr_minmax_by_block() {
    compile_ok("x = ['apple', 'fig', 'cherry'].minmax_by { |s| s.length }\n");
}

// ── each_slice ────────────────────────────────────────────────────────────────

#[test]
fn arr_each_slice_chunks() {
    compile_ok("[1, 2, 3, 4, 5, 6].each_slice(2) { |s| puts s.length }\n");
}

// ── each_cons ─────────────────────────────────────────────────────────────────

#[test]
fn arr_each_cons_sliding_window() {
    compile_ok("[1, 2, 3, 4, 5].each_cons(3) { |w| puts w.first }\n");
}

// ── combination ───────────────────────────────────────────────────────────────

#[test]
fn arr_combination_n() {
    compile_ok("x = [1, 2, 3, 4].combination(2).to_a\n");
}

// ── permutation ───────────────────────────────────────────────────────────────

#[test]
fn arr_permutation_n() {
    compile_ok("x = [1, 2, 3].permutation(2).to_a\n");
}

// ── product ───────────────────────────────────────────────────────────────────

#[test]
fn arr_product_cartesian() {
    compile_ok("x = [1, 2].product([3, 4])\n");
}

// ── transpose ─────────────────────────────────────────────────────────────────

#[test]
fn arr_transpose_matrix() {
    compile_ok("x = [[1, 2], [3, 4], [5, 6]].transpose\n");
}

// ── & intersection ────────────────────────────────────────────────────────────

#[test]
fn arr_intersection_operator() {
    compile_ok("x = [1, 2, 3, 4] & [2, 4, 6]\n");
}

#[test]
fn arr_intersection_runtime() {
    assert_eq!(
        run_ruby_one("puts ([1, 2, 3, 4] & [2, 4, 6]).length\n"),
        "2"
    );
}

// ── | union ───────────────────────────────────────────────────────────────────

#[test]
fn arr_union_operator() {
    compile_ok("x = [1, 2, 3] | [2, 3, 4]\n");
}

#[test]
fn arr_union_runtime() {
    assert_eq!(run_ruby_one("puts ([1, 2, 3] | [2, 3, 4]).length\n"), "4");
}

// ── - difference ──────────────────────────────────────────────────────────────

#[test]
fn arr_difference_operator() {
    compile_ok("x = [1, 2, 3, 4] - [2, 4]\n");
}

#[test]
fn arr_difference_runtime() {
    assert_eq!(run_ruby_one("puts ([1, 2, 3, 4] - [2, 4]).length\n"), "2");
}

// ── + concatenation ───────────────────────────────────────────────────────────

#[test]
fn arr_concat_operator() {
    compile_ok("x = [1, 2] + [3, 4]\n");
}

#[test]
fn arr_concat_runtime() {
    assert_eq!(run_ruby_one("puts ([1, 2] + [3, 4]).length\n"), "4");
}

// ── * repetition ─────────────────────────────────────────────────────────────

#[test]
fn arr_repetition_operator() {
    compile_ok("x = [1, 2] * 3\n");
}

#[test]
fn arr_repetition_runtime() {
    assert_eq!(run_ruby_one("puts ([1, 2] * 3).length\n"), "6");
}

// ── flatten(depth) ────────────────────────────────────────────────────────────

#[test]
fn arr_flatten_to_depth() {
    compile_ok("x = [1, [2, [3, [4]]]].flatten(1)\n");
}

#[test]
fn arr_flatten_depth_runtime() {
    assert_eq!(run_ruby_one("puts [1, [2, [3]]].flatten(1).length\n"), "3");
}

// ── compact new array ─────────────────────────────────────────────────────────

#[test]
fn arr_compact_returns_new_array() {
    compile_ok("x = [1, nil, 2, nil, 3].compact\n");
}

// ── delete(val) ───────────────────────────────────────────────────────────────

#[test]
fn arr_delete_matching_value() {
    compile_ok("a = [1, 2, 3, 2, 1]\na.delete(2)\n");
}

// ── delete_if ────────────────────────────────────────────────────────────────

#[test]
fn arr_delete_if_block() {
    compile_ok("a = [1, 2, 3, 4]\na.delete_if { |n| n.even? }\n");
}

// ── keep_if ───────────────────────────────────────────────────────────────────

#[test]
fn arr_keep_if_block() {
    compile_ok("a = [1, 2, 3, 4]\na.keep_if { |n| n.even? }\n");
}

// ── collect_concat (flat_map alias) ──────────────────────────────────────────

#[test]
fn arr_collect_concat_alias() {
    compile_ok("x = [[1, 2], [3]].collect_concat { |a| a }\n");
}

// ── inject with symbol ────────────────────────────────────────────────────────

#[test]
fn arr_inject_with_symbol() {
    compile_ok("x = [1, 2, 3, 4].inject(:+)\n");
}

#[test]
fn arr_inject_symbol_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3, 4].inject(:+)\n"), "10");
}

// ── inject with initial value ─────────────────────────────────────────────────

#[test]
fn arr_inject_with_initial_value() {
    compile_ok("x = [1, 2, 3].inject(10) { |sum, n| sum + n }\n");
}

#[test]
fn arr_inject_initial_runtime() {
    assert_eq!(
        run_ruby_one("puts [1, 2, 3].inject(10) { |sum, n| sum + n }\n"),
        "16"
    );
}

// ── chunk ─────────────────────────────────────────────────────────────────────

#[test]
fn arr_chunk_consecutive_groups() {
    compile_ok("[1, 1, 2, 2, 3].chunk { |n| n }.to_a\n");
}

// ── chunk_while ───────────────────────────────────────────────────────────────

#[test]
fn arr_chunk_while_consecutive_pairs() {
    compile_ok("[1, 2, 3, 5, 6, 10].chunk_while { |a, b| b == a + 1 }.to_a\n");
}

// ── each_with_index returning enumerator ─────────────────────────────────────

#[test]
fn arr_each_with_index_values() {
    let out = run_ruby("['a', 'b', 'c'].each_with_index { |v, i| puts v }\n");
    assert_eq!(out, vec!["a", "b", "c"]);
}

// ── filter_map ────────────────────────────────────────────────────────────────

#[test]
fn arr_filter_map_combined() {
    compile_ok("x = [1, 2, 3, 4, 5].filter_map { |n| n * 2 if n.odd? }\n");
}

// ── sum with initial value ────────────────────────────────────────────────────

#[test]
fn arr_sum_with_initial() {
    compile_ok("x = [1, 2, 3].sum(10)\n");
}

#[test]
fn arr_sum_initial_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3].sum(10)\n"), "16");
}

// ── uniq with block ───────────────────────────────────────────────────────────

#[test]
fn arr_uniq_with_block() {
    compile_ok("x = ['apple', 'Banana', 'cherry'].uniq { |s| s.downcase[0] }\n");
}

// ── Array.new with size and default ───────────────────────────────────────────

#[test]
fn arr_new_with_size_and_default() {
    compile_ok("x = Array.new(5, 0)\n");
}

#[test]
fn arr_new_default_runtime() {
    assert_eq!(run_ruby_one("puts Array.new(5, 0).length\n"), "5");
}

// ── Array.new with size and block ─────────────────────────────────────────────

#[test]
fn arr_new_with_size_and_block() {
    compile_ok("x = Array.new(5) { |i| i * 2 }\n");
}

#[test]
fn arr_new_block_runtime() {
    let out = run_ruby("puts Array.new(3) { |i| i * 2 }.first\n");
    assert_eq!(out, vec!["0"]);
}
