use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── reduce / inject ─────────────────────────────────────────────────────────

#[test]
fn reduce_symbol_plus() {
    compile_ok("[1, 2, 3, 4].reduce(:+)\n");
}

#[test]
fn reduce_initial_and_symbol() {
    compile_ok("[1, 2, 3].reduce(10, :+)\n");
}

#[test]
fn reduce_with_block() {
    compile_ok("x = [1, 2, 3, 4].reduce { |sum, n| sum + n }\n");
}

#[test]
fn inject_alias() {
    compile_ok("x = [1, 2, 3].inject(0) { |acc, n| acc + n }\n");
}

// ── flat_map ────────────────────────────────────────────────────────────────

#[test]
fn flat_map_basic() {
    compile_ok("x = [1, 2, 3].flat_map { |n| [n, n * 2] }\n");
}

#[test]
fn flat_map_nested_array() {
    compile_ok("x = [[1, 2], [3, 4]].flat_map { |a| a }\n");
}

// ── each_with_object ────────────────────────────────────────────────────────

#[test]
fn each_with_object_hash() {
    compile_ok("x = ['a', 'b', 'c'].each_with_object({}) { |s, h| h[s] = s.upcase }\n");
}

// ── group_by ────────────────────────────────────────────────────────────────

#[test]
fn group_by_even_odd() {
    compile_ok("x = [1, 2, 3, 4, 5, 6].group_by { |n| n % 2 == 0 ? 'even' : 'odd' }\n");
}

// ── sort_by ─────────────────────────────────────────────────────────────────

#[test]
fn sort_by_string_length() {
    compile_ok("x = ['banana', 'apple', 'fig'].sort_by { |s| s.length }\n");
}

// ── min_by / max_by ─────────────────────────────────────────────────────────

#[test]
fn min_by_length() {
    compile_ok("x = ['banana', 'fig', 'apple'].min_by { |s| s.length }\n");
}

#[test]
fn max_by_length() {
    compile_ok("x = ['banana', 'fig', 'apple'].max_by { |s| s.length }\n");
}

// ── minmax / minmax_by ──────────────────────────────────────────────────────

#[test]
fn minmax_array() {
    compile_ok("x = [3, 1, 4, 1, 5].minmax\n");
}

#[test]
fn minmax_by_block() {
    compile_ok("x = ['banana', 'fig', 'apple'].minmax_by { |s| s.length }\n");
}

// ── each_slice / each_cons ──────────────────────────────────────────────────

#[test]
fn each_slice_chunks() {
    compile_ok("[1, 2, 3, 4, 5].each_slice(2) { |s| puts s.length }\n");
}

#[test]
fn each_cons_window() {
    compile_ok("[1, 2, 3, 4, 5].each_cons(3) { |c| puts c.length }\n");
}

// ── chunk / chunk_while ─────────────────────────────────────────────────────

#[test]
fn chunk_consecutive() {
    compile_ok("x = [1, 1, 2, 2, 3].chunk { |n| n }.to_a\n");
}

#[test]
fn chunk_while_consecutive() {
    compile_ok("x = [1, 2, 3, 5, 6, 10].chunk_while { |a, b| b == a + 1 }.to_a\n");
}

// ── zip ─────────────────────────────────────────────────────────────────────

#[test]
fn zip_two_arrays() {
    compile_ok("x = [1, 2, 3].zip([4, 5, 6])\n");
}

// ── take / take_while ───────────────────────────────────────────────────────

#[test]
fn take_n_elements() {
    compile_ok("x = [1, 2, 3, 4, 5].take(3)\n");
}

#[test]
fn take_while_condition() {
    compile_ok("x = [1, 2, 3, 4, 5].take_while { |n| n < 4 }\n");
}

// ── drop / drop_while ───────────────────────────────────────────────────────

#[test]
fn drop_n_elements() {
    compile_ok("x = [1, 2, 3, 4, 5].drop(2)\n");
}

#[test]
fn drop_while_condition() {
    compile_ok("x = [1, 2, 3, 4, 5].drop_while { |n| n < 3 }\n");
}

// ── find / detect ───────────────────────────────────────────────────────────

#[test]
fn find_first_match() {
    compile_ok("x = [1, 2, 3, 4].find { |n| n > 2 }\n");
}

#[test]
fn detect_alias() {
    compile_ok("x = [1, 2, 3, 4].detect { |n| n.even? }\n");
}

// ── find_index / index ──────────────────────────────────────────────────────

#[test]
fn find_index_by_value() {
    compile_ok("x = [10, 20, 30].find_index(20)\n");
}

#[test]
fn find_index_by_block() {
    compile_ok("x = [10, 20, 30].find_index { |n| n > 15 }\n");
}

// ── count ───────────────────────────────────────────────────────────────────

#[test]
fn count_no_args() {
    compile_ok("x = [1, 2, 3].count\n");
}

#[test]
fn count_with_block() {
    compile_ok("x = [1, 2, 3, 4, 5].count { |n| n.odd? }\n");
}

// ── tally ───────────────────────────────────────────────────────────────────

#[test]
fn tally_occurrences() {
    compile_ok("x = ['a', 'b', 'a', 'c', 'b', 'a'].tally\n");
}

// ── any? / all? / none? / one? ──────────────────────────────────────────────

#[test]
fn any_with_block() {
    compile_ok("x = [1, 2, 3].any? { |n| n > 2 }\n");
}

#[test]
fn all_with_block() {
    compile_ok("x = [2, 4, 6].all? { |n| n.even? }\n");
}

#[test]
fn none_with_block() {
    compile_ok("x = [1, 3, 5].none? { |n| n.even? }\n");
}

#[test]
fn one_exactly_one() {
    compile_ok("x = [1, 2, 3].one? { |n| n == 2 }\n");
}

// ── include? / member? ──────────────────────────────────────────────────────

#[test]
fn include_membership() {
    compile_ok("x = [1, 2, 3].include?(2)\n");
}

#[test]
fn member_alias() {
    compile_ok("x = [1, 2, 3].member?(4)\n");
}

// ── sum ─────────────────────────────────────────────────────────────────────

#[test]
fn sum_numeric() {
    compile_ok("x = [1, 2, 3, 4, 5].sum\n");
}

#[test]
fn sum_with_block_transform() {
    compile_ok("x = [1, 2, 3].sum { |n| n * 2 }\n");
}

// ── filter_map ──────────────────────────────────────────────────────────────

#[test]
fn filter_map_combined() {
    compile_ok("x = [1, 2, 3, 4, 5].filter_map { |n| n * 2 if n.odd? }\n");
}

// ── first / to_a / entries ──────────────────────────────────────────────────

#[test]
fn first_element() {
    compile_ok("x = [10, 20, 30].first\n");
}

#[test]
fn first_n_elements() {
    compile_ok("x = [10, 20, 30, 40].first(2)\n");
}

#[test]
fn to_a_conversion() {
    compile_ok("x = [1, 2, 3].to_a\n");
}

#[test]
fn entries_alias() {
    compile_ok("x = [1, 2, 3].entries\n");
}

// ── Custom class with Enumerable ────────────────────────────────────────────

#[test]
fn custom_class_enumerable() {
    compile_ok(
        "module Enumerable\nend\n\
         class NumberBag\n  include Enumerable\n\
         def initialize(arr)\n    @data = arr\n  end\n\
         def each(&block)\n    @data.each(&block)\n  end\nend\n\
         bag = NumberBag.new([1, 2, 3])\nbag.each { |n| puts n }\n",
    );
}

// ── Runtime smoke tests ─────────────────────────────────────────────────────

#[test]
fn reduce_block_runtime() {
    assert_eq!(
        run_ruby_one(
            "[1, 2, 3, 4].reduce(0) { |sum, n| sum + n }\nputs [1, 2, 3, 4].reduce(0) { |sum, n| sum + n }\n"
        ),
        "10"
    );
}

#[test]
fn take_runtime() {
    let out = run_ruby("[1, 2, 3, 4, 5].take(3).each { |n| puts n }\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn count_with_block_runtime() {
    assert_eq!(
        run_ruby_one("puts [1, 2, 3, 4, 5].count { |n| n.odd? }\n"),
        "3"
    );
}

#[test]
fn find_runtime() {
    assert_eq!(run_ruby_one("puts [1, 2, 3, 4].find { |n| n > 2 }\n"), "3");
}
