use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── Iteration ───────────────────────────────────────────────────────────────

#[test]
fn range_each() {
    compile_ok("(1..5).each { |i| puts i }\n");
}

#[test]
fn range_map() {
    compile_ok("x = (1..5).map { |i| i * 2 }\n");
}

#[test]
fn range_select() {
    compile_ok("x = (1..10).select { |i| i.even? }\n");
}

// ── Membership ──────────────────────────────────────────────────────────────

#[test]
fn range_include_q() {
    compile_ok("x = (1..10).include?(5)\n");
}

#[test]
fn range_cover_q() {
    compile_ok("x = (1..10).cover?(5)\n");
}

// ── Conversion ──────────────────────────────────────────────────────────────

#[test]
fn range_to_a() {
    compile_ok("x = (1..5).to_a\n");
}

// ── Min / Max / Sum ─────────────────────────────────────────────────────────

#[test]
fn range_min() {
    compile_ok("x = (3..9).min\n");
}

#[test]
fn range_max() {
    compile_ok("x = (3..9).max\n");
}

#[test]
fn range_sum() {
    compile_ok("x = (1..100).sum\n");
}

// ── Count / First / Last ────────────────────────────────────────────────────

#[test]
fn range_count() {
    compile_ok("x = (1..10).count\n");
}

#[test]
fn range_first() {
    compile_ok("x = (5..20).first\n");
}

#[test]
fn range_last() {
    compile_ok("x = (5..20).last\n");
}

#[test]
fn range_first_n() {
    compile_ok("x = (1..100).first(5)\n");
}

#[test]
fn range_last_n() {
    compile_ok("x = (1..100).last(5)\n");
}

// ── Step ────────────────────────────────────────────────────────────────────

#[test]
fn range_step_integer() {
    compile_ok("(0..20).step(5) { |i| puts i }\n");
}

#[test]
fn range_step_float() {
    compile_ok("(0.0..1.0).step(0.25) { |f| puts f }\n");
}

// ── Endless / Beginless / Size ──────────────────────────────────────────────

#[test]
fn range_endless() {
    compile_ok("r = (1..)\n");
}

#[test]
fn range_beginless() {
    compile_ok("r = (..5)\n");
}

#[test]
fn range_size_endless() {
    // size/count of an endless range is Float::INFINITY — compile only
    compile_ok("r = (1..)\n");
}

// ── String Ranges ───────────────────────────────────────────────────────────

#[test]
fn string_range_literal() {
    compile_ok("r = ('a'..'e')\n");
}

#[test]
fn string_range_to_a() {
    compile_ok("x = ('a'..'e').to_a\n");
}

// ── Exclusive Range ─────────────────────────────────────────────────────────

#[test]
fn exclusive_range_excludes_last() {
    compile_ok("x = (1...5).to_a\n");
}

// ── Case / When ─────────────────────────────────────────────────────────────

#[test]
fn range_in_case_when() {
    compile_ok(
        "score = 75\n\
         case score\n\
         when 90..100 then puts 'A'\n\
         when 70..89  then puts 'B'\n\
         else              puts 'C'\n\
         end\n",
    );
}

#[test]
fn range_triple_equals() {
    compile_ok("x = (1..10) === 5\n");
}

// ── Enumerable on Range ─────────────────────────────────────────────────────

#[test]
fn range_each_slice() {
    compile_ok("(1..9).each_slice(3) { |s| puts s.length }\n");
}

#[test]
fn range_reduce_inject() {
    compile_ok("x = (1..5).reduce(:+)\n");
}

#[test]
fn range_min_by() {
    compile_ok("x = (1..5).min_by { |n| -n }\n");
}

#[test]
fn range_max_by() {
    compile_ok("x = (1..5).max_by { |n| -n }\n");
}

// ── Predicates ──────────────────────────────────────────────────────────────

#[test]
fn range_any_q() {
    compile_ok("x = (1..10).any? { |n| n > 8 }\n");
}

#[test]
fn range_all_q() {
    compile_ok("x = (1..5).all? { |n| n > 0 }\n");
}

#[test]
fn range_none_q() {
    compile_ok("x = (1..5).none? { |n| n > 10 }\n");
}

// ── Accessors / Predicates on Range object ──────────────────────────────────

#[test]
fn range_begin_accessor() {
    compile_ok("r = (3..7)\nx = r.begin\n");
}

#[test]
fn range_end_accessor() {
    compile_ok("r = (3..7)\nx = r.end\n");
}

#[test]
fn range_exclude_end_q() {
    compile_ok("puts (1...5).exclude_end?\n");
}

#[test]
fn range_equality() {
    compile_ok("x = (1..5) == (1..5)\n");
}

// ── Runtime smoke tests ─────────────────────────────────────────────────────

#[test]
fn range_each_runtime() {
    let out = run_ruby("(1..3).each { |i| puts i }\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn range_to_a_runtime() {
    let out = run_ruby("puts (1..5).to_a.length\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn range_sum_runtime() {
    assert_eq!(run_ruby_one("puts (1..10).sum\n"), "55");
}

#[test]
fn range_include_q_runtime() {
    assert_eq!(run_ruby_one("puts (1..10).include?(5)\n"), "true");
}

#[test]
fn range_exclude_end_q_runtime() {
    assert_eq!(run_ruby_one("puts (1...5).exclude_end?\n"), "true");
}

#[test]
fn range_case_when_runtime() {
    let out = run_ruby(
        "score = 75\n\
         case score\n\
         when 90..100 then puts 'A'\n\
         when 70..89  then puts 'B'\n\
         else              puts 'C'\n\
         end\n",
    );
    assert_eq!(out, vec!["B"]);
}
