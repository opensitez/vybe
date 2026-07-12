use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── Iteration ───────────────────────────────────────────────────────────────

#[test]
fn hash_each() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nh.each { |k, v| puts k }\n");
}

#[test]
fn hash_each_pair() {
    compile_ok("h = {'x' => 10}\nh.each_pair { |k, v| puts v }\n");
}

#[test]
fn hash_each_key() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nh.each_key { |k| puts k }\n");
}

#[test]
fn hash_each_value() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nh.each_value { |v| puts v }\n");
}

// ── Transformation ──────────────────────────────────────────────────────────

#[test]
fn hash_map_collect() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.map { |k, v| [k, v * 2] }\n");
}

#[test]
fn hash_select_filter() {
    compile_ok("h = {'a' => 1, 'b' => 2, 'c' => 3}\nx = h.select { |k, v| v > 1 }\n");
}

#[test]
fn hash_reject() {
    compile_ok("h = {'a' => 1, 'b' => 2, 'c' => 3}\nx = h.reject { |k, v| v == 2 }\n");
}

#[test]
fn hash_transform_keys() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.transform_keys { |k| k.upcase }\n");
}

#[test]
fn hash_transform_values() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.transform_values { |v| v * 10 }\n");
}

#[test]
fn hash_filter_map() {
    compile_ok(
        "h = {'a' => 1, 'b' => 2, 'c' => 3}\nx = h.filter_map { |k, v| [k, v * 2] if v > 1 }\n",
    );
}

// ── Predicates ──────────────────────────────────────────────────────────────

#[test]
fn hash_any_q() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.any? { |k, v| v > 1 }\n");
}

#[test]
fn hash_all_q() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.all? { |k, v| v > 0 }\n");
}

#[test]
fn hash_none_q() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.none? { |k, v| v > 5 }\n");
}

#[test]
fn hash_empty_q() {
    compile_ok("h = {}\nx = h.empty?\n");
}

#[test]
fn hash_include_q() {
    compile_ok("h = {'a' => 1}\nx = h.include?('a')\n");
}

#[test]
fn hash_has_value_q() {
    compile_ok("h = {'a' => 1}\nx = h.has_value?(1)\n");
}

#[test]
fn hash_value_q() {
    compile_ok("h = {'a' => 42}\nx = h.value?(42)\n");
}

#[test]
fn hash_key_q() {
    compile_ok("h = {'a' => 1}\nx = h.key?('a')\n");
}

#[test]
fn hash_member_q() {
    compile_ok("h = {'a' => 1}\nx = h.member?('a')\n");
}

// ── Counting / Size ─────────────────────────────────────────────────────────

#[test]
fn hash_count_no_args() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.count\n");
}

#[test]
fn hash_count_with_block() {
    compile_ok("h = {'a' => 1, 'b' => 2, 'c' => 3}\nx = h.count { |k, v| v > 1 }\n");
}

#[test]
fn hash_length() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.length\n");
}

#[test]
fn hash_size() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.size\n");
}

// ── Conversion ──────────────────────────────────────────────────────────────

#[test]
fn hash_to_a() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.to_a\n");
}

#[test]
fn hash_flatten() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.flatten\n");
}

#[test]
fn hash_invert() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.invert\n");
}

#[test]
fn hash_to_s() {
    compile_ok("h = {'a' => 1}\nx = h.to_s\n");
}

// ── Search / Lookup ─────────────────────────────────────────────────────────

#[test]
fn hash_min_by() {
    compile_ok("h = {'a' => 3, 'b' => 1, 'c' => 2}\nx = h.min_by { |k, v| v }\n");
}

#[test]
fn hash_max_by() {
    compile_ok("h = {'a' => 3, 'b' => 1, 'c' => 2}\nx = h.max_by { |k, v| v }\n");
}

#[test]
fn hash_sort_by() {
    compile_ok("h = {'b' => 2, 'a' => 1}\nx = h.sort_by { |k, v| k }\n");
}

#[test]
fn hash_dig() {
    compile_ok("h = {'a' => {'b' => 42}}\nx = h.dig('a', 'b')\n");
}

#[test]
fn hash_slice() {
    compile_ok("h = {'a' => 1, 'b' => 2, 'c' => 3}\nx = h.slice('a', 'c')\n");
}

#[test]
fn hash_assoc() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.assoc('a')\n");
}

#[test]
fn hash_rassoc() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.rassoc(2)\n");
}

#[test]
fn hash_key_by_value() {
    compile_ok("h = {'a' => 1, 'b' => 2}\nx = h.key(2)\n");
}

// ── Mutation ────────────────────────────────────────────────────────────────

#[test]
fn hash_compact() {
    compile_ok("h = {'a' => 1, 'b' => nil, 'c' => 3}\nx = h.compact\n");
}

#[test]
fn hash_merge_with_block() {
    compile_ok(
        "h1 = {'a' => 1}\nh2 = {'a' => 2}\nx = h1.merge(h2) { |key, old, new_v| old + new_v }\n",
    );
}

#[test]
fn hash_merge_bang() {
    compile_ok("h = {'a' => 1}\nh.merge!({'b' => 2})\n");
}

#[test]
fn hash_update() {
    compile_ok("h = {'a' => 1}\nh.update({'b' => 2})\n");
}

#[test]
fn hash_store() {
    compile_ok("h = {}\nh.store('key', 'value')\n");
}

#[test]
fn hash_freeze() {
    compile_ok("h = {'a' => 1}\nh.freeze\n");
}

// ── Construction ────────────────────────────────────────────────────────────

#[test]
fn hash_new_default_value() {
    compile_ok("h = Hash.new(0)\nh['missing']\n");
}

#[test]
fn hash_new_default_proc() {
    compile_ok("h = Hash.new { |hash, key| hash[key] = key.upcase }\nh['hello']\n");
}

// ── Nested access / Equality ────────────────────────────────────────────────

#[test]
fn hash_nested_access() {
    compile_ok("h = {'outer' => {'inner' => 42}}\nx = h['outer']['inner']\n");
}

#[test]
fn hash_equality() {
    compile_ok("h1 = {'a' => 1}\nh2 = {'a' => 1}\nx = h1 == h2\n");
}

// ── Runtime smoke tests ─────────────────────────────────────────────────────

#[test]
fn hash_each_runtime() {
    let out = run_ruby("h = {'a' => 1, 'b' => 2}\nh.each_key { |k| puts k }\n");
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn hash_count_runtime() {
    assert_eq!(
        run_ruby_one("h = {'a' => 1, 'b' => 2, 'c' => 3}\nputs h.count\n"),
        "3"
    );
}

#[test]
fn hash_empty_q_runtime() {
    assert_eq!(run_ruby_one("puts({}.empty?)\n"), "true");
}

#[test]
fn hash_length_runtime() {
    assert_eq!(
        run_ruby_one("h = {'x' => 1, 'y' => 2}\nputs h.length\n"),
        "2"
    );
}
