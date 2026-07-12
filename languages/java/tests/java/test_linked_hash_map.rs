use crate::helpers::{run_in_main, run_main};

#[test]
fn linkedhashmap_insertion_order_preserves_first_inserted_key() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"second\", 2); map.put(\"first\", 1); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["second"]);
}

#[test]
fn linkedhashmap_insertion_order_iterates_three_keys_in_sequence() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"alpha\", 1); map.put(\"beta\", 2); map.put(\"gamma\", 3); for (String key : map.keySet()) { System.out.println(key); }",
    );
    assert_eq!(out, vec!["alpha", "beta", "gamma"]);
}

#[test]
fn linkedhashmap_values_iterator_follows_insertion_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 10); map.put(\"b\", 20); map.put(\"c\", 30); java.util.Iterator<Integer> it = map.values().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn linkedhashmap_entry_set_iterator_preserves_insertion_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"x\", 1); map.put(\"y\", 2); java.util.Iterator<java.util.Map.Entry<String, Integer>> it = map.entrySet().iterator(); java.util.Map.Entry<String, Integer> first = it.next(); System.out.println(first.getKey()); System.out.println(first.getValue());",
    );
    assert_eq!(out, vec!["x", "1"]);
}

#[test]
fn linkedhashmap_get_does_not_reorder_in_insertion_mode() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"first\", 1); map.put(\"second\", 2); map.get(\"first\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn linkedhashmap_put_overwrite_retains_original_key_position() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"a\", 9); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn linkedhashmap_remove_then_iterate_shows_remaining_insertion_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"keep\", 1); map.put(\"drop\", 2); map.put(\"also\", 3); map.remove(\"drop\"); for (String key : map.keySet()) { System.out.println(key); }",
    );
    assert_eq!(out, vec!["keep", "also"]);
}

#[test]
fn linkedhashmap_reinsert_removed_key_appends_at_end() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.remove(\"a\"); map.put(\"a\", 3); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "a"]);
}

#[test]
fn linkedhashmap_access_order_get_moves_touched_key_to_end() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); map.get(\"a\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "c", "a"]);
}

#[test]
fn linkedhashmap_access_order_put_on_existing_key_moves_to_end() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); map.put(\"a\", 9); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "c", "a"]);
}

#[test]
fn linkedhashmap_access_order_two_gets_move_both_keys_to_tail() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); map.get(\"a\"); map.get(\"b\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["c", "a", "b"]);
}

#[test]
fn linkedhashmap_access_order_lru_touch_sequence_reorders_keys() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"one\", 1); map.put(\"two\", 2); map.put(\"three\", 3); map.get(\"one\"); map.get(\"two\"); map.get(\"three\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["one", "two", "three"]);
}

#[test]
fn linkedhashmap_access_order_get_middle_entry_promotes_it_last() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"left\", 1); map.put(\"mid\", 2); map.put(\"right\", 3); map.get(\"mid\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["left", "right", "mid"]);
}

#[test]
fn linkedhashmap_insertion_order_integer_keys_follow_put_sequence() {
    let out = run_main(
        "java.util.LinkedHashMap<Integer, String> map = new java.util.LinkedHashMap<Integer, String>(); map.put(30, \"c\"); map.put(10, \"a\"); map.put(20, \"b\"); java.util.Iterator<Integer> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["30", "10", "20"]);
}

#[test]
fn linkedhashmap_access_order_integer_get_reorders_to_tail() {
    let out = run_main(
        "java.util.LinkedHashMap<Integer, String> map = new java.util.LinkedHashMap<Integer, String>(16, 0.75f, true); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); map.get(1); java.util.Iterator<Integer> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["2", "3", "1"]);
}

#[test]
fn linkedhashmap_put_all_appends_source_entries_in_iteration_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> base = new java.util.LinkedHashMap<String, Integer>(); base.put(\"a\", 1); java.util.LinkedHashMap<String, Integer> extra = new java.util.LinkedHashMap<String, Integer>(); extra.put(\"b\", 2); extra.put(\"c\", 3); base.putAll(extra); java.util.Iterator<String> it = base.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn linkedhashmap_clear_then_reinsert_starts_fresh_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"old\", 1); map.clear(); map.put(\"new\", 2); System.out.println(map.keySet().iterator().next());",
    );
    assert_eq!(out, vec!["new"]);
}

#[test]
fn linkedhashmap_insertion_order_for_each_visits_puts_sequence() {
    let out = run_in_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"p\", 1); map.put(\"q\", 2); map.put(\"r\", 3); System.out.println(Joiner.keys(map));",
        "static class Joiner { static String keys(java.util.LinkedHashMap<String, Integer> map) { final String[] out = {\"\"}; map.forEach((k, v) -> { out[0] = out[0] + k; }); return out[0]; } }",
    );
    assert_eq!(out, vec!["pqr"]);
}

#[test]
fn linkedhashmap_access_order_contains_key_does_not_reorder() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.containsKey(\"a\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn linkedhashmap_access_order_put_if_absent_on_hit_does_not_reorder() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.putIfAbsent(\"a\", 9); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn linkedhashmap_access_order_put_if_absent_on_miss_inserts_at_end() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.putIfAbsent(\"b\", 2); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn linkedhashmap_insertion_order_after_remove_middle_keeps_relative_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); map.remove(\"b\"); for (String key : map.keySet()) { System.out.println(key); }",
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn linkedhashmap_access_order_replace_all_does_not_change_iteration_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.replaceAll((k, v) -> v + 10); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn linkedhashmap_access_order_get_or_default_on_hit_moves_key() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.getOrDefault(\"a\", 0); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "a"]);
}

#[test]
fn linkedhashmap_insertion_order_single_entry_iterator() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"only\", 1); System.out.println(map.keySet().iterator().next()); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["only", "1"]);
}

#[test]
fn linkedhashmap_access_order_single_get_keeps_lone_entry() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"solo\", 1); map.get(\"solo\"); System.out.println(map.keySet().iterator().next());",
    );
    assert_eq!(out, vec!["solo"]);
}

#[test]
fn linkedhashmap_insertion_order_remove_first_then_iterate() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"first\", 1); map.put(\"second\", 2); map.remove(\"first\"); System.out.println(map.keySet().iterator().next());",
    );
    assert_eq!(out, vec!["second"]);
}

#[test]
fn linkedhashmap_access_order_remove_then_get_reinserts_at_end() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.remove(\"a\"); map.put(\"a\", 3); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "a"]);
}

#[test]
fn linkedhashmap_values_order_tracks_key_insertion_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"k1\", 100); map.put(\"k2\", 200); map.put(\"k3\", 300); java.util.Iterator<Integer> it = map.values().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["100", "200", "300"]);
}

#[test]
fn linkedhashmap_access_order_chain_abca_produces_lru_tail() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); map.get(\"a\"); map.get(\"b\"); map.get(\"c\"); map.get(\"a\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "c", "a"]);
}

#[test]
fn linkedhashmap_insertion_order_null_value_preserves_key_position() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", null); map.put(\"c\", 3); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn linkedhashmap_access_order_compute_if_present_on_hit_moves_key() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.computeIfPresent(\"a\", (k, v) -> v + 1); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "a"]);
}

#[test]
fn linkedhashmap_access_order_merge_on_existing_key_moves_entry() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.merge(\"a\", 5, (oldVal, newVal) -> oldVal + newVal); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "a"]);
}

#[test]
fn linkedhashmap_insertion_order_new_key_after_many_puts_goes_last() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); map.put(\"d\", 4); java.util.Iterator<String> it = map.keySet().iterator(); String last = \"\"; while (it.hasNext()) { last = it.next(); } System.out.println(last);",
    );
    assert_eq!(out, vec!["d"]);
}

#[test]
fn linkedhashmap_access_order_touch_oldest_then_iterate_shows_promoted_last() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"oldest\", 1); map.put(\"middle\", 2); map.put(\"newest\", 3); map.get(\"oldest\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["middle", "newest", "oldest"]);
}

#[test]
fn linkedhashmap_insertion_order_key_set_size_matches_distinct_puts() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"a\", 9); System.out.println(map.keySet().size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linkedhashmap_access_order_lru_pattern_after_four_operations() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"w\", 1); map.put(\"x\", 2); map.put(\"y\", 3); map.put(\"z\", 4); map.get(\"w\"); map.get(\"x\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["y", "z", "w", "x"]);
}

#[test]
fn linkedhashmap_insertion_order_entry_set_size_matches_puts() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"one\", 1); map.put(\"two\", 2); System.out.println(map.entrySet().size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linkedhashmap_access_order_put_new_after_reads_keeps_untouched_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"a\", 1); map.put(\"b\", 2); map.get(\"a\"); map.put(\"c\", 3); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b", "a", "c"]);
}

#[test]
fn linkedhashmap_access_order_eldest_key_after_three_gets_is_first_iterator() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(16, 0.75f, true); map.put(\"eldest\", 1); map.put(\"mid\", 2); map.put(\"young\", 3); map.get(\"mid\"); map.get(\"young\"); map.get(\"eldest\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["mid"]);
}
