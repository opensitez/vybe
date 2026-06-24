use crate::helpers::{run_in_main, run_main};

#[test]
fn hashmap_compute_inserts_when_key_absent() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.compute(\"k\", (key, val) -> 10); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn hashmap_compute_replaces_when_key_present() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); map.compute(\"k\", (key, val) -> val + 4); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn hashmap_compute_return_value_is_function_result() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.compute(\"k\", (key, val) -> 3));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashmap_compute_removing_entry_when_function_returns_null() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); map.compute(\"k\", (key, val) -> null); System.out.println(map.containsKey(\"k\"));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashmap_compute_if_present_updates_existing_mapping() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 2); map.computeIfPresent(\"k\", (key, val) -> val * 3); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn hashmap_compute_if_present_absent_key_is_noop() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.computeIfPresent(\"k\", (key, val) -> 9); System.out.println(map.containsKey(\"k\"));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashmap_compute_if_present_removes_entry_when_function_returns_null() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); map.computeIfPresent(\"k\", (key, val) -> null); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn hashmap_merge_inserts_when_key_absent() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.merge(\"k\", 5, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn hashmap_merge_combines_existing_values_with_remapping_function() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 2); map.merge(\"k\", 3, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn hashmap_merge_return_value_reflects_final_mapping() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.merge(\"k\", 7, (a, b) -> a + b));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn hashmap_replace_key_old_new_succeeds_when_value_matches() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 4); System.out.println(map.replace(\"k\", 4, 9)); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["true", "9"]);
}

#[test]
fn hashmap_replace_key_old_new_fails_on_value_mismatch() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 4); System.out.println(map.replace(\"k\", 1, 9)); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["false", "4"]);
}

#[test]
fn hashmap_replace_key_old_new_false_when_key_missing() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.replace(\"k\", 1, 9));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashmap_replace_all_updates_every_stored_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.replaceAll((key, val) -> val + 10); System.out.println(map.get(\"a\")); System.out.println(map.get(\"b\"));",
    );
    assert_eq!(out, vec!["11", "12"]);
}

#[test]
fn hashmap_replace_all_can_rename_values_from_keys() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"x\", 1); map.put(\"y\", 2); map.replaceAll((key, val) -> key.length()); System.out.println(map.get(\"x\")); System.out.println(map.get(\"y\"));",
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn hashmap_for_each_accumulates_values_via_biconsumer() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); final int[] sum = {0}; map.forEach((key, val) -> { sum[0] += val; }); System.out.println(sum[0]);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashmap_for_each_reads_keys_in_insertion_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"first\", 1); map.put(\"second\", 2); final String[] seen = {\"\"}; map.forEach((key, val) -> { seen[0] = seen[0] + key + \",\"; }); System.out.println(seen[0]);",
    );
    assert_eq!(out, vec!["first,second,"]);
}

#[test]
fn hashmap_entry_set_iterator_has_next_on_nonempty_map() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"alpha\", 1); java.util.Iterator<java.util.Map.Entry<String, Integer>> it = map.entrySet().iterator(); System.out.println(it.hasNext());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashmap_entry_set_get_key_from_first_entry() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"key\", 9); java.util.Map.Entry<String, Integer> entry = map.entrySet().iterator().next(); System.out.println(entry.getKey());",
    );
    assert_eq!(out, vec!["key"]);
}

#[test]
fn hashmap_entry_set_get_value_from_first_entry() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"key\", 9); java.util.Map.Entry<String, Integer> entry = map.entrySet().iterator().next(); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn hashmap_entry_set_set_value_updates_map() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); java.util.Map.Entry<String, Integer> entry = map.entrySet().iterator().next(); entry.setValue(8); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn linkedhashmap_preserves_insertion_order_on_first_key() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"b\", 2); map.put(\"a\", 1); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn linkedhashmap_iteration_follows_insertion_sequence() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"one\", 1); map.put(\"two\", 2); map.put(\"three\", 3); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["one", "two", "three"]);
}

#[test]
fn linkedhashmap_get_does_not_reorder_insertion_mode() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.get(\"a\"); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn treemap_natural_order_first_key_is_minimum() {
    let out = run_main(
        "java.util.TreeMap<String, Integer> map = new java.util.TreeMap<String, Integer>(); map.put(\"c\", 3); map.put(\"a\", 1); map.put(\"b\", 2); System.out.println(map.firstKey());",
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn treemap_natural_order_last_key_is_maximum() {
    let out = run_main(
        "java.util.TreeMap<String, Integer> map = new java.util.TreeMap<String, Integer>(); map.put(\"c\", 3); map.put(\"a\", 1); map.put(\"b\", 2); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["c"]);
}

#[test]
fn treemap_get_returns_value_for_existing_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(2, \"two\"); map.put(1, \"one\"); System.out.println(map.get(2));",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn treemap_key_set_iterator_returns_sorted_order() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(30, \"c\"); map.put(10, \"a\"); map.put(20, \"b\"); java.util.Iterator<Integer> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn navigable_map_submap_includes_lower_inclusive_bound() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); map.put(4, \"d\"); java.util.SortedMap<Integer, String> sub = map.subMap(2, 4); System.out.println(sub.containsKey(2)); System.out.println(sub.containsKey(4));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn navigable_map_submap_excludes_outside_keys() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); java.util.SortedMap<Integer, String> sub = map.subMap(2, 4); System.out.println(sub.containsKey(1)); System.out.println(sub.containsKey(3));",
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn navigable_map_head_map_excludes_to_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); java.util.SortedMap<Integer, String> head = map.headMap(3); System.out.println(head.containsKey(2)); System.out.println(head.containsKey(3));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn navigable_map_tail_map_includes_from_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); java.util.SortedMap<Integer, String> tail = map.tailMap(2); System.out.println(tail.containsKey(2)); System.out.println(tail.containsKey(1));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn treemap_ceiling_entry_finds_next_equal_or_higher_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(10, \"ten\"); map.put(20, \"twenty\"); java.util.Map.Entry<Integer, String> entry = map.ceilingEntry(15); System.out.println(entry.getKey());",
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn treemap_floor_entry_finds_next_equal_or_lower_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(10, \"ten\"); map.put(20, \"twenty\"); java.util.Map.Entry<Integer, String> entry = map.floorEntry(15); System.out.println(entry.getKey());",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn treemap_higher_entry_finds_strictly_greater_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(10, \"ten\"); map.put(20, \"twenty\"); java.util.Map.Entry<Integer, String> entry = map.higherEntry(10); System.out.println(entry.getKey());",
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn treemap_lower_entry_finds_strictly_smaller_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(10, \"ten\"); map.put(20, \"twenty\"); java.util.Map.Entry<Integer, String> entry = map.lowerEntry(20); System.out.println(entry.getKey());",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn hashmap_compute_if_absent_not_invoked_when_key_exists() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 5); map.computeIfAbsent(\"k\", key -> 99); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn hashmap_merge_with_zero_delta_keeps_existing_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 4); map.merge(\"k\", 0, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn hashmap_entry_set_size_matches_put_count() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); System.out.println(map.entrySet().size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linkedhashmap_values_collection_reflects_insertion_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 10); map.put(\"b\", 20); java.util.Iterator<Integer> it = map.values().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn treemap_submap_size_counts_in_range_entries() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); map.put(4, \"d\"); System.out.println(map.subMap(2, 4).size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treemap_head_map_size_counts_keys_below_bound() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); System.out.println(map.headMap(3).size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treemap_tail_map_size_counts_keys_from_bound_upward() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); System.out.println(map.tailMap(2).size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashmap_for_each_can_build_concatenated_key_list() {
    let out = run_in_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"x\", 1); map.put(\"y\", 2); System.out.println(Keys.join(map));",
        "static class Keys { static String join(java.util.Map<String, Integer> map) { final String[] out = {\"\"}; map.forEach((k, v) -> { out[0] = out[0] + k; }); return out[0]; } }",
    );
    assert_eq!(out, vec!["xy"]);
}

#[test]
fn treemap_first_entry_returns_lowest_key_value_pair() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(5, \"five\"); map.put(1, \"one\"); java.util.Map.Entry<Integer, String> entry = map.firstEntry(); System.out.println(entry.getKey()); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["1", "one"]);
}

#[test]
fn treemap_last_entry_returns_highest_key_value_pair() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(5, \"five\"); map.put(1, \"one\"); java.util.Map.Entry<Integer, String> entry = map.lastEntry(); System.out.println(entry.getKey()); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["5", "five"]);
}

#[test]
fn hashmap_replace_all_on_empty_map_is_noop() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.replaceAll((key, val) -> val + 1); System.out.println(map.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linkedhashmap_put_overwrite_does_not_change_key_order() {
    let out = run_main(
        "java.util.LinkedHashMap<String, Integer> map = new java.util.LinkedHashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"a\", 9); java.util.Iterator<String> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn treemap_navigable_submap_get_reads_interior_value() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); System.out.println(map.subMap(1, 3).get(2));",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn hashmap_merge_removes_entry_when_remapping_returns_null() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); map.merge(\"k\", 2, (oldVal, newVal) -> null); System.out.println(map.containsKey(\"k\"));",
    );
    assert_eq!(out, vec!["false"]);
}
