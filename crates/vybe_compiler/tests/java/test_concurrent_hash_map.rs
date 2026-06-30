/// java.util.concurrent.ConcurrentHashMap — thread-safe map operations.
use crate::helpers::{run_in_main, run_main};

#[test]
fn concurrent_hash_map_put_and_get_round_trip() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("alpha", 1); System.out.println(map.get("alpha"));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn concurrent_hash_map_new_map_is_empty() {
    let out = run_main(
        "java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); System.out.println(map.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_size_after_two_puts() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn concurrent_hash_map_contains_key_after_put() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 9); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_contains_value_after_put() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 9); System.out.println(map.containsValue(9));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_remove_returns_previous_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 4); System.out.println(map.remove("k"));"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn concurrent_hash_map_remove_leaves_map_empty() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 4); map.remove("k"); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_put_if_absent_inserts_when_missing() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); System.out.println(map.putIfAbsent("k", 7)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["null", "7"]);
}

#[test]
fn concurrent_hash_map_put_if_absent_skips_when_present() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.putIfAbsent("k", 9)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn concurrent_hash_map_compute_if_absent_inserts_when_missing() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.computeIfAbsent("k", key -> 11); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn concurrent_hash_map_compute_if_absent_leaves_existing_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 3); map.computeIfAbsent("k", key -> 99); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn concurrent_hash_map_compute_if_present_updates_existing() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 2); map.computeIfPresent("k", (key, val) -> val + 5); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn concurrent_hash_map_compute_if_present_noop_when_absent() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.computeIfPresent("k", (key, val) -> 9); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn concurrent_hash_map_compute_inserts_when_absent() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.compute("k", (key, val) -> 10); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn concurrent_hash_map_compute_updates_existing_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); map.compute("k", (key, val) -> val + 4); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn concurrent_hash_map_merge_inserts_when_key_absent() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.merge("k", 5, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn concurrent_hash_map_merge_combines_existing_values() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 2); map.merge("k", 3, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn concurrent_hash_map_replace_existing_key() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.replace("k", 8)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["1", "8"]);
}

#[test]
fn concurrent_hash_map_replace_conditional_succeeds_when_value_matches() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.replace("k", 1, 6)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["true", "6"]);
}

#[test]
fn concurrent_hash_map_replace_conditional_fails_when_value_differs() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.replace("k", 2, 6)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["false", "1"]);
}

#[test]
fn concurrent_hash_map_get_or_default_returns_default_when_absent() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); System.out.println(map.getOrDefault("missing", 42));"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn concurrent_hash_map_get_or_default_returns_existing_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 3); System.out.println(map.getOrDefault("k", 42));"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn concurrent_hash_map_replace_all_updates_every_entry() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); map.replaceAll((key, val) -> val * 10); System.out.println(map.get("a")); System.out.println(map.get("b"));"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn concurrent_hash_map_key_set_contains_inserted_key() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("x", 1); System.out.println(map.keySet().contains("x"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_values_contains_inserted_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("x", 5); System.out.println(map.values().contains(5));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_entry_set_iterator_reads_key_and_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 2); java.util.Map.Entry<String, Integer> entry = map.entrySet().iterator().next(); System.out.println(entry.getKey()); System.out.println(entry.getValue());"#,
    );
    assert_eq!(out, vec!["k", "2"]);
}

#[test]
fn concurrent_hash_map_clear_empties_all_entries() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); map.clear(); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn concurrent_hash_map_put_overwrites_existing_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); map.put("k", 9); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn concurrent_hash_map_put_returns_previous_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.put("k", 2));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn concurrent_hash_map_initial_capacity_constructor_accepts_entries() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(8); map.put("k", 1); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn concurrent_hash_map_for_each_prints_each_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); map.forEach((key, val) -> System.out.println(val));"#,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"1".to_string()));
    assert!(out.contains(&"2".to_string()));
}

#[test]
fn concurrent_hash_map_compute_removes_entry_when_function_returns_null() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); map.compute("k", (key, val) -> null); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn concurrent_hash_map_compute_if_present_removes_when_function_returns_null() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); map.computeIfPresent("k", (key, val) -> null); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn concurrent_hash_map_merge_removes_when_remapping_returns_null() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 1); map.merge("k", 2, (oldVal, newVal) -> null); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn concurrent_hash_map_mapping_count_reflects_size() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); System.out.println(map.mappingCount());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn concurrent_hash_map_new_key_set_behaves_as_set() {
    let out = run_main(
        r#"java.util.Set<String> set = java.util.concurrent.ConcurrentHashMap.newKeySet(); set.add("one"); set.add("two"); System.out.println(set.size()); System.out.println(set.contains("one"));"#,
    );
    assert_eq!(out, vec!["2", "true"]);
}

#[test]
fn concurrent_hash_map_new_key_set_remove_eliminates_element() {
    let out = run_main(
        r#"java.util.Set<String> set = java.util.concurrent.ConcurrentHashMap.newKeySet(); set.add("x"); set.remove("x"); System.out.println(set.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_reduce_values_sums_integers() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 2); map.put("b", 3); System.out.println(map.reduceValues(1L, (a, b) -> a + b));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn concurrent_hash_map_reduce_entries_concatenates_keys() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); String joined = map.reduceEntries(1L, (e1, e2) -> e1.getKey() + e2.getKey()); System.out.println(joined.length() == 2);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_search_values_finds_matching_entry() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 10); map.put("b", 20); Integer found = map.searchValues(1L, val -> val == 20 ? val : null); System.out.println(found);"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn concurrent_hash_map_search_keys_finds_matching_key() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("target", 1); map.put("other", 2); String found = map.searchKeys(1L, key -> key.equals("target") ? key : null); System.out.println(found);"#,
    );
    assert_eq!(out, vec!["target"]);
}

#[test]
fn concurrent_hash_map_two_threads_increment_same_key_atomically() {
    let types = r#"
        static java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>();
    "#;
    let out = run_in_main(
        r#"map.put("n", 0); Thread t1 = new Thread(() -> { for (int i = 0; i < 50; i++) map.compute("n", (k, v) -> v + 1); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 50; i++) map.compute("n", (k, v) -> v + 1); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(map.get("n"));"#,
        types,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn concurrent_hash_map_two_threads_put_distinct_keys() {
    let types = r#"
        static java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>();
    "#;
    let out = run_in_main(
        r#"Thread t1 = new Thread(() -> map.put("left", 1)); Thread t2 = new Thread(() -> map.put("right", 2)); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(map.size());"#,
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn concurrent_hash_map_put_all_merges_other_map() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); java.util.HashMap<String, Integer> other = new java.util.HashMap<String, Integer>(); other.put("a", 1); other.put("b", 2); map.putAll(other); System.out.println(map.size()); System.out.println(map.get("b"));"#,
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn concurrent_hash_map_remove_key_value_pair_conditional() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 5); System.out.println(map.remove("k", 5)); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn concurrent_hash_map_remove_key_value_pair_fails_on_mismatch() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", 5); System.out.println(map.remove("k", 9)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["false", "5"]);
}

#[test]
fn concurrent_hash_map_null_value_allowed() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("k", null); System.out.println(map.containsKey("k")); System.out.println(map.get("k") == null);"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn concurrent_hash_map_null_key_throws_null_pointer_exception() {
    let out = run_in_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); try { map.put(null, 1); System.out.println("ok"); } catch (NullPointerException e) { System.out.println("npe"); }"#,
        "",
    );
    assert_eq!(out, vec!["npe"]);
}

#[test]
fn concurrent_hash_map_compute_if_absent_return_value_is_inserted_mapping() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); System.out.println(map.computeIfAbsent("k", key -> 15));"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn concurrent_hash_map_merge_return_value_reflects_final_mapping() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); System.out.println(map.merge("k", 7, (a, b) -> a + b));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn concurrent_hash_map_for_each_key_prints_keys() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("p", 1); map.put("q", 2); map.forEachKey(1L, key -> System.out.println(key));"#,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"p".to_string()));
    assert!(out.contains(&"q".to_string()));
}

#[test]
fn concurrent_hash_map_for_each_value_sums_with_side_effect() {
    let types = r#"
        static int total = 0;
    "#;
    let out = run_in_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 3); map.put("b", 4); map.forEachValue(1L, val -> { total += val; }); System.out.println(total);"#,
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn concurrent_hash_map_reduce_keys_to_count_entries() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); map.put("c", 3); int count = map.reduceKeys(1L, (k1, k2) -> k1 + k2).length(); System.out.println(count > 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn concurrent_hash_map_search_entries_finds_by_value() {
    let out = run_main(
        r#"java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>(); map.put("a", 100); java.util.Map.Entry<String, Integer> hit = map.searchEntries(1L, e -> e.getValue() == 100 ? e : null); System.out.println(hit.getKey());"#,
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn concurrent_hash_map_thread_safe_put_if_absent_only_one_wins() {
    let types = r#"
        static java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>();
        static int wins = 0;
    "#;
    let out = run_in_main(
        r#"Thread t1 = new Thread(() -> { if (map.putIfAbsent("k", 1) == null) wins++; }); Thread t2 = new Thread(() -> { if (map.putIfAbsent("k", 2) == null) wins++; }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(wins); System.out.println(map.size());"#,
        types,
    );
    assert_eq!(out, vec!["1", "1"]);
}
