/// java.util.WeakHashMap — weak-key map semantics.
use crate::helpers::{run_in_main, run_main};

#[test]
fn weak_hash_map_put_and_get_round_trip() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn weak_hash_map_new_map_is_empty() {
    let out = run_main(
        "java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); System.out.println(map.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_size_after_put() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn weak_hash_map_contains_key_after_put() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 9); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_contains_value_after_put() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 9); System.out.println(map.containsValue(9));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_remove_returns_previous_value() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 4); System.out.println(map.remove("k"));"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn weak_hash_map_remove_leaves_map_empty() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 4); map.remove("k"); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_put_overwrites_existing_value() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); map.put("k", 9); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn weak_hash_map_put_returns_previous_value() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.put("k", 2));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn weak_hash_map_get_or_default_returns_default_when_absent() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); System.out.println(map.getOrDefault("missing", 42));"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn weak_hash_map_get_or_default_returns_existing_value() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 3); System.out.println(map.getOrDefault("k", 42));"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn weak_hash_map_clear_empties_all_entries() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); map.clear(); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn weak_hash_map_key_set_contains_inserted_key() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("x", 1); System.out.println(map.keySet().contains("x"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_values_contains_inserted_value() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("x", 5); System.out.println(map.values().contains(5));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_entry_set_iterator_reads_pair() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 2); java.util.Map.Entry<String, Integer> entry = map.entrySet().iterator().next(); System.out.println(entry.getKey()); System.out.println(entry.getValue());"#,
    );
    assert_eq!(out, vec!["k", "2"]);
}

#[test]
fn weak_hash_map_put_all_merges_other_map() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); java.util.HashMap<String, Integer> other = new java.util.HashMap<String, Integer>(); other.put("a", 1); other.put("b", 2); map.putAll(other); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn weak_hash_map_null_key_allowed() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put(null, 7); System.out.println(map.get(null));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn weak_hash_map_null_value_allowed() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", null); System.out.println(map.containsKey("k")); System.out.println(map.get("k") == null);"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn weak_hash_map_initial_capacity_constructor() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(16); map.put("k", 1); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn weak_hash_map_remove_key_value_conditional_succeeds() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 5); System.out.println(map.remove("k", 5)); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn weak_hash_map_remove_key_value_conditional_fails_on_mismatch() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 5); System.out.println(map.remove("k", 9)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["false", "5"]);
}

#[test]
fn weak_hash_map_replace_existing_key() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.replace("k", 8)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["1", "8"]);
}

#[test]
fn weak_hash_map_replace_conditional_succeeds() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.replace("k", 1, 6)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["true", "6"]);
}

#[test]
fn weak_hash_map_compute_if_absent_inserts_when_missing() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.computeIfAbsent("k", key -> 11); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn weak_hash_map_compute_if_present_updates_existing() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 2); map.computeIfPresent("k", (key, val) -> val + 5); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn weak_hash_map_merge_combines_existing_values() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 2); map.merge("k", 3, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn weak_hash_map_for_each_visits_each_entry() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); map.forEach((key, val) -> System.out.println(val));"#,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"1".to_string()));
    assert!(out.contains(&"2".to_string()));
}

#[test]
fn weak_hash_map_equal_content_maps_are_equal() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> a = new java.util.WeakHashMap<String, Integer>(); a.put("k", 1); java.util.WeakHashMap<String, Integer> b = new java.util.WeakHashMap<String, Integer>(); b.put("k", 1); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_key_set_size_matches_map_size() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); System.out.println(map.keySet().size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn weak_hash_map_values_size_matches_map_size() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); System.out.println(map.values().size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn weak_hash_map_entry_set_size_matches_map_size() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); System.out.println(map.entrySet().size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn weak_hash_map_put_if_absent_inserts_when_missing() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); System.out.println(map.putIfAbsent("k", 7)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["null", "7"]);
}

#[test]
fn weak_hash_map_put_if_absent_skips_when_present() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.putIfAbsent("k", 9)); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn weak_hash_map_compute_inserts_when_absent() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.compute("k", (key, val) -> 10); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn weak_hash_map_compute_removes_when_function_returns_null() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); map.compute("k", (key, val) -> null); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn weak_hash_map_replace_all_updates_values() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("a", 1); map.put("b", 2); map.replaceAll((key, val) -> val * 10); System.out.println(map.get("a")); System.out.println(map.get("b"));"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn weak_hash_map_retain_key_while_strong_reference_held() {
    let types = r#"
        static String strongKey = "alive";
    "#;
    let out = run_in_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put(strongKey, 99); System.out.println(map.get(strongKey));"#,
        types,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn weak_hash_map_different_canonical_equal_strings_share_entry() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); String a = "lit"; String b = "lit"; map.put(a, 1); System.out.println(map.get(b));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn weak_hash_map_new_string_equal_but_distinct_key_is_separate_until_interned() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); String key = new String("x"); map.put(key, 5); System.out.println(map.get("x"));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn weak_hash_map_remove_nonexistent_returns_null() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); System.out.println(map.remove("missing"));"#,
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn weak_hash_map_get_absent_returns_null() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); System.out.println(map.get("nope") == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn weak_hash_map_contains_key_false_for_absent() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); System.out.println(map.containsKey("absent"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn weak_hash_map_contains_value_false_for_absent() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); System.out.println(map.containsValue(9));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn weak_hash_map_merge_inserts_when_key_absent() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.merge("k", 5, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get("k"));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn weak_hash_map_compute_if_present_noop_when_absent() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.computeIfPresent("k", (key, val) -> 9); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn weak_hash_map_key_set_remove_updates_map() {
    let out = run_main(
        r#"java.util.WeakHashMap<String, Integer> map = new java.util.WeakHashMap<String, Integer>(); map.put("k", 1); map.keySet().remove("k"); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}
