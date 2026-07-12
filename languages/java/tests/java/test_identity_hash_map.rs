/// java.util.IdentityHashMap — identity-based key equality.
use crate::helpers::run_main;

#[test]
fn identity_hash_map_put_and_get_same_reference() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn identity_hash_map_equal_but_distinct_keys_are_separate_entries() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String a = new String("x"); String b = new String("x"); map.put(a, 1); map.put(b, 2); System.out.println(map.size()); System.out.println(map.get(a)); System.out.println(map.get(b));"#,
    );
    assert_eq!(out, vec!["2", "1", "2"]);
}

#[test]
fn identity_hash_map_literal_equal_string_misses_distinct_key() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("lit"); map.put(key, 5); System.out.println(map.get("lit"));"#,
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn identity_hash_map_new_map_is_empty() {
    let out = run_main(
        "java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); System.out.println(map.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn identity_hash_map_size_after_two_distinct_puts() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.put(new String("b"), 2); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn identity_hash_map_contains_key_uses_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); System.out.println(map.containsKey(key)); System.out.println(map.containsKey("k"));"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn identity_hash_map_contains_value_finds_matching_value() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 9); System.out.println(map.containsValue(9));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn identity_hash_map_remove_by_identity_returns_value() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 4); System.out.println(map.remove(key)); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["4", "true"]);
}

#[test]
fn identity_hash_map_remove_equal_literal_does_not_remove_distinct_key() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 4); System.out.println(map.remove("k")); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["null", "1"]);
}

#[test]
fn identity_hash_map_put_overwrites_same_reference() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); map.put(key, 9); System.out.println(map.get(key)); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["9", "1"]);
}

#[test]
fn identity_hash_map_put_returns_previous_for_same_reference() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); System.out.println(map.put(key, 2));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn identity_hash_map_get_or_default_uses_identity_lookup() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 3); System.out.println(map.getOrDefault(key, 42)); System.out.println(map.getOrDefault("k", 42));"#,
    );
    assert_eq!(out, vec!["3", "42"]);
}

#[test]
fn identity_hash_map_clear_empties_entries() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.clear(); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn identity_hash_map_key_set_contains_identity_reference() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("x"); map.put(key, 1); System.out.println(map.keySet().contains(key));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn identity_hash_map_values_contains_inserted_value() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("x"), 5); System.out.println(map.values().contains(5));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn identity_hash_map_entry_set_iterator_reads_identity_pair() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 2); java.util.Map.Entry<String, Integer> entry = map.entrySet().iterator().next(); System.out.println(entry.getKey() == key); System.out.println(entry.getValue());"#,
    );
    assert_eq!(out, vec!["true", "2"]);
}

#[test]
fn identity_hash_map_put_all_copies_by_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> src = new java.util.IdentityHashMap<String, Integer>(); String key = new String("a"); src.put(key, 1); java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.putAll(src); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn identity_hash_map_null_key_allowed() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(null, 7); System.out.println(map.get(null));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn identity_hash_map_null_value_allowed() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, null); System.out.println(map.containsKey(key)); System.out.println(map.get(key) == null);"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn identity_hash_map_initial_capacity_constructor() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(32); String key = new String("k"); map.put(key, 1); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn identity_hash_map_remove_key_value_conditional_by_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 5); System.out.println(map.remove(key, 5)); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn identity_hash_map_replace_existing_by_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); System.out.println(map.replace(key, 8)); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["1", "8"]);
}

#[test]
fn identity_hash_map_replace_conditional_by_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); System.out.println(map.replace(key, 1, 6)); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["true", "6"]);
}

#[test]
fn identity_hash_map_integer_keys_distinguish_boxed_instances() {
    let out = run_main(
        "java.util.IdentityHashMap<Integer, String> map = new java.util.IdentityHashMap<Integer, String>(); Integer a = new Integer(1); Integer b = new Integer(1); map.put(a, \"first\"); map.put(b, \"second\"); System.out.println(map.size()); System.out.println(map.get(a)); System.out.println(map.get(b));",
    );
    assert_eq!(out, vec!["2", "first", "second"]);
}

#[test]
fn identity_hash_map_same_integer_reference_single_entry() {
    let out = run_main(
        "java.util.IdentityHashMap<Integer, String> map = new java.util.IdentityHashMap<Integer, String>(); Integer key = new Integer(7); map.put(key, \"a\"); map.put(key, \"b\"); System.out.println(map.size()); System.out.println(map.get(key));",
    );
    assert_eq!(out, vec!["1", "b"]);
}

#[test]
fn identity_hash_map_object_array_keys_by_identity() {
    let out = run_main(
        "java.util.IdentityHashMap<int[], String> map = new java.util.IdentityHashMap<int[], String>(); int[] a = new int[]{1}; int[] b = new int[]{1}; map.put(a, \"a\"); map.put(b, \"b\"); System.out.println(map.size()); System.out.println(map.get(a));",
    );
    assert_eq!(out, vec!["2", "a"]);
}

#[test]
fn identity_hash_map_for_each_visits_each_identity_entry() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.put(new String("b"), 2); map.forEach((key, val) -> System.out.println(val));"#,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"1".to_string()));
    assert!(out.contains(&"2".to_string()));
}

#[test]
fn identity_hash_map_key_set_size_matches_entry_count() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.put(new String("b"), 2); System.out.println(map.keySet().size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn identity_hash_map_values_size_matches_entry_count() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.put(new String("b"), 2); System.out.println(map.values().size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn identity_hash_map_entry_set_size_matches_entry_count() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.put(new String("b"), 2); System.out.println(map.entrySet().size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn identity_hash_map_put_if_absent_inserts_for_identity_key() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); System.out.println(map.putIfAbsent(key, 7)); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["null", "7"]);
}

#[test]
fn identity_hash_map_put_if_absent_skips_same_reference() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); System.out.println(map.putIfAbsent(key, 9)); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn identity_hash_map_compute_if_absent_uses_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.computeIfAbsent(key, k -> 11); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn identity_hash_map_compute_if_present_updates_by_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 2); map.computeIfPresent(key, (k, val) -> val + 5); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn identity_hash_map_merge_combines_by_identity_key() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 2); map.merge(key, 3, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn identity_hash_map_compute_inserts_for_identity_key() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.compute(key, (k, val) -> 10); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn identity_hash_map_compute_removes_when_function_returns_null() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); map.compute(key, (k, val) -> null); System.out.println(map.containsKey(key));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn identity_hash_map_replace_all_updates_all_values() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.put(new String("b"), 2); map.replaceAll((key, val) -> val * 10); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn identity_hash_map_clone_preserves_identity_mappings() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 3); java.util.IdentityHashMap<String, Integer> copy = (java.util.IdentityHashMap<String, Integer>) map.clone(); System.out.println(copy.get(key));"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn identity_hash_map_get_absent_identity_returns_null() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); System.out.println(map.get(key) == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn identity_hash_map_remove_nonexistent_identity_returns_null() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); System.out.println(map.remove(key));"#,
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn identity_hash_map_contains_value_false_when_absent() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("k"), 1); System.out.println(map.containsValue(9));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn identity_hash_map_key_set_remove_by_identity() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.put(key, 1); map.keySet().remove(key); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn identity_hash_map_two_canonical_equal_strings_one_entry_if_same_reference() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = "shared"; map.put(key, 1); System.out.println(map.get(key)); System.out.println(map.size());"#,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn identity_hash_map_merge_inserts_when_identity_key_absent() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); String key = new String("k"); map.merge(key, 5, (oldVal, newVal) -> oldVal + newVal); System.out.println(map.get(key));"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn identity_hash_map_identity_hash_code_differs_from_regular_hashmap_size() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> idMap = new java.util.IdentityHashMap<String, Integer>(); idMap.put(new String("a"), 1); idMap.put(new String("a"), 2); java.util.HashMap<String, Integer> hashMap = new java.util.HashMap<String, Integer>(); hashMap.put(new String("a"), 1); hashMap.put(new String("a"), 2); System.out.println(idMap.size()); System.out.println(hashMap.size());"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn identity_hash_map_entry_set_clear_empties_map() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.entrySet().clear(); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn identity_hash_map_values_remove_all_clears_map() {
    let out = run_main(
        r#"java.util.IdentityHashMap<String, Integer> map = new java.util.IdentityHashMap<String, Integer>(); map.put(new String("a"), 1); map.values().clear(); System.out.println(map.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}
