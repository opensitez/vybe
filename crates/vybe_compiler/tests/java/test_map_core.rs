use crate::helpers::run_main;

#[test]
fn hashmap_put_then_get_returns_stored_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"alpha\", 7); System.out.println(map.get(\"alpha\"));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn hashmap_put_overwrites_existing_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); map.put(\"k\", 9); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn hashmap_get_missing_key_returns_null() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.get(\"missing\"));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn hashmap_remove_existing_key_returns_previous_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"x\", 4); System.out.println(map.remove(\"x\"));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn hashmap_remove_missing_key_returns_null() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.remove(\"ghost\"));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn hashmap_contains_key_true_when_entry_present() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); System.out.println(map.containsKey(\"k\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashmap_contains_key_false_when_entry_absent() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.containsKey(\"k\"));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashmap_contains_value_true_when_value_present() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 42); System.out.println(map.containsValue(42));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashmap_contains_value_false_when_value_absent() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); System.out.println(map.containsValue(99));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashmap_size_counts_stored_entries() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashmap_new_instance_starts_with_zero_size() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn hashmap_is_empty_on_fresh_instance() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashmap_is_empty_false_after_insert() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); System.out.println(map.isEmpty());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashmap_clear_empties_all_entries() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.clear(); System.out.println(map.size()); System.out.println(map.containsKey(\"a\"));",
    );
    assert_eq!(out, vec!["0", "false"]);
}

#[test]
fn hashmap_get_returns_null_after_clear() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 5); map.clear(); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn hashmap_put_if_absent_inserts_when_key_missing() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.putIfAbsent(\"k\", 11); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn hashmap_put_if_absent_preserves_existing_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 3); map.putIfAbsent(\"k\", 99); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashmap_put_if_absent_returns_existing_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 8); System.out.println(map.putIfAbsent(\"k\", 1));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn hashmap_get_or_default_returns_stored_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 6); System.out.println(map.getOrDefault(\"k\", 0));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn hashmap_get_or_default_returns_fallback_when_missing() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.getOrDefault(\"k\", -1));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn hashmap_key_set_size_matches_entry_count() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); System.out.println(map.keySet().size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashmap_values_collection_size_matches_entry_count() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.put(\"c\", 3); System.out.println(map.values().size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashmap_entry_set_size_matches_entry_count() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"x\", 1); map.put(\"y\", 2); System.out.println(map.entrySet().size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashmap_put_all_copies_entries_from_source_map() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); java.util.HashMap<String, Integer> other = new java.util.HashMap<String, Integer>(); other.put(\"a\", 10); other.put(\"b\", 20); map.putAll(other); System.out.println(map.get(\"a\")); System.out.println(map.get(\"b\"));",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn hashmap_put_all_merges_into_existing_entries() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); java.util.HashMap<String, Integer> other = new java.util.HashMap<String, Integer>(); other.put(\"b\", 2); map.putAll(other); System.out.println(map.size()); System.out.println(map.get(\"a\")); System.out.println(map.get(\"b\"));",
    );
    assert_eq!(out, vec!["2", "1", "2"]);
}

#[test]
fn hashmap_replace_updates_existing_mapping() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); map.replace(\"k\", 5); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn hashmap_replace_returns_null_for_missing_key() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.replace(\"missing\", 9));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn hashmap_replace_returns_previous_value_when_key_exists() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 2); System.out.println(map.replace(\"k\", 7)); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["2", "7"]);
}

#[test]
fn hashmap_compute_if_absent_inserts_when_key_missing() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.computeIfAbsent(\"k\", key -> 15); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn hashmap_compute_if_absent_keeps_existing_value() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 4); map.computeIfAbsent(\"k\", key -> 99); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn map_interface_reference_accepts_hashmap_instance() {
    let out = run_main(
        "java.util.Map<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 12); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn map_interface_put_and_size_through_polymorphic_reference() {
    let out = run_main(
        "java.util.Map<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); System.out.println(map.size()); System.out.println(map.containsKey(\"b\"));",
    );
    assert_eq!(out, vec!["2", "true"]);
}

#[test]
fn hashmap_put_returns_null_for_new_key() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); System.out.println(map.put(\"k\", 1));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn hashmap_put_returns_previous_value_on_overwrite() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); System.out.println(map.put(\"k\", 2));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn hashmap_remove_reduces_size_by_one() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.remove(\"a\"); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn hashmap_remove_then_contains_key_is_false() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 1); map.remove(\"k\"); System.out.println(map.containsKey(\"k\"));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashmap_stores_multiple_independent_keys() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"one\", 1); map.put(\"two\", 2); map.put(\"three\", 3); System.out.println(map.get(\"one\")); System.out.println(map.get(\"three\"));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn hashmap_integer_keyed_entries_are_retrievable() {
    let out = run_main(
        "java.util.HashMap<Integer, String> map = new java.util.HashMap<Integer, String>(); map.put(1, \"one\"); map.put(2, \"two\"); System.out.println(map.get(1)); System.out.println(map.get(2));",
    );
    assert_eq!(out, vec!["one", "two"]);
}

#[test]
fn hashmap_contains_value_finds_duplicate_values_across_keys() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 5); map.put(\"b\", 5); System.out.println(map.containsValue(5));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashmap_key_set_shrinks_after_remove() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.remove(\"a\"); System.out.println(map.keySet().size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn hashmap_values_collection_empty_after_clear() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.clear(); System.out.println(map.values().size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn hashmap_entry_set_empty_after_clear() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.clear(); System.out.println(map.entrySet().size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn hashmap_string_values_concatenate_from_get() {
    let out = run_main(
        "java.util.HashMap<String, String> map = new java.util.HashMap<String, String>(); map.put(\"left\", \"vy\"); map.put(\"right\", \"be\"); System.out.println(map.get(\"left\") + map.get(\"right\"));",
    );
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn hashmap_chained_mutations_leave_expected_final_state() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.put(\"b\", 2); map.remove(\"a\"); map.put(\"c\", 3); System.out.println(map.size()); System.out.println(map.get(\"b\")); System.out.println(map.get(\"c\"));",
    );
    assert_eq!(out, vec!["2", "2", "3"]);
}

#[test]
fn map_interface_remove_through_hashmap_implementation() {
    let out = run_main(
        "java.util.Map<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"k\", 9); map.remove(\"k\"); System.out.println(map.size()); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["0", "null"]);
}
