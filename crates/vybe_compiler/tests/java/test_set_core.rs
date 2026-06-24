use crate::helpers::run_main;

#[test]
fn hashset_add_first_element_increments_size() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(10); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn hashset_add_duplicate_is_ignored() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(1); set.add(1); set.add(2); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashset_add_three_distinct_values_reports_size_three() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(4); set.add(5); set.add(6); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashset_contains_finds_added_integer() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(42); System.out.println(set.contains(42));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashset_contains_rejects_absent_integer() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(1); System.out.println(set.contains(9));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashset_remove_drops_existing_member() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(7); set.add(8); set.remove(7); System.out.println(set.size()); System.out.println(set.contains(7));",
    );
    assert_eq!(out, vec!["1", "false"]);
}

#[test]
fn hashset_remove_absent_member_leaves_size_unchanged() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(3); set.remove(99); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn hashset_size_counts_unique_members() {
    let out = run_main(
        "java.util.HashSet<String> set = new java.util.HashSet<String>(); set.add(\"a\"); set.add(\"b\"); set.add(\"a\"); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashset_is_empty_true_on_construction() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); System.out.println(set.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashset_is_empty_false_after_add() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(0); System.out.println(set.isEmpty());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn hashset_clear_empties_all_members() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(1); set.add(2); set.clear(); System.out.println(set.size()); System.out.println(set.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn hashset_enhanced_for_loop_accumulates_all_elements() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<Integer>(); set.add(1); set.add(2); set.add(3); int sum = 0; for (int v : set) { sum += v; } System.out.println(sum);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn linkedhashset_add_preserves_insertion_order_in_iteration() {
    let out = run_main(
        "java.util.LinkedHashSet<Integer> set = new java.util.LinkedHashSet<Integer>(); set.add(2); set.add(1); set.add(3); for (int v : set) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["2", "1", "3"]);
}

#[test]
fn linkedhashset_add_duplicate_keeps_size_stable() {
    let out = run_main(
        "java.util.LinkedHashSet<Integer> set = new java.util.LinkedHashSet<Integer>(); set.add(5); set.add(5); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn linkedhashset_contains_finds_string_member() {
    let out = run_main(
        "java.util.LinkedHashSet<String> set = new java.util.LinkedHashSet<String>(); set.add(\"vybe\"); System.out.println(set.contains(\"vybe\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linkedhashset_remove_updates_size() {
    let out = run_main(
        "java.util.LinkedHashSet<Integer> set = new java.util.LinkedHashSet<Integer>(); set.add(10); set.add(20); set.remove(10); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn linkedhashset_is_empty_on_new_instance() {
    let out = run_main(
        "java.util.LinkedHashSet<String> set = new java.util.LinkedHashSet<String>(); System.out.println(set.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linkedhashset_clear_resets_size_to_zero() {
    let out = run_main(
        "java.util.LinkedHashSet<Integer> set = new java.util.LinkedHashSet<Integer>(); set.add(1); set.clear(); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn linkedhashset_size_after_multiple_adds() {
    let out = run_main(
        "java.util.LinkedHashSet<Integer> set = new java.util.LinkedHashSet<Integer>(); set.add(1); set.add(2); set.add(3); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn linkedhashset_contains_false_for_missing_key() {
    let out = run_main(
        "java.util.LinkedHashSet<String> set = new java.util.LinkedHashSet<String>(); set.add(\"one\"); System.out.println(set.contains(\"two\"));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn linkedhashset_iteration_prints_elements_in_add_sequence() {
    let out = run_main(
        "java.util.LinkedHashSet<String> set = new java.util.LinkedHashSet<String>(); set.add(\"first\"); set.add(\"second\"); for (String s : set) { System.out.println(s); }",
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn linkedhashset_contains_null_after_explicit_null_add() {
    let out = run_main(
        "java.util.LinkedHashSet<String> set = new java.util.LinkedHashSet<String>(); set.add(null); System.out.println(set.contains(null));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linkedhashset_remove_then_contains_false() {
    let out = run_main(
        "java.util.LinkedHashSet<Integer> set = new java.util.LinkedHashSet<Integer>(); set.add(4); set.remove(4); System.out.println(set.contains(4));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn linkedhashset_readd_after_remove() {
    let out = run_main(
        "java.util.LinkedHashSet<Integer> set = new java.util.LinkedHashSet<Integer>(); set.add(9); set.remove(9); set.add(9); System.out.println(set.size()); System.out.println(set.contains(9));",
    );
    assert_eq!(out, vec!["1", "true"]);
}

#[test]
fn treeset_add_sorts_on_iteration() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(3); set.add(1); set.add(2); for (int v : set) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn treeset_duplicate_add_ignored() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(2); set.add(2); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn treeset_contains_member() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(11); System.out.println(set.contains(11));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn treeset_remove_reduces_size() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.remove(1); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn treeset_is_empty_initially() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); System.out.println(set.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn treeset_clear_removes_all() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(5); set.add(6); set.clear(); System.out.println(set.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn treeset_size_reports_unique_count() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(10); set.add(20); set.add(10); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treeset_first_element_is_smallest_after_adds() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(30); set.add(10); set.add(20); int first = -1; for (int v : set) { first = v; break; } System.out.println(first);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn treeset_contains_absent_returns_false() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); System.out.println(set.contains(2));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn treeset_remove_absent_member_size_unchanged() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(8); set.remove(1); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn treeset_string_members_iterate_lexicographically() {
    let out = run_main(
        "java.util.TreeSet<String> set = new java.util.TreeSet<String>(); set.add(\"cherry\"); set.add(\"apple\"); set.add(\"banana\"); for (String s : set) { System.out.println(s); }",
    );
    assert_eq!(out, vec!["apple", "banana", "cherry"]);
}
