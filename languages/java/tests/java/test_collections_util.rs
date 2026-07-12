use crate::helpers::run_main;

#[test]
fn collections_sort_orders_unsorted_integers_ascending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(30); list.add(10); list.add(20); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn collections_sort_orders_unsorted_strings_lexicographically() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"cherry\"); list.add(\"apple\"); list.add(\"banana\"); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["apple", "cherry"]);
}

#[test]
fn collections_sort_leaves_already_sorted_list_unchanged() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn collections_sort_handles_single_element_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(42); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["42", "1"]);
}

#[test]
fn collections_sort_after_reverse_restores_ascending_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.Collections.reverse(list); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn collections_reverse_flips_integer_list_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.Collections.reverse(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn collections_reverse_twice_restores_original_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(6); list.add(7); java.util.Collections.reverse(list); java.util.Collections.reverse(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["5", "7"]);
}

#[test]
fn collections_reverse_on_singleton_list_is_noop() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(99); java.util.Collections.reverse(list); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn collections_shuffle_preserves_list_size() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); list.add(4); java.util.Collections.shuffle(list); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn collections_shuffle_preserves_element_sum() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); int before = 0; for (int v : list) { before += v; } java.util.Collections.shuffle(list); int after = 0; for (int v : list) { after += v; } System.out.println(before); System.out.println(after);",
    );
    assert_eq!(out, vec!["6", "6"]);
}

#[test]
fn collections_shuffle_on_empty_list_stays_empty() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); java.util.Collections.shuffle(list); System.out.println(list.size()); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn collections_fill_sets_every_list_slot_to_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(0); list.add(0); list.add(0); java.util.Collections.fill(list, 7); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["7", "7"]);
}

#[test]
fn collections_fill_on_empty_list_is_noop() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); java.util.Collections.fill(list, \"x\"); System.out.println(list.size()); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn collections_copy_transfers_source_elements_into_destination() {
    let out = run_main(
        "java.util.ArrayList<Integer> src = new java.util.ArrayList<Integer>(); src.add(1); src.add(2); java.util.ArrayList<Integer> dest = new java.util.ArrayList<Integer>(); dest.add(9); dest.add(9); java.util.Collections.copy(dest, src); System.out.println(dest.get(0)); System.out.println(dest.get(1));",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn collections_copy_leaves_source_list_unchanged() {
    let out = run_main(
        "java.util.ArrayList<Integer> src = new java.util.ArrayList<Integer>(); src.add(4); src.add(5); java.util.ArrayList<Integer> dest = new java.util.ArrayList<Integer>(); dest.add(0); dest.add(0); java.util.Collections.copy(dest, src); System.out.println(src.get(0)); System.out.println(src.get(1));",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn collections_min_finds_smallest_integer_in_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(8); list.add(3); list.add(11); System.out.println(java.util.Collections.min(list));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collections_min_finds_smallest_string_lexicographically() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"zebra\"); list.add(\"ant\"); list.add(\"moth\"); System.out.println(java.util.Collections.min(list));",
    );
    assert_eq!(out, vec!["ant"]);
}

#[test]
fn collections_min_on_singleton_list_returns_only_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(-4); System.out.println(java.util.Collections.min(list));",
    );
    assert_eq!(out, vec!["-4"]);
}

#[test]
fn collections_max_finds_largest_integer_in_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(8); list.add(3); list.add(11); System.out.println(java.util.Collections.max(list));",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn collections_max_among_negative_integers_picks_closest_to_zero() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(-9); list.add(-2); list.add(-6); System.out.println(java.util.Collections.max(list));",
    );
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn collections_max_on_two_equal_values_returns_that_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(5); System.out.println(java.util.Collections.max(list));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn collections_frequency_counts_matching_elements() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"a\"); list.add(\"b\"); list.add(\"a\"); System.out.println(java.util.Collections.frequency(list, \"a\"));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collections_frequency_returns_zero_for_absent_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); System.out.println(java.util.Collections.frequency(list, 9));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn collections_frequency_counts_all_duplicate_occurrences() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(7); list.add(7); list.add(7); System.out.println(java.util.Collections.frequency(list, 7));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collections_disjoint_true_when_no_shared_elements() {
    let out = run_main(
        "java.util.ArrayList<Integer> a = new java.util.ArrayList<Integer>(); a.add(1); a.add(2); java.util.ArrayList<Integer> b = new java.util.ArrayList<Integer>(); b.add(3); b.add(4); System.out.println(java.util.Collections.disjoint(a, b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn collections_disjoint_false_when_lists_share_an_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> a = new java.util.ArrayList<Integer>(); a.add(1); a.add(2); java.util.ArrayList<Integer> b = new java.util.ArrayList<Integer>(); b.add(2); b.add(3); System.out.println(java.util.Collections.disjoint(a, b));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn collections_disjoint_true_for_two_empty_lists() {
    let out = run_main(
        "java.util.ArrayList<Integer> a = new java.util.ArrayList<Integer>(); java.util.ArrayList<Integer> b = new java.util.ArrayList<Integer>(); System.out.println(java.util.Collections.disjoint(a, b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn collections_ncopies_repeats_value_requested_number_of_times() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Collections.nCopies(4, 8); System.out.println(list.size()); System.out.println(list.get(3));",
    );
    assert_eq!(out, vec!["4", "8"]);
}

#[test]
fn collections_ncopies_with_zero_count_yields_empty_list() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Collections.nCopies(0, 1); System.out.println(list.size()); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn collections_ncopies_single_copy_matches_singleton_list_size() {
    let out = run_main(
        "java.util.List<String> copied = java.util.Collections.nCopies(1, \"solo\"); java.util.List<String> single = java.util.Collections.singletonList(\"solo\"); System.out.println(copied.size()); System.out.println(single.size());",
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn collections_singleton_list_holds_exactly_one_element() {
    let out = run_main(
        "java.util.List<String> list = java.util.Collections.singletonList(\"solo\"); System.out.println(list.get(0)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["solo", "1"]);
}

#[test]
fn collections_singleton_list_reports_nonempty() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Collections.singletonList(42); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn collections_singleton_set_holds_one_member() {
    let out = run_main(
        "java.util.Set<String> set = java.util.Collections.singleton(\"only\"); System.out.println(set.size()); System.out.println(set.contains(\"only\"));",
    );
    assert_eq!(out, vec!["1", "true"]);
}

#[test]
fn collections_singleton_map_holds_one_entry() {
    let out = run_main(
        "java.util.Map<String, Integer> map = java.util.Collections.singletonMap(\"k\", 9); System.out.println(map.get(\"k\")); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["9", "1"]);
}

#[test]
fn collections_unmodifiable_list_allows_indexed_reads() {
    let out = run_main(
        "java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); backing.add(2); backing.add(4); java.util.List<Integer> view = java.util.Collections.unmodifiableList(backing); System.out.println(view.get(0)); System.out.println(view.get(1));",
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn collections_unmodifiable_list_reports_same_size_as_source() {
    let out = run_main(
        "java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); backing.add(\"a\"); backing.add(\"b\"); backing.add(\"c\"); java.util.List<String> view = java.util.Collections.unmodifiableList(backing); System.out.println(view.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collections_synchronized_list_accepts_add_operations() {
    let out = run_main(
        "java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(10); sync.add(20); System.out.println(sync.get(0)); System.out.println(sync.get(1));",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn collections_synchronized_list_preserves_insertion_order() {
    let out = run_main(
        "java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add(\"first\"); sync.add(\"second\"); System.out.println(sync.get(0)); System.out.println(sync.get(1));",
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn collections_empty_list_has_zero_size() {
    let out = run_main(
        "java.util.List<Object> list = java.util.Collections.emptyList(); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn collections_empty_list_is_empty_predicate() {
    let out = run_main(
        "java.util.List<String> list = java.util.Collections.emptyList(); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn collections_empty_map_has_zero_entries() {
    let out = run_main(
        "java.util.Map<String, Integer> map = java.util.Collections.emptyMap(); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn collections_empty_map_is_empty_predicate() {
    let out = run_main(
        "java.util.Map<String, String> map = java.util.Collections.emptyMap(); System.out.println(map.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn collections_reverse_then_min_reads_new_head_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(5); list.add(3); java.util.Collections.reverse(list); System.out.println(java.util.Collections.min(list));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn collections_frequency_on_singleton_list_is_one() {
    let out = run_main(
        "java.util.List<String> list = java.util.Collections.singletonList(\"x\"); System.out.println(java.util.Collections.frequency(list, \"x\"));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn collections_fill_then_max_reads_uniform_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(0); list.add(0); java.util.Collections.fill(list, 12); System.out.println(java.util.Collections.max(list)); System.out.println(java.util.Collections.min(list));",
    );
    assert_eq!(out, vec!["12", "12"]);
}
