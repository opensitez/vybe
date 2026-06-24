use crate::helpers::run_main;

#[test]
fn arraylist_add_appends_element_to_end() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(20); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn arraylist_add_at_index_inserts_without_dropping_tail() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); list.add(1, 2); System.out.println(list.get(1)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn arraylist_get_reads_element_by_zero_based_index() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"alpha\"); list.add(\"beta\"); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["beta"]);
}

#[test]
fn arraylist_set_replaces_existing_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.set(0, 9); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["9", "2"]);
}

#[test]
fn arraylist_remove_by_index_shifts_later_elements_left() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(20); list.add(30); list.remove(1); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn arraylist_remove_by_object_drops_first_equal_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(6); list.remove(Integer.valueOf(5)); System.out.println(list.size()); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["1", "6"]);
}

#[test]
fn arraylist_size_counts_current_elements() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"a\"); list.add(\"b\"); list.add(\"c\"); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraylist_is_empty_true_on_construction() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_is_empty_false_after_add() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_clear_removes_every_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.clear(); System.out.println(list.size()); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn arraylist_contains_finds_present_value() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"java\"); list.add(\"rust\"); System.out.println(list.contains(\"java\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_contains_rejects_missing_value() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"one\"); System.out.println(list.contains(\"two\"));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_index_of_returns_first_matching_position() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(7); list.add(4); System.out.println(list.indexOf(4));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arraylist_index_of_returns_negative_one_when_absent() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); System.out.println(list.indexOf(9));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn arraylist_last_index_of_returns_final_matching_position() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(7); list.add(4); System.out.println(list.lastIndexOf(4));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arraylist_sub_list_exposes_range_view_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.List<Integer> sub = list.subList(1, 3); System.out.println(sub.get(0)); System.out.println(sub.get(1));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn arraylist_sub_list_size_matches_requested_span() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); list.add(4); java.util.List<Integer> sub = list.subList(1, 4); System.out.println(sub.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraylist_add_all_appends_every_source_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> left = new java.util.ArrayList<Integer>(); left.add(1); java.util.ArrayList<Integer> right = new java.util.ArrayList<Integer>(); right.add(2); right.add(3); left.addAll(right); System.out.println(left.size()); System.out.println(left.get(2));",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn arraylist_add_all_at_index_inserts_block_in_middle() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(4); java.util.ArrayList<Integer> mid = new java.util.ArrayList<Integer>(); mid.add(2); mid.add(3); list.addAll(1, mid); System.out.println(list.get(1)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn arraylist_remove_all_deletes_values_present_in_other_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.ArrayList<Integer> drop = new java.util.ArrayList<Integer>(); drop.add(2); list.removeAll(drop); System.out.println(list.size()); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn arraylist_retain_all_keeps_only_intersection_members() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.ArrayList<Integer> keep = new java.util.ArrayList<Integer>(); keep.add(2); keep.add(4); list.retainAll(keep); System.out.println(list.size()); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn arraylist_sort_orders_elements_ascending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); list.add(2); list.sort(null); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn collections_sort_orders_arraylist_in_place() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(30); list.add(10); list.add(20); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn list_of_factory_builds_immutable_sequence() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.List.of(10, 20, 30); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn list_of_with_no_arguments_is_empty() {
    let out = run_main(
        "java.util.List<String> list = java.util.List.of(); System.out.println(list.size()); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn arrays_as_list_wraps_array_literal_elements() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Arrays.asList(5, 6, 7); System.out.println(list.get(1)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["6", "3"]);
}

#[test]
fn arrays_as_list_from_array_variable_preserves_order() {
    let out = run_main(
        "Integer[] data = {1, 2, 3}; java.util.List<Integer> list = java.util.Arrays.asList(data); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn enhanced_for_loop_iterates_arraylist_elements() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); for (int v : list) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn arraylist_peek_first_reads_head_without_removing() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(11); list.add(22); System.out.println(list.peekFirst()); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["11", "2"]);
}

#[test]
fn arraylist_peek_last_reads_tail_without_removing() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(11); list.add(22); System.out.println(list.peekLast()); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["22", "22"]);
}

#[test]
fn arraylist_add_first_inserts_at_head_position() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(2); list.addFirst(1); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn arraylist_add_last_appends_at_tail_position() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.addLast(2); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arraylist_remove_first_drops_head_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(6); list.removeFirst(); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arraylist_remove_last_drops_tail_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(6); list.removeLast(); System.out.println(list.get(0)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["5", "1"]);
}

#[test]
fn list_reference_type_accepts_arraylist_instance() {
    let out = run_main(
        "java.util.List<String> list = new java.util.ArrayList<String>(); list.add(\"vybe\"); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn arraylist_set_returns_previous_element_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(7); System.out.println(list.set(0, 9)); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["7", "9"]);
}

#[test]
fn arraylist_add_single_argument_reports_success() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); System.out.println(list.add(42));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_remove_all_on_empty_operand_is_noop() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.ArrayList<Integer> empty = new java.util.ArrayList<Integer>(); list.removeAll(empty); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arraylist_retain_all_with_disjoint_filter_clears_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.ArrayList<Integer> other = new java.util.ArrayList<Integer>(); other.add(9); list.retainAll(other); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn collections_singleton_list_holds_one_element() {
    let out = run_main(
        "java.util.List<String> list = java.util.Collections.singletonList(\"solo\"); System.out.println(list.get(0)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["solo", "1"]);
}

#[test]
fn collections_empty_list_has_zero_length() {
    let out = run_main(
        "java.util.List<Object> list = java.util.Collections.emptyList(); System.out.println(list.size()); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn collections_n_copies_repeats_single_value() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Collections.nCopies(3, 8); System.out.println(list.size()); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "8"]);
}

#[test]
fn arraylist_index_of_and_last_index_of_differ_with_duplicates() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"x\"); list.add(\"y\"); list.add(\"x\"); System.out.println(list.indexOf(\"x\")); System.out.println(list.lastIndexOf(\"x\"));",
    );
    assert_eq!(out, vec!["0", "2"]);
}

#[test]
fn arraylist_remove_at_zero_promotes_next_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(100); list.add(200); list.remove(0); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["200"]);
}

#[test]
fn arraylist_foreach_loop_accumulates_running_total() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); int sum = 0; for (int n : list) { sum += n; } System.out.println(sum);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arraylist_to_array_call_leaves_list_size_unchanged() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"a\"); list.add(\"b\"); list.toArray(); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arraylist_get_first_element_via_head_access() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(4); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraylist_get_last_element_via_tail_index() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(4); System.out.println(list.get(list.size() - 1));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn arraylist_contains_null_after_explicit_null_add() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(null); System.out.println(list.contains(null));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn list_of_three_element_sequence_reads_middle_entry() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.List.of(2, 4, 6); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn arraylist_add_multiple_strings_preserves_insertion_order() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"first\"); list.add(\"second\"); list.add(\"third\"); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["first", "third"]);
}

#[test]
fn arraylist_clear_after_multiple_adds_resets_is_empty() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); list.clear(); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_sublist_first_element_matches_parent_index() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(20); list.add(30); java.util.List<Integer> sub = list.subList(0, 2); System.out.println(sub.get(0)); System.out.println(sub.get(1));",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn collections_sort_after_adding_strings_orders_lexicographically() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"cherry\"); list.add(\"apple\"); list.add(\"banana\"); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["apple", "cherry"]);
}

#[test]
fn arraylist_add_all_at_end_matches_sequential_add() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.ArrayList<Integer> extra = new java.util.ArrayList<Integer>(); extra.add(2); extra.add(3); list.addAll(extra); System.out.println(list.get(1)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn arraylist_remove_by_index_reduces_size_by_one() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.remove(1); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["1"]);
}
