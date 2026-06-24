use crate::helpers::{run_in_main, run_main};

#[test]
fn list_iterator_forward_has_next_on_nonempty_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.ListIterator<Integer> it = list.listIterator(); System.out.println(it.hasNext());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn list_iterator_next_reads_elements_in_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(20); java.util.ListIterator<Integer> it = list.listIterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn list_iterator_previous_walks_backward_after_next() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(6); java.util.ListIterator<Integer> it = list.listIterator(); it.next(); it.next(); System.out.println(it.previous());",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn list_iterator_set_replaces_last_returned_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.ListIterator<Integer> it = list.listIterator(); it.next(); it.set(9); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn list_iterator_add_inserts_at_cursor_position() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); java.util.ListIterator<Integer> it = list.listIterator(); it.next(); it.add(2); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn list_iterator_next_index_reports_cursor_position() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"a\"); list.add(\"b\"); java.util.ListIterator<String> it = list.listIterator(); it.next(); System.out.println(it.nextIndex());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn linkedlist_list_iterator_at_index_starts_mid_list() {
    let out = run_main(
        "java.util.LinkedList<Integer> list = new java.util.LinkedList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.ListIterator<Integer> it = list.listIterator(1); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn sublist_remove_from_middle_shrinks_parent_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.List<Integer> sub = list.subList(1, 3); sub.remove(0); System.out.println(list.size()); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn sublist_clear_empties_parent_range_only() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); list.subList(1, 3).clear(); System.out.println(list.size()); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn sublist_add_at_end_extends_parent_tail() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.subList(1, 2).add(9); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn add_all_at_index_zero_prepends_entire_block() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); java.util.ArrayList<Integer> block = new java.util.ArrayList<Integer>(); block.add(1); block.add(2); list.addAll(0, block); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn add_all_at_index_at_end_appends_block() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.ArrayList<Integer> block = new java.util.ArrayList<Integer>(); block.add(2); block.add(3); list.addAll(1, block); System.out.println(list.size()); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn add_all_at_index_on_empty_list_builds_sequence() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); java.util.ArrayList<Integer> block = new java.util.ArrayList<Integer>(); block.add(4); block.add(5); list.addAll(0, block); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn add_all_empty_collection_leaves_list_unchanged() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(7); java.util.ArrayList<Integer> empty = new java.util.ArrayList<Integer>(); list.addAll(0, empty); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn remove_all_deletes_every_occurrence_in_argument() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(2); list.add(3); java.util.ArrayList<Integer> drop = new java.util.ArrayList<Integer>(); drop.add(2); list.removeAll(drop); System.out.println(list.size()); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn remove_all_when_all_elements_match_clears_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.ArrayList<Integer> drop = new java.util.ArrayList<Integer>(); drop.add(1); drop.add(2); list.removeAll(drop); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn retain_all_empty_collection_clears_every_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.ArrayList<Integer> empty = new java.util.ArrayList<Integer>(); list.retainAll(empty); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn retain_all_with_identical_collection_preserves_all() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(5); java.util.ArrayList<Integer> keep = new java.util.ArrayList<Integer>(); keep.add(4); keep.add(5); list.retainAll(keep); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn retain_all_no_overlap_yields_empty_list() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"a\"); list.add(\"b\"); java.util.ArrayList<String> other = new java.util.ArrayList<String>(); other.add(\"z\"); list.retainAll(other); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn collections_shuffle_with_seeded_random_is_reproducible() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.Collections.shuffle(list, new java.util.Random(42)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collections_min_finds_smallest_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(9); list.add(2); list.add(7); System.out.println(java.util.Collections.min(list));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collections_max_finds_largest_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(9); list.add(2); list.add(7); System.out.println(java.util.Collections.max(list));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn collections_frequency_counts_matching_occurrences() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"x\"); list.add(\"y\"); list.add(\"x\"); System.out.println(java.util.Collections.frequency(list, \"x\"));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collections_frequency_zero_for_absent_value() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); System.out.println(java.util.Collections.frequency(list, 9));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn list_copy_of_creates_independent_immutable_snapshot() {
    let out = run_main(
        "java.util.ArrayList<Integer> source = new java.util.ArrayList<Integer>(); source.add(1); source.add(2); java.util.List<Integer> copy = java.util.List.copyOf(source); System.out.println(copy.get(0)); System.out.println(copy.size());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn list_copy_of_preserves_element_order() {
    let out = run_main(
        "java.util.List<String> copy = java.util.List.copyOf(java.util.Arrays.asList(\"a\", \"b\", \"c\")); System.out.println(copy.get(2));",
    );
    assert_eq!(out, vec!["c"]);
}

#[test]
fn list_of_add_attempt_is_caught_as_unsupported() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.List.of(1, 2); try { list.add(3); System.out.println(\"mutated\"); } catch (UnsupportedOperationException e) { System.out.println(\"immutable\"); }",
    );
    assert_eq!(out, vec!["immutable"]);
}

#[test]
fn list_of_set_attempt_is_caught_as_unsupported() {
    let out = run_main(
        "java.util.List<String> list = java.util.List.of(\"a\"); try { list.set(0, \"b\"); System.out.println(\"mutated\"); } catch (UnsupportedOperationException e) { System.out.println(\"immutable\"); }",
    );
    assert_eq!(out, vec!["immutable"]);
}

#[test]
fn list_of_remove_attempt_is_caught_as_unsupported() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.List.of(5); try { list.remove(0); System.out.println(\"mutated\"); } catch (UnsupportedOperationException e) { System.out.println(\"immutable\"); }",
    );
    assert_eq!(out, vec!["immutable"]);
}

#[test]
fn sort_with_comparator_orders_integers_ascending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); list.add(2); list.sort((a, b) -> a.compareTo(b)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn sort_with_comparator_orders_strings_descending() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"b\"); list.add(\"a\"); list.add(\"c\"); list.sort((a, b) -> b.compareTo(a)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["c", "a"]);
}

#[test]
fn collections_sort_with_comparator_reorders_custom_metric() {
    let out = run_in_main(
        "java.util.ArrayList<Item> list = new java.util.ArrayList<Item>(); list.add(new Item(3)); list.add(new Item(1)); list.add(new Item(2)); java.util.Collections.sort(list, (a, b) -> Integer.compare(a.rank, b.rank)); System.out.println(list.get(0).rank); System.out.println(list.get(2).rank);",
        "static class Item { int rank; Item(int rank) { this.rank = rank; } }",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn binary_search_finds_exact_match_index() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(20); list.add(30); System.out.println(java.util.Collections.binarySearch(list, 20));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn binary_search_returns_negative_insertion_point_minus_one() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(30); System.out.println(java.util.Collections.binarySearch(list, 20));",
    );
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn binary_search_on_singleton_list_finds_only_element() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(42); System.out.println(java.util.Collections.binarySearch(list, 42));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn binary_search_absent_on_singleton_returns_negative_two() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(42); System.out.println(java.util.Collections.binarySearch(list, 7));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn linkedlist_peek_first_reads_head_without_removal() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); deque.add(11); deque.add(22); System.out.println(deque.peekFirst()); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["11", "2"]);
}

#[test]
fn linkedlist_peek_last_reads_tail_without_removal() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); deque.add(11); deque.add(22); System.out.println(deque.peekLast());",
    );
    assert_eq!(out, vec!["22"]);
}

#[test]
fn linkedlist_poll_first_removes_and_returns_head() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); deque.add(5); deque.add(6); System.out.println(deque.pollFirst()); System.out.println(deque.get(0));",
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn linkedlist_poll_last_removes_and_returns_tail() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); deque.add(5); deque.add(6); System.out.println(deque.pollLast()); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["6", "1"]);
}

#[test]
fn linkedlist_peek_first_on_empty_returns_null() {
    let out = run_main(
        "java.util.LinkedList<String> deque = new java.util.LinkedList<String>(); System.out.println(deque.peekFirst());",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn linkedlist_poll_first_on_empty_returns_null() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); System.out.println(deque.pollFirst());",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn arraydeque_peek_first_without_removing_head() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offerLast(3); deque.offerLast(4); System.out.println(deque.peekFirst()); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn arraydeque_poll_first_removes_head_element() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offerLast(8); deque.offerLast(9); System.out.println(deque.pollFirst()); System.out.println(deque.peekFirst());",
    );
    assert_eq!(out, vec!["8", "9"]);
}

#[test]
fn arraydeque_offer_last_appends_to_tail() {
    let out = run_main(
        "java.util.ArrayDeque<String> deque = new java.util.ArrayDeque<String>(); deque.offerLast(\"a\"); deque.offerLast(\"b\"); System.out.println(deque.peekLast());",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn linkedlist_offer_first_inserts_at_head() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); deque.add(2); deque.offerFirst(1); System.out.println(deque.get(0));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn list_iterator_previous_index_before_first_next() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.ListIterator<Integer> it = list.listIterator(); System.out.println(it.previousIndex());",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn sublist_last_element_matches_parent_index() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(20); list.add(30); java.util.List<Integer> sub = list.subList(1, 3); System.out.println(sub.get(1));",
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn add_all_at_index_middle_splits_original_sequence() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(4); java.util.ArrayList<Integer> mid = new java.util.ArrayList<Integer>(); mid.add(2); mid.add(3); list.addAll(1, mid); System.out.println(list.get(2)); System.out.println(list.get(3));",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn remove_all_self_reference_clears_entire_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.removeAll(list); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["0"]);
}
