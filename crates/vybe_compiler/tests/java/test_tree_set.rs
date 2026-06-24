use crate::helpers::{run_in_main, run_main};

#[test]
fn treeset_natural_order_iteration_prints_ascending_integers() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(5); set.add(1); set.add(3); for (int v : set) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["1", "3", "5"]);
}

#[test]
fn treeset_first_returns_minimum_element() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(8); set.add(2); set.add(6); System.out.println(set.first());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treeset_last_returns_maximum_element() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(8); set.add(2); set.add(6); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn treeset_head_set_excludes_to_element() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.add(3); set.add(4); java.util.SortedSet<Integer> head = set.headSet(3); System.out.println(head.contains(2)); System.out.println(head.contains(3));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn treeset_tail_set_includes_from_element() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.add(3); set.add(4); java.util.SortedSet<Integer> tail = set.tailSet(3); System.out.println(tail.contains(3)); System.out.println(tail.contains(2));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn treeset_sub_set_includes_lower_bound_and_excludes_upper() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(10); set.add(20); set.add(30); set.add(40); java.util.SortedSet<Integer> sub = set.subSet(20, 40); System.out.println(sub.contains(20)); System.out.println(sub.contains(40)); System.out.println(sub.contains(30));",
    );
    assert_eq!(out, vec!["true", "false", "true"]);
}

#[test]
fn treeset_ceiling_finds_equal_or_next_higher() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(4); set.add(12); set.add(20); System.out.println(set.ceiling(12));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn treeset_ceiling_between_elements_returns_next_higher() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(4); set.add(12); set.add(20); System.out.println(set.ceiling(7));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn treeset_floor_finds_equal_or_next_lower() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(4); set.add(12); set.add(20); System.out.println(set.floor(12));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn treeset_floor_between_elements_returns_next_lower() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(4); set.add(12); set.add(20); System.out.println(set.floor(7));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn treeset_higher_skips_equal_value() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(6); set.add(14); System.out.println(set.higher(6));",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn treeset_lower_skips_equal_value() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(6); set.add(14); System.out.println(set.lower(14));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn treeset_ceiling_returns_null_when_above_maximum() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); System.out.println(set.ceiling(100));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn treeset_floor_returns_null_when_below_minimum() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); System.out.println(set.floor(-5));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn treeset_custom_comparator_orders_descending_on_iteration() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>((a, b) -> b - a); set.add(1); set.add(3); set.add(2); for (int v : set) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn treeset_reverse_comparator_first_is_largest() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>((a, b) -> b - a); set.add(10); set.add(30); set.add(20); System.out.println(set.first()); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["30", "10"]);
}

#[test]
fn treeset_comparator_by_string_length_sorts_shortest_first() {
    let out = run_main(
        "java.util.TreeSet<String> set = new java.util.TreeSet<String>((a, b) -> Integer.compare(a.length(), b.length())); set.add(\"longword\"); set.add(\"a\"); set.add(\"mid\"); System.out.println(set.first()); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["a", "longword"]);
}

#[test]
fn treeset_string_members_iterate_lexicographically() {
    let out = run_main(
        "java.util.TreeSet<String> set = new java.util.TreeSet<String>(); set.add(\"dog\"); set.add(\"ant\"); set.add(\"cat\"); for (String s : set) { System.out.println(s); }",
    );
    assert_eq!(out, vec!["ant", "cat", "dog"]);
}

#[test]
fn treeset_head_set_size_counts_elements_below_bound() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.add(3); set.add(4); System.out.println(set.headSet(4).size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn treeset_tail_set_size_counts_elements_from_bound_upward() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.add(3); set.add(4); System.out.println(set.tailSet(2).size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn treeset_sub_set_size_counts_inclusive_lower_exclusive_upper() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.add(3); set.add(4); System.out.println(set.subSet(2, 4).size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treeset_poll_first_removes_and_returns_minimum() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(9); set.add(1); set.add(5); System.out.println(set.pollFirst()); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn treeset_poll_last_removes_and_returns_maximum() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(9); set.add(1); set.add(5); System.out.println(set.pollLast()); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["9", "2"]);
}

#[test]
fn treeset_descending_set_iterates_from_max_to_min() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(2); set.add(4); set.add(6); java.util.Iterator<Integer> it = set.descendingSet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["6", "4"]);
}

#[test]
fn treeset_descending_set_first_is_original_last() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(2); set.add(4); set.add(6); System.out.println(set.descendingSet().first());",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn treeset_negative_integers_sort_numerically() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(-1); set.add(0); set.add(-3); System.out.println(set.first()); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["-3", "0"]);
}

#[test]
fn treeset_add_duplicate_does_not_increase_size() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(7); set.add(7); set.add(7); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn treeset_remove_smallest_leaves_next_as_first() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.add(3); set.remove(1); System.out.println(set.first());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treeset_sub_set_iteration_stays_sorted_within_view() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(1); set.add(2); set.add(3); set.add(4); java.util.SortedSet<Integer> sub = set.subSet(2, 4); for (int v : sub) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn treeset_head_set_first_matches_parent_minimum_in_view() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(5); set.add(10); set.add(15); System.out.println(set.headSet(12).first());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn treeset_tail_set_last_matches_parent_maximum_in_view() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(5); set.add(10); set.add(15); System.out.println(set.tailSet(10).last());",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn treeset_comparator_natural_order_matches_default_iteration() {
    let out = run_main(
        "java.util.TreeSet<Integer> natural = new java.util.TreeSet<Integer>(); natural.add(3); natural.add(1); java.util.TreeSet<Integer> explicit = new java.util.TreeSet<Integer>(java.util.Comparator.naturalOrder()); explicit.add(3); explicit.add(1); System.out.println(natural.first()); System.out.println(explicit.first());",
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn treeset_comparator_reverse_order_puts_largest_first() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(java.util.Comparator.reverseOrder()); set.add(1); set.add(3); set.add(2); System.out.println(set.first()); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn treeset_higher_on_maximum_returns_null() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(5); set.add(10); System.out.println(set.higher(10));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn treeset_lower_on_minimum_returns_null() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(5); set.add(10); System.out.println(set.lower(5));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn treeset_head_set_ceiling_within_view_finds_next() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(2); set.add(4); set.add(6); System.out.println(set.headSet(6).ceiling(3));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn treeset_tail_set_floor_within_view_finds_previous() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); set.add(2); set.add(4); set.add(6); System.out.println(set.tailSet(4).floor(5));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn treeset_add_all_preserves_sorted_unique_membership() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(); java.util.ArrayList<Integer> batch = new java.util.ArrayList<Integer>(); batch.add(3); batch.add(1); batch.add(3); batch.add(2); set.addAll(batch); System.out.println(set.size()); System.out.println(set.first());",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn treeset_anonymous_comparator_class_reverses_order() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(new java.util.Comparator<Integer>() { public int compare(Integer a, Integer b) { return b - a; } }); set.add(2); set.add(4); System.out.println(set.first());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn treeset_custom_comparator_with_record_values_orders_by_score() {
    let out = run_in_main(
        "java.util.TreeSet<Scored> set = new java.util.TreeSet<Scored>((a, b) -> Integer.compare(a.points, b.points)); set.add(new Scored(50)); set.add(new Scored(10)); set.add(new Scored(30)); System.out.println(set.first().points); System.out.println(set.last().points);",
        "static class Scored { int points; Scored(int points) { this.points = points; } }",
    );
    assert_eq!(out, vec!["10", "50"]);
}
