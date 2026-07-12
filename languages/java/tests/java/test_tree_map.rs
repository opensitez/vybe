use crate::helpers::{run_in_main, run_main};

#[test]
fn treemap_integer_keys_iterate_in_natural_ascending_order() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(30, \"c\"); map.put(10, \"a\"); map.put(20, \"b\"); java.util.Iterator<Integer> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn treemap_string_keys_sort_lexicographically_on_iteration() {
    let out = run_main(
        "java.util.TreeMap<String, Integer> map = new java.util.TreeMap<String, Integer>(); map.put(\"delta\", 4); map.put(\"alpha\", 1); map.put(\"charlie\", 3); map.put(\"bravo\", 2); for (String key : map.keySet()) { System.out.println(key); }",
    );
    assert_eq!(out, vec!["alpha", "bravo", "charlie", "delta"]);
}

#[test]
fn treemap_first_key_returns_smallest_natural_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(7, \"seven\"); map.put(2, \"two\"); map.put(5, \"five\"); System.out.println(map.firstKey());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treemap_last_key_returns_largest_natural_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(7, \"seven\"); map.put(2, \"two\"); map.put(5, \"five\"); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn treemap_first_key_after_single_insertion() {
    let out = run_main(
        "java.util.TreeMap<String, Integer> map = new java.util.TreeMap<String, Integer>(); map.put(\"solo\", 1); System.out.println(map.firstKey());",
    );
    assert_eq!(out, vec!["solo"]);
}

#[test]
fn treemap_last_key_after_single_insertion() {
    let out = run_main(
        "java.util.TreeMap<String, Integer> map = new java.util.TreeMap<String, Integer>(); map.put(\"solo\", 1); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["solo"]);
}

#[test]
fn treemap_first_entry_pairs_minimum_key_with_value() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(9, \"nine\"); map.put(3, \"three\"); java.util.Map.Entry<Integer, String> entry = map.firstEntry(); System.out.println(entry.getKey()); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["3", "three"]);
}

#[test]
fn treemap_last_entry_pairs_maximum_key_with_value() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(9, \"nine\"); map.put(3, \"three\"); java.util.Map.Entry<Integer, String> entry = map.lastEntry(); System.out.println(entry.getKey()); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["9", "nine"]);
}

#[test]
fn treemap_custom_comparator_orders_keys_descending() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>((a, b) -> b - a); map.put(1, \"a\"); map.put(3, \"c\"); map.put(2, \"b\"); System.out.println(map.firstKey()); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn treemap_reverse_comparator_iteration_reads_high_to_low() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>((a, b) -> b - a); map.put(10, \"ten\"); map.put(30, \"thirty\"); map.put(20, \"twenty\"); java.util.Iterator<Integer> it = map.keySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["30", "20"]);
}

#[test]
fn treemap_comparator_by_string_length_sorts_shortest_key_first() {
    let out = run_main(
        "java.util.TreeMap<String, Integer> map = new java.util.TreeMap<String, Integer>((a, b) -> Integer.compare(a.length(), b.length())); map.put(\"longer\", 2); map.put(\"a\", 1); map.put(\"mid\", 3); System.out.println(map.firstKey()); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["a", "longer"]);
}

#[test]
fn treemap_anonymous_comparator_class_reverses_integer_order() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(new java.util.Comparator<Integer>() { public int compare(Integer a, Integer b) { return b - a; } }); map.put(4, \"four\"); map.put(1, \"one\"); System.out.println(map.firstKey());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn treemap_ceiling_key_returns_exact_match_when_present() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(5, \"five\"); map.put(15, \"fifteen\"); System.out.println(map.ceilingKey(15));",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn treemap_ceiling_key_returns_next_higher_when_between_keys() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(5, \"five\"); map.put(15, \"fifteen\"); System.out.println(map.ceilingKey(9));",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn treemap_floor_key_returns_exact_match_when_present() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(5, \"five\"); map.put(15, \"fifteen\"); System.out.println(map.floorKey(5));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn treemap_floor_key_returns_next_lower_when_between_keys() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(5, \"five\"); map.put(15, \"fifteen\"); System.out.println(map.floorKey(9));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn treemap_higher_key_skips_equal_candidate() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(8, \"eight\"); map.put(16, \"sixteen\"); System.out.println(map.higherKey(8));",
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn treemap_lower_key_skips_equal_candidate() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(8, \"eight\"); map.put(16, \"sixteen\"); System.out.println(map.lowerKey(16));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn treemap_ceiling_entry_returns_matching_value() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(12, \"dozen\"); map.put(24, \"two-dozen\"); java.util.Map.Entry<Integer, String> entry = map.ceilingEntry(18); System.out.println(entry.getKey()); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["24", "two-dozen"]);
}

#[test]
fn treemap_floor_entry_returns_matching_value() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(12, \"dozen\"); map.put(24, \"two-dozen\"); java.util.Map.Entry<Integer, String> entry = map.floorEntry(18); System.out.println(entry.getKey()); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["12", "dozen"]);
}

#[test]
fn treemap_higher_entry_finds_strictly_greater_pair() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"one\"); map.put(3, \"three\"); java.util.Map.Entry<Integer, String> entry = map.higherEntry(1); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["three"]);
}

#[test]
fn treemap_lower_entry_finds_strictly_smaller_pair() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"one\"); map.put(3, \"three\"); java.util.Map.Entry<Integer, String> entry = map.lowerEntry(3); System.out.println(entry.getValue());",
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn treemap_submap_contains_lower_bound_inclusive() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); map.put(4, \"d\"); java.util.SortedMap<Integer, String> sub = map.subMap(2, 4); System.out.println(sub.firstKey()); System.out.println(sub.lastKey());",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn treemap_submap_excludes_upper_exclusive_bound() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); map.put(4, \"d\"); java.util.SortedMap<Integer, String> sub = map.subMap(2, 4); System.out.println(sub.containsKey(4));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn treemap_head_map_includes_keys_strictly_below_to_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); java.util.SortedMap<Integer, String> head = map.headMap(3); System.out.println(head.size()); System.out.println(head.containsKey(3));",
    );
    assert_eq!(out, vec!["2", "false"]);
}

#[test]
fn treemap_tail_map_includes_from_key_inclusive() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); java.util.SortedMap<Integer, String> tail = map.tailMap(2); System.out.println(tail.firstKey()); System.out.println(tail.containsKey(1));",
    );
    assert_eq!(out, vec!["2", "false"]);
}

#[test]
fn treemap_navigable_submap_get_reads_interior_mapping() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(10, \"ten\"); map.put(20, \"twenty\"); map.put(30, \"thirty\"); System.out.println(map.subMap(10, 30).get(20));",
    );
    assert_eq!(out, vec!["twenty"]);
}

#[test]
fn treemap_remove_minimum_updates_first_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); map.remove(1); System.out.println(map.firstKey());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treemap_remove_maximum_updates_last_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); map.remove(3); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn treemap_put_overwrite_keeps_single_entry_per_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(5, \"old\"); map.put(5, \"new\"); System.out.println(map.size()); System.out.println(map.get(5));",
    );
    assert_eq!(out, vec!["1", "new"]);
}

#[test]
fn treemap_negative_integer_keys_sort_numerically() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(-1, \"neg\"); map.put(0, \"zero\"); map.put(-5, \"min\"); System.out.println(map.firstKey()); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["-5", "0"]);
}

#[test]
fn treemap_poll_first_entry_removes_smallest_pair() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(2, \"two\"); map.put(1, \"one\"); java.util.Map.Entry<Integer, String> entry = map.pollFirstEntry(); System.out.println(entry.getKey()); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn treemap_poll_last_entry_removes_largest_pair() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(2, \"two\"); map.put(1, \"one\"); java.util.Map.Entry<Integer, String> entry = map.pollLastEntry(); System.out.println(entry.getKey()); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn treemap_descending_key_set_starts_at_maximum() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); map.put(3, \"c\"); java.util.Iterator<Integer> it = map.descendingKeySet().iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn treemap_descending_map_first_key_is_original_last() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(4, \"d\"); map.put(1, \"a\"); map.put(3, \"c\"); java.util.NavigableMap<Integer, String> reversed = map.descendingMap(); System.out.println(reversed.firstKey());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn treemap_ceiling_key_returns_null_above_maximum() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); System.out.println(map.ceilingKey(99));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn treemap_floor_key_returns_null_below_minimum() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(); map.put(1, \"a\"); map.put(2, \"b\"); System.out.println(map.floorKey(-1));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn treemap_comparator_natural_order_matches_default_ordering() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> natural = new java.util.TreeMap<Integer, String>(); natural.put(2, \"b\"); natural.put(1, \"a\"); java.util.TreeMap<Integer, String> explicit = new java.util.TreeMap<Integer, String>(java.util.Comparator.naturalOrder()); explicit.put(2, \"b\"); explicit.put(1, \"a\"); System.out.println(natural.firstKey()); System.out.println(explicit.firstKey());",
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn treemap_comparator_reverse_order_puts_largest_key_first() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(java.util.Comparator.reverseOrder()); map.put(1, \"a\"); map.put(3, \"c\"); map.put(2, \"b\"); System.out.println(map.firstKey()); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn treemap_custom_comparator_with_record_values_orders_by_rank() {
    let out = run_in_main(
        "java.util.TreeMap<Ranked, String> map = new java.util.TreeMap<Ranked, String>((a, b) -> Integer.compare(a.score, b.score)); map.put(new Ranked(30), \"high\"); map.put(new Ranked(10), \"low\"); map.put(new Ranked(20), \"mid\"); System.out.println(map.firstKey().score); System.out.println(map.lastKey().score);",
        "static class Ranked { int score; Ranked(int score) { this.score = score; } }",
    );
    assert_eq!(out, vec!["10", "30"]);
}
