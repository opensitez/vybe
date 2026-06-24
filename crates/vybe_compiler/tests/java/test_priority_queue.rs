use crate::helpers::{run_in_main, run_main};

#[test]
fn priorityqueue_empty_poll_returns_null() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn priorityqueue_min_heap_poll_returns_smallest_offer() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(50); pq.offer(10); pq.offer(30); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn priorityqueue_offer_then_poll_drains_in_ascending_order() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(9); pq.offer(1); pq.offer(5); pq.offer(3); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "3", "5", "9"]);
}

#[test]
fn priorityqueue_peek_reads_minimum_without_removing() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(8); pq.offer(2); pq.offer(6); System.out.println(pq.peek()); System.out.println(pq.size());",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn priorityqueue_peek_then_poll_return_same_minimum() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(15); pq.offer(3); pq.offer(11); System.out.println(pq.peek()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn priorityqueue_single_element_peek_equals_poll() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(99); System.out.println(pq.peek()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["99", "99"]);
}

#[test]
fn priorityqueue_negative_integers_poll_smallest_first() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(-1); pq.offer(-5); pq.offer(0); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["-5", "-1"]);
}

#[test]
fn priorityqueue_size_counts_all_offered_elements() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(3); pq.offer(1); pq.offer(2); System.out.println(pq.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn priorityqueue_size_decrements_after_each_poll() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(4); pq.offer(2); pq.poll(); System.out.println(pq.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn priorityqueue_poll_after_full_drain_returns_null() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(1); pq.poll(); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn priorityqueue_clear_empties_all_elements() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(1); pq.offer(2); pq.clear(); System.out.println(pq.size()); System.out.println(pq.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn priorityqueue_add_method_behaves_like_offer_for_ordering() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.add(7); pq.add(2); pq.add(9); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["2", "7"]);
}

#[test]
fn priorityqueue_element_returns_minimum_like_peek() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(12); pq.offer(4); System.out.println(pq.element()); System.out.println(pq.size());",
    );
    assert_eq!(out, vec!["4", "2"]);
}

#[test]
fn priorityqueue_offer_after_partial_poll_reorders_heap() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(5); pq.offer(1); pq.offer(9); pq.poll(); pq.offer(0); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn priorityqueue_natural_order_large_spread_polls_ascending() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(1000); pq.offer(-100); pq.offer(0); pq.offer(500); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["-100", "0"]);
}

#[test]
fn priorityqueue_duplicate_offers_poll_equal_values_in_some_order() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(5); pq.offer(5); pq.offer(5); System.out.println(pq.poll()); System.out.println(pq.size());",
    );
    assert_eq!(out, vec!["5", "2"]);
}

#[test]
fn priorityqueue_contains_finds_offered_value() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(3); pq.offer(7); System.out.println(pq.contains(7));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn priorityqueue_contains_rejects_absent_value() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(3); System.out.println(pq.contains(9));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn priorityqueue_remove_specific_element_updates_size() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(1); pq.offer(2); pq.offer(3); pq.remove(Integer.valueOf(2)); System.out.println(pq.size()); System.out.println(pq.peek());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn priorityqueue_strings_natural_order_polls_lexicographic_minimum() {
    let out = run_main(
        "java.util.PriorityQueue<String> pq = new java.util.PriorityQueue<String>(); pq.offer(\"cherry\"); pq.offer(\"apple\"); pq.offer(\"banana\"); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["apple"]);
}

#[test]
fn priorityqueue_custom_comparator_polls_largest_integer_first() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>((a, b) -> b - a); pq.offer(1); pq.offer(3); pq.offer(2); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn priorityqueue_reverse_comparator_makes_max_heap() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>((a, b) -> b - a); pq.offer(10); pq.offer(40); pq.offer(20); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["40", "20", "10"]);
}

#[test]
fn priorityqueue_comparator_reverse_order_polls_descending() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(java.util.Comparator.reverseOrder()); pq.offer(2); pq.offer(8); pq.offer(5); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn priorityqueue_anonymous_comparator_class_builds_max_heap() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(new java.util.Comparator<Integer>() { public int compare(Integer a, Integer b) { return b - a; } }); pq.offer(4); pq.offer(1); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn priorityqueue_custom_comparator_by_string_length() {
    let out = run_main(
        "java.util.PriorityQueue<String> pq = new java.util.PriorityQueue<String>((a, b) -> Integer.compare(a.length(), b.length())); pq.offer(\"longer\"); pq.offer(\"a\"); pq.offer(\"mid\"); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["a", "mid"]);
}

#[test]
fn priorityqueue_custom_comparator_modulo_priority() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>((a, b) -> Integer.compare(a % 10, b % 10)); pq.offer(17); pq.offer(4); pq.offer(23); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["4", "17"]);
}

#[test]
fn priorityqueue_initial_capacity_still_orders_correctly() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(2); pq.offer(6); pq.offer(1); pq.offer(3); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn priorityqueue_comparator_constructor_with_initial_capacity() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(4, (a, b) -> b - a); pq.offer(2); pq.offer(5); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn priorityqueue_interleaved_offer_poll_maintains_heap_property() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(5); System.out.println(pq.poll()); pq.offer(1); pq.offer(9); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["5", "1", "9"]);
}

#[test]
fn priorityqueue_peek_after_partial_drain_reads_new_minimum() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(7); pq.offer(2); pq.offer(9); pq.poll(); System.out.println(pq.peek());",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn priorityqueue_peek_on_empty_returns_null() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); System.out.println(pq.peek());",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn priorityqueue_offer_zero_and_negative_orders_correctly() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(0); pq.offer(-2); pq.offer(2); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["-2", "0", "2"]);
}

#[test]
fn priorityqueue_batch_offer_then_full_drain_ascending() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(6); pq.offer(2); pq.offer(8); pq.offer(1); pq.offer(4); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "2", "4", "6", "8"]);
}

#[test]
fn priorityqueue_reoffer_after_drain_rebuilds_min_heap() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(3); pq.poll(); pq.offer(1); pq.offer(2); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn priorityqueue_remove_on_empty_returns_false() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); System.out.println(pq.remove(Integer.valueOf(1)));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn priorityqueue_custom_comparator_with_record_priority_field() {
    let out = run_in_main(
        "java.util.PriorityQueue<Task> pq = new java.util.PriorityQueue<Task>((a, b) -> Integer.compare(a.priority, b.priority)); pq.offer(new Task(30)); pq.offer(new Task(5)); pq.offer(new Task(15)); System.out.println(pq.poll().priority); System.out.println(pq.poll().priority);",
        "static class Task { int priority; Task(int priority) { this.priority = priority; } }",
    );
    assert_eq!(out, vec!["5", "15"]);
}

#[test]
fn priorityqueue_natural_order_five_element_poll_sequence() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(7); pq.offer(2); pq.offer(9); pq.offer(4); pq.offer(1); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "2", "4"]);
}

#[test]
fn priorityqueue_comparator_natural_order_matches_default_min_heap() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> def = new java.util.PriorityQueue<Integer>(); def.offer(3); def.offer(1); java.util.PriorityQueue<Integer> explicit = new java.util.PriorityQueue<Integer>(java.util.Comparator.naturalOrder()); explicit.offer(3); explicit.offer(1); System.out.println(def.poll()); System.out.println(explicit.poll());",
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn priorityqueue_to_array_preserves_size_not_order() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(3); pq.offer(1); pq.offer(2); System.out.println(pq.toArray().length);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn priorityqueue_offer_many_then_peek_still_shows_global_minimum() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(20); pq.offer(5); pq.offer(15); pq.offer(1); pq.offer(10); System.out.println(pq.peek()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "1"]);
}
