use crate::helpers::run_main;

#[test]
fn arraydeque_offer_appends_element_to_tail() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(10); deque.offer(20); System.out.println(deque.peek()); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["10", "2"]);
}

#[test]
fn arraydeque_poll_removes_element_from_head() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(5); deque.offer(6); System.out.println(deque.poll()); System.out.println(deque.poll());",
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn arraydeque_peek_reads_head_without_removing() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(11); deque.offer(22); System.out.println(deque.peek()); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["11", "2"]);
}

#[test]
fn arraydeque_add_on_unbounded_deque_appends_like_offer() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.add(3); deque.add(4); System.out.println(deque.poll()); System.out.println(deque.poll());",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn arraydeque_offer_poll_fifo_order_preserved() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(1); deque.offer(2); deque.offer(3); System.out.println(deque.poll()); System.out.println(deque.poll()); System.out.println(deque.poll());",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn arraydeque_push_prepends_for_stack_usage() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.push(2); deque.push(1); System.out.println(deque.peek());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arraydeque_pop_removes_front_stack_top() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.push(2); deque.push(1); System.out.println(deque.pop()); System.out.println(deque.pop());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn arraydeque_push_pop_lifo_order() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.push(10); deque.push(20); deque.push(30); System.out.println(deque.pop()); System.out.println(deque.pop()); System.out.println(deque.pop());",
    );
    assert_eq!(out, vec!["30", "20", "10"]);
}

#[test]
fn arraydeque_peek_after_push_reads_stack_top() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.push(7); deque.push(8); System.out.println(deque.peek()); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["8", "2"]);
}

#[test]
fn arraydeque_size_counts_offered_elements() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(1); deque.offer(2); deque.offer(3); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraydeque_poll_on_empty_returns_null() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); System.out.println(deque.poll());",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn arraydeque_offer_multiple_then_poll_sequence() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(100); deque.offer(200); deque.offer(300); System.out.println(deque.poll()); System.out.println(deque.peek());",
    );
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn linkedlist_as_queue_offer_poll_order() {
    let out = run_main(
        "java.util.LinkedList<Integer> queue = new java.util.LinkedList<Integer>(); queue.offer(4); queue.offer(5); System.out.println(queue.poll()); System.out.println(queue.poll());",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn linkedlist_as_deque_add_first_and_offer_tail() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); deque.addFirst(2); deque.offer(3); System.out.println(deque.poll()); System.out.println(deque.poll());",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn priorityqueue_poll_returns_smallest_value_first() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(30); pq.offer(10); pq.offer(20); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn priorityqueue_offer_inserts_then_orders_on_poll() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(9); pq.offer(1); pq.offer(5); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn priorityqueue_peek_views_minimum_without_removing() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(8); pq.offer(2); pq.offer(6); System.out.println(pq.peek()); System.out.println(pq.size());",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn priorityqueue_multiple_polls_drain_in_sorted_order() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(4); pq.offer(1); pq.offer(3); pq.offer(2); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn arraydeque_remove_first_drops_head_element() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(5); deque.offer(6); deque.removeFirst(); System.out.println(deque.peek());",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arraydeque_remove_last_drops_tail_element() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(5); deque.offer(6); deque.removeLast(); System.out.println(deque.peekLast());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn arraydeque_peek_first_reads_head_element() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(12); deque.offer(13); System.out.println(deque.peekFirst());",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn arraydeque_peek_last_reads_tail_element() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(12); deque.offer(13); System.out.println(deque.peekLast());",
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn arraydeque_add_last_appends_at_tail() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.addLast(1); deque.addLast(2); System.out.println(deque.peekLast());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arraydeque_add_first_inserts_at_head() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.addLast(2); deque.addFirst(1); System.out.println(deque.peekFirst());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn queue_interface_offer_poll_peek_chain() {
    let out = run_main(
        "java.util.Queue<Integer> queue = new java.util.ArrayDeque<Integer>(); queue.offer(7); queue.offer(8); System.out.println(queue.peek()); System.out.println(queue.poll()); System.out.println(queue.peek());",
    );
    assert_eq!(out, vec!["7", "7", "8"]);
}

#[test]
fn deque_interface_push_pop_stack_operations() {
    let out = run_main(
        "java.util.Deque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.push(1); deque.push(2); System.out.println(deque.pop()); System.out.println(deque.pop());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn priorityqueue_larger_values_poll_after_smaller_ones() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(50); pq.offer(10); pq.offer(40); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["10", "40"]);
}

#[test]
fn priorityqueue_single_element_peek_equals_poll() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(99); System.out.println(pq.peek()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["99", "99"]);
}

#[test]
fn arraydeque_element_reads_head_like_peek() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(21); deque.offer(22); System.out.println(deque.element()); System.out.println(deque.size());",
    );
    assert_eq!(out, vec!["21", "2"]);
}

#[test]
fn arraydeque_remove_object_drops_matching_entry() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(1); deque.offer(2); deque.offer(3); deque.remove(Integer.valueOf(2)); System.out.println(deque.size()); System.out.println(deque.peek());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn priorityqueue_offer_with_negative_integers_orders_correctly() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(-1); pq.offer(-5); pq.offer(0); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["-5", "-1"]);
}

#[test]
fn arraydeque_clear_empties_all_elements() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(1); deque.offer(2); deque.clear(); System.out.println(deque.size()); System.out.println(deque.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn linkedlist_offer_poll_on_sequential_fifo_usage() {
    let out = run_main(
        "java.util.LinkedList<Integer> queue = new java.util.LinkedList<Integer>(); queue.offer(9); queue.offer(8); queue.offer(7); System.out.println(queue.poll()); System.out.println(queue.poll()); System.out.println(queue.poll());",
    );
    assert_eq!(out, vec!["9", "8", "7"]);
}

#[test]
fn arraydeque_poll_after_clear_returns_null() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(1); deque.clear(); System.out.println(deque.poll());",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn priorityqueue_size_after_multiple_offers() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(3); pq.offer(1); pq.offer(2); System.out.println(pq.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraydeque_remove_on_queue_interface_drops_head() {
    let out = run_main(
        "java.util.Queue<Integer> queue = new java.util.ArrayDeque<Integer>(); queue.offer(4); queue.offer(5); System.out.println(queue.remove()); System.out.println(queue.peek());",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn deque_add_first_then_peek_first_shows_new_head() {
    let out = run_main(
        "java.util.Deque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.addLast(2); deque.addFirst(1); System.out.println(deque.peekFirst());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn priorityqueue_natural_ordering_polls_ascending_values() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(7); pq.offer(2); pq.offer(9); pq.offer(4); System.out.println(pq.poll()); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["2", "4", "7"]);
}

#[test]
fn arraydeque_as_stack_push_push_pop_pop_order() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> stack = new java.util.ArrayDeque<Integer>(); stack.push(1); stack.push(2); stack.push(3); System.out.println(stack.pop()); System.out.println(stack.pop()); System.out.println(stack.pop());",
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn linkedlist_push_pop_deque_stack_operations() {
    let out = run_main(
        "java.util.LinkedList<Integer> deque = new java.util.LinkedList<Integer>(); deque.push(10); deque.push(20); System.out.println(deque.pop()); System.out.println(deque.pop());",
    );
    assert_eq!(out, vec!["20", "10"]);
}

#[test]
fn priorityqueue_peek_then_poll_return_same_minimum() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(15); pq.offer(3); pq.offer(11); System.out.println(pq.peek()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn arraydeque_offer_at_tail_poll_from_head_three_step_sequence() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(100); deque.offer(200); deque.offer(300); System.out.println(deque.poll()); System.out.println(deque.poll()); System.out.println(deque.poll());",
    );
    assert_eq!(out, vec!["100", "200", "300"]);
}

#[test]
fn queue_add_method_appends_element_for_fifo_usage() {
    let out = run_main(
        "java.util.Queue<Integer> queue = new java.util.LinkedList<Integer>(); queue.add(6); queue.add(7); System.out.println(queue.remove()); System.out.println(queue.element());",
    );
    assert_eq!(out, vec!["6", "7"]);
}

#[test]
fn arraydeque_remove_last_on_single_element_deque_empties_it() {
    let out = run_main(
        "java.util.ArrayDeque<Integer> deque = new java.util.ArrayDeque<Integer>(); deque.offer(42); deque.removeLast(); System.out.println(deque.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn priorityqueue_poll_after_draining_returns_null() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(); pq.offer(1); pq.poll(); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["null"]);
}
