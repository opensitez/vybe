/// java.util.concurrent.LinkedBlockingQueue — bounded blocking FIFO queue.
use crate::helpers::{run_in_main, run_main};

#[test]
fn linked_blocking_queue_offer_adds_element_when_capacity_available() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.offer(10)); System.out.println(q.peek());",
    );
    assert_eq!(out, vec!["true", "10"]);
}

#[test]
fn linked_blocking_queue_poll_removes_head_element() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(5); q.offer(6); System.out.println(q.poll()); System.out.println(q.poll());",
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn linked_blocking_queue_peek_reads_head_without_removing() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(11); System.out.println(q.peek()); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["11", "1"]);
}

#[test]
fn linked_blocking_queue_put_and_take_round_trip() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<String> q = new java.util.concurrent.LinkedBlockingQueue<String>(); q.put(\"item\"); System.out.println(q.take());",
    );
    assert_eq!(out, vec!["item"]);
}

#[test]
fn linked_blocking_queue_fifo_order_preserved() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); q.offer(3); System.out.println(q.poll()); System.out.println(q.poll()); System.out.println(q.poll());",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn linked_blocking_queue_new_queue_is_empty() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_size_reflects_element_count() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linked_blocking_queue_bounded_capacity_constructor() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(2); q.offer(1); q.offer(2); System.out.println(q.remainingCapacity());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn linked_blocking_queue_remaining_capacity_on_empty_unbounded() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.remainingCapacity() > 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_add_throws_when_bounded_full() {
    let out = run_in_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(1); q.add(1); try { q.add(2); System.out.println(\"ok\"); } catch (IllegalStateException e) { System.out.println(\"full\"); }",
        "",
    );
    assert_eq!(out, vec!["full"]);
}

#[test]
fn linked_blocking_queue_offer_fails_when_bounded_full() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(1); q.offer(1); System.out.println(q.offer(2));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn linked_blocking_queue_poll_on_empty_returns_null() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.poll() == null);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_peek_on_empty_returns_null() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.peek() == null);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_element_on_empty_throws() {
    let out = run_in_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); try { q.element(); System.out.println(\"ok\"); } catch (java.util.NoSuchElementException e) { System.out.println(\"empty\"); }",
        "",
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn linked_blocking_queue_remove_on_empty_throws() {
    let out = run_in_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); try { q.remove(); System.out.println(\"ok\"); } catch (java.util.NoSuchElementException e) { System.out.println(\"empty\"); }",
        "",
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn linked_blocking_queue_element_returns_head_without_removing() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(7); System.out.println(q.element()); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["7", "1"]);
}

#[test]
fn linked_blocking_queue_remove_returns_head() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(8); System.out.println(q.remove()); System.out.println(q.isEmpty());",
    );
    assert_eq!(out, vec!["8", "true"]);
}

#[test]
fn linked_blocking_queue_drain_to_moves_elements_to_list() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); java.util.ArrayList<Integer> dest = new java.util.ArrayList<Integer>(); System.out.println(q.drainTo(dest)); System.out.println(dest.size());",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn linked_blocking_queue_drain_to_with_max_moves_partial() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); q.offer(3); java.util.ArrayList<Integer> dest = new java.util.ArrayList<Integer>(); System.out.println(q.drainTo(dest, 2)); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn linked_blocking_queue_contains_finds_offered_element() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(42); System.out.println(q.contains(42));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_remove_object_eliminates_matching() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); System.out.println(q.remove(Integer.valueOf(1))); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["true", "1"]);
}

#[test]
fn linked_blocking_queue_iterator_traverses_fifo_order() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); java.util.Iterator<Integer> it = q.iterator(); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn linked_blocking_queue_clear_empties_all_elements() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); q.clear(); System.out.println(q.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_to_array_has_correct_length() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); Object[] arr = q.toArray(); System.out.println(arr.length);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linked_blocking_queue_offer_with_timeout_succeeds_when_space_available() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(2); System.out.println(q.offer(1, 10, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_poll_with_timeout_returns_null_when_empty() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.poll(1, java.util.concurrent.TimeUnit.MILLISECONDS) == null);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_poll_with_timeout_retrieves_element() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(9); System.out.println(q.poll(10, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn linked_blocking_queue_producer_consumer_with_put_take() {
    let types = r#"
        static String consumed;
    "#;
    let out = run_in_main(
        r#"java.util.concurrent.LinkedBlockingQueue<String> q = new java.util.concurrent.LinkedBlockingQueue<String>(); Thread producer = new Thread(() -> { try { q.put("data"); } catch (InterruptedException e) {} }); Thread consumer = new Thread(() -> { try { consumed = q.take(); } catch (InterruptedException e) {} }); producer.start(); consumer.start(); producer.join(); consumer.join(); System.out.println(consumed);"#,
        types,
    );
    assert_eq!(out, vec!["data"]);
}

#[test]
fn linked_blocking_queue_two_producers_both_enqueue() {
    let types = r#"
        static java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>();
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { q.put(1); } catch (InterruptedException e) {} }); Thread t2 = new Thread(() -> { try { q.put(2); } catch (InterruptedException e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(q.size());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linked_blocking_queue_capacity_constructor_with_initial_collection() {
    let out = run_main(
        "java.util.ArrayList<Integer> seed = new java.util.ArrayList<Integer>(); seed.add(4); seed.add(5); java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(seed); System.out.println(q.size()); System.out.println(q.poll());",
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn linked_blocking_queue_offer_timeout_fails_when_full() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(1); q.offer(1); System.out.println(q.offer(2, 1, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn linked_blocking_queue_remove_object_returns_false_when_absent() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); System.out.println(q.remove(Integer.valueOf(9)));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn linked_blocking_queue_contains_false_for_absent() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.contains(99));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn linked_blocking_queue_add_returns_true_on_success() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); System.out.println(q.add(3));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn linked_blocking_queue_put_multiple_then_drain_to_empties() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.put(1); q.put(2); java.util.ArrayList<Integer> dest = new java.util.ArrayList<Integer>(); q.drainTo(dest); System.out.println(q.isEmpty()); System.out.println(dest.size());",
    );
    assert_eq!(out, vec!["true", "2"]);
}

#[test]
fn linked_blocking_queue_peek_after_poll_reads_new_head() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); q.offer(2); q.poll(); System.out.println(q.peek());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linked_blocking_queue_size_after_drain_to_is_zero() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); java.util.ArrayList<Integer> dest = new java.util.ArrayList<Integer>(); q.drainTo(dest); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn linked_blocking_queue_unbounded_accepts_many_offers() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); for (int i = 0; i < 5; i++) q.offer(i); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn linked_blocking_queue_iterator_does_not_remove_on_next() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(1); java.util.Iterator<Integer> it = q.iterator(); it.next(); System.out.println(q.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn linked_blocking_queue_to_array_typed_copies_elements() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<String> q = new java.util.concurrent.LinkedBlockingQueue<String>(); q.offer(\"a\"); q.offer(\"b\"); String[] arr = q.toArray(new String[0]); System.out.println(arr[1]);",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn linked_blocking_queue_bounded_capacity_three_accepts_three() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(3); q.offer(1); q.offer(2); q.offer(3); System.out.println(q.size()); System.out.println(q.remainingCapacity());",
    );
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn linked_blocking_queue_consumer_polls_in_fifo_from_two_producers() {
    let types = r#"
        static java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>();
        static int first;
        static int second;
    "#;
    let out = run_in_main(
        "q.offer(1); q.offer(2); Thread t = new Thread(() -> { first = q.poll(); second = q.poll(); }); t.start(); t.join(); System.out.println(first); System.out.println(second);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn linked_blocking_queue_put_interruptibly_accepts_element() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.put(55); System.out.println(q.poll());",
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn linked_blocking_queue_take_then_offer_maintains_fifo() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.put(1); q.put(2); System.out.println(q.take()); q.put(3); System.out.println(q.take()); System.out.println(q.take());",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn linked_blocking_queue_drain_to_empty_queue_returns_zero() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); java.util.ArrayList<Integer> dest = new java.util.ArrayList<Integer>(); System.out.println(q.drainTo(dest));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn linked_blocking_queue_offer_poll_single_element_round_trip() {
    let out = run_main(
        "java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(); q.offer(100); System.out.println(q.poll()); System.out.println(q.isEmpty());",
    );
    assert_eq!(out, vec!["100", "true"]);
}

#[test]
fn linked_blocking_queue_take_blocks_until_element_available() {
    let types = r#"
        static java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>();
        static int taken = -1;
    "#;
    let out = run_in_main(
        "Thread consumer = new Thread(() -> { try { taken = q.take(); } catch (InterruptedException e) {} }); consumer.start(); Thread.sleep(5); q.put(99); consumer.join(); System.out.println(taken);",
        types,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn linked_blocking_queue_bounded_put_blocks_until_space() {
    let types = r#"
        static java.util.concurrent.LinkedBlockingQueue<Integer> q = new java.util.concurrent.LinkedBlockingQueue<Integer>(1);
        static boolean putCompleted = false;
    "#;
    let out = run_in_main(
        "q.put(1); Thread consumer = new Thread(() -> { try { q.take(); putCompleted = true; } catch (InterruptedException e) {} }); consumer.start(); Thread.sleep(5); q.put(2); consumer.join(); System.out.println(putCompleted); System.out.println(q.size());",
        types,
    );
    assert_eq!(out, vec!["true", "1"]);
}
