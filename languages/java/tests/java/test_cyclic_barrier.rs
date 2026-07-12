/// java.util.concurrent.CyclicBarrier — reusable barrier synchronization.
use crate::helpers::{run_in_main, run_main};

#[test]
fn cyclic_barrier_get_parties_returns_constructor_value() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(3); System.out.println(barrier.getParties());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn cyclic_barrier_get_number_waiting_initially_zero() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2); System.out.println(barrier.getNumberWaiting());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn cyclic_barrier_await_with_single_party_returns_immediately() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(1); System.out.println(barrier.await());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn cyclic_barrier_two_threads_both_reach_barrier() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int arrivals = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); arrivals++; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); arrivals++; } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(arrivals);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn cyclic_barrier_await_returns_arrival_index() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int index = -1;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { index = barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(index >= 0);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_reset_clears_waiting_count() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2); barrier.reset(); System.out.println(barrier.getNumberWaiting());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn cyclic_barrier_is_broken_false_initially() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2); System.out.println(barrier.isBroken());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn cyclic_barrier_reusable_after_all_parties_arrive() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int rounds = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); rounds++; barrier.await(); rounds++; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(rounds);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn cyclic_barrier_three_parties_all_proceed_together() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(3);
        static int done = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); done++; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); done++; } catch (Exception e) {} }); Thread t3 = new Thread(() -> { try { barrier.await(); done++; } catch (Exception e) {} }); t1.start(); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(done);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn cyclic_barrier_await_with_timeout_succeeds_when_parties_arrive() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static boolean timedAwaitOk = false;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(100, java.util.concurrent.TimeUnit.MILLISECONDS); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { timedAwaitOk = barrier.await(100, java.util.concurrent.TimeUnit.MILLISECONDS) >= 0; } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(timedAwaitOk);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_await_with_timeout_times_out_when_incomplete() {
    let out = run_in_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2); try { barrier.await(1, java.util.concurrent.TimeUnit.MILLISECONDS); System.out.println(\"ok\"); } catch (java.util.concurrent.TimeoutException e) { System.out.println(\"timeout\"); } catch (Exception e) { System.out.println(\"other\"); }",
        "",
    );
    assert_eq!(out, vec!["timeout"]);
}

#[test]
fn cyclic_barrier_barrier_action_runs_when_tripped() {
    let types = r#"
        static boolean actionRan = false;
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2, () -> { actionRan = true; });
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(actionRan);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_get_number_waiting_increases_before_trip() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int waiting = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { waiting = barrier.getNumberWaiting(); barrier.await(); } catch (Exception e) {} }); Thread.sleep(5); Thread t2 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(waiting);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn cyclic_barrier_reset_while_threads_waiting_breaks_barrier() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static boolean broken = false;
    "#;
    let out = run_in_main(
        "Thread waiter = new Thread(() -> { try { barrier.await(); } catch (java.util.concurrent.BrokenBarrierException e) { broken = true; } catch (Exception e) {} }); waiter.start(); Thread.sleep(5); barrier.reset(); waiter.join(); System.out.println(barrier.isBroken());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_single_party_multiple_awaits_in_sequence() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(1); System.out.println(barrier.await()); System.out.println(barrier.await());",
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn cyclic_barrier_two_parties_exchange_data_at_barrier() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int shared = 0;
    "#;
    let out = run_in_main(
        "Thread writer = new Thread(() -> { shared = 99; try { barrier.await(); } catch (Exception e) {} }); Thread reader = new Thread(() -> { try { barrier.await(); System.out.println(shared); } catch (Exception e) {} }); writer.start(); reader.start(); writer.join(); reader.join();",
        types,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn cyclic_barrier_parties_two_get_parties_stable() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2); barrier.reset(); System.out.println(barrier.getParties());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn cyclic_barrier_four_parties_all_complete_round() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(4);
        static int count = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); count++; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); count++; } catch (Exception e) {} }); Thread t3 = new Thread(() -> { try { barrier.await(); count++; } catch (Exception e) {} }); Thread t4 = new Thread(() -> { try { barrier.await(); count++; } catch (Exception e) {} }); t1.start(); t2.start(); t3.start(); t4.start(); t1.join(); t2.join(); t3.join(); t4.join(); System.out.println(count);",
        types,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn cyclic_barrier_await_after_reset_starts_fresh_cycle() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static boolean secondRound = false;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); barrier.reset(); barrier.await(); secondRound = true; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); Thread.sleep(5); barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(secondRound);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_main_thread_plus_worker_reach_barrier() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
    "#;
    let out = run_in_main(
        "Thread worker = new Thread(() -> { try { barrier.await(); System.out.println(\"worker\"); } catch (Exception e) {} }); worker.start(); barrier.await(); System.out.println(\"main\"); worker.join();",
        types,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"worker".to_string()));
    assert!(out.contains(&"main".to_string()));
}

#[test]
fn cyclic_barrier_get_number_waiting_zero_after_trip() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int waitingAfter = -1;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); waitingAfter = barrier.getNumberWaiting(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(waitingAfter);",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn cyclic_barrier_barrier_action_runs_once_per_trip() {
    let types = r#"
        static int actionCount = 0;
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2, () -> { actionCount++; });
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); barrier.await(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(actionCount);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn cyclic_barrier_await_index_zero_for_last_arriver_in_two_party() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int lastIndex = -1;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { lastIndex = barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(lastIndex);",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn cyclic_barrier_parties_one_always_trips_immediately() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(1); System.out.println(barrier.await()); System.out.println(barrier.getNumberWaiting());",
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn cyclic_barrier_two_rounds_increment_shared_counter() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int counter = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); counter++; barrier.await(); counter++; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(counter);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn cyclic_barrier_await_timeout_returns_negative_on_timeout() {
    let out = run_in_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2); try { int idx = barrier.await(1, java.util.concurrent.TimeUnit.MILLISECONDS); System.out.println(idx); } catch (java.util.concurrent.TimeoutException e) { System.out.println(\"timeout\"); } catch (Exception e) { System.out.println(\"err\"); }",
        "",
    );
    assert_eq!(out, vec!["timeout"]);
}

#[test]
fn cyclic_barrier_reset_on_fresh_barrier_is_noop() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(3); barrier.reset(); System.out.println(barrier.isBroken());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn cyclic_barrier_three_parties_two_rounds_all_complete() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(3);
        static int total = 0;
    "#;
    let out = run_in_main(
        "Runnable task = () -> { try { barrier.await(); total++; barrier.await(); total++; } catch (Exception e) {} }; Thread t1 = new Thread(task); Thread t2 = new Thread(task); Thread t3 = new Thread(task); t1.start(); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(total);",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn cyclic_barrier_is_broken_after_forced_reset_with_waiter() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(3);
    "#;
    let out = run_in_main(
        "Thread w = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); w.start(); Thread.sleep(5); barrier.reset(); w.join(); System.out.println(barrier.isBroken());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_get_parties_large_value() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(10); System.out.println(barrier.getParties());",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn cyclic_barrier_staggered_arrivals_all_release_together() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int phase = 0;
    "#;
    let out = run_in_main(
        "Thread slow = new Thread(() -> { try { Thread.sleep(5); barrier.await(); phase = 1; } catch (Exception e) {} }); Thread fast = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); slow.start(); fast.start(); slow.join(); fast.join(); System.out.println(phase);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn cyclic_barrier_action_sets_flag_visible_to_waiters() {
    let types = r#"
        static boolean ready = false;
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2, () -> { ready = true; });
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); System.out.println(ready); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join();",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_two_party_first_arrival_waits_for_second() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static boolean firstWaiting = false;
    "#;
    let out = run_in_main(
        "Thread first = new Thread(() -> { try { firstWaiting = true; barrier.await(); System.out.println(\"first done\"); } catch (Exception e) {} }); first.start(); Thread.sleep(5); Thread second = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); second.start(); first.join(); second.join(); System.out.println(firstWaiting);",
        types,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"first done".to_string()));
    assert!(out.contains(&"true".to_string()));
}

#[test]
fn cyclic_barrier_await_after_full_trip_resets_waiting_count() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(barrier.getNumberWaiting());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn cyclic_barrier_parties_two_constructor_with_runnable() {
    let types = r#"
        static int ran = 0;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2, () -> { ran++; }); Thread t = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t.start(); barrier.await(); t.join(); System.out.println(ran);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn cyclic_barrier_multiple_awaits_same_thread_on_single_party() {
    let out = run_main(
        "java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(1); barrier.await(); barrier.await(); barrier.await(); System.out.println(barrier.getParties());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn cyclic_barrier_worker_exception_during_await_may_break() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static boolean sawBroken = false;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); } catch (Exception e) { sawBroken = barrier.isBroken(); } }); Thread t2 = new Thread(() -> { try { Thread.sleep(5); barrier.await(); } catch (Exception e) {} }); t1.start(); barrier.reset(); t2.join(); t1.join(); System.out.println(barrier.isBroken());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_await_timed_with_seconds_on_complete_barrier() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static boolean ok = false;
    "#;
    let out = run_in_main(
        "Thread t = new Thread(() -> { try { barrier.await(1, java.util.concurrent.TimeUnit.SECONDS); } catch (Exception e) {} }); t.start(); ok = barrier.await(1, java.util.concurrent.TimeUnit.SECONDS) >= 0; t.join(); System.out.println(ok);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn cyclic_barrier_five_parties_single_round() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(5);
        static int arrived = 0;
    "#;
    let out = run_in_main(
        "for (int i = 0; i < 5; i++) { Thread t = new Thread(() -> { try { barrier.await(); arrived++; } catch (Exception e) {} }); t.start(); t.join(); } System.out.println(arrived);",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn cyclic_barrier_get_number_waiting_before_second_arrival() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(3);
        static int midWait = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t1.start(); Thread.sleep(5); midWait = barrier.getNumberWaiting(); Thread t2 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); Thread t3 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(midWait);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn cyclic_barrier_reset_after_trip_allows_new_cycle() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int cycles = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); cycles++; barrier.reset(); barrier.await(); cycles++; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); Thread.sleep(5); barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(cycles);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn cyclic_barrier_action_runnable_runs_on_barrier_thread() {
    let types = r#"
        static String marker = "";
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2, () -> { marker = "tripped"; });
    "#;
    let out = run_in_main(
        "Thread t = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t.start(); barrier.await(); t.join(); System.out.println(marker);",
        types,
    );
    assert_eq!(out, vec!["tripped"]);
}

#[test]
fn cyclic_barrier_two_threads_both_print_after_await() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { barrier.await(); System.out.println(\"a\"); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); System.out.println(\"b\"); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join();",
        types,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"a".to_string()));
    assert!(out.contains(&"b".to_string()));
}

#[test]
fn cyclic_barrier_parties_three_await_index_less_than_parties() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(3);
        static int maxIndex = -1;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { int i = barrier.await(); if (i > maxIndex) maxIndex = i; } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { int i = barrier.await(); if (i > maxIndex) maxIndex = i; } catch (Exception e) {} }); Thread t3 = new Thread(() -> { try { int i = barrier.await(); if (i > maxIndex) maxIndex = i; } catch (Exception e) {} }); t1.start(); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(maxIndex);",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn cyclic_barrier_is_not_broken_after_successful_await() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
    "#;
    let out = run_in_main(
        "Thread t = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t.start(); barrier.await(); t.join(); System.out.println(barrier.isBroken());",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn cyclic_barrier_phased_computation_with_two_barriers() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier phase1 = new java.util.concurrent.CyclicBarrier(2);
        static java.util.concurrent.CyclicBarrier phase2 = new java.util.concurrent.CyclicBarrier(2);
        static int value = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { value = 1; phase1.await(); value = value + 10; phase2.await(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { phase1.await(); value = value + 100; phase2.await(); System.out.println(value); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join();",
        types,
    );
    assert_eq!(out, vec!["111"]);
}

#[test]
fn cyclic_barrier_get_parties_unchanged_after_await() {
    let types = r#"
        static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
    "#;
    let out = run_in_main(
        "Thread t = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t.start(); barrier.await(); t.join(); System.out.println(barrier.getParties());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}
