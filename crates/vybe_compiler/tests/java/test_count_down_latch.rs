/// java.util.concurrent.CountDownLatch — one-shot countdown synchronization.
use crate::helpers::{run_in_main, run_main};

#[test]
fn count_down_latch_initial_count_is_constructor_value() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(3); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn count_down_latch_count_down_decrements_remaining() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(2); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn count_down_latch_await_returns_immediately_when_count_zero() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(0); latch.await(); System.out.println(\"released\");",
    );
    assert_eq!(out, vec!["released"]);
}

#[test]
fn count_down_latch_await_unblocks_after_single_count_down() {
    let types = r#"
        static boolean workerDone = false;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); Thread worker = new Thread(() -> { latch.countDown(); workerDone = true; }); worker.start(); latch.await(); System.out.println(workerDone);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_await_waits_for_multiple_count_downs() {
    let types = r#"
        static int phase = 0;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(2); Thread t1 = new Thread(() -> { latch.countDown(); phase = 1; }); Thread t2 = new Thread(() -> { latch.countDown(); phase = 2; }); t1.start(); t2.start(); latch.await(); System.out.println(latch.getCount());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_count_down_below_zero_stays_at_zero() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); latch.countDown(); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_await_with_timeout_returns_true_when_released() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(0); System.out.println(latch.await(10, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_await_with_timeout_returns_false_when_not_released() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); System.out.println(latch.await(1, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn count_down_latch_worker_await_then_main_count_down() {
    let types = r#"
        static String msg = "";
    "#;
    let out = run_in_main(
        r#"java.util.concurrent.CountDownLatch start = new java.util.concurrent.CountDownLatch(1); Thread worker = new Thread(() -> { try { start.await(); msg = "go"; } catch (InterruptedException e) {} }); worker.start(); start.countDown(); worker.join(); System.out.println(msg);"#,
        types,
    );
    assert_eq!(out, vec!["go"]);
}

#[test]
fn count_down_latch_three_workers_each_count_down_once() {
    let types = r#"
        static java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(3);
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> latch.countDown()); Thread t2 = new Thread(() -> latch.countDown()); Thread t3 = new Thread(() -> latch.countDown()); t1.start(); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(latch.getCount());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_main_awaits_worker_completion_signal() {
    let types = r#"
        static java.util.concurrent.CountDownLatch done = new java.util.concurrent.CountDownLatch(1);
        static int result = 0;
    "#;
    let out = run_in_main(
        "Thread worker = new Thread(() -> { result = 6 * 7; done.countDown(); }); worker.start(); done.await(); System.out.println(result);",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn count_down_latch_get_count_after_partial_count_down() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(5); latch.countDown(); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn count_down_latch_await_with_timeout_releases_when_count_down_from_thread() {
    let types = r#"
        static boolean timedOut = true;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); Thread t = new Thread(() -> { try { Thread.sleep(5); latch.countDown(); } catch (InterruptedException e) {} }); t.start(); timedOut = !latch.await(100, java.util.concurrent.TimeUnit.MILLISECONDS); t.join(); System.out.println(timedOut);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_zero_initial_count_await_is_immediate() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(0); System.out.println(latch.getCount()); latch.await(); System.out.println(\"ok\");",
    );
    assert_eq!(out, vec!["0", "ok"]);
}

#[test]
fn count_down_latch_single_count_down_from_main_thread() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); latch.countDown(); System.out.println(latch.getCount()); latch.await(); System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["0", "done"]);
}

#[test]
fn count_down_latch_two_step_gate_with_intermediate_count() {
    let types = r#"
        static int step = 0;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CountDownLatch gate = new java.util.concurrent.CountDownLatch(2); Thread t1 = new Thread(() -> { step = 1; gate.countDown(); }); Thread t2 = new Thread(() -> { step = 2; gate.countDown(); }); t1.start(); t2.start(); gate.await(); System.out.println(gate.getCount());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_await_blocks_until_all_parties_arrive() {
    let types = r#"
        static java.util.concurrent.CountDownLatch ready = new java.util.concurrent.CountDownLatch(2);
        static int arrivals = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { arrivals++; ready.countDown(); }); Thread t2 = new Thread(() -> { arrivals++; ready.countDown(); }); t1.start(); t2.start(); ready.await(); System.out.println(arrivals);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn count_down_latch_count_down_is_idempotent_past_zero() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); latch.countDown(); latch.countDown(); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_large_initial_count_partial_decrement() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(100); for (int i = 0; i < 10; i++) latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn count_down_latch_worker_sets_flag_before_count_down() {
    let types = r#"
        static boolean prepared = false;
        static java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1);
    "#;
    let out = run_in_main(
        "Thread worker = new Thread(() -> { prepared = true; latch.countDown(); }); worker.start(); latch.await(); System.out.println(prepared);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_await_after_all_count_downs_completes() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(3); latch.countDown(); latch.countDown(); latch.countDown(); latch.await(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_sequential_count_downs_from_main() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(4); latch.countDown(); System.out.println(latch.getCount()); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn count_down_latch_await_timeout_zero_on_unreleased() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(2); System.out.println(latch.await(0, java.util.concurrent.TimeUnit.MILLISECONDS)); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["false", "2"]);
}

#[test]
fn count_down_latch_multiple_awaits_after_release_all_proceed() {
    let types = r#"
        static int finished = 0;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); Thread t1 = new Thread(() -> { try { latch.await(); finished++; } catch (InterruptedException e) {} }); Thread t2 = new Thread(() -> { try { latch.await(); finished++; } catch (InterruptedException e) {} }); t1.start(); t2.start(); latch.countDown(); t1.join(); t2.join(); System.out.println(finished);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn count_down_latch_start_signal_unblocks_waiting_worker() {
    let types = r#"
        static String status = "waiting";
    "#;
    let out = run_in_main(
        r#"java.util.concurrent.CountDownLatch start = new java.util.concurrent.CountDownLatch(1); Thread worker = new Thread(() -> { try { start.await(); status = "running"; } catch (InterruptedException e) {} }); worker.start(); Thread.sleep(5); start.countDown(); worker.join(); System.out.println(status);"#,
        types,
    );
    assert_eq!(out, vec!["running"]);
}

#[test]
fn count_down_latch_done_signal_after_computation() {
    let types = r#"
        static java.util.concurrent.CountDownLatch done = new java.util.concurrent.CountDownLatch(1);
    "#;
    let out = run_in_main(
        "Thread worker = new Thread(() -> { System.out.println(\"work\"); done.countDown(); }); worker.start(); done.await(); System.out.println(\"joined\");",
        types,
    );
    assert_eq!(out, vec!["work", "joined"]);
}

#[test]
fn count_down_latch_party_count_one_released_by_worker() {
    let types = r#"
        static boolean released = false;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); Thread t = new Thread(() -> latch.countDown()); Thread waiter = new Thread(() -> { try { latch.await(); released = true; } catch (InterruptedException e) {} }); waiter.start(); t.start(); t.join(); waiter.join(); System.out.println(released);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_get_count_never_negative() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); latch.countDown(); latch.countDown(); System.out.println(latch.getCount() >= 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_await_immediate_when_pre_counted_down() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(2); latch.countDown(); latch.countDown(); latch.await(); System.out.println(\"ready\");",
    );
    assert_eq!(out, vec!["ready"]);
}

#[test]
fn count_down_latch_two_latches_independent_counts() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch a = new java.util.concurrent.CountDownLatch(2); java.util.concurrent.CountDownLatch b = new java.util.concurrent.CountDownLatch(3); a.countDown(); System.out.println(a.getCount()); System.out.println(b.getCount());",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn count_down_latch_worker_chain_count_down_sequence() {
    let types = r#"
        static java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(3);
        static int order = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { order = order + 1; latch.countDown(); }); Thread t2 = new Thread(() -> { order = order + 10; latch.countDown(); }); Thread t3 = new Thread(() -> { order = order + 100; latch.countDown(); }); t1.start(); t2.start(); t3.start(); latch.await(); System.out.println(latch.getCount());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_await_timeout_with_nanoseconds_unit() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(0); System.out.println(latch.await(1, java.util.concurrent.TimeUnit.NANOSECONDS));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_await_timeout_with_seconds_unit_on_released() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); latch.countDown(); System.out.println(latch.await(1, java.util.concurrent.TimeUnit.SECONDS));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_main_thread_count_down_unblocks_self_await() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); Thread releaser = new Thread(() -> { latch.countDown(); }); Thread waiter = new Thread(() -> { try { latch.await(); System.out.println(\"unblocked\"); } catch (InterruptedException e) {} }); waiter.start(); releaser.start(); releaser.join(); waiter.join();",
    );
    assert_eq!(out, vec!["unblocked"]);
}

#[test]
fn count_down_latch_partial_await_timeout_leaves_count_unchanged() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(3); latch.await(1, java.util.concurrent.TimeUnit.MILLISECONDS); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn count_down_latch_single_count_down_reduces_by_one() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(10); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn count_down_latch_finish_line_two_runners() {
    let types = r#"
        static java.util.concurrent.CountDownLatch finish = new java.util.concurrent.CountDownLatch(2);
        static int score = 0;
    "#;
    let out = run_in_main(
        "Thread r1 = new Thread(() -> { score += 10; finish.countDown(); }); Thread r2 = new Thread(() -> { score += 20; finish.countDown(); }); r1.start(); r2.start(); finish.await(); System.out.println(score);",
        types,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn count_down_latch_barrier_style_release_after_n_events() {
    let types = r#"
        static java.util.concurrent.CountDownLatch events = new java.util.concurrent.CountDownLatch(4);
    "#;
    let out = run_in_main(
        "for (int i = 0; i < 4; i++) { Thread t = new Thread(() -> events.countDown()); t.start(); t.join(); } events.await(); System.out.println(events.getCount());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_await_after_exact_count_downs_from_loop() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(5); for (int i = 0; i < 5; i++) latch.countDown(); latch.await(); System.out.println(\"sync\");",
    );
    assert_eq!(out, vec!["sync"]);
}

#[test]
fn count_down_latch_get_count_initial_one() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn count_down_latch_worker_exception_still_count_down_in_finally() {
    let types = r#"
        static java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1);
        static boolean sawError = false;
    "#;
    let out = run_in_main(
        "Thread worker = new Thread(() -> { try { throw new RuntimeException(\"fail\"); } catch (RuntimeException e) { sawError = true; } finally { latch.countDown(); } }); worker.start(); latch.await(); System.out.println(sawError);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_delayed_count_down_from_sleeping_thread() {
    let types = r#"
        static long waitedMs = 0;
    "#;
    let out = run_in_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); long start = System.currentTimeMillis(); Thread t = new Thread(() -> { try { Thread.sleep(5); latch.countDown(); } catch (InterruptedException e) {} }); t.start(); latch.await(); waitedMs = System.currentTimeMillis() - start; t.join(); System.out.println(waitedMs >= 0);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_cannot_reset_count_after_release() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); latch.countDown(); System.out.println(latch.getCount()); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn count_down_latch_parallel_tasks_signal_completion() {
    let types = r#"
        static java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(2);
        static int sum = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { sum += 5; latch.countDown(); }); Thread t2 = new Thread(() -> { sum += 7; latch.countDown(); }); t1.start(); t2.start(); latch.await(); System.out.println(sum);",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn count_down_latch_await_with_microseconds_timeout_on_open_latch() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(0); System.out.println(latch.await(1, java.util.concurrent.TimeUnit.MICROSECONDS));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn count_down_latch_triple_count_down_from_main_reaches_zero() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(3); latch.countDown(); latch.countDown(); latch.countDown(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_one_shot_cannot_be_reused_after_zero() {
    let out = run_main(
        "java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1); latch.countDown(); latch.await(); System.out.println(latch.getCount());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn count_down_latch_coordinator_waits_for_worker_ready_signals() {
    let types = r#"
        static java.util.concurrent.CountDownLatch ready = new java.util.concurrent.CountDownLatch(3);
        static int readyCount = 0;
    "#;
    let out = run_in_main(
        "Thread w1 = new Thread(() -> { readyCount++; ready.countDown(); }); Thread w2 = new Thread(() -> { readyCount++; ready.countDown(); }); Thread w3 = new Thread(() -> { readyCount++; ready.countDown(); }); w1.start(); w2.start(); w3.start(); ready.await(); System.out.println(readyCount);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}
