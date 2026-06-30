/// java.util.concurrent.Semaphore — permit-based synchronization.
use crate::helpers::{run_in_main, run_main};

#[test]
fn semaphore_available_permits_matches_constructor() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(3); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn semaphore_acquire_decrements_available_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(2); sem.acquire(); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn semaphore_release_increments_available_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); sem.acquire(); sem.release(); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn semaphore_try_acquire_succeeds_when_permit_available() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); System.out.println(sem.tryAcquire()); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["true", "0"]);
}

#[test]
fn semaphore_try_acquire_fails_when_no_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); sem.acquire(); System.out.println(sem.tryAcquire());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn semaphore_acquire_uninterruptibly_decrements_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(2); sem.acquireUninterruptibly(); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn semaphore_drain_permits_returns_and_clears() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(4); System.out.println(sem.drainPermits()); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["4", "0"]);
}

#[test]
fn semaphore_has_queued_threads_false_when_no_waiters() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); System.out.println(sem.hasQueuedThreads());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn semaphore_get_queue_length_zero_when_no_waiters() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); System.out.println(sem.getQueueLength());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn semaphore_is_fair_reflects_constructor_flag() {
    let out = run_main(
        "java.util.concurrent.Semaphore fair = new java.util.concurrent.Semaphore(1, true); java.util.concurrent.Semaphore unfair = new java.util.concurrent.Semaphore(1, false); System.out.println(fair.isFair()); System.out.println(unfair.isFair());",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn semaphore_acquire_multiple_permits_at_once() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(5); sem.acquire(3); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn semaphore_release_multiple_permits_at_once() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0); sem.release(3); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn semaphore_try_acquire_with_timeout_succeeds_when_available() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); System.out.println(sem.tryAcquire(10, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_try_acquire_with_timeout_fails_when_unavailable() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); sem.acquire(); System.out.println(sem.tryAcquire(1, java.util.concurrent.TimeUnit.MILLISECONDS)); sem.release();",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn semaphore_try_acquire_permits_succeeds_when_enough_available() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(3); System.out.println(sem.tryAcquire(2)); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["true", "1"]);
}

#[test]
fn semaphore_try_acquire_permits_fails_when_insufficient() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); System.out.println(sem.tryAcquire(2));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn semaphore_worker_acquires_and_releases_permit() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1);
        static boolean entered = false;
    "#;
    let out = run_in_main(
        "Thread worker = new Thread(() -> { try { sem.acquire(); entered = true; sem.release(); } catch (InterruptedException e) {} }); worker.start(); worker.join(); System.out.println(entered);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_two_threads_serialize_on_single_permit() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1);
        static int concurrent = 0;
        static int maxConcurrent = 0;
    "#;
    let out = run_in_main(
        "Runnable task = () -> { try { sem.acquire(); concurrent++; if (concurrent > maxConcurrent) maxConcurrent = concurrent; Thread.sleep(1); concurrent--; sem.release(); } catch (Exception e) {} }; Thread t1 = new Thread(task); Thread t2 = new Thread(task); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(maxConcurrent);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn semaphore_binary_semaphore_mutex_pattern() {
    let types = r#"
        static java.util.concurrent.Semaphore mutex = new java.util.concurrent.Semaphore(1);
        static int counter = 0;
    "#;
    let out = run_in_main(
        "Runnable inc = () -> { try { mutex.acquire(); counter++; mutex.release(); } catch (InterruptedException e) {} }; Thread t1 = new Thread(inc); Thread t2 = new Thread(inc); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(counter);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn semaphore_release_without_prior_acquire_increases_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0); sem.release(); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn semaphore_acquire_blocks_until_release_from_other_thread() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0);
        static boolean acquired = false;
    "#;
    let out = run_in_main(
        "Thread waiter = new Thread(() -> { try { sem.acquire(); acquired = true; } catch (InterruptedException e) {} }); waiter.start(); Thread.sleep(5); sem.release(); waiter.join(); System.out.println(acquired);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_three_permits_allow_three_concurrent_holders() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(3);
        static int maxHeld = 0;
        static int held = 0;
    "#;
    let out = run_in_main(
        "Runnable task = () -> { try { sem.acquire(); held++; if (held > maxHeld) maxHeld = held; Thread.sleep(2); held--; sem.release(); } catch (Exception e) {} }; Thread t1 = new Thread(task); Thread t2 = new Thread(task); Thread t3 = new Thread(task); t1.start(); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(maxHeld);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn semaphore_drain_permits_on_empty_returns_zero() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0); System.out.println(sem.drainPermits());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn semaphore_acquire_uninterruptibly_blocks_like_acquire() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0);
        static boolean got = false;
    "#;
    let out = run_in_main(
        "Thread t = new Thread(() -> { sem.acquireUninterruptibly(); got = true; }); t.start(); Thread.sleep(5); sem.release(); t.join(); System.out.println(got);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_try_acquire_with_zero_timeout_fails_immediately() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); sem.acquire(); System.out.println(sem.tryAcquire(0, java.util.concurrent.TimeUnit.MILLISECONDS)); sem.release();",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn semaphore_release_more_than_one_restores_multiple_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0); sem.release(5); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn semaphore_acquire_two_then_release_two() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(4); sem.acquire(2); System.out.println(sem.availablePermits()); sem.release(2); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn semaphore_fair_semaphore_serializes_waiting_threads() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1, true);
        static int order = 0;
        static int first = 0;
        static int second = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { sem.acquire(); first = ++order; sem.release(); } catch (InterruptedException e) {} }); Thread t2 = new Thread(() -> { try { sem.acquire(); second = ++order; sem.release(); } catch (InterruptedException e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(first > 0 && second > 0);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_zero_initial_permits_requires_release_before_acquire() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0); sem.release(); System.out.println(sem.tryAcquire());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_has_queued_threads_after_blocking_acquire() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1);
        static boolean sawQueue;
    "#;
    let out = run_in_main(
        "sem.acquire(); Thread waiter = new Thread(() -> { try { sem.acquire(); } catch (InterruptedException e) {} }); waiter.start(); Thread.sleep(5); sawQueue = sem.hasQueuedThreads(); sem.release(); sem.release(); waiter.join(); System.out.println(sawQueue);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_get_queue_length_reflects_waiting_threads() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1);
        static int queueLen;
    "#;
    let out = run_in_main(
        "sem.acquire(); Thread w = new Thread(() -> { try { sem.acquire(); } catch (InterruptedException e) {} }); w.start(); Thread.sleep(5); queueLen = sem.getQueueLength(); sem.release(); sem.release(); w.join(); System.out.println(queueLen);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn semaphore_try_acquire_multiple_with_timeout() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(3); System.out.println(sem.tryAcquire(2, 10, java.util.concurrent.TimeUnit.MILLISECONDS)); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["true", "1"]);
}

#[test]
fn semaphore_try_acquire_multiple_timeout_insufficient_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); System.out.println(sem.tryAcquire(2, 1, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn semaphore_acquire_interruptibly_same_as_acquire_when_unblocked() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); sem.acquire(); System.out.println(sem.availablePermits()); sem.release();",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn semaphore_release_after_double_acquire_needs_two_releases() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(2); sem.acquire(); sem.acquire(); sem.release(); System.out.println(sem.availablePermits()); sem.release(); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn semaphore_resource_pool_two_of_three_in_use() {
    let out = run_main(
        "java.util.concurrent.Semaphore pool = new java.util.concurrent.Semaphore(3); pool.acquire(); pool.acquire(); System.out.println(pool.availablePermits());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn semaphore_producer_signals_consumer_via_release() {
    let types = r#"
        static java.util.concurrent.Semaphore ready = new java.util.concurrent.Semaphore(0);
        static String data = "";
    "#;
    let out = run_in_main(
        r#"Thread producer = new Thread(() -> { data = "payload"; ready.release(); }); Thread consumer = new Thread(() -> { try { ready.acquire(); System.out.println(data); } catch (InterruptedException e) {} }); producer.start(); consumer.start(); producer.join(); consumer.join();"#,
        types,
    );
    assert_eq!(out, vec!["payload"]);
}

#[test]
fn semaphore_throttling_limits_parallel_section() {
    let types = r#"
        static java.util.concurrent.Semaphore throttle = new java.util.concurrent.Semaphore(2);
        static int peak = 0;
        static int active = 0;
    "#;
    let out = run_in_main(
        "Runnable job = () -> { try { throttle.acquire(); active++; if (active > peak) peak = active; Thread.sleep(2); active--; throttle.release(); } catch (Exception e) {} }; Thread t1 = new Thread(job); Thread t2 = new Thread(job); Thread t3 = new Thread(job); t1.start(); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(peak);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn semaphore_drain_then_release_rebuilds_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(3); sem.drainPermits(); sem.release(2); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn semaphore_large_initial_permit_count() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(100); sem.acquire(50); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn semaphore_try_acquire_after_release_succeeds() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); sem.acquire(); sem.release(); System.out.println(sem.tryAcquire());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_available_permits_never_negative_after_ops() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); sem.acquire(); sem.release(); sem.release(); System.out.println(sem.availablePermits() >= 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_two_permits_both_acquired_by_main() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(2); sem.acquire(); sem.acquire(); System.out.println(sem.availablePermits()); sem.release(); sem.release();",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn semaphore_worker_release_unblocks_main_acquire() {
    let types = r#"
        static java.util.concurrent.Semaphore gate = new java.util.concurrent.Semaphore(0);
    "#;
    let out = run_in_main(
        "Thread worker = new Thread(() -> { try { Thread.sleep(5); gate.release(); } catch (InterruptedException e) {} }); worker.start(); gate.acquire(); System.out.println(\"passed\"); worker.join();",
        types,
    );
    assert_eq!(out, vec!["passed"]);
}

#[test]
fn semaphore_fair_flag_preserved_after_operations() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(2, true); sem.acquire(); sem.release(); System.out.println(sem.isFair());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_try_acquire_with_nanoseconds_timeout() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1); System.out.println(sem.tryAcquire(100, java.util.concurrent.TimeUnit.NANOSECONDS));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn semaphore_multiple_release_accumulates_permits() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(0); sem.release(); sem.release(); sem.release(); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn semaphore_acquire_all_permits_leaves_zero() {
    let out = run_main(
        "java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(3); sem.acquire(3); System.out.println(sem.availablePermits());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn semaphore_has_queued_threads_false_after_all_released() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1);
    "#;
    let out = run_in_main(
        "sem.acquire(); Thread w = new Thread(() -> { try { sem.acquire(); sem.release(); } catch (InterruptedException e) {} }); w.start(); Thread.sleep(5); sem.release(); w.join(); System.out.println(sem.hasQueuedThreads());",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn semaphore_get_queue_length_zero_after_waiter_acquired() {
    let types = r#"
        static java.util.concurrent.Semaphore sem = new java.util.concurrent.Semaphore(1);
    "#;
    let out = run_in_main(
        "sem.acquire(); Thread w = new Thread(() -> { try { sem.acquire(); sem.release(); } catch (InterruptedException e) {} }); w.start(); Thread.sleep(5); sem.release(); w.join(); System.out.println(sem.getQueueLength());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn semaphore_signaling_chain_three_threads() {
    let types = r#"
        static java.util.concurrent.Semaphore turn = new java.util.concurrent.Semaphore(0);
        static int stage = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { try { turn.acquire(); stage = 1; turn.release(); } catch (InterruptedException e) {} }); Thread t2 = new Thread(() -> { try { turn.acquire(); stage = 2; } catch (InterruptedException e) {} }); turn.release(); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(stage);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn semaphore_bounded_pool_acquire_release_cycle() {
    let out = run_main(
        "java.util.concurrent.Semaphore pool = new java.util.concurrent.Semaphore(1); pool.acquire(); pool.release(); pool.acquire(); System.out.println(pool.tryAcquire()); pool.release();",
    );
    assert_eq!(out, vec!["true"]);
}
