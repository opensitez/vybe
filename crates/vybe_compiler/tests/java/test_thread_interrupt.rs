/// Thread interrupt flag semantics.
use crate::helpers::run_main;

#[test]
fn thread_interrupt_sets_interrupted_flag() {
    let out = run_main(
        "Thread t = Thread.currentThread(); t.interrupt(); System.out.println(t.isInterrupted()); t.interrupted();",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_interrupted_clears_flag_and_returns_previous_state() {
    let out = run_main(
        "Thread t = Thread.currentThread(); t.interrupt(); System.out.println(Thread.interrupted()); System.out.println(t.isInterrupted());",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn sleeping_thread_interrupt_surfaces_as_interrupted_exception() {
    let out = run_main(
        "Thread t = new Thread(() -> { try { Thread.sleep(1000); } catch (InterruptedException e) { System.out.println(\"caught\"); } }); t.start(); t.interrupt(); t.join();",
    );
    assert_eq!(out, vec!["caught"]);
}

#[test]
fn worker_checks_interrupted_status_in_loop() {
    let out = run_main(
        r#"Thread t = new Thread(() -> { int n = 0; while (!Thread.currentThread().isInterrupted() && n < 5) { n++; } System.out.println(n); }); t.start(); t.interrupt(); t.join();"#,
    );
    assert_eq!(out, vec!["1"]);
}
