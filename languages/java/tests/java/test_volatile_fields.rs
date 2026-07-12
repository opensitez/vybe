/// volatile field visibility between threads.
use crate::helpers::run_in_main;

#[test]
fn volatile_flag_visible_to_reader_thread_after_writer_sets_it() {
    let out = run_in_main(
        "Flag f = new Flag(); Thread writer = new Thread(() -> { f.ready = true; }); Thread reader = new Thread(() -> { while (!f.ready) { } System.out.println(\"seen\"); }); writer.start(); reader.start(); writer.join(); reader.join();",
        r#"static class Flag { volatile boolean ready = false; }"#,
    );
    assert_eq!(out, vec!["seen"]);
}

#[test]
fn volatile_counter_updated_from_background_thread() {
    let out = run_in_main(
        "Counter c = new Counter(); Thread t = new Thread(() -> { c.value = 42; }); t.start(); t.join(); System.out.println(c.value);",
        r#"static class Counter { volatile int value = 0; }"#,
    );
    assert_eq!(out, vec!["42"]);
}
