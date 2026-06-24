use crate::helpers::{run_in_main, run_main};

#[test]
fn thread_subclass_run_invoked_directly_without_start() {
    let types = r#"
        static class Worker extends Thread {
            public void run() { System.out.println("direct"); }
        }
    "#;
    let out = run_in_main("Worker w = new Worker(); w.run();", types);
    assert_eq!(out, vec!["direct"]);
}

#[test]
fn thread_subclass_start_join_executes_run_body() {
    let types = r#"
        static class Worker extends Thread {
            public void run() { System.out.println("started"); }
        }
    "#;
    let out = run_in_main(
        "Worker w = new Worker(); w.start(); w.join(); System.out.println(\"joined\");",
        types,
    );
    assert_eq!(out, vec!["started", "joined"]);
}

#[test]
fn runnable_lambda_run_prints_message() {
    let out = run_main("Runnable r = () -> System.out.println(\"lambda\"); r.run();");
    assert_eq!(out, vec!["lambda"]);
}

#[test]
fn runnable_lambda_wrapped_in_thread_start_join() {
    let out = run_main(
        "Thread t = new Thread(() -> System.out.println(\"worker\")); t.start(); t.join(); System.out.println(\"main\");",
    );
    assert_eq!(out, vec!["worker", "main"]);
}

#[test]
fn anonymous_runnable_thread_start_join() {
    let out = run_main(
        "Thread t = new Thread(new Runnable() { public void run() { System.out.println(\"anon\"); } }); t.start(); t.join();",
    );
    assert_eq!(out, vec!["anon"]);
}

#[test]
fn thread_sleep_allows_main_to_continue_after_delay() {
    let out = run_main(
        "System.out.println(\"before\"); Thread.sleep(1); System.out.println(\"after\");",
    );
    assert_eq!(out, vec!["before", "after"]);
}

#[test]
fn thread_sleep_zero_completes_without_error() {
    let out = run_main("Thread.sleep(0); System.out.println(\"ok\");");
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn thread_sleep_inside_worker_does_not_abort_join() {
    let out = run_main(
        "Thread t = new Thread(() -> { Thread.sleep(1); System.out.println(\"slept\"); }); t.start(); t.join(); System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["slept", "done"]);
}

#[test]
fn current_thread_get_name_returns_nonempty_default() {
    let out = run_main(
        "String name = Thread.currentThread().getName(); System.out.println(name.length() > 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn current_thread_set_name_visible_via_get_name() {
    let out = run_main(
        "Thread t = Thread.currentThread(); t.setName(\"alpha\"); System.out.println(t.getName());",
    );
    assert_eq!(out, vec!["alpha"]);
}

#[test]
fn current_thread_set_name_twice_keeps_latest() {
    let out = run_main(
        "Thread t = Thread.currentThread(); t.setName(\"one\"); t.setName(\"two\"); System.out.println(t.getName());",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn worker_thread_set_name_before_start() {
    let out = run_main(
        "Thread t = new Thread(() -> System.out.println(Thread.currentThread().getName())); t.setName(\"worker\"); t.start(); t.join();",
    );
    assert_eq!(out, vec!["worker"]);
}

#[test]
fn thread_min_priority_constant_is_one() {
    let out = run_main("System.out.println(Thread.MIN_PRIORITY);");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn thread_max_priority_constant_is_ten() {
    let out = run_main("System.out.println(Thread.MAX_PRIORITY);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn thread_norm_priority_between_min_and_max() {
    let out = run_main(
        "System.out.println(Thread.NORM_PRIORITY >= Thread.MIN_PRIORITY); System.out.println(Thread.NORM_PRIORITY <= Thread.MAX_PRIORITY);",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn thread_is_alive_false_before_start() {
    let out = run_main(
        "Thread t = new Thread(() -> {}); System.out.println(t.isAlive());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn thread_is_alive_false_after_join() {
    let out = run_main(
        "Thread t = new Thread(() -> System.out.println(\"x\")); t.start(); t.join(); System.out.println(t.isAlive());",
    );
    assert_eq!(out, vec!["x", "false"]);
}

#[test]
fn thread_is_alive_true_while_worker_runs() {
    let types = r#"
        static int[] gate = {0};
        static class GateThread extends Thread {
            public void run() {
                gate[0] = 1;
                try { Thread.sleep(50); } catch (InterruptedException e) { }
                gate[0] = 2;
            }
        }
    "#;
    let out = run_in_main(
        "GateThread t = new GateThread(); t.start(); while (gate[0] == 0) { Thread.sleep(0); } System.out.println(t.isAlive()); t.join(); System.out.println(t.isAlive());",
        types,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn thread_constructor_accepts_runnable_argument() {
    let out = run_main(
        "Runnable r = () -> System.out.println(\"ctor\"); Thread t = new Thread(r); t.start(); t.join();",
    );
    assert_eq!(out, vec!["ctor"]);
}

#[test]
fn thread_subclass_constructor_passes_name_to_super() {
    let types = r#"
        static class NamedWorker extends Thread {
            NamedWorker(String name) { super(name); }
            public void run() { System.out.println(getName()); }
        }
    "#;
    let out = run_in_main(
        "NamedWorker w = new NamedWorker(\"named\"); w.start(); w.join();",
        types,
    );
    assert_eq!(out, vec!["named"]);
}

#[test]
fn two_threads_both_complete_before_main_continues() {
    let out = run_main(
        "Thread a = new Thread(() -> System.out.println(\"a\")); Thread b = new Thread(() -> System.out.println(\"b\")); a.start(); b.start(); a.join(); b.join(); System.out.println(\"end\");",
    );
    assert_eq!(out.len(), 3);
    assert!(out.contains(&"a".to_string()));
    assert!(out.contains(&"b".to_string()));
    assert_eq!(out.last().map(String::as_str), Some("end"));
}

#[test]
fn thread_join_waits_for_worker_output_order() {
    let out = run_main(
        "Thread t = new Thread(() -> { System.out.println(\"first\"); }); t.start(); t.join(); System.out.println(\"second\");",
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn thread_subclass_run_can_read_constructor_field() {
    let types = r#"
        static class EchoThread extends Thread {
            String msg;
            EchoThread(String msg) { this.msg = msg; }
            public void run() { System.out.println(msg); }
        }
    "#;
    let out = run_in_main(
        "EchoThread t = new EchoThread(\"payload\"); t.start(); t.join();",
        types,
    );
    assert_eq!(out, vec!["payload"]);
}

#[test]
fn runnable_captures_effectively_final_local() {
    let out = run_main(
        "int base = 4; Runnable r = () -> System.out.println(base + 1); Thread t = new Thread(r); t.start(); t.join();",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn thread_current_thread_same_before_and_after_set_name() {
    let out = run_main(
        "Thread a = Thread.currentThread(); Thread b = Thread.currentThread(); a.setName(\"main\"); System.out.println(a == b); System.out.println(b.getName());",
    );
    assert_eq!(out, vec!["true", "main"]);
}

#[test]
fn thread_priority_get_default_within_range() {
    let out = run_main(
        "Thread t = Thread.currentThread(); int p = t.getPriority(); System.out.println(p >= Thread.MIN_PRIORITY); System.out.println(p <= Thread.MAX_PRIORITY);",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn thread_set_priority_to_max_is_reflected() {
    let out = run_main(
        "Thread t = new Thread(() -> {}); t.setPriority(Thread.MAX_PRIORITY); System.out.println(t.getPriority());",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn thread_set_priority_to_min_is_reflected() {
    let out = run_main(
        "Thread t = new Thread(() -> {}); t.setPriority(Thread.MIN_PRIORITY); System.out.println(t.getPriority());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn thread_subclass_run_invoked_multiple_times_prints_each_time() {
    let types = r#"
        static class Repeat extends Thread {
            public void run() { System.out.println("tick"); }
        }
    "#;
    let out = run_in_main(
        "Repeat r = new Repeat(); r.run(); r.run(); System.out.println(\"done\");",
        types,
    );
    assert_eq!(out, vec!["tick", "tick", "done"]);
}

#[test]
fn thread_start_join_sequence_runs_exactly_once() {
    let types = r#"
        static int[] hits = {0};
        static class Once extends Thread {
            public void run() { hits[0]++; System.out.println(hits[0]); }
        }
    "#;
    let out = run_in_main(
        "Once t = new Once(); t.start(); t.join(); System.out.println(hits[0]);",
        types,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn thread_nested_start_join_inside_runnable() {
    let out = run_main(
        "Thread outer = new Thread(() -> { Thread inner = new Thread(() -> System.out.println(\"inner\")); inner.start(); try { inner.join(); } catch (InterruptedException e) { } }); outer.start(); outer.join(); System.out.println(\"outer\");",
    );
    assert_eq!(out, vec!["inner", "outer"]);
}

#[test]
fn thread_sleep_in_main_between_two_worker_starts() {
    let out = run_main(
        "Thread t1 = new Thread(() -> System.out.println(\"one\")); Thread t2 = new Thread(() -> System.out.println(\"two\")); t1.start(); Thread.sleep(1); t2.start(); t1.join(); t2.join();",
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"one".to_string()));
    assert!(out.contains(&"two".to_string()));
}

#[test]
fn runnable_anonymous_class_run_without_thread() {
    let out = run_main(
        "Runnable r = new Runnable() { public void run() { System.out.println(\"plain\"); } }; r.run();",
    );
    assert_eq!(out, vec!["plain"]);
}

#[test]
fn thread_subclass_overrides_run_not_start() {
    let types = r#"
        static class Safe extends Thread {
            public void run() { System.out.println("body"); }
        }
    "#;
    let out = run_in_main(
        "Safe s = new Safe(); s.start(); s.join();",
        types,
    );
    assert_eq!(out, vec!["body"]);
}

#[test]
fn thread_join_on_current_thread_is_noop() {
    let out = run_main(
        "Thread t = Thread.currentThread(); t.join(); System.out.println(\"alive\");",
    );
    assert_eq!(out, vec!["alive"]);
}

#[test]
fn thread_worker_can_call_current_thread_get_name() {
    let out = run_main(
        "Thread t = new Thread(() -> { Thread self = Thread.currentThread(); System.out.println(self.getName().length() > 0); }); t.setName(\"self\"); t.start(); t.join();",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_multiple_join_calls_are_safe() {
    let out = run_main(
        "Thread t = new Thread(() -> System.out.println(\"once\")); t.start(); t.join(); t.join(); System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["once", "done"]);
}

#[test]
fn thread_runnable_lambda_returns_after_join_before_next_print() {
    let out = run_main(
        "Thread t = new Thread(() -> System.out.println(\"mid\")); System.out.println(\"begin\"); t.start(); t.join(); System.out.println(\"finish\");",
    );
    assert_eq!(out, vec!["begin", "mid", "finish"]);
}

#[test]
fn thread_subclass_with_instance_field_mutation_in_run() {
    let types = r#"
        static class Acc extends Thread {
            int total = 0;
            public void run() { total = 7; }
        }
    "#;
    let out = run_in_main(
        "Acc a = new Acc(); a.start(); a.join(); System.out.println(a.total);",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn thread_sleep_does_not_skip_subsequent_print() {
    let out = run_main(
        "System.out.println(1); Thread.sleep(1); System.out.println(2); Thread.sleep(1); System.out.println(3);",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn thread_min_and_max_priority_constants_differ() {
    let out = run_main(
        "System.out.println(Thread.MIN_PRIORITY < Thread.MAX_PRIORITY);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_is_alive_false_for_fresh_runnable_thread() {
    let out = run_main(
        "Runnable r = () -> System.out.println(\"noop\"); Thread t = new Thread(r); System.out.println(t.isAlive());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn thread_start_then_join_with_empty_runnable_completes() {
    let out = run_main(
        "Thread t = new Thread(() -> {}); t.start(); t.join(); System.out.println(\"ok\");",
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn thread_subclass_run_prints_before_join_unblocks_main() {
    let types = r#"
        static class Ordered extends Thread {
            public void run() { System.out.println("child"); }
        }
    "#;
    let out = run_in_main(
        "System.out.println(\"parent\"); Ordered o = new Ordered(); o.start(); o.join();",
        types,
    );
    assert_eq!(out, vec!["parent", "child"]);
}

#[test]
fn thread_current_thread_set_name_does_not_affect_other_thread_name() {
    let out = run_main(
        "Thread main = Thread.currentThread(); main.setName(\"main\"); Thread t = new Thread(() -> System.out.println(Thread.currentThread().getName())); t.setName(\"other\"); t.start(); t.join(); System.out.println(main.getName());",
    );
    assert_eq!(out, vec!["other", "main"]);
}
