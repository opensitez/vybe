use crate::helpers::{run_in_main, run_main};

#[test]
fn synchronized_method_serializes_two_thread_increments() {
    let types = r#"
        static class Counter {
            int value = 0;
            synchronized void inc() { value++; }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); Thread t1 = new Thread(() -> { for (int i = 0; i < 50; i++) c.inc(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 50; i++) c.inc(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn synchronized_block_on_this_serializes_updates() {
    let types = r#"
        static class Counter {
            int value = 0;
            void inc() {
                synchronized (this) { value++; }
            }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); Thread t1 = new Thread(() -> { for (int i = 0; i < 40; i++) c.inc(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 60; i++) c.inc(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn synchronized_block_on_shared_lock_object_serializes_updates() {
    let types = r#"
        static class Counter {
            final Object lock = new Object();
            int value = 0;
            void inc() {
                synchronized (lock) { value++; }
            }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); Thread t1 = new Thread(() -> { for (int i = 0; i < 25; i++) c.inc(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 25; i++) c.inc(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn static_synchronized_method_serializes_class_level_counter() {
    let types = r#"
        static class Counter {
            static int value = 0;
            static synchronized void inc() { value++; }
        }
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { for (int i = 0; i < 30; i++) Counter.inc(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 70; i++) Counter.inc(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(Counter.value);",
        types,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn synchronized_method_allows_single_thread_full_count() {
    let types = r#"
        static class Counter {
            int value = 0;
            synchronized void inc() { value++; }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); for (int i = 0; i < 5; i++) c.inc(); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn synchronized_block_reads_consistent_value() {
    let types = r#"
        static class Holder {
            int value = 0;
            synchronized int read() { return value; }
            synchronized void write(int v) { value = v; }
        }
    "#;
    let out = run_in_main(
        "Holder h = new Holder(); h.write(9); System.out.println(h.read());",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn synchronized_method_on_two_instances_does_not_block_each_other() {
    let types = r#"
        static class Counter {
            int value = 0;
            synchronized void inc() { value++; }
        }
    "#;
    let out = run_in_main(
        "Counter a = new Counter(); Counter b = new Counter(); Thread t1 = new Thread(() -> { a.inc(); a.inc(); }); Thread t2 = new Thread(() -> { b.inc(); b.inc(); b.inc(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(a.value); System.out.println(b.value);",
        types,
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn synchronized_block_nested_on_same_lock_is_reentrant() {
    let types = r#"
        static class Reentrant {
            int depth = 0;
            void enter() {
                synchronized (this) {
                    depth++;
                    synchronized (this) { depth++; }
                }
            }
        }
    "#;
    let out = run_in_main("Reentrant r = new Reentrant(); r.enter(); System.out.println(r.depth);", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_method_reentrant_from_same_thread() {
    let types = r#"
        static class Reentrant {
            int hits = 0;
            synchronized void outer() {
                hits++;
                inner();
            }
            synchronized void inner() { hits++; }
        }
    "#;
    let out = run_in_main("Reentrant r = new Reentrant(); r.outer(); System.out.println(r.hits);", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn static_synchronized_block_on_class_literal() {
    let types = r#"
        static class Counter {
            static int value = 0;
            static void inc() {
                synchronized (Counter.class) { value++; }
            }
        }
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { for (int i = 0; i < 10; i++) Counter.inc(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 15; i++) Counter.inc(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(Counter.value);",
        types,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn synchronized_method_preserves_order_of_single_thread_calls() {
    let types = r#"
        static class Seq {
            int last = 0;
            synchronized void step(int n) { last = n; }
        }
    "#;
    let out = run_in_main(
        "Seq s = new Seq(); s.step(1); s.step(2); s.step(3); System.out.println(s.last);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn synchronized_block_protects_list_append_size() {
    let types = r#"
        static class Bucket {
            java.util.ArrayList<Integer> items = new java.util.ArrayList<Integer>();
            synchronized void add(int v) { items.add(v); }
            synchronized int size() { return items.size(); }
        }
    "#;
    let out = run_in_main(
        "Bucket b = new Bucket(); Thread t1 = new Thread(() -> { for (int i = 0; i < 3; i++) b.add(i); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 2; i++) b.add(i); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(b.size());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn synchronized_method_two_instances_independent_locks() {
    let types = r#"
        static class Box {
            int value = 0;
            synchronized void set(int v) { value = v; }
            synchronized int get() { return value; }
        }
    "#;
    let out = run_in_main(
        "Box a = new Box(); Box b = new Box(); a.set(1); b.set(2); System.out.println(a.get()); System.out.println(b.get());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn synchronized_block_on_explicit_lock_field() {
    let types = r#"
        static class Gate {
            final Object mutex = new Object();
            boolean open = false;
            void openGate() {
                synchronized (mutex) { open = true; }
            }
            boolean isOpen() {
                synchronized (mutex) { return open; }
            }
        }
    "#;
    let out = run_in_main("Gate g = new Gate(); g.openGate(); System.out.println(g.isOpen());", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn static_synchronized_method_resets_via_synchronized_setter() {
    let types = r#"
        static class Total {
            static int n = 0;
            static synchronized void add(int v) { n += v; }
            static synchronized void clear() { n = 0; }
        }
    "#;
    let out = run_in_main(
        "Total.add(5); Total.clear(); System.out.println(Total.n);",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn synchronized_method_runs_without_contention_single_threaded() {
    let types = r#"
        static class Safe {
            synchronized String label() { return "ok"; }
        }
    "#;
    let out = run_in_main("Safe s = new Safe(); System.out.println(s.label());", types);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn synchronized_block_around_string_builder_append() {
    let types = r#"
        static class Log {
            StringBuilder sb = new StringBuilder();
            synchronized void append(String part) { sb.append(part); }
            synchronized String text() { return sb.toString(); }
        }
    "#;
    let out = run_in_main(
        "Log log = new Log(); Thread t1 = new Thread(() -> log.append(\"a\")); Thread t2 = new Thread(() -> log.append(\"b\")); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(log.text().length());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_method_three_threads_sum_counter() {
    let types = r#"
        static class Counter {
            int value = 0;
            synchronized void inc() { value++; }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); Thread t1 = new Thread(() -> { for (int i = 0; i < 10; i++) c.inc(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 10; i++) c.inc(); }); Thread t3 = new Thread(() -> { for (int i = 0; i < 10; i++) c.inc(); }); t1.start(); t2.start(); t3.start(); t1.join(); t2.join(); t3.join(); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn synchronized_block_same_object_used_by_multiple_methods() {
    let types = r#"
        static class Account {
            final Object lock = new Object();
            int balance = 0;
            void deposit(int amount) {
                synchronized (lock) { balance += amount; }
            }
            int snapshot() {
                synchronized (lock) { return balance; }
            }
        }
    "#;
    let out = run_in_main(
        "Account a = new Account(); Thread t1 = new Thread(() -> a.deposit(5)); Thread t2 = new Thread(() -> a.deposit(7)); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(a.snapshot());",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn static_synchronized_method_visible_across_instances() {
    let types = r#"
        static class Shared {
            static int hits = 0;
            static synchronized void hit() { hits++; }
        }
    "#;
    let out = run_in_main(
        "Shared a = new Shared(); Shared b = new Shared(); Thread t1 = new Thread(() -> a.hit()); Thread t2 = new Thread(() -> b.hit()); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(Shared.hits);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_method_returns_value_atomically() {
    let types = r#"
        static class IdGen {
            int next = 1;
            synchronized int take() { return next++; }
        }
    "#;
    let out = run_in_main(
        "IdGen g = new IdGen(); Thread t1 = new Thread(() -> System.out.println(g.take())); Thread t2 = new Thread(() -> System.out.println(g.take())); t1.start(); t2.start(); t1.join(); t2.join();",
        types,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"1".to_string()));
    assert!(out.contains(&"2".to_string()));
}

#[test]
fn synchronized_block_prevents_lost_updates_on_int_field() {
    let types = r#"
        static class Tallier {
            int sum = 0;
            void add(int v) {
                synchronized (this) { sum += v; }
            }
        }
    "#;
    let out = run_in_main(
        "Tallier t = new Tallier(); Thread a = new Thread(() -> t.add(10)); Thread b = new Thread(() -> t.add(20)); a.start(); b.start(); a.join(); b.join(); System.out.println(t.sum);",
        types,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn synchronized_method_on_subclass_inherits_lock() {
    let types = r#"
        static class Base {
            int n = 0;
            synchronized void bump() { n++; }
        }
        static class Derived extends Base { }
    "#;
    let out = run_in_main(
        "Derived d = new Derived(); Thread t1 = new Thread(() -> { for (int i = 0; i < 4; i++) d.bump(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 6; i++) d.bump(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(d.n);",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn synchronized_block_with_local_lock_variable_per_instance() {
    let types = r#"
        static class Cell {
            final Object gate = new Object();
            String value = "";
            void set(String v) {
                synchronized (gate) { value = v; }
            }
            String get() {
                synchronized (gate) { return value; }
            }
        }
    "#;
    let out = run_in_main(
        "Cell c = new Cell(); c.set(\"vybe\"); System.out.println(c.get());",
        types,
    );
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn static_synchronized_method_and_instance_method_do_not_share_lock() {
    let types = r#"
        static class Mixed {
            static int staticCount = 0;
            int instanceCount = 0;
            static synchronized void staticInc() { staticCount++; }
            synchronized void instanceInc() { instanceCount++; }
        }
    "#;
    let out = run_in_main(
        "Mixed m = new Mixed(); Thread t1 = new Thread(() -> Mixed.staticInc()); Thread t2 = new Thread(() -> m.instanceInc()); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(Mixed.staticCount); System.out.println(m.instanceCount);",
        types,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn synchronized_method_called_from_run_method_in_thread() {
    let types = r#"
        static class Counter {
            int value = 0;
            synchronized void inc() { value++; }
        }
        static class Worker extends Thread {
            Counter counter;
            Worker(Counter counter) { this.counter = counter; }
            public void run() { counter.inc(); }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); Worker w = new Worker(c); w.start(); w.join(); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_block_allows_read_after_write_in_same_thread() {
    let types = r#"
        static class Flag {
            boolean ready = false;
            void publish() {
                synchronized (this) { ready = true; }
            }
            boolean isReady() {
                synchronized (this) { return ready; }
            }
        }
    "#;
    let out = run_in_main(
        "Flag f = new Flag(); f.publish(); System.out.println(f.isReady());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_method_serializes_negative_and_positive_increments() {
    let types = r#"
        static class Balance {
            int amount = 100;
            synchronized void adjust(int delta) { amount += delta; }
        }
    "#;
    let out = run_in_main(
        "Balance b = new Balance(); Thread t1 = new Thread(() -> b.adjust(-30)); Thread t2 = new Thread(() -> b.adjust(20)); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(b.amount);",
        types,
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn synchronized_block_on_shared_array_length_guard() {
    let types = r#"
        static class Buffer {
            int[] data = new int[3];
            int size = 0;
            synchronized void push(int v) {
                if (size < data.length) { data[size++] = v; }
            }
            synchronized int length() { return size; }
        }
    "#;
    let out = run_in_main(
        "Buffer b = new Buffer(); Thread t1 = new Thread(() -> { b.push(1); b.push(2); }); Thread t2 = new Thread(() -> { b.push(3); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(b.length());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn static_synchronized_method_blocks_second_thread_until_first_finishes() {
    let types = r#"
        static class Serial {
            static int last = 0;
            static synchronized void mark(int n) {
                last = n;
                try { Thread.sleep(1); } catch (InterruptedException e) { }
            }
        }
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> Serial.mark(1)); Thread t2 = new Thread(() -> Serial.mark(2)); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(Serial.last);",
        types,
    );
    assert_eq!(out.len(), 1);
    assert!(out[0] == "1" || out[0] == "2");
}

#[test]
fn synchronized_method_empty_body_still_acquires_lock() {
    let types = r#"
        static class Touch {
            synchronized void touch() { }
        }
    "#;
    let out = run_in_main("Touch t = new Touch(); t.touch(); System.out.println(\"ok\");", types);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn synchronized_block_two_nested_scopes_same_monitor() {
    let types = r#"
        static class Track {
            int a = 0;
            int b = 0;
            void run() {
                synchronized (this) { a = 1; }
                synchronized (this) { b = 2; }
            }
        }
    "#;
    let out = run_in_main("Track t = new Track(); t.run(); System.out.println(t.a + t.b);", types);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn synchronized_method_on_private_method_via_public_wrapper() {
    let types = r#"
        static class Vault {
            int secret = 0;
            synchronized void store(int v) { secret = v; }
            synchronized int load() { return secret; }
        }
    "#;
    let out = run_in_main(
        "Vault v = new Vault(); v.store(42); System.out.println(v.load());",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn synchronized_block_lock_field_shared_between_two_instances_fails_without_sharing() {
    let types = r#"
        static class Node {
            static final Object shared = new Object();
            int value = 0;
            void set(int v) {
                synchronized (shared) { value = v; }
            }
            int get() {
                synchronized (shared) { return value; }
            }
        }
    "#;
    let out = run_in_main(
        "Node a = new Node(); Node b = new Node(); Thread t1 = new Thread(() -> a.set(4)); Thread t2 = new Thread(() -> b.set(9)); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(a.get() + b.get());",
        types,
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn static_synchronized_method_reset_then_increment() {
    let types = r#"
        static class Score {
            static int points = 0;
            static synchronized void reset() { points = 0; }
            static synchronized void add(int v) { points += v; }
        }
    "#;
    let out = run_in_main(
        "Score.add(3); Score.reset(); Score.add(2); System.out.println(Score.points);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_method_interleaved_reads_remain_consistent() {
    let types = r#"
        static class Pair {
            int x = 0;
            int y = 0;
            synchronized void set(int a, int b) { x = a; y = b; }
            synchronized int sum() { return x + y; }
        }
    "#;
    let out = run_in_main(
        "Pair p = new Pair(); Thread writer = new Thread(() -> p.set(2, 3)); Thread reader = new Thread(() -> { try { Thread.sleep(1); } catch (InterruptedException e) { } }); writer.start(); reader.start(); writer.join(); reader.join(); System.out.println(p.sum());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn synchronized_block_on_this_matches_synchronized_method_lock() {
    let types = r#"
        static class Dual {
            int value = 0;
            synchronized void incMethod() { value++; }
            void incBlock() {
                synchronized (this) { value++; }
            }
        }
    "#;
    let out = run_in_main(
        "Dual d = new Dual(); Thread t1 = new Thread(() -> { for (int i = 0; i < 5; i++) d.incMethod(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 5; i++) d.incBlock(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(d.value);",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn synchronized_method_called_recursively_from_same_thread() {
    let types = r#"
        static class Factorial {
            synchronized int fact(int n) {
                if (n <= 1) return 1;
                return n * fact(n - 1);
            }
        }
    "#;
    let out = run_in_main("Factorial f = new Factorial(); System.out.println(f.fact(4));", types);
    assert_eq!(out, vec!["24"]);
}

#[test]
fn synchronized_block_guarding_boolean_toggle() {
    let types = r#"
        static class Toggle {
            boolean on = false;
            synchronized void flip() { on = !on; }
            synchronized boolean state() { return on; }
        }
    "#;
    let out = run_in_main("Toggle t = new Toggle(); t.flip(); t.flip(); System.out.println(t.state());", types);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn static_synchronized_method_high_contention_still_reaches_total() {
    let types = r#"
        static class Hits {
            static int count = 0;
            static synchronized void ping() { count++; }
        }
    "#;
    let out = run_in_main(
        "java.util.ArrayList<Thread> threads = new java.util.ArrayList<Thread>(); for (int i = 0; i < 8; i++) { threads.add(new Thread(() -> { for (int j = 0; j < 5; j++) Hits.ping(); })); } for (int i = 0; i < threads.size(); i++) threads.get(i).start(); for (int i = 0; i < threads.size(); i++) threads.get(i).join(); System.out.println(Hits.count);",
        types,
    );
    assert_eq!(out, vec!["40"]);
}

#[test]
fn synchronized_method_after_join_sees_final_value() {
    let types = r#"
        static class Done {
            int value = 0;
            synchronized void set(int v) { value = v; }
            synchronized int get() { return value; }
        }
    "#;
    let out = run_in_main(
        "Done d = new Done(); Thread t = new Thread(() -> d.set(11)); t.start(); t.join(); System.out.println(d.get());",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn synchronized_block_on_distinct_lock_objects_allow_parallel_updates() {
    let types = r#"
        static class Split {
            final Object left = new Object();
            final Object right = new Object();
            int a = 0;
            int b = 0;
            void incA() { synchronized (left) { a++; } }
            void incB() { synchronized (right) { b++; } }
        }
    "#;
    let out = run_in_main(
        "Split s = new Split(); Thread t1 = new Thread(() -> { for (int i = 0; i < 3; i++) s.incA(); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 4; i++) s.incB(); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(s.a); System.out.println(s.b);",
        types,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn synchronized_method_zero_initial_increment_to_one() {
    let types = r#"
        static class Once {
            int n = 0;
            synchronized void once() { if (n == 0) n = 1; }
        }
    "#;
    let out = run_in_main(
        "Once o = new Once(); Thread t1 = new Thread(() -> o.once()); Thread t2 = new Thread(() -> o.once()); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(o.n);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_block_string_concat_under_lock() {
    let types = r#"
        static class Joiner {
            String parts = "";
            synchronized void add(String s) { parts = parts + s; }
            synchronized String all() { return parts; }
        }
    "#;
    let out = run_in_main(
        "Joiner j = new Joiner(); Thread t1 = new Thread(() -> j.add(\"x\")); Thread t2 = new Thread(() -> j.add(\"y\")); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(j.all().length());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn static_synchronized_method_called_from_multiple_class_references() {
    let types = r#"
        static class Ref {
            static int n = 0;
            static synchronized void add() { n++; }
        }
    "#;
    let out = run_in_main(
        "Ref r1 = new Ref(); Ref r2 = new Ref(); Thread t1 = new Thread(() -> r1.add()); Thread t2 = new Thread(() -> r2.add()); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(Ref.n);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}
