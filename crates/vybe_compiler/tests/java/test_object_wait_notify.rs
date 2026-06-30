use crate::helpers::{run_in_main, run_main};

#[test]
fn object_wait_notify_single_producer_consumer() {
    let types = r#"
        static class Box {
            int value = 0;
            synchronized void produce() { value = 42; notify(); }
            synchronized void consume() throws InterruptedException {
                while (value == 0) wait();
                System.out.println(value);
            }
        }
    "#;
    let out = run_in_main(
        "Box box = new Box(); Thread consumer = new Thread(() -> { try { box.consume(); } catch (InterruptedException e) {} }); consumer.start(); Thread.sleep(10); box.produce(); consumer.join();",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn object_notify_awakens_waiting_thread() {
    let types = r#"
        static class Gate {
            boolean open = false;
            synchronized void awaitOpen() throws InterruptedException {
                while (!open) wait();
                System.out.println("open");
            }
            synchronized void open() { open = true; notify(); }
        }
    "#;
    let out = run_in_main(
        "Gate g = new Gate(); Thread t = new Thread(() -> { try { g.awaitOpen(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); g.open(); t.join();",
        types,
    );
    assert_eq!(out, vec!["open"]);
}

#[test]
fn object_wait_notify_all_wakes_multiple_waiters() {
    let types = r#"
        static class Latch {
            int count = 0;
            synchronized void awaitReady() throws InterruptedException {
                while (count < 2) wait();
            }
            synchronized void signal() { count++; notifyAll(); }
        }
    "#;
    let out = run_in_main(
        "Latch l = new Latch(); Thread t1 = new Thread(() -> { try { l.awaitReady(); } catch (InterruptedException e) {} System.out.println(\"a\"); }); Thread t2 = new Thread(() -> { try { l.awaitReady(); } catch (InterruptedException e) {} System.out.println(\"b\"); }); t1.start(); t2.start(); Thread.sleep(10); l.signal(); l.signal(); t1.join(); t2.join();",
        types,
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"a".to_string()));
    assert!(out.contains(&"b".to_string()));
}

#[test]
fn object_wait_with_timeout_returns_when_notified() {
    let types = r#"
        static class Timed {
            boolean done = false;
            synchronized void waitBriefly() throws InterruptedException {
                wait(500);
                if (done) System.out.println("notified");
            }
            synchronized void complete() { done = true; notify(); }
        }
    "#;
    let out = run_in_main(
        "Timed t = new Timed(); Thread w = new Thread(() -> { try { t.waitBriefly(); } catch (InterruptedException e) {} }); w.start(); Thread.sleep(10); t.complete(); w.join();",
        types,
    );
    assert_eq!(out, vec!["notified"]);
}

#[test]
fn object_wait_notify_producer_consumer_queue() {
    let types = r#"
        static class Queue1 {
            int item = 0;
            boolean hasItem = false;
            synchronized void put(int v) { while (hasItem) { try { wait(); } catch (InterruptedException e) {} } item = v; hasItem = true; notify(); }
            synchronized int take() throws InterruptedException { while (!hasItem) wait(); hasItem = false; notify(); return item; }
        }
    "#;
    let out = run_in_main(
        "Queue1 q = new Queue1(); Thread consumer = new Thread(() -> { try { System.out.println(q.take()); } catch (InterruptedException e) {} }); consumer.start(); Thread.sleep(10); q.put(7); consumer.join();",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn object_notify_only_one_waiter_awakened() {
    let types = r#"
        static class Counter {
            int n = 0;
            synchronized void inc() throws InterruptedException {
                while (n == 0) wait();
                n--;
                System.out.println("go");
            }
            synchronized void signal() { n = 1; notify(); }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); Thread t = new Thread(() -> { try { c.inc(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); c.signal(); t.join();",
        types,
    );
    assert_eq!(out, vec!["go"]);
}

#[test]
fn object_wait_releases_lock_for_other_thread() {
    let types = r#"
        static class LockOrder {
            boolean flag = false;
            synchronized void setter() { flag = true; notify(); }
            synchronized void getter() throws InterruptedException {
                while (!flag) wait();
                System.out.println(flag);
            }
        }
    "#;
    let out = run_in_main(
        "LockOrder lo = new LockOrder(); Thread t = new Thread(() -> { try { lo.getter(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); lo.setter(); t.join();",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn object_wait_notify_handoff_between_threads() {
    let types = r#"
        static class Handoff {
            String msg = null;
            synchronized void send(String m) { msg = m; notify(); }
            synchronized String recv() throws InterruptedException { while (msg == null) wait(); return msg; }
        }
    "#;
    let out = run_in_main(
        r#"Handoff h = new Handoff(); Thread t = new Thread(() -> { try { System.out.println(h.recv()); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); h.send("ping"); t.join();"#,
        types,
    );
    assert_eq!(out, vec!["ping"]);
}

#[test]
fn object_wait_spurious_wakeup_guarded_by_while() {
    let types = r#"
        static class Guard {
            int ready = 0;
            synchronized void await() throws InterruptedException { while (ready == 0) wait(); System.out.println(ready); }
            synchronized void setReady(int v) { ready = v; notifyAll(); }
        }
    "#;
    let out = run_in_main(
        "Guard g = new Guard(); Thread t = new Thread(() -> { try { g.await(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); g.setReady(5); t.join();",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn object_notify_all_after_two_waits() {
    let types = r#"
        static class Release {
            boolean go = false;
            synchronized void pause() throws InterruptedException { while (!go) wait(); }
            synchronized void release() { go = true; notifyAll(); }
        }
    "#;
    let out = run_in_main(
        "Release r = new Release(); Thread t1 = new Thread(() -> { try { r.pause(); System.out.println(\"1\"); } catch (InterruptedException e) {} }); Thread t2 = new Thread(() -> { try { r.pause(); System.out.println(\"2\"); } catch (InterruptedException e) {} }); t1.start(); t2.start(); Thread.sleep(10); r.release(); t1.join(); t2.join();",
        types,
    );
    assert_eq!(out.len(), 2);
}

#[test]
fn object_wait_notify_inherited_monitor() {
    let types = r#"
        static class Parent {
            synchronized void wake() { notify(); }
        }
        static class Child extends Parent {
            synchronized void sleep() throws InterruptedException { wait(100); System.out.println("up"); }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); Thread t = new Thread(() -> { try { c.sleep(); } catch (InterruptedException e) { System.out.println(\"up\"); } }); t.start(); Thread.sleep(10); c.wake(); t.join();",
        types,
    );
    assert_eq!(out.len(), 1);
}

#[test]
fn object_wait_zero_timeout_returns_after_notify() {
    let types = r#"
        static class Zero {
            boolean set = false;
            synchronized void waitZero() throws InterruptedException { while (!set) wait(0); System.out.println("ok"); }
            synchronized void mark() { set = true; notify(); }
        }
    "#;
    let out = run_in_main(
        "Zero z = new Zero(); Thread t = new Thread(() -> { try { z.waitZero(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); z.mark(); t.join();",
        types,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn object_notify_before_wait_misses_signal() {
    let types = r#"
        static class Miss {
            boolean signaled = false;
            synchronized void signal() { signaled = true; notify(); }
            synchronized void check() throws InterruptedException {
                if (!signaled) wait(50);
                System.out.println(signaled);
            }
        }
    "#;
    let out = run_in_main(
        "Miss m = new Miss(); m.signal(); Thread t = new Thread(() -> { try { m.check(); } catch (InterruptedException e) {} }); t.start(); t.join();",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn object_wait_notify_two_stage_pipeline() {
    let types = r#"
        static class Stage {
            int stage = 0;
            synchronized void advance() throws InterruptedException { while (stage < 1) wait(); stage = 2; notify(); }
            synchronized void start() { stage = 1; notify(); }
            synchronized void finish() throws InterruptedException { while (stage < 2) wait(); System.out.println(stage); }
        }
    "#;
    let out = run_in_main(
        "Stage s = new Stage(); Thread mid = new Thread(() -> { try { s.advance(); } catch (InterruptedException e) {} }); Thread end = new Thread(() -> { try { s.finish(); } catch (InterruptedException e) {} }); mid.start(); end.start(); Thread.sleep(10); s.start(); mid.join(); end.join();",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn object_wait_notify_bounded_buffer_size_one() {
    let types = r#"
        static class Slot {
            Integer val = null;
            synchronized void put(int v) throws InterruptedException { while (val != null) wait(); val = v; notify(); }
            synchronized int get() throws InterruptedException { while (val == null) wait(); int r = val; val = null; notify(); return r; }
        }
    "#;
    let out = run_in_main(
        "Slot s = new Slot(); Thread c = new Thread(() -> { try { System.out.println(s.get()); } catch (InterruptedException e) {} }); c.start(); Thread.sleep(10); try { s.put(99); } catch (InterruptedException e) {} c.join();",
        types,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn object_notify_reentrant_lock_same_thread() {
    let types = r#"
        static class Reentrant {
            int depth = 0;
            synchronized void outer() throws InterruptedException {
                depth++;
                if (depth < 2) { wait(10); depth++; }
                notify();
                System.out.println(depth);
            }
        }
    "#;
    let out = run_in_main(
        "Reentrant r = new Reentrant(); Thread t = new Thread(() -> { try { r.outer(); } catch (InterruptedException e) {} }); t.start(); t.join();",
        types,
    );
    assert_eq!(out.len(), 1);
}

#[test]
fn object_wait_notify_flag_toggle() {
    let types = r#"
        static class Toggle {
            boolean on = false;
            synchronized void turnOn() { on = true; notifyAll(); }
            synchronized void waitOn() throws InterruptedException { while (!on) wait(); System.out.println(on); }
        }
    "#;
    let out = run_in_main(
        "Toggle t = new Toggle(); Thread w = new Thread(() -> { try { t.waitOn(); } catch (InterruptedException e) {} }); w.start(); Thread.sleep(10); t.turnOn(); w.join();",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn object_wait_long_timeout_with_early_notify() {
    let types = r#"
        static class Early {
            boolean hit = false;
            synchronized void waitLong() throws InterruptedException { while (!hit) wait(1000); System.out.println("early"); }
            synchronized void ping() { hit = true; notify(); }
        }
    "#;
    let out = run_in_main(
        "Early e = new Early(); Thread t = new Thread(() -> { try { e.waitLong(); } catch (InterruptedException e2) {} }); t.start(); Thread.sleep(10); e.ping(); t.join();",
        types,
    );
    assert_eq!(out, vec!["early"]);
}

#[test]
fn object_notify_all_clears_wait_set() {
    let types = r#"
        static class Broadcast {
            int gen = 0;
            synchronized void waitGen(int g) throws InterruptedException { while (gen < g) wait(); System.out.println(gen); }
            synchronized void bump() { gen++; notifyAll(); }
        }
    "#;
    let out = run_in_main(
        "Broadcast b = new Broadcast(); Thread t = new Thread(() -> { try { b.waitGen(1); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); b.bump(); t.join();",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn object_wait_notify_prints_sequence() {
    let types = r#"
        static class Seq {
            int step = 0;
            synchronized void waitStep(int s) throws InterruptedException { while (step < s) wait(); System.out.println(s); }
            synchronized void reach(int s) { if (step < s) { step = s; notifyAll(); } }
        }
    "#;
    let out = run_in_main(
        "Seq s = new Seq(); Thread t1 = new Thread(() -> { try { s.waitStep(1); } catch (InterruptedException e) {} }); Thread t2 = new Thread(() -> { try { s.waitStep(2); } catch (InterruptedException e) {} }); t1.start(); t2.start(); Thread.sleep(10); s.reach(1); Thread.sleep(10); s.reach(2); t1.join(); t2.join();",
        types,
    );
    assert_eq!(out.len(), 2);
}

#[test]
fn object_wait_on_private_lock_object() {
    let types = r#"
        static class PrivateLock {
            final Object lock = new Object();
            boolean ready = false;
            void await() throws InterruptedException {
                synchronized (lock) { while (!ready) lock.wait(); System.out.println("ready"); }
            }
            void setReady() { synchronized (lock) { ready = true; lock.notify(); } }
        }
    "#;
    let out = run_in_main(
        "PrivateLock p = new PrivateLock(); Thread t = new Thread(() -> { try { p.await(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); p.setReady(); t.join();",
        types,
    );
    assert_eq!(out, vec!["ready"]);
}

#[test]
fn object_notify_wakes_one_of_two_competing_waiters() {
    let types = r#"
        static class Compete {
            int tickets = 0;
            synchronized void take() throws InterruptedException { while (tickets == 0) wait(); tickets--; System.out.println("got"); }
            synchronized void give() { tickets++; notify(); }
        }
    "#;
    let out = run_in_main(
        "Compete c = new Compete(); Thread t = new Thread(() -> { try { c.take(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); c.give(); t.join();",
        types,
    );
    assert_eq!(out, vec!["got"]);
}

#[test]
fn object_wait_notify_integer_state_machine() {
    let types = r#"
        static class FSM {
            int state = 0;
            synchronized void toState(int s) throws InterruptedException { while (state < s - 1) wait(); state = s; notifyAll(); }
            synchronized void report(int s) throws InterruptedException { while (state < s) wait(); System.out.println(state); }
        }
    "#;
    let out = run_in_main(
        "FSM f = new FSM(); Thread t = new Thread(() -> { try { f.report(2); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); try { f.toState(1); f.toState(2); } catch (InterruptedException e) {} t.join();",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn object_wait_notify_three_thread_barrier() {
    let types = r#"
        static class Barrier3 {
            int arrived = 0;
            synchronized void arrive() throws InterruptedException { arrived++; if (arrived < 3) { while (arrived < 3) wait(); } else { notifyAll(); } }
        }
    "#;
    let out = run_in_main(
        "Barrier3 b = new Barrier3(); Thread t1 = new Thread(() -> { try { b.arrive(); System.out.println(\"x\"); } catch (InterruptedException e) {} }); Thread t2 = new Thread(() -> { try { b.arrive(); System.out.println(\"y\"); } catch (InterruptedException e) {} }); t1.start(); t2.start(); Thread.sleep(10); try { b.arrive(); System.out.println(\"z\"); } catch (InterruptedException e) {} t1.join(); t2.join();",
        types,
    );
    assert_eq!(out.len(), 3);
}

#[test]
fn object_wait_notify_string_payload_exchange() {
    let types = r#"
        static class Mailbox {
            String letter = null;
            synchronized void deposit(String s) { letter = s; notify(); }
            synchronized String pickup() throws InterruptedException { while (letter == null) wait(); String s = letter; letter = null; return s; }
        }
    "#;
    let out = run_in_main(
        r#"Mailbox m = new Mailbox(); Thread t = new Thread(() -> { try { System.out.println(m.pickup()); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); m.deposit("note"); t.join();"#,
        types,
    );
    assert_eq!(out, vec!["note"]);
}

#[test]
fn object_wait_notify_counter_increment_while_waiting() {
    let types = r#"
        static class Acc {
            int total = 0;
            synchronized void add(int n) { total += n; if (total >= 10) notify(); }
            synchronized void waitTen() throws InterruptedException { while (total < 10) wait(); System.out.println(total); }
        }
    "#;
    let out = run_in_main(
        "Acc a = new Acc(); Thread t = new Thread(() -> { try { a.waitTen(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); for (int i = 0; i < 5; i++) a.add(2); t.join();",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn object_notify_without_waiter_is_noop() {
    let types = r#"
        static class NoWaiter {
            synchronized void ping() { notify(); System.out.println("ping"); }
        }
    "#;
    let out = run_in_main("NoWaiter n = new NoWaiter(); n.ping();", types);
    assert_eq!(out, vec!["ping"]);
}

#[test]
fn object_wait_notify_all_resets_condition() {
    let types = r#"
        static class Reset {
            boolean active = true;
            synchronized void deactivate() { active = false; notifyAll(); }
            synchronized void waitInactive() throws InterruptedException { while (active) wait(); System.out.println("inactive"); }
        }
    "#;
    let out = run_in_main(
        "Reset r = new Reset(); Thread t = new Thread(() -> { try { r.waitInactive(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); r.deactivate(); t.join();",
        types,
    );
    assert_eq!(out, vec!["inactive"]);
}

#[test]
fn object_wait_notify_ping_pong_twice() {
    let types = r#"
        static class PingPong {
            int turn = 0;
            synchronized void waitTurn(int t) throws InterruptedException { while (turn != t) wait(); System.out.println(t); turn++; notifyAll(); }
        }
    "#;
    let out = run_in_main(
        "PingPong pp = new PingPong(); Thread t = new Thread(() -> { try { pp.waitTurn(1); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); try { pp.waitTurn(0); } catch (InterruptedException e) {} t.join();",
        types,
    );
    assert_eq!(out.len(), 2);
}

#[test]
fn object_wait_notify_even_odd_coordination() {
    let types = r#"
        static class EvenOdd {
            int n = 0;
            synchronized void printEven() throws InterruptedException { while (n % 2 != 0) wait(); System.out.println(n); n++; notify(); }
            synchronized void printOdd() throws InterruptedException { while (n % 2 == 0) wait(); System.out.println(n); n++; notify(); }
        }
    "#;
    let out = run_in_main(
        "EvenOdd eo = new EvenOdd(); Thread t = new Thread(() -> { try { eo.printOdd(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); try { eo.printEven(); } catch (InterruptedException e) {} t.join();",
        types,
    );
    assert_eq!(out.len(), 2);
}
