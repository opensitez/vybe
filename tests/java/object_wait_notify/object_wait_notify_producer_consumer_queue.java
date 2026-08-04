// vybe-test: java/object_wait_notify/object_wait_notify_producer_consumer_queue
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {

    // A static String, NOT a StringBuilder. Calling a method on a bare static
    // FIELD receiver fails under Vybe with "undefined is not callable"
    // (measured): `SB.append(x)` throws while `StringBuilder l = SB;
    // l.append(x)` works, so the method is resolved from the receiver's
    // declared type at the call site and a static field carries none. String
    // concatenation onto a static field has no such problem.
    static String __buf = "";

    static void __p(Object o) {
        __buf = __buf + String.valueOf(o) + "\n";
    }

    static void __pr(Object o) {
        __buf = __buf + String.valueOf(o);
    }

    static void __check(String want) {
        String got = __buf;
        // The final `println` contributes a trailing newline that the expected
        // line vector never carried, so it is not part of the comparison.
        if (got.endsWith("\n")) {
            got = got.substring(0, got.length() - 1);
        }
        if (!got.equals(want)) {
            System.out.println("FAIL: want [" + want + "] got [" + got + "]");
            throw new RuntimeException("assertion failed");
        }
    }

static class Queue1 {
            int item = 0;
            boolean hasItem = false;
            synchronized void put(int v) { while (hasItem) { try { wait(); } catch (InterruptedException e) {} } item = v; hasItem = true; notify(); }
            synchronized int take() throws InterruptedException { while (!hasItem) wait(); hasItem = false; notify(); return item; }
        }
    public static void main(String[] args) {
Queue1 q = new Queue1(); Thread consumer = new Thread(() -> { try { __p(q.take()); } catch (InterruptedException e) {} }); consumer.start(); Thread.sleep(10); q.put(7); consumer.join();
__check("7");
    }
}

