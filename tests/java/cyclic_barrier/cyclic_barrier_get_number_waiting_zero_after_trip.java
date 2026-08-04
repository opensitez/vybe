// vybe-test: java/cyclic_barrier/cyclic_barrier_get_number_waiting_zero_after_trip
// origin: languages/java/tests/java/test_cyclic_barrier.rs

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

static java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2);
        static int waitingAfter = -1;
    public static void main(String[] args) {
Thread t1 = new Thread(() -> { try { barrier.await(); waitingAfter = barrier.getNumberWaiting(); } catch (Exception e) {} }); Thread t2 = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t1.start(); t2.start(); t1.join(); t2.join(); __p(waitingAfter);
__check("0");
    }
}

