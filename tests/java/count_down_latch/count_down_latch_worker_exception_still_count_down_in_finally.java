// vybe-test: java/count_down_latch/count_down_latch_worker_exception_still_count_down_in_finally
// origin: languages/java/tests/java/test_count_down_latch.rs

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

static java.util.concurrent.CountDownLatch latch = new java.util.concurrent.CountDownLatch(1);
        static boolean sawError = false;
    public static void main(String[] args) {
Thread worker = new Thread(() -> { try { throw new RuntimeException("fail"); } catch (RuntimeException e) { sawError = true; } finally { latch.countDown(); } }); worker.start(); latch.await(); __p(sawError);
__check("true");
    }
}

