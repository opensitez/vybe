// vybe-test: java/cyclic_barrier/cyclic_barrier_parties_two_constructor_with_runnable
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

static int ran = 0;
    public static void main(String[] args) {
java.util.concurrent.CyclicBarrier barrier = new java.util.concurrent.CyclicBarrier(2, () -> { ran++; }); Thread t = new Thread(() -> { try { barrier.await(); } catch (Exception e) {} }); t.start(); barrier.await(); t.join(); __p(ran);
__check("1");
    }
}

