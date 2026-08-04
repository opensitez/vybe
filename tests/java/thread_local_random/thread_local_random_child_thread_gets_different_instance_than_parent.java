// vybe-test: java/thread_local_random/thread_local_random_child_thread_gets_different_instance_than_parent
// origin: languages/java/tests/java/test_thread_local_random.rs

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

static java.util.concurrent.ThreadLocalRandom parentRng;
        static java.util.concurrent.ThreadLocalRandom childRng;
        static boolean same;
    public static void main(String[] args) {
parentRng = java.util.concurrent.ThreadLocalRandom.current(); Thread t = new Thread(() -> { childRng = java.util.concurrent.ThreadLocalRandom.current(); same = parentRng == childRng; }); t.start(); t.join(); __p(same);
__check("false");
    }
}

