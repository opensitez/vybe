// vybe-test: java/synchronized_core/static_synchronized_method_high_contention_still_reaches_total
// origin: languages/java/tests/java/test_synchronized_core.rs

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

static class Hits {
            static int count = 0;
            static synchronized void ping() { count++; }
        }
    public static void main(String[] args) {
java.util.ArrayList<Thread> threads = new java.util.ArrayList<Thread>(); for (int i = 0; i < 8; i++) { threads.add(new Thread(() -> { for (int j = 0; j < 5; j++) Hits.ping(); })); } for (int i = 0; i < threads.size(); i++) threads.get(i).start(); for (int i = 0; i < threads.size(); i++) threads.get(i).join(); __p(Hits.count);
__check("40");
    }
}

