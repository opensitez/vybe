// vybe-test: java/synchronized_core/synchronized_method_interleaved_reads_remain_consistent
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

static class Pair {
            int x = 0;
            int y = 0;
            synchronized void set(int a, int b) { x = a; y = b; }
            synchronized int sum() { return x + y; }
        }
    public static void main(String[] args) {
Pair p = new Pair(); Thread writer = new Thread(() -> p.set(2, 3)); Thread reader = new Thread(() -> { try { Thread.sleep(1); } catch (InterruptedException e) { } }); writer.start(); reader.start(); writer.join(); reader.join(); __p(p.sum());
__check("5");
    }
}

