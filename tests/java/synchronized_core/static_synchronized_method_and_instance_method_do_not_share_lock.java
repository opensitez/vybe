// vybe-test: java/synchronized_core/static_synchronized_method_and_instance_method_do_not_share_lock
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

static class Mixed {
            static int staticCount = 0;
            int instanceCount = 0;
            static synchronized void staticInc() { staticCount++; }
            synchronized void instanceInc() { instanceCount++; }
        }
    public static void main(String[] args) {
Mixed m = new Mixed(); Thread t1 = new Thread(() -> Mixed.staticInc()); Thread t2 = new Thread(() -> m.instanceInc()); t1.start(); t2.start(); t1.join(); t2.join(); __p(Mixed.staticCount); __p(m.instanceCount);
__check("1\n1");
    }
}

