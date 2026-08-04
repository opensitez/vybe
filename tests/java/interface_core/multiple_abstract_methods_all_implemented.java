// vybe-test: java/interface_core/multiple_abstract_methods_all_implemented
// origin: languages/java/tests/java/test_interface_core.rs

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

interface PairOps {
            int left();
            int right();
            default int sum() { return left() + right(); }
        }
        static class TwoInts implements PairOps {
            int a; int b;
            TwoInts(int a, int b) { this.a = a; this.b = b; }
            public int left() { return a; }
            public int right() { return b; }
        }
    public static void main(String[] args) {
PairOps p = new TwoInts(3, 4); __p(p.sum());
__check("7");
    }
}

