// vybe-test: java/inheritance_core/sibling_instances_do_not_share_override_state
// origin: languages/java/tests/java/test_inheritance_core.rs

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

static class Counter { int n = 0; void inc() { n++; } }
        static class NamedCounter extends Counter { String name; NamedCounter(String name) { this.name = name; } }
    public static void main(String[] args) {
NamedCounter a = new NamedCounter("a"); NamedCounter b = new NamedCounter("b"); a.inc(); __p(a.n); __p(b.n);
__check("1\n0");
    }
}

