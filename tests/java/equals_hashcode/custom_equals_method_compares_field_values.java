// vybe-test: java/equals_hashcode/custom_equals_method_compares_field_values
// origin: languages/java/tests/java/test_equals_hashcode.rs

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
            int x;
            int y;
            Pair(int x, int y) { this.x = x; this.y = y; }
            boolean equals(Pair other) {
                return other != null && other.x == x && other.y == y;
            }
        }
    public static void main(String[] args) {
Pair a = new Pair(1, 2); Pair b = new Pair(1, 2); Pair c = new Pair(1, 3); __p(a.equals(b)); __p(a.equals(c));
__check("true\nfalse");
    }
}

