// vybe-test: java/constructors/overloaded_constructors_three_distinct_signatures
// origin: languages/java/tests/java/test_constructors.rs

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

static class Flex {
            int tag;
            Flex() { tag = 0; }
            Flex(int n) { tag = n; }
            Flex(int a, int b) { tag = a * 10 + b; }
        }
    public static void main(String[] args) {
Flex a = new Flex(); Flex b = new Flex(7); Flex c = new Flex(2, 3); __p(a.tag); __p(b.tag); __p(c.tag);
__check("0\n7\n23");
    }
}

