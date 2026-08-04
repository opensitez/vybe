// vybe-test: java/string_joiner/string_joiner_merge_chain_three_joiners
// origin: languages/java/tests/java/test_string_joiner.rs

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

    public static void main(String[] args) {
java.util.StringJoiner a = new java.util.StringJoiner(","); a.add("1"); java.util.StringJoiner b = new java.util.StringJoiner(","); b.add("2"); java.util.StringJoiner c = new java.util.StringJoiner(","); c.add("3"); a.merge(b); a.merge(c); __p(a.toString());
__check("1,2,3");
    }
}

