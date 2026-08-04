// vybe-test: java/comparator_core/comparator_comparing_double_field_via_then_comparing
// origin: languages/java/tests/java/test_comparator_core.rs

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

static class Score { double value; Score(double value) { this.value = value; } }
    public static void main(String[] args) {
java.util.ArrayList<Score> list = new java.util.ArrayList<Score>(); list.add(new Score(2.5)); list.add(new Score(1.1)); list.add(new Score(3.0)); list.sort(java.util.Comparator.comparing((Score s) -> s.value)); __p(list.get(0).value); __p(list.get(2).value);
__check("1.1\n3.0");
    }
}

