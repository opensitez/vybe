// vybe-test: java/java_generics_and_interfaces/generic_bound_compare
// origin: languages/java/tests/java/test_java_generics_and_interfaces.rs

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

static class Num implements Comparable<Num> { int v; Num(int v) { this.v = v; } public int compareTo(Num o) { return v - o.v; } } static class Compares { static <T extends Comparable<T>> int less(T a, T b) { return a.compareTo(b) < 0 ? 1 : 0; } }
    public static void main(String[] args) {
__p(Compares.less(new Num(3), new Num(5)));
__check("1");
    }
}

