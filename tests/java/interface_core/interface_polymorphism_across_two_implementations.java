// vybe-test: java/interface_core/interface_polymorphism_across_two_implementations
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

interface Op { int run(int n); }
        static class AddOne implements Op { public int run(int n) { return n + 1; } }
        static class TimesTwo implements Op { public int run(int n) { return n * 2; } }
    public static void main(String[] args) {
Op a = new AddOne(); Op b = new TimesTwo(); __p(a.run(4)); __p(b.run(4));
__check("5\n8");
    }
}

