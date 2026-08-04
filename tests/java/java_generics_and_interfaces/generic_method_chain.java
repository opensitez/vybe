// vybe-test: java/java_generics_and_interfaces/generic_method_chain
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

static class Util { static <T> Holder<T> wrap(T x) { return new Holder<>(x); } static int wrapAndUnwrap(Integer x) { return wrap(x).v; } static class Holder<T> { T v; Holder(T v){this.v=v;} } }
    public static void main(String[] args) {
__p(Util.wrapAndUnwrap(2));
__check("2");
    }
}

