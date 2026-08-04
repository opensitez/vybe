// vybe-test: java/constructors/copy_constructor_pattern_via_this_fields
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

static class Copy {
            String a;
            String b;
            Copy(String a, String b) { this.a = a; this.b = b; }
            Copy(Copy other) { this(other.a, other.b); }
        }
    public static void main(String[] args) {
Copy src = new Copy("hi", "bye"); Copy dst = new Copy(src); __p(dst.a); __p(dst.b);
__check("hi\nbye");
    }
}

