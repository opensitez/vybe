// vybe-test: java/inner_classes/local_class_declared_inside_if_branch
// origin: languages/java/tests/java/test_inner_classes.rs

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

static class Util {
            int pick(boolean flag) {
                if (flag) {
                    class Local { int value = 1; }
                    return new Local().value;
                } else {
                    class Local { int value = 2; }
                    return new Local().value;
                }
            }
        }
    public static void main(String[] args) {
Util util = new Util(); __p(util.pick(true)); __p(util.pick(false));
__check("1\n2");
    }
}

