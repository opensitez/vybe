// vybe-test: java/abstract_classes/abstract_expression_eval_add
// origin: languages/java/tests/java/test_abstract_classes.rs

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

static abstract class Expr { abstract int eval(); }
        static class Add extends Expr {
            Expr left;
            Expr right;
            Add(Expr left, Expr right) { this.left = left; this.right = right; }
            int eval() { return left.eval() + right.eval(); }
        }
        static class Num extends Expr {
            int value;
            Num(int value) { this.value = value; }
            int eval() { return value; }
        }
    public static void main(String[] args) {
Expr e = new Add(new Num(2), new Num(3)); __p(e.eval());
__check("5");
    }
}

