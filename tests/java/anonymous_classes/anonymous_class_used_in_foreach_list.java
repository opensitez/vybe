// vybe-test: java/anonymous_classes/anonymous_class_used_in_foreach_list
// origin: languages/java/tests/java/test_anonymous_classes.rs

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

static interface Label { String text(); }
    public static void main(String[] args) {
java.util.ArrayList<Label> list = new java.util.ArrayList<Label>(); list.add(new Label() { public String text() { return "a"; } }); list.add(new Label() { public String text() { return "b"; } }); __p(list.get(0).text()); __p(list.get(1).text());
__check("a\nb");
    }
}

