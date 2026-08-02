// Vybe test harness — Java.
//
// Real Java, exactly as test262's `assert.js` is real JavaScript: this file
// compiles with `javac` on its own, so it can be read, formatted and debugged
// with Java's own tools.
//
// Java has no top-level functions, so the harness must be a static member of
// the test's class rather than prepended to the file. The emitter splices the
// members between the outermost braces below into `Main`, which is also where
// `run_main`/`run_in_main` put the program.
//
// Output is COLLECTED, not paired. The emitter rewrites every
// `System.out.println(x)` into `__p(x)` and compares the whole output once, so
// a program whose print count is not static — a loop — still asserts. Pairing
// the i-th print with the i-th expected line left 659 of 7,395 Java cases
// without any assertion at all.
//
// It prints its own diagnostic BEFORE throwing, on the real `System.out` rather
// than into the buffer. That is not decoration: an uncaught error renders as
// `RuntimeError: [object]` under Vybe, so the expected and actual values would
// otherwise be lost entirely.
public class Check {
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
}
