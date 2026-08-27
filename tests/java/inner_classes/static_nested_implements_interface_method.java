import java.util.*;
import java.util.stream.*;
import java.util.function.*;
import java.util.concurrent.*;
import java.time.*;
import java.time.format.*;
import java.net.*;
import java.io.*;
import java.nio.file.*;
import java.lang.reflect.*;

// vybe-test: java/inner_classes/static_nested_implements_interface_method
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

static interface Named { String name(); }
        static class Outer {
            static class Inner implements Named { public String name() { return "inner"; } }
        }
    public static void main(String[] args) throws Throwable {
Named n = new Outer.Inner(); __p(n.name());
__check("inner");
    }
}

