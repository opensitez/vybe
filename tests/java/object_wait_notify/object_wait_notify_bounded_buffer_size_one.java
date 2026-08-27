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

// vybe-test: java/object_wait_notify/object_wait_notify_bounded_buffer_size_one
// origin: languages/java/tests/java/test_object_wait_notify.rs

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

static class Slot {
            Integer val = null;
            synchronized void put(int v) throws InterruptedException { while (val != null) wait(); val = v; notify(); }
            synchronized int get() throws InterruptedException { while (val == null) wait(); int r = val; val = null; notify(); return r; }
        }
    public static void main(String[] args) throws Throwable {
Slot s = new Slot(); Thread c = new Thread(() -> { try { __p(s.get()); } catch (InterruptedException e) {} }); c.start(); Thread.sleep(10); try { s.put(99); } catch (InterruptedException e) {} c.join();
__check("99");
    }
}

