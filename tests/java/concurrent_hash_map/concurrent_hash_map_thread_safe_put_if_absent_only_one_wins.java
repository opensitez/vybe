// vybe-test: java/concurrent_hash_map/concurrent_hash_map_thread_safe_put_if_absent_only_one_wins
// origin: languages/java/tests/java/test_concurrent_hash_map.rs

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

static java.util.concurrent.ConcurrentHashMap<String, Integer> map = new java.util.concurrent.ConcurrentHashMap<String, Integer>();
        static int wins = 0;
    public static void main(String[] args) {
Thread t1 = new Thread(() -> { if (map.putIfAbsent("k", 1) == null) wins++; }); Thread t2 = new Thread(() -> { if (map.putIfAbsent("k", 2) == null) wins++; }); t1.start(); t2.start(); t1.join(); t2.join(); __p(wins); __p(map.size());
__check("1\n1");
    }
}

