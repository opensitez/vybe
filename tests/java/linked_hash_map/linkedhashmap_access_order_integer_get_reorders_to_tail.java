// vybe-test: java/linked_hash_map/linkedhashmap_access_order_integer_get_reorders_to_tail
// origin: languages/java/tests/java/test_linked_hash_map.rs

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

    public static void main(String[] args) {
java.util.LinkedHashMap<Integer, String> map = new java.util.LinkedHashMap<Integer, String>(16, 0.75f, true); map.put(1, "a"); map.put(2, "b"); map.put(3, "c"); map.get(1); java.util.Iterator<Integer> it = map.keySet().iterator(); __p(it.next()); __p(it.next()); __p(it.next());
__check("2\n3\n1");
    }
}

