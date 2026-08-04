// vybe-test: java/map_advanced/hashmap_for_each_can_build_concatenated_key_list
// origin: languages/java/tests/java/test_map_advanced.rs

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

static class Keys { static String join(java.util.Map<String, Integer> map) { final String[] out = {""}; map.forEach((k, v) -> { out[0] = out[0] + k; }); return out[0]; } }
    public static void main(String[] args) {
java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put("x", 1); map.put("y", 2); __p(Keys.join(map));
__check("xy");
    }
}

