// vybe-test: java/tree_map/treemap_custom_comparator_with_record_values_orders_by_rank
// origin: languages/java/tests/java/test_tree_map.rs

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

static class Ranked { int score; Ranked(int score) { this.score = score; } }
    public static void main(String[] args) {
java.util.TreeMap<Ranked, String> map = new java.util.TreeMap<Ranked, String>((a, b) -> Integer.compare(a.score, b.score)); map.put(new Ranked(30), "high"); map.put(new Ranked(10), "low"); map.put(new Ranked(20), "mid"); __p(map.firstKey().score); __p(map.lastKey().score);
__check("10\n30");
    }
}

