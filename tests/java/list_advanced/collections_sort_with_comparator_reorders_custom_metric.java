// vybe-test: java/list_advanced/collections_sort_with_comparator_reorders_custom_metric
// origin: languages/java/tests/java/test_list_advanced.rs

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

static class Item { int rank; Item(int rank) { this.rank = rank; } }
    public static void main(String[] args) {
java.util.ArrayList<Item> list = new java.util.ArrayList<Item>(); list.add(new Item(3)); list.add(new Item(1)); list.add(new Item(2)); java.util.Collections.sort(list, (a, b) -> Integer.compare(a.rank, b.rank)); __p(list.get(0).rank); __p(list.get(2).rank);
__check("1\n3");
    }
}

