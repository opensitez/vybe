// vybe-test: java/priority_queue/priorityqueue_comparator_natural_order_matches_default_min_heap
// origin: languages/java/tests/java/test_priority_queue.rs

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
java.util.PriorityQueue<Integer> def = new java.util.PriorityQueue<Integer>(); def.offer(3); def.offer(1); java.util.PriorityQueue<Integer> explicit = new java.util.PriorityQueue<Integer>(java.util.Comparator.naturalOrder()); explicit.offer(3); explicit.offer(1); __p(def.poll()); __p(explicit.poll());
__check("1\n1");
    }
}

