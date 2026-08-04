// vybe-test: kotlin/java_collections_queue/test_priority_queue_orders_by_natural_order
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pq = java.util.PriorityQueue<Int>()
            pq.add(5)
            pq.add(1)
            pq.add(3)
            __p((pq.peek()).toString())
            __p((pq.poll()).toString())
            __p((pq.peek()).toString())
            __p((pq.size).toString())
        
__check("1\n1\n3\n2")
}
