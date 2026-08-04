// vybe-test: kotlin/java_collections_queue/test_arrays_deque_bidirectional_views
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
            val q = java.util.ArrayDeque<Int>()
            q.addFirst(1)
            q.addLast(2)
            q.addFirst(0)
            __p((q.toString()).toString())
            __p((q.pop()).toString())
            __p((q.removeLast()).toString())
            __p((q.removeFirst()).toString())
        
__check("[0, 1, 2]\n0\n2\n1")
}
