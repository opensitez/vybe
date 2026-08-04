// vybe-test: kotlin/java_collections_queue/test_linked_list_as_queue
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
            val list = java.util.LinkedList<String>()
            list.offer("first")
            list.offer("second")
            __p((list.peek()).toString())
            __p((list.poll()).toString())
            __p((list.peek()).toString())
            __p((list.removeFirst()).toString())
        
__check("first\nfirst\nsecond\nsecond")
}
