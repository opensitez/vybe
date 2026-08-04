// vybe-test: kotlin/iterator_protocol/test_iterator_multiple_iterables_independent
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

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
            val list = listOf(1, 2)
            val a = list.iterator()
            val b = list.iterator()
            __p((a.next()).toString())
            __p((b.next()).toString())
            __p((a.next()).toString())
            __p((b.next()).toString())
        
__check("1\n1\n2\n2")
}
