// vybe-test: kotlin/iterator_protocol/test_iterator_collect_to_mutable_list
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
            val src = listOf(4, 5, 6).iterator()
            val dst = mutableListOf<Int>()
            while (src.hasNext()) {
                dst.add(src.next())
            }
            __p((dst.joinToString(",")).toString())
        
__check("4,5,6")
}
