// vybe-test: kotlin/iterator_protocol/test_iterator_peeking_behavior
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

class P : Iterator<Int> {
            private val data = listOf(7, 8)
            private var index = 0
            override fun hasNext() = index < data.size
            override fun next() = data[index++]
        }

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
            val it = P()
            __p((it.hasNext()).toString())
            __p((it.next()).toString())
            __p((it.hasNext()).toString())
            __p((it.next()).toString())
            __p((it.hasNext()).toString())
        
__check("true\n7\ntrue\n8\nfalse")
}
