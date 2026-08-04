// vybe-test: kotlin/iterator_protocol/test_custom_iterator_implementing_interface
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

class RangeIterator : Iterator<Int> {
            private var i = 0
            private val end = 3
            override fun hasNext() = i < end
            override fun next(): Int {
                val value = i
                i += 1
                return value
            }
        }

        class RangeIterable : Iterable<Int> {
            override fun iterator(): Iterator<Int> = RangeIterator()
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
            val it = RangeIterable().iterator()
            var sum = 0
            for (value in RangeIterable()) {
                sum += value
            }
            __p((sum).toString())
            __p((it.hasNext()).toString())
            __p((it.next()).toString())
            __p((it.next()).toString())
        
__check("3\ntrue\n1\n2")
}
