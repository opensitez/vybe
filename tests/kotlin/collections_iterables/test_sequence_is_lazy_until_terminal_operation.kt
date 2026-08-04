// vybe-test: kotlin/collections_iterables/test_sequence_is_lazy_until_terminal_operation
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

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
            var called = 0
            val source = sequenceOf(1, 2, 3, 4, 5)
            val transformed = source.map {
                called += 1
                it * 2
            }
            __p((called).toString())
            val first = transformed.first()
            __p((called).toString())
            __p((first).toString())
            val rest = transformed.take(2).toList()
            __p((rest.joinToString(",")).toString())
            __p((called).toString())
        
__check("0\n1\n2\n4,6\n3")
}
