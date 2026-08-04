// vybe-test: kotlin/collections_iterables/test_sequence_take_and_take_while_are_short_circuiting
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
            var mapped = 0
            val seq = sequenceOf(1, 2, 3, 4, 5).map {
                mapped += 1
                it
            }
            val taken = seq.take(3).toList().joinToString(",")
            __p((mapped).toString())
            val bounded = sequenceOf(1, 2, 3, 4, 5)
                .map { it }
                .takeWhile { it < 4 }
                .toList()
                .joinToString(",")
            __p((bounded).toString())
            __p((mapped).toString())
        
__check("3\n1,2,3\n3")
}
