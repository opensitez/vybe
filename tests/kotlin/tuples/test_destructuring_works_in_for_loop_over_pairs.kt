// vybe-test: kotlin/tuples/test_destructuring_works_in_for_loop_over_pairs
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

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
            val rows = listOf(Pair("a", 1), Pair("b", 2), Pair("c", 3))
            var total = ""
            for ((label, value) in rows) {
                total += "$label$value-"
            }
            __p((total).toString())
        
__check("a1-b2-c3-")
}
