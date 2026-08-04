// vybe-test: kotlin/range_projection/test_when_with_range_subject
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

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
            val v = 4
            val label = when (v) {
                in 1..3 -> "small"
                in 4..6 -> "mid"
                else -> "big"
            }
            __p((label).toString())
            __p((v in 1..10).toString())
        
__check("mid\ntrue")
}
