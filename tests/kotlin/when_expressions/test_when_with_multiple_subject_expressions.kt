// vybe-test: kotlin/when_expressions/test_when_with_multiple_subject_expressions
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun classify(value: Int): String {
            return when (value) {
                in 1..3 -> "low"
                in 4..10 -> "mid"
                !in 1..10 -> "outside"
                else -> "other"
            }
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
            __p((classify(2)).toString())
            __p((classify(10)).toString())
            __p((classify(20)).toString())
        
__check("low\nmid\noutside")
}
