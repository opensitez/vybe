// vybe-test: kotlin/when_expressions/test_when_with_fallback_expressions_using_arithmetic
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun score(value: Int): String {
            return when (value % 3) {
                0 -> "triple"
                1 -> "plus"
                2 -> "plus2"
                else -> "?"
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
            __p((score(10)).toString())
            __p((score(11)).toString())
            __p((score(12)).toString())
        
__check("plus\nplus2\ntriple")
}
