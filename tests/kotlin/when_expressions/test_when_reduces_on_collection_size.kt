// vybe-test: kotlin/when_expressions/test_when_reduces_on_collection_size
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun sizeLabel(values: List<Int>): String {
            return when (values.size) {
                0 -> "empty"
                in 1..2 -> "small"
                in 3..4 -> "mid"
                else -> "large"
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
            __p((sizeLabel(listOf())).toString())
            __p((sizeLabel(listOf(1))).toString())
            __p((sizeLabel(listOf(1, 2, 3))).toString())
            __p((sizeLabel(listOf(1, 2, 3, 4, 5))).toString())
        
__check("empty\nsmall\nmid\nlarge")
}
