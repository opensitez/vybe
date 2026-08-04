// vybe-test: kotlin/when_expressions/test_when_nested_scoping_and_binding
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun describe(a: Int, b: Int): String {
            return when (a) {
                0 -> when (b) {
                    0 -> "a0b0"
                    else -> "a0bN"
                }
                else -> when {
                    b == 0 -> "aNb0"
                    b > 10 -> "aNbH"
                    else -> "aNbL"
                }
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
            __p((describe(0, 0)).toString())
            __p((describe(0, 4)).toString())
            __p((describe(5, 12)).toString())
        
__check("a0b0\na0bN\naNbH")
}
