// vybe-test: kotlin/infix/test_infix_contains_from_extension
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Window(val values: Set<Int>) {
            operator fun contains(value: Int): Boolean = values.contains(value)
        }

        infix fun Window.has(value: Int): Boolean = value in this

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
            val setWindow = Window(setOf(2, 4, 6))
            __p((setWindow has 4).toString())
            __p((setWindow has 3).toString())
        
__check("true\nfalse")
}
