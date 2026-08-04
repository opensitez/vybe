// vybe-test: kotlin/destructuring/test_destructuring_custom_component_functions
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

class Holder(private val a: Int, private val b: Int) {
            operator fun component1() = a
            operator fun component2() = b
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
            val value = Holder(7, 8)
            val (left, right) = value
            __p((left).toString())
            __p((right).toString())
            __p((left + right).toString())
        
__check("7\n8\n15")
}
