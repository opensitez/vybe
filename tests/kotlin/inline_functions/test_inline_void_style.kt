// vybe-test: kotlin/inline_functions/test_inline_void_style
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun tap(value: Int, action: (Int) -> Unit): Int {
            action(value)
            return value
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
            var seen = 0
            val out = tap(3) { v -> seen += v }
            __p((out).toString())
            __p((seen).toString())
        
__check("3\n3")
}
