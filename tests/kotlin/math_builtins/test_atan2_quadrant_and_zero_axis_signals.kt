// vybe-test: kotlin/math_builtins/test_atan2_quadrant_and_zero_axis_signals
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

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
            val a = kotlin.math.atan2(0.0, -1.0)
            val b = kotlin.math.atan2(1.0, 0.0)
            __p((a == kotlin.math.PI).toString())
            __p((b > 1.5).toString())
            __p((b < 2.0).toString())
        
__check("true\ntrue\ntrue")
}
