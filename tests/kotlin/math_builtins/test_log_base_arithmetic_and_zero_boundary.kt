// vybe-test: kotlin/math_builtins/test_log_base_arithmetic_and_zero_boundary
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
            val ten = kotlin.math.log(1000.0, 10.0)
            val two = kotlin.math.log(8.0, 2.0)
            val tiny = kotlin.math.log10(1.0)
            __p((kotlin.math.round(ten)).toString())
            __p((kotlin.math.round(two)).toString())
            __p((tiny).toString())
        
__check("3\n3\n0.0")
}
