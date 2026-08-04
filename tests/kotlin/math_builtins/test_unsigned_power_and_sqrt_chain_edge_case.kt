// vybe-test: kotlin/math_builtins/test_unsigned_power_and_sqrt_chain_edge_case
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
            val squared = kotlin.math.sqrt(2.0) * kotlin.math.sqrt(2.0)
            val closeToTwo = kotlin.math.abs(squared - 2.0) < 0.0000001
            __p((closeToTwo).toString())
            __p((kotlin.math.pow(2.0, 0.0)).toString())
            __p((kotlin.math.pow(0.0, 1.0)).toString())
            __p((kotlin.math.pow(0.0, 0.0)).toString())
        
__check("true\n1.0\n0.0\n1.0")
}
