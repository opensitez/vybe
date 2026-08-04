// vybe-test: kotlin/builtins/test_math_atan_and_tan_inverse_identity
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

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
            val angle = 0.75
            __p((round((tan(atan(angle)) - angle) * 1e9)).toString())
            __p((sign(0.0)).toString())
            __p((sign(-5.0)).toString())
            __p((sign(5.0)).toString())
        
__check("0\n0.0\n-1.0\n1.0")
}
