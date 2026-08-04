// vybe-test: kotlin/numeric_types/test_int_and_double_mix_rounds_through_double
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

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
            val value = 5
            __p((value + 2.5).toString())
            __p((value * 1.5).toString())
            __p((value / 2.0).toString())
            __p((10 / 4 + 0.5).toString())
        
__check("7.5\n7.5\n2.5\n3.5")
}
