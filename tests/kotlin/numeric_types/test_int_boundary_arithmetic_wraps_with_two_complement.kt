// vybe-test: kotlin/numeric_types/test_int_boundary_arithmetic_wraps_with_two_complement
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
            __p((Int.MAX_VALUE + 1).toString())
            __p((Int.MIN_VALUE - 1).toString())
            __p((Long.MAX_VALUE + 1).toString())
            __p((Long.MIN_VALUE - 1).toString())
        
__check("-2147483648\n2147483647\n-9223372036854775808\n9223372036854775807")
}
