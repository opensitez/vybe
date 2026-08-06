// vybe-test: kotlin/numeric_types/test_long_basic_arithmetic
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
            val a: Long = 1_000_000_000_000
            val b: Long = 250
            __p((a + b).toString())
            __p((a - b).toString())
            __p((a * 2).toString())
            __p((a / b).toString())
        
// Real Kotlin agrees: 1_000_000_000_000 / 250 is 4_000_000_000 — the old
// value dropped three zeros.
__check("1000000000250\n999999999750\n2000000000000\n4000000000")
}
