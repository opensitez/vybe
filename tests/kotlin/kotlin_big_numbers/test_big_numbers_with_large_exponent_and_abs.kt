// vybe-test: kotlin/kotlin_big_numbers/test_big_numbers_with_large_exponent_and_abs
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

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
            val value = java.math.BigDecimal("-123456")
            __p((value.abs().toPlainString()).toString())
            val scaled = java.math.BigInteger("2").pow(20)
            __p((scaled.toString()).toString())
            __p((scaled.toString().length).toString())
        
__check("123456\n1048576\n7")
}
