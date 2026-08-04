// vybe-test: kotlin/kotlin_big_numbers/test_big_decimal_scale_and_precision
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

import java.math.RoundingMode

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
            val value = java.math.BigDecimal("12.3456")
            val reduced = value.setScale(2, RoundingMode.HALF_UP)
            __p((reduced.toPlainString()).toString())
            __p((reduced.scale()).toString())
            val up = value.setScale(1, RoundingMode.CEILING)
            __p((up.toPlainString()).toString())
        
__check("12.35\n2\n12.4")
}
