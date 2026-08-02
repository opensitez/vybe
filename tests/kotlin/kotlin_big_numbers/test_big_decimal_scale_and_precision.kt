// vybe-test: kotlin/kotlin_big_numbers/test_big_decimal_scale_and_precision
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

import java.math.RoundingMode

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.math.BigDecimal("12.3456")
            val reduced = value.setScale(2, RoundingMode.HALF_UP)
            __check((reduced.toPlainString()).toString(), "12.35")
            __check((reduced.scale()).toString(), "2")
            val up = value.setScale(1, RoundingMode.CEILING)
            __check((up.toPlainString()).toString(), "12.4")
        }
