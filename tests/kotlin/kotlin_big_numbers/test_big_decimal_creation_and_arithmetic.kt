// vybe-test: kotlin/kotlin_big_numbers/test_big_decimal_creation_and_arithmetic
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

import java.math.RoundingMode

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.math.BigDecimal("10.5")
            val b = java.math.BigDecimal("4")
            __check((a.add(b).toPlainString()).toString(), "14.5")
            __check((a.subtract(b).toPlainString()).toString(), "6.5")
            __check((a.multiply(b).toPlainString()).toString(), "42.0")
            __check((a.divide(b, 2, RoundingMode.HALF_UP).toPlainString()).toString(), "2.62")
        }
