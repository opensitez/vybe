// vybe-test: kotlin/kotlin_big_numbers/test_big_decimal_compare_and_sign
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.math.BigDecimal("-2")
            val b = java.math.BigDecimal("3")
            __check((a.compareTo(b)).toString(), "-1")
            __check((a.signum()).toString(), "-1")
            __check((b.signum()).toString(), "1")
            __check((java.math.BigDecimal.ZERO.compareTo(java.math.BigDecimal("0.00"))).toString(), "0")
        }
