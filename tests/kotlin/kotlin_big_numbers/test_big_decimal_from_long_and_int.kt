// vybe-test: kotlin/kotlin_big_numbers/test_big_decimal_from_long_and_int
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.math.BigDecimal(123L)
            val b = java.math.BigDecimal.valueOf(45L)
            __check((a + b).toString(), "168")
            __check((java.math.BigDecimal("1.5") + java.math.BigDecimal.valueOf(1)).toString(), "2.5")
        }
