// vybe-test: kotlin/kotlin_big_numbers/test_big_numbers_with_large_exponent_and_abs
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.math.BigDecimal("-123456")
            __check((value.abs().toPlainString()).toString(), "123456")
            val scaled = java.math.BigInteger("2").pow(20)
            __check((scaled.toString()).toString(), "1048576")
            __check((scaled.toString().length).toString(), "7")
        }
