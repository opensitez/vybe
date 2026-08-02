// vybe-test: kotlin/kotlin_operator_overflow/test_power_by_repeated_multiplication
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 2 * 3 * 4
            __check((x).toString(), "24")
            val y = 2L * 3L * 4L
            __check((y).toString(), "24")
        }
