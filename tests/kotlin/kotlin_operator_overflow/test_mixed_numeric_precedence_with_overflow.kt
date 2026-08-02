// vybe-test: kotlin/kotlin_operator_overflow/test_mixed_numeric_precedence_with_overflow
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Int.MAX_VALUE
            val b = a + 1 - 1
            __check((b).toString(), "2147483647")
        }
