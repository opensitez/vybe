// vybe-test: kotlin/kotlin_operator_overflow/test_unary_minus_min_bound
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((-Int.MIN_VALUE).toString(), "-2147483648")
        }
