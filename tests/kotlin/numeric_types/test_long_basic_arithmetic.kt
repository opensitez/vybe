// vybe-test: kotlin/numeric_types/test_long_basic_arithmetic
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Long = 1_000_000_000_000
            val b: Long = 250
            __check((a + b).toString(), "1000000000250")
            __check((a - b).toString(), "999999999750")
            __check((a * 2).toString(), "2000000000000")
            __check((a / b).toString(), "4000000")
        }
