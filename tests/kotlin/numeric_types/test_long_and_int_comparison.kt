// vybe-test: kotlin/numeric_types/test_long_and_int_comparison
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val intValue = 100
            val longValue = 100L
            __check((intValue == longValue).toString(), "true")
            __check((intValue < longValue + 1).toString(), "true")
            __check(((intValue + 1) >= longValue).toString(), "true")
        }
