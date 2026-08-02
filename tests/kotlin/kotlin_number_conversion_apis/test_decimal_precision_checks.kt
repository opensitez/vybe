// vybe-test: kotlin/kotlin_number_conversion_apis/test_decimal_precision_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0.1f + 0.2f).toString(), "0.30000001")
            __check(((0.1 + 0.2) == 0.3).toString(), "false")
            __check(((1_000_000_000_000.0).toLong()).toString(), "1000000000000")
            __check(((1.0 / 0.0).isInfinite()).toString(), "true")
            __check(((-1.0 / 0.0).isInfinite()).toString(), "true")
        }
