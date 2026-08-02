// vybe-test: kotlin/builtins/test_rounding_threshold_boundaries
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(2.499_999)).toString(), "2")
            __check((round(2.5)).toString(), "3")
            __check((round(2.500_001)).toString(), "3")
            __check((round(-2.500_001)).toString(), "-3")
        }
