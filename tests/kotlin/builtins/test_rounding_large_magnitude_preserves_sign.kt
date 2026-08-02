// vybe-test: kotlin/builtins/test_rounding_large_magnitude_preserves_sign
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(123456.49)).toString(), "123456")
            __check((round(123456.50)).toString(), "123457")
            __check((round(-123456.51)).toString(), "-123457")
        }
