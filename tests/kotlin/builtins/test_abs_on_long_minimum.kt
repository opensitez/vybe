// vybe-test: kotlin/builtins/test_abs_on_long_minimum
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(Long.MIN_VALUE)).toString(), "-9223372036854775808")
        }
