// vybe-test: kotlin/builtins/test_abs_nan_input_stays_not_a_number
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0.0 / 0.0
            __check((abs(value).isNaN()).toString(), "true")
        }
