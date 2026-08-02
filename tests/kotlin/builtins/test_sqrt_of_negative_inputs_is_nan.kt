// vybe-test: kotlin/builtins/test_sqrt_of_negative_inputs_is_nan
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = sqrt(-4.0)
            __check((value.isNaN()).toString(), "true")
        }
