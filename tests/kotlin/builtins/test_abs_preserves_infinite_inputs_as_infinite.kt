// vybe-test: kotlin/builtins/test_abs_preserves_infinite_inputs_as_infinite
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val positive = abs(1.0 / 0.0)
            val negative = abs(-1.0 / 0.0)
            __check((positive.isInfinite()).toString(), "true")
            __check((negative.isInfinite()).toString(), "true")
            __check((negative > 0.0).toString(), "true")
        }
