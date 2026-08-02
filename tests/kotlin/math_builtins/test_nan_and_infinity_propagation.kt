// vybe-test: kotlin/math_builtins/test_nan_and_infinity_propagation
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nan = kotlin.math.sqrt(-1.0)
            val inf = 1.0 / 0.0
            val finite = kotlin.math.isFinite(inf)
            val notFinite = kotlin.math.isFinite(nan)
            __check((nan.isNaN()).toString(), "true")
            __check((finite).toString(), "false")
            __check((notFinite).toString(), "false")
        }
