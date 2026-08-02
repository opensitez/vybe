// vybe-test: kotlin/math_builtins/test_atan2_quadrant_and_zero_axis_signals
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.math.atan2(0.0, -1.0)
            val b = kotlin.math.atan2(1.0, 0.0)
            __check((a == kotlin.math.PI).toString(), "true")
            __check((b > 1.5).toString(), "true")
            __check((b < 2.0).toString(), "true")
        }
