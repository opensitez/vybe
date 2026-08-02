// vybe-test: kotlin/math_builtins/test_log_base_arithmetic_and_zero_boundary
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ten = kotlin.math.log(1000.0, 10.0)
            val two = kotlin.math.log(8.0, 2.0)
            val tiny = kotlin.math.log10(1.0)
            __check((kotlin.math.round(ten)).toString(), "3")
            __check((kotlin.math.round(two)).toString(), "3")
            __check((tiny).toString(), "0.0")
        }
