// vybe-test: kotlin/math_builtins/test_abs_plus_pow_and_rounding_stability
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = abs(pow(-2.0, 2.0) - 3.0)
            val y = round(2.5)
            val z = floor(5.9 - 1.2)
            __check((x).toString(), "1")
            __check((y).toString(), "3")
            __check((z).toString(), "4")
            __check((x + y + z).toString(), "8")
        }
