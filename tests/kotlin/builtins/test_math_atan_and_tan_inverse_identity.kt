// vybe-test: kotlin/builtins/test_math_atan_and_tan_inverse_identity
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val angle = 0.75
            __check((round((tan(atan(angle)) - angle) * 1e9)).toString(), "0")
            __check((sign(0.0)).toString(), "0.0")
            __check((sign(-5.0)).toString(), "-1.0")
            __check((sign(5.0)).toString(), "1.0")
        }
