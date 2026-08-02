// vybe-test: kotlin/math_builtins/test_pow_zero_base_cases
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(0.0, 5.0)).toString(), "0")
            __check((pow(0.0, 0.0)).toString(), "1")
            __check((pow(0.0, 1.0)).toString(), "0")
        }
