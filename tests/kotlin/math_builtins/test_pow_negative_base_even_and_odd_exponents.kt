// vybe-test: kotlin/math_builtins/test_pow_negative_base_even_and_odd_exponents
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(-2.0, 2.0)).toString(), "4")
            __check((pow(-2.0, 3.0)).toString(), "-8")
            __check((pow(-3.0, 4.0)).toString(), "81")
            __check((pow(-3.0, 5.0)).toString(), "-243")
        }
