// vybe-test: kotlin/math_builtins/test_pow_positive_exponent_integers_as_doubles
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(2.0, 0.0)).toString(), "1")
            __check((pow(2.0, 1.0)).toString(), "1")
            __check((pow(2.0, 2.0)).toString(), "4")
            __check((pow(2.0, 3.0)).toString(), "8")
            __check((pow(2.0, 4.0)).toString(), "16")
        }
