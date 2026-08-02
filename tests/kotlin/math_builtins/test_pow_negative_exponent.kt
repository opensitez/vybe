// vybe-test: kotlin/math_builtins/test_pow_negative_exponent
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(2.0, -1.0)).toString(), "0.5")
            __check((pow(4.0, -2.0)).toString(), "0.0625")
            __check((pow(10.0, -1.0)).toString(), "0.1")
        }
