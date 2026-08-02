// vybe-test: kotlin/builtins/test_pow_zero_one_and_negative_base_sign
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(0.0, 0.0)).toString(), "1")
            __check((pow(1.0, 9.0)).toString(), "1")
            __check((pow(-3.0, 2.0)).toString(), "9")
            __check((pow(-3.0, 3.0)).toString(), "-27")
        }
