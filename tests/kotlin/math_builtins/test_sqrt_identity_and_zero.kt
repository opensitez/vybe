// vybe-test: kotlin/math_builtins/test_sqrt_identity_and_zero
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sqrt(0.0)).toString(), "0")
            __check((sqrt(1.0)).toString(), "1")
            __check((sqrt(4.0)).toString(), "2")
            __check((sqrt(81.0)).toString(), "9")
        }
