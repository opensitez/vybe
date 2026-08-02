// vybe-test: kotlin/builtins/test_pow_nested_with_abs_and_sign
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = pow(abs(-12.0), 2.0)
            val signed = pow(-2.0, 4.0)
            __check((value).toString(), "144")
            __check((signed).toString(), "16")
        }
