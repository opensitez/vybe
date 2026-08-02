// vybe-test: kotlin/builtins/test_pow_fractional_root_path
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(9.0, 0.5)).toString(), "3")
        }
