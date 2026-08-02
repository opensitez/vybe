// vybe-test: kotlin/builtins/test_sqrt_zero_and_one
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sqrt(0.0)).toString(), "0")
            __check((sqrt(1.0)).toString(), "1")
        }
