// vybe-test: kotlin/builtins/test_abs_and_zero_behavior
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(0)).toString(), "0")
            __check((abs(1)).toString(), "1")
            __check((abs(-1)).toString(), "1")
        }
