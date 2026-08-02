// vybe-test: kotlin/builtins/test_abs_on_int_minimum
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(Int.MIN_VALUE)).toString(), "-2147483648")
        }
