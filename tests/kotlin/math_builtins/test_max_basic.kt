// vybe-test: kotlin/math_builtins/test_max_basic
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((max(3, 7)).toString(), "7")
            __check((max(-5, 2)).toString(), "2")
            __check((max(9, 9)).toString(), "9")
        }
