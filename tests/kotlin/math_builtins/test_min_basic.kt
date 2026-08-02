// vybe-test: kotlin/math_builtins/test_min_basic
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((min(3, 7)).toString(), "3")
            __check((min(-5, 2)).toString(), "-5")
            __check((min(9, 9)).toString(), "9")
        }
