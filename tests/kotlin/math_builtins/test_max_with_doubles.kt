// vybe-test: kotlin/math_builtins/test_max_with_doubles
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((max(1.5, 2.5)).toString(), "2.5")
            __check((min(1.5, -2.5)).toString(), "-2.5")
            __check((min(-1.2, -3.4)).toString(), "-3.4")
            __check((max(-1.2, -3.4)).toString(), "-1.2")
        }
