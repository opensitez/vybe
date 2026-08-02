// vybe-test: kotlin/math_builtins/test_min_max_chain
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((max(1, max(3, 9))).toString(), "9")
            __check((min(1, min(3, 9))).toString(), "1")
            __check((max(min(10, 2), max(-4, 12))).toString(), "12")
            __check((min(max(10, 2), min(-4, 12))).toString(), "-4")
        }
