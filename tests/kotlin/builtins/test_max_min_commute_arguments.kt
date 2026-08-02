// vybe-test: kotlin/builtins/test_max_min_commute_arguments
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((max(9, -4)).toString(), "9")
            __check((min(9, -4)).toString(), "-4")
            __check((max(-4, 9)).toString(), "9")
            __check((min(-4, 9)).toString(), "-4")
        }
