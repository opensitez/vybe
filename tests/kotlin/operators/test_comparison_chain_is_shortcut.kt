// vybe-test: kotlin/operators/test_comparison_chain_is_shortcut
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((2 + 3 > 3 * 1 && 4 <= 4).toString(), "true")
            __check((2 + 3 > 3 * 2 || 4 < 1).toString(), "false")
            __check((5 > 4 && 4 > 3).toString(), "true")
            __check((5 > 4 && 4 > 3 && 2 > 1).toString(), "true")
        }
