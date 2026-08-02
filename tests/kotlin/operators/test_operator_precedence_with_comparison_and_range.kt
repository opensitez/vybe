// vybe-test: kotlin/operators/test_operator_precedence_with_comparison_and_range
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 4
            __check((value + 2 * 3 > 12).toString(), "false")
            __check(((value + 2) * 3 > 12).toString(), "true")
            __check((value in 1..(2 + 1) * 2).toString(), "true")
            __check((value in 1..2 + 1 * 2).toString(), "false")
        }
