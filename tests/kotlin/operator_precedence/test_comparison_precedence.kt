// vybe-test: kotlin/operator_precedence/test_comparison_precedence
// origin: languages/kotlin/tests/kotlin/test_operator_precedence.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3 + 2 > 4 && 2 * 2 == 4).toString(), "true")
        }
