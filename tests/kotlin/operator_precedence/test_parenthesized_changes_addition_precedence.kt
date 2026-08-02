// vybe-test: kotlin/operator_precedence/test_parenthesized_changes_addition_precedence
// origin: languages/kotlin/tests/kotlin/test_operator_precedence.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((2 + 3) * 4).toString(), "20")
        }
