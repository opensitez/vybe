// vybe-test: kotlin/operator_precedence/test_not_before_equality
// origin: languages/kotlin/tests/kotlin/test_operator_precedence.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3
            __check((! (value == 4)).toString(), "true")
        }
