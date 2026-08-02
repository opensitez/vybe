// vybe-test: kotlin/operator_precedence/test_boolean_and_before_or
// origin: languages/kotlin/tests/kotlin/test_operator_precedence.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true || false && false).toString(), "true")
        }
