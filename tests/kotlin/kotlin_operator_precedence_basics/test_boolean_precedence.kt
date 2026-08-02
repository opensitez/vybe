// vybe-test: kotlin/kotlin_operator_precedence_basics/test_boolean_precedence
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_precedence_basics.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true || false && false).toString(), "true")
            __check(((true || false) && false).toString(), "false")
        }
