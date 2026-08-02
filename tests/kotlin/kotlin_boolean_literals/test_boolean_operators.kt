// vybe-test: kotlin/kotlin_boolean_literals/test_boolean_operators
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true).toString(), "true")
            __check((false).toString(), "false")
            __check((true && false).toString(), "false")
            __check((true || false).toString(), "true")
            __check((!true).toString(), "false")
        }
