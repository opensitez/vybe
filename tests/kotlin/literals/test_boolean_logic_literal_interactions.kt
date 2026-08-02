// vybe-test: kotlin/literals/test_boolean_logic_literal_interactions
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true && false).toString(), "false")
            __check((false || true).toString(), "true")
            __check((!false).toString(), "true")
            __check((true && (1 > 0)).toString(), "true")
        }
