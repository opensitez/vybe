// vybe-test: kotlin/boolean_logic/test_boolean_or_and_precedence
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true || false && false).toString(), "true")
            __check(((true || false) && false).toString(), "false")
            __check((false || true && true).toString(), "true")
            __check(((false || true) && true).toString(), "true")
        }
