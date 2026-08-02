// vybe-test: kotlin/boolean_logic/test_boolean_literals_and_basic_ops
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

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
            __check((!false).toString(), "true")
        }
