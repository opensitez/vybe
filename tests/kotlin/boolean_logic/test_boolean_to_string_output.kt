// vybe-test: kotlin/boolean_logic/test_boolean_to_string_output
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true.toString()).toString(), "true")
            __check((false.toString()).toString(), "false")
            __check(((true && true).toString()).toString(), "true")
        }
