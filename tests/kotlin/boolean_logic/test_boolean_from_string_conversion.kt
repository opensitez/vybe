// vybe-test: kotlin/boolean_logic/test_boolean_from_string_conversion
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("true".toBoolean()).toString(), "true")
            __check(("false".toBoolean()).toString(), "false")
            __check(("TRUE".toBoolean()).toString(), "false")
            __check(("junk".toBoolean()).toString(), "false")
        }
