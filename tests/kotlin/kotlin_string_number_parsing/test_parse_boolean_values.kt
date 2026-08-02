// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_boolean_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

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
            __check(("true".toBooleanStrictOrNull()).toString(), "true")
            __check(("TRUE".toBooleanStrictOrNull()).toString(), "null")
            __check(("x".toBooleanStrictOrNull()).toString(), "null")
        }
