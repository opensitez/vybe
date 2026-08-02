// vybe-test: kotlin/conversions/test_boolean_parse_with_numeric_aliases_is_false
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("1".toBoolean()).toString(), "false")
            __check(("0".toBoolean()).toString(), "false")
            __check(("TRUE ".toBoolean()).toString(), "false")
            __check((" false ".toBoolean()).toString(), "false")
        }
