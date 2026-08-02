// vybe-test: kotlin/conversions/test_boolean_case_insensitive_parse
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("TRUE".toBoolean()).toString(), "true")
            __check(("False".toBoolean()).toString(), "false")
            __check(("fAlSe".toBoolean()).toString(), "false")
        }
