// vybe-test: kotlin/characters/test_upper_and_lower_case_checks
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('a'.isLowerCase()).toString(), "true")
            __check(('B'.isUpperCase()).toString(), "true")
            __check(('9'.isLowerCase()).toString(), "false")
            __check(('9'.isUpperCase()).toString(), "false")
        }
