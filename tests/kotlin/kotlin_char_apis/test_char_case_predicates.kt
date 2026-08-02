// vybe-test: kotlin/kotlin_char_apis/test_char_case_predicates
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('a'.isUpperCase()).toString(), "false")
            __check(('A'.isUpperCase()).toString(), "true")
            __check(('a'.isLowerCase()).toString(), "true")
            __check(('A'.isLowerCase()).toString(), "false")
            __check(('9'.isUpperCase()).toString(), "false")
        }
