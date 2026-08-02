// vybe-test: kotlin/in_keyword/test_in_string_char_membership
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "kotlin"
            __check(('o' in text).toString(), "true")
            __check(('x' !in text).toString(), "true")
        }
