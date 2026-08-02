// vybe-test: kotlin/kotlin_multiline_strings/test_raw_string_leading_spaces_trimmed
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
            left
            right
            """.trimIndent()
            __check((text).toString(), "left\nright")
        }
