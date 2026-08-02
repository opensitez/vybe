// vybe-test: kotlin/raw_strings/test_raw_string_trim_indent_complex
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
                    a
                      b
                    c
            """.trimIndent()
            __check((text.length).toString(), "9")
        }
