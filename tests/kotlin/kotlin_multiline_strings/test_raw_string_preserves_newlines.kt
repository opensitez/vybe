// vybe-test: kotlin/kotlin_multiline_strings/test_raw_string_preserves_newlines
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
line1
line2
line3
"""
            __check((text.lines().size).toString(), "4")
            __check((text.lines()[1]).toString(), "line2")
        }
