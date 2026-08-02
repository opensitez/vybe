// vybe-test: kotlin/kotlin_multiline_strings/test_raw_string_with_quoted_marker
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
"quoted"
not quoted
"""
            __check((text.trim().split("\n")[0]).toString(), "\"quoted\"")
        }
