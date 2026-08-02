// vybe-test: kotlin/kotlin_multiline_strings/test_raw_string_boolean_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = true
            val text = """
status=${'$'}{if (ok) "yes" else "no"}
"""
            __check((text.trim()).toString(), "status=yes")
        }
