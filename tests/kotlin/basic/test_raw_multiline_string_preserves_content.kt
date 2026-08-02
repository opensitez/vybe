// vybe-test: kotlin/basic/test_raw_multiline_string_preserves_content
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
first
second
third
"""
            __check((text.trim()).toString(), "first\nsecond\nthird")
        }
