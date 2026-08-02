// vybe-test: kotlin/kotlin_multiline_strings/test_raw_string_with_tabs_preserved
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """a\tb\tc"""
            __check((text.length).toString(), "5")
        }
