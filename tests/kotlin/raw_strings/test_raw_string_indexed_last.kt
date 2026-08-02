// vybe-test: kotlin/raw_strings/test_raw_string_indexed_last
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """abc"""
            __check((text[text.length - 1]).toString(), "c")
        }
