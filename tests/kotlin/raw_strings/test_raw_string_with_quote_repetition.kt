// vybe-test: kotlin/raw_strings/test_raw_string_with_quote_repetition
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val open = "\"\"\""
            val text = """contains """ + open + """ inside"""
            __check((text).toString(), "contains \"\"\" inside")
        }
