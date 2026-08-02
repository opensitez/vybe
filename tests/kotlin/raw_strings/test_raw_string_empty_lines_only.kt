// vybe-test: kotlin/raw_strings/test_raw_string_empty_lines_only
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
                
                
            """
            __check((text.lines().size).toString(), "4")
        }
