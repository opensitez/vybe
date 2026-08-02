// vybe-test: kotlin/string_search_apis/test_line_and_trimmed_queries
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val block = "\n a \n b\n"
            val lines = block.lines()
            val trimmed = block.trim()
            __check((lines.size).toString(), "3")
            __check((trimmed).toString(), "a \n b")
            __check((trimmed.isNotEmpty()).toString(), "true")
        }
