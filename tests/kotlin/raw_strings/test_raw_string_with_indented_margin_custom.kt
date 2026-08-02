// vybe-test: kotlin/raw_strings/test_raw_string_with_indented_margin_custom
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
                >a
                >b
            """.trimMargin(">")
            __check((text.lines().size).toString(), "2")
        }
