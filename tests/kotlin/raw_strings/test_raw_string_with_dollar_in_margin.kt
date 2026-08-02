// vybe-test: kotlin/raw_strings/test_raw_string_with_dollar_in_margin
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
                |$x = 1
            """.trimMargin()
            __check((text).toString(), "\$x = 1")
        }
