// vybe-test: kotlin/literals/test_multiline_comment_style_string_content
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val raw = """
                /*
                one
                */
                """.trimIndent()
            __check((raw.contains("/*")).toString(), "true")
            __check((raw.lines().size).toString(), "3")
        }
