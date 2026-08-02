// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_multi_line_template_concatenation
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lines = """
                a
                b
            """.trimIndent()
            __check(("${'$'}{lines.lines().size}").toString(), "2")
            __check((lines[0]).toString(), "a")
        }
