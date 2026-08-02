// vybe-test: kotlin/kotlin_string_line_ops/test_trim_margin_and_indent
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
                |a
                |b
                |c
            """.trimMargin()
            __check((text).toString(), "a\nb\nc")
            val raw = """\n    one\n    two\n""".trimIndent()
            __check((raw.startsWith("one")).toString(), "true")
            __check((raw.lines().size).toString(), "2")
        }
