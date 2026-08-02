// vybe-test: kotlin/kotlin_multiline_strings/test_trim_margin_with_custom_delimiter
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """
>one
>two
>three
""".trimMargin(">")
            __check((text).toString(), "one\ntwo\nthree")
        }
