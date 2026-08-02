// vybe-test: kotlin/strings/test_string_line_splitting_retains_trailing_empty
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "a\nb\n"
            val lines = value.lines()
            __check((lines.size).toString(), "3")
            __check((lines[0]).toString(), "a")
            __check((lines[1]).toString(), "b")
            __check((lines[2]).toString(), "")
        }
