// vybe-test: kotlin/kotlin_string_line_ops/test_lines_without_trailing
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "a\n"
            val lines = value.lines()
            __check((lines.size).toString(), "2")
            __check((lines[0]).toString(), "a")
            __check((lines[1]).toString(), "")
        }
