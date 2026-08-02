// vybe-test: kotlin/strings/test_lines_and_trim_with_blank_lines
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "a\n\nb\n"
            val raw = value.lines()
            __check((raw.size).toString(), "3")
            __check((raw[1]).toString(), "")
            __check((raw[2].isEmpty()).toString(), "true")
            __check((value.lines().filter { it.isNotEmpty() }.joinToString("|")).toString(), "a|b")
        }
