// vybe-test: kotlin/strings/test_string_filter_and_counted_predicates
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "a1b2c3d4"
            val digits = value.count { it.isDigit() }
            val letters = value.count { it.isLetter() }
            val filtered = value.filterIndexed { index, ch -> index % 2 == 0 && ch.isLetter() }
            __check((digits).toString(), "4")
            __check((letters).toString(), "4")
            __check((filtered).toString(), "ac")
        }
