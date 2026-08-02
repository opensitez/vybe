// vybe-test: kotlin/strings/test_empty_and_blank_predicates
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = ""
            val blanks = "  \n\t"
            val word = "k"
            __check((empty.isEmpty()).toString(), "true")
            __check((empty.isBlank()).toString(), "true")
            __check((blanks.isEmpty()).toString(), "false")
            __check((blanks.isBlank()).toString(), "true")
            __check((word.isBlank()).toString(), "false")
        }
