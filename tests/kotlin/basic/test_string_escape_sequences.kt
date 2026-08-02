// vybe-test: kotlin/basic/test_string_escape_sequences
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("a\\nb\\n").toString(), "a\\nb\\n")
            __check(("tab\\tend").toString(), "tab\tend")
            __check(("quote: \"").toString(), "quote: \"")
            __check(('c').toString(), "c")
        }
