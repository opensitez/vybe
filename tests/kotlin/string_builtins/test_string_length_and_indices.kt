// vybe-test: kotlin/string_builtins/test_string_length_and_indices
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "kotlin"
            __check((text.length).toString(), "6")
            __check((text[text.length - 1]).toString(), "n")
        }
