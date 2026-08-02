// vybe-test: kotlin/string_builtins/test_string_is_blank_and_empty
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("".isEmpty()).toString(), "true")
            __check(("   ".isBlank()).toString(), "true")
            __check(("x".isBlank()).toString(), "false")
        }
