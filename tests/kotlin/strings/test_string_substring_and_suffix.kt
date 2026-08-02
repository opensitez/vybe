// vybe-test: kotlin/strings/test_string_substring_and_suffix
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "abcdef"
            __check((source.substring(1, 4)).toString(), "bcd")
            __check((source.substring(3)).toString(), "def")
        }
