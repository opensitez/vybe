// vybe-test: kotlin/strings/test_compare_to_on_empty_and_singleton
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("".isEmpty()).toString(), "true")
            __check(("".isNotEmpty()).toString(), "false")
            __check(("".compareTo("")).toString(), "0")
            __check(("a".compareTo("")).toString(), "1")
        }
