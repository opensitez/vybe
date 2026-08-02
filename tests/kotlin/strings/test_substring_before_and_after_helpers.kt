// vybe-test: kotlin/strings/test_substring_before_and_after_helpers
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "name=value"
            __check((value.substringAfter("=")).toString(), "value")
            __check((value.substringBefore("=")).toString(), "name")
            __check(("plain".substringAfter("=", "missing")).toString(), "missing")
        }
