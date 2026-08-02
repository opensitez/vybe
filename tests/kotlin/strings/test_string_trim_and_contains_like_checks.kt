// vybe-test: kotlin/strings/test_string_trim_and_contains_like_checks
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "  Kotlin  "
            val trimmed = value.trim()
            __check((trimmed).toString(), "Kotlin")
            __check((trimmed.startsWith("Kot")).toString(), "true")
            __check((trimmed.endsWith("lin")).toString(), "true")
            __check((trimmed.contains("tin")).toString(), "true")
        }
