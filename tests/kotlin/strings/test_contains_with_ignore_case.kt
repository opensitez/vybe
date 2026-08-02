// vybe-test: kotlin/strings/test_contains_with_ignore_case
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "Kotlin"
            __check((word.contains("kot")).toString(), "false")
            __check((word.contains("kot", true)).toString(), "true")
            __check((word.startsWith("Ko")).toString(), "true")
            __check((word.endsWith("IN", ignoreCase = true)).toString(), "true")
        }
