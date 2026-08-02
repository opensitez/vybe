// vybe-test: kotlin/when_expressions/test_when_with_string_subject_patterns
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun bucket(word: String): String {
            return when (word.lowercase()) {
                "yes", "y", "oui" -> "affirmative"
                "no", "n", "non" -> "negative"
                else -> "unknown"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((bucket("YES")).toString(), "affirmative")
            __check((bucket("No")).toString(), "negative")
            __check((bucket("maybe")).toString(), "unknown")
        }
