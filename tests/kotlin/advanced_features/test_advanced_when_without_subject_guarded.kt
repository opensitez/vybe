// vybe-test: kotlin/advanced_features/test_advanced_when_without_subject_guarded
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 0
            val label = when {
                x > 0 -> "positive"
                x == 0 -> "zero"
                else -> "negative"
            }
            __check((label).toString(), "zero")
        }
