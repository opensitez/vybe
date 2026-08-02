// vybe-test: kotlin/when_subjects/test_when_with_multiple_checks
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun classify(v: Int): String = when (v) {
            1, 2, 3 -> "low"
            4, 5, 6 -> "mid"
            else -> "high"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(2)).toString(), "low")
            __check((classify(5)).toString(), "mid")
            __check((classify(9)).toString(), "high")
        }
