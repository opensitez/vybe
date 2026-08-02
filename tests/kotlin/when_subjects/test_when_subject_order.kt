// vybe-test: kotlin/when_subjects/test_when_subject_order
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun pick(x: Int): String = when (x) {
            in 1..5 -> "low"
            6 -> "six"
            in 6..9 -> "mid"
            else -> "high"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(2)).toString(), "low")
            __check((pick(6)).toString(), "six")
            __check((pick(8)).toString(), "mid")
        }
