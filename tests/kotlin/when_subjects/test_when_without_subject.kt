// vybe-test: kotlin/when_subjects/test_when_without_subject
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = 7
            val out = when {
                v < 0 -> "neg"
                v < 5 -> "low"
                v == 7 -> "seven"
                else -> "other"
            }
            __check((out).toString(), "seven")
        }
