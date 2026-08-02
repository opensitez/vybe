// vybe-test: kotlin/when_subjects/test_when_subject_with_fallback_pair
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = Pair(9, 1)
            val out = when (pair) {
                Pair(1, 1) -> "11"
                Pair(9, 1) -> "91"
                else -> "other"
            }
            __check((out).toString(), "91")
        }
