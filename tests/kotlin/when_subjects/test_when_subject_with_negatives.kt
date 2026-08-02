// vybe-test: kotlin/when_subjects/test_when_subject_with_negatives
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = -3
            val out = when (x) {
                in Int.MIN_VALUE..-1 -> "neg"
                in 0..9 -> "small"
                else -> "other"
            }
            __check((out).toString(), "neg")
        }
