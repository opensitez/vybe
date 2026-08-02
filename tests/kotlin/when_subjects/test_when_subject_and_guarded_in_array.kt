// vybe-test: kotlin/when_subjects/test_when_subject_and_guarded_in_array
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun classify(x: Int): String = when (x) {
            2, 4, 6 -> "even-basic"
            in 7..9 -> "high-seven-nine"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(2)).toString(), "even-basic")
            __check((classify(8)).toString(), "high-seven-nine")
            __check((classify(12)).toString(), "other")
        }
