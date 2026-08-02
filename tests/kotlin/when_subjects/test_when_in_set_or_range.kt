// vybe-test: kotlin/when_subjects/test_when_in_set_or_range
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun classify(x: Int): String = when (x) {
            1, 2, 3 -> "small"
            in 4..6 -> "mid"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(2)).toString(), "small")
            __check((classify(5)).toString(), "mid")
            __check((classify(10)).toString(), "other")
        }
