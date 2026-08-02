// vybe-test: kotlin/when_subjects/test_when_range_membership
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun inRange(x: Int): String = when (x) {
            in 1..3 -> "small"
            in 4..6 -> "mid"
            else -> "big"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((inRange(2)).toString(), "small")
            __check((inRange(5)).toString(), "mid")
            __check((inRange(9)).toString(), "big")
        }
