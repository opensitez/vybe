// vybe-test: kotlin/when_guards/test_when_guard_with_range
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun score(v: Int): String = when {
            v in 1..3 -> "low"
            v in 4..6 -> "mid"
            v in 7..9 -> "high"
            else -> "out"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(2)).toString(), "low")
            __check((score(6)).toString(), "mid")
            __check((score(12)).toString(), "out")
        }
