// vybe-test: kotlin/when_guards/test_when_guard_fallback
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = -1
            val out = when {
                v > 10 -> "high"
                v < 0 -> "low"
                else -> "mid"
            }
            __check((out).toString(), "low")
        }
