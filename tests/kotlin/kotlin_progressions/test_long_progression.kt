// vybe-test: kotlin/kotlin_progressions/test_long_progression
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = 1L..5L
            __check((values.start).toString(), "1")
            __check((values.endInclusive).toString(), "5")
            __check((values.step).toString(), "1")
        }
