// vybe-test: kotlin/kotlin_progressions/test_range_reversal
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = (1..5).reversed()
            __check((r.first()).toString(), "5")
            __check((r.last()).toString(), "1")
        }
