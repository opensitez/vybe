// vybe-test: kotlin/ranges/test_long_range_step_and_contains
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val timeline = 1000L downTo 995L step 2
            __check((timeline.first).toString(), "1000")
            __check((timeline.last).toString(), "996")
            __check((999L in timeline).toString(), "true")
            __check((998L in timeline).toString(), "false")
            __check((timeline.count()).toString(), "3")
        }
