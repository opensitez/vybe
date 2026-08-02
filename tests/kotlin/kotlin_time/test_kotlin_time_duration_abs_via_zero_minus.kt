// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_abs_via_zero_minus
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = -(4.toDuration(DurationUnit.SECONDS))
            val inverted = -value
            __check((inverted.inWholeMilliseconds).toString(), "4000")
            __check((inverted == 4.toDuration(DurationUnit.SECONDS)).toString(), "true")
        }
