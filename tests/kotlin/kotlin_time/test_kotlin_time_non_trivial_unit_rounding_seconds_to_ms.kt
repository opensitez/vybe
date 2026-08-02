// vybe-test: kotlin/kotlin_time/test_kotlin_time_non_trivial_unit_rounding_seconds_to_ms
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 2.toDuration(DurationUnit.SECONDS)
            __check((value.toLong(DurationUnit.MILLISECONDS)).toString(), "2000")
            __check((value.toLong(DurationUnit.MICROSECONDS)).toString(), "2000000")
        }
