// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_to_long_nanoseconds
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 250.toDuration(DurationUnit.MILLISECONDS)
            __check((value.toLong(DurationUnit.NANOSECONDS)).toString(), "250000000")
            __check((value.toLong(DurationUnit.MICROSECONDS)).toString(), "250000")
        }
