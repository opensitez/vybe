// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_components_like_whole_minutes
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3700.toDuration(DurationUnit.SECONDS)
            __check((value.inWholeMinutes).toString(), "61")
            __check((value.inWholeHours).toString(), "1")
            __check((value.inWholeDays).toString(), "0")
        }
