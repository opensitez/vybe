// vybe-test: kotlin/kotlin_time/test_kotlin_time_days_unit_conversion
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1.toDuration(DurationUnit.DAYS)
            __check((value.inWholeHours).toString(), "24")
            __check((value.inWholeMinutes).toString(), "1440")
            __check((value.inWholeSeconds).toString(), "86400")
        }
