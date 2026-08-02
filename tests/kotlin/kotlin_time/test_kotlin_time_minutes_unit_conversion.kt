// vybe-test: kotlin/kotlin_time/test_kotlin_time_minutes_unit_conversion
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 2.toDuration(DurationUnit.MINUTES)
            __check((value.inWholeSeconds).toString(), "120")
            __check((value.inWholeMilliseconds).toString(), "120000")
            __check((value.inWholeMinutes).toString(), "2")
        }
