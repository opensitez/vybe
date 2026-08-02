// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_truncation_floor
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1501.toDuration(DurationUnit.MILLISECONDS)
            __check((value.inWholeSeconds).toString(), "1")
            __check((value.inWholeMilliseconds).toString(), "1501")
        }
