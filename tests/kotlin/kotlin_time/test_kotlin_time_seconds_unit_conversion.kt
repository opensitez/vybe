// vybe-test: kotlin/kotlin_time/test_kotlin_time_seconds_unit_conversion
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 5.toDuration(DurationUnit.SECONDS)
            __check((value.inWholeMilliseconds).toString(), "5000")
            __check((value.inWholeSeconds).toString(), "5")
        }
