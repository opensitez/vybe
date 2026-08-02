// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_to_string_is_string_like
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 90.toDuration(DurationUnit.SECONDS)
            __check((value.toString().contains("s")).toString(), "true")
            __check((value.inWholeMinutes).toString(), "1")
        }
