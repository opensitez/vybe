// vybe-test: kotlin/kotlin_time/test_kotlin_time_negative_plus_positive_cancel_to_zero
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = -(3.toDuration(DurationUnit.SECONDS)) + 3.toDuration(DurationUnit.SECONDS)
            __check((value == Duration.ZERO).toString(), "true")
            __check((value.inWholeMilliseconds).toString(), "0")
        }
