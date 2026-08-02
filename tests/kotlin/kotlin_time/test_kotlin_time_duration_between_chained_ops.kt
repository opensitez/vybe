// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_between_chained_ops
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = (1.toDuration(DurationUnit.MINUTES) + 30.toDuration(DurationUnit.SECONDS)) - 1.toDuration(DurationUnit.SECONDS)
            __check((value.inWholeSeconds).toString(), "89")
        }
