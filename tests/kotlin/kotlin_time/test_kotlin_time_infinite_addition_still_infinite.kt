// vybe-test: kotlin/kotlin_time/test_kotlin_time_infinite_addition_still_infinite
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Duration.INFINITE + 10.toDuration(DurationUnit.SECONDS)
            __check((value.isInfinite()).toString(), "true")
            __check((Duration.INFINITE - 10.toDuration(DurationUnit.SECONDS) == Duration.INFINITE).toString(), "true")
        }
