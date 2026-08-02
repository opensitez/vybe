// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_compare_with_zero
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val positive = 1.toDuration(DurationUnit.SECONDS)
            val zero = Duration.ZERO
            val negative = -(500.toDuration(DurationUnit.MILLISECONDS))
            __check((positive > zero).toString(), "true")
            __check((zero > negative).toString(), "true")
            __check((negative < zero).toString(), "true")
        }
