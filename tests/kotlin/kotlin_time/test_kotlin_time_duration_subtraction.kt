// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_subtraction
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = 5.toDuration(DurationUnit.SECONDS)
            val right = 2.toDuration(DurationUnit.SECONDS)
            __check(((left - right).inWholeSeconds).toString(), "3")
        }
