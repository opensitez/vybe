// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_addition
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = 2.toDuration(DurationUnit.SECONDS)
            val right = 500.toDuration(DurationUnit.MILLISECONDS)
            __check(((left + right).inWholeMilliseconds).toString(), "2500")
        }
