// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_multiplication
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3.toDuration(DurationUnit.SECONDS)
            __check(((value * 2).inWholeMilliseconds).toString(), "6000")
            __check(((value * 3).inWholeMilliseconds).toString(), "9000")
        }
