// vybe-test: kotlin/kotlin_time/test_kotlin_time_zero_duration_is_zero
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Duration.ZERO.toDouble(DurationUnit.SECONDS)).toString(), "0.0")
            __check((Duration.ZERO.inWholeMilliseconds).toString(), "0")
        }
