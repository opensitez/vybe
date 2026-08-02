// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_negation_sign
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = -(2.toDuration(DurationUnit.SECONDS))
            __check((value.inWholeSeconds).toString(), "-2")
            __check((value.isNegative()).toString(), "true")
            __check((value < Duration.ZERO).toString(), "true")
        }
