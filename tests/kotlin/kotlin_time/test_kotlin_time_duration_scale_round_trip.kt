// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_scale_round_trip
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 12.toDuration(DurationUnit.MILLISECONDS)
            __check(((base * 5).toLong(DurationUnit.MILLISECONDS)).toString(), "60")
            __check((((base * 5) / 5).inWholeMilliseconds).toString(), "12")
            __check((base.inWholeMilliseconds).toString(), "12")
        }
