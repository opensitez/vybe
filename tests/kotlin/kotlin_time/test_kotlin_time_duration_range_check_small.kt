// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_range_check_small
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 500.toDuration(DurationUnit.MILLISECONDS)
            val b = 1.toDuration(DurationUnit.SECONDS)
            val c = 1500.toDuration(DurationUnit.MILLISECONDS)
            __check((a < b).toString(), "true")
            __check((c > b).toString(), "true")
            __check((c > a).toString(), "true")
        }
