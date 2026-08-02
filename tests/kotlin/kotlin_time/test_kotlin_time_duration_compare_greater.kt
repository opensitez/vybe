// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_compare_greater
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 90.toDuration(DurationUnit.SECONDS)
            val b = 1.toDuration(DurationUnit.MINUTES)
            __check((a > b).toString(), "true")
            __check((a >= b).toString(), "true")
            __check((a < b).toString(), "false")
        }
