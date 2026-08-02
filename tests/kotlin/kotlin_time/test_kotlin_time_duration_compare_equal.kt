// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_compare_equal
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 1.toDuration(DurationUnit.MINUTES)
            val b = 60.toDuration(DurationUnit.SECONDS)
            __check((a == b).toString(), "true")
            __check((a != b).toString(), "false")
        }
