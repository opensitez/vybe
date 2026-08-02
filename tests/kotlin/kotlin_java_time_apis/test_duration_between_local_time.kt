// vybe-test: kotlin/kotlin_java_time_apis/test_duration_between_local_time
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val start = java.time.LocalTime.parse("08:00")
            val end = java.time.LocalTime.parse("10:30")
            val d = java.time.Duration.between(start, end)
            __check((d.toMinutes()).toString(), "150")
            __check((d.seconds).toString(), "9000")
        }
