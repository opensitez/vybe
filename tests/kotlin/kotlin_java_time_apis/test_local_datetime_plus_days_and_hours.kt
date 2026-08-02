// vybe-test: kotlin/kotlin_java_time_apis/test_local_datetime_plus_days_and_hours
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDateTime.parse("2024-07-01T10:00:00").plusDays(2).plusHours(5)
            __check((value.toString()).toString(), "2024-07-03T15:00")
        }
