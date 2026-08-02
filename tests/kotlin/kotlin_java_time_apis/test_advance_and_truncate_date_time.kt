// vybe-test: kotlin/kotlin_java_time_apis/test_advance_and_truncate_date_time
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDateTime.parse("2024-01-01T10:59:59")
            __check((value.plusSeconds(62).toString()).toString(), "2024-01-01T11:01:01")
            __check((value.with(java.time.temporal.ChronoField.HOUR_OF_DAY, 0).toString()).toString(), "2024-01-01T00:59:59")
        }
