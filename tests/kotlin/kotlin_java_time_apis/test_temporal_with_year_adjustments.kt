// vybe-test: kotlin/kotlin_java_time_apis/test_temporal_with_year_adjustments
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDate.parse("2024-02-29").withYear(2025)
            __check((value.toString()).toString(), "2025-02-28")
            val atStart = value.withMonth(1).withDayOfMonth(1)
            __check((atStart.toString()).toString(), "2025-01-01")
        }
