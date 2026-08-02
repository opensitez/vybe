// vybe-test: kotlin/kotlin_java_time_apis/test_local_date_minus_months_and_days
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDate.parse("2024-03-31").minusMonths(1)
            __check((value.toString()).toString(), "2024-02-29")
            val shifted = value.minusDays(1)
            __check((shifted.toString()).toString(), "2024-02-28")
        }
