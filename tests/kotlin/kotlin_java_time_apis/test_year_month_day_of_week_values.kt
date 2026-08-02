// vybe-test: kotlin/kotlin_java_time_apis/test_year_month_day_of_week_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDate.parse("2024-07-30")
            __check((value.dayOfWeek.value).toString(), "2")
            __check((value.dayOfWeek.name).toString(), "TUESDAY")
        }
