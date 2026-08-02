// vybe-test: kotlin/kotlin_java_time_apis/test_week_fields_for_specific_date
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDate.parse("2024-01-02")
            val week = value.get(java.time.temporal.IsoFields.WEEK_OF_WEEK_BASED_YEAR)
            val weekYear = value.get(java.time.temporal.IsoFields.WEEK_BASED_YEAR)
            __check((week).toString(), "1")
            __check((weekYear).toString(), "2024")
            __check((value.dayOfWeek.value).toString(), "2")
        }
