// vybe-test: kotlin/kotlin_java_time_apis/test_year_month_parse_and_plus
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.YearMonth.parse("2024-05").plusMonths(8)
            __check((value.year).toString(), "2025")
            __check((value.monthValue).toString(), "1")
        }
