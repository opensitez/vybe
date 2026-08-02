// vybe-test: kotlin/kotlin_java_time_apis/test_month_day_parse_and_fields
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.MonthDay.parse("--12-25")
            __check((value.monthValue).toString(), "12")
            __check((value.dayOfMonth).toString(), "25")
        }
