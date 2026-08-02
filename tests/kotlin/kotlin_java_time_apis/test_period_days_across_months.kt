// vybe-test: kotlin/kotlin_java_time_apis/test_period_days_across_months
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val start = java.time.LocalDate.parse("2023-11-30")
            val end = java.time.LocalDate.parse("2023-12-01")
            val p = java.time.Period.between(start, end)
            __check((p.days).toString(), "1")
            __check((p.months).toString(), "0")
            __check((p.years).toString(), "0")
        }
