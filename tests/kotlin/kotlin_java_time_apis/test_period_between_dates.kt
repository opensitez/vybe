// vybe-test: kotlin/kotlin_java_time_apis/test_period_between_dates
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val start = java.time.LocalDate.parse("2024-01-01")
            val end = java.time.LocalDate.parse("2024-03-11")
            val p = java.time.Period.between(start, end)
            __check((p.months).toString(), "2")
            __check((p.days).toString(), "10")
            __check((p.years).toString(), "0")
        }
