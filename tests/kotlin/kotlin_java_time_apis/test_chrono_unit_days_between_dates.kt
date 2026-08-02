// vybe-test: kotlin/kotlin_java_time_apis/test_chrono_unit_days_between_dates
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.time.LocalDate.parse("2024-01-01")
            val b = java.time.LocalDate.parse("2024-01-10")
            __check((java.time.temporal.ChronoUnit.DAYS.between(a, b)).toString(), "9")
            __check((java.time.temporal.ChronoUnit.MONTHS.between(a, b)).toString(), "0")
            __check((java.time.temporal.ChronoUnit.WEEKS.between(a, b)).toString(), "1")
        }
